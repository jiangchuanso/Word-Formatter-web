# -*- coding: utf-8 -*-
"""Core document formatting engine for Word Formatter Pro.

This module is intentionally independent from Tkinter so GUI, CLI, and skills
can reuse the same formatter implementation. This copy is maintained with the
main release.
"""

import copy
from copy import deepcopy
import gc
import itertools
import logging
import math
import os
from collections import Counter
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import uuid


from docx import Document
from docx.document import Document as _Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.styles import StyleFactory
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Pt, Cm, RGBColor
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from wfp_process import converter_profile_uri, extended_path, remove_owned_temp_dir, run_isolated

try:
    import jieba
except ImportError:  # jieba 缺失时新增强功能降级为仅按标点判断，不影响主流程
    jieba = None


IS_WINDOWS = sys.platform.startswith('win')
IS_LINUX = sys.platform.startswith('linux')
if IS_WINDOWS:
    try:
        import win32com.client
    except ImportError:
        win32com = None
    try:
        import pythoncom
    except ImportError:
        pythoncom = None
else:
    win32com = None
    pythoncom = None

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

BLANK_LINE_MODE_PRESERVE = '不改动任何空行'
BLANK_LINE_MODE_DELETE_SINGLE = '删除单个空行，多个空行保留至1个空行'
BLANK_LINE_MODE_KEEP_SINGLE = '保留单个空行，多个空行保留至1个空行'
BLANK_LINE_MODE_OPTIONS = [
    BLANK_LINE_MODE_PRESERVE,
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
]
DEFAULT_BLANK_LINE_MODE = BLANK_LINE_MODE_DELETE_SINGLE
SUPPORTED_FILE_EXTENSIONS = ('.docx', '.doc', '.wps', '.txt', '.md')
LARGE_FOLDER_FILE_CONFIRM_THRESHOLD = 1000

RE_SAFE_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
RE_HAS_CHINESE = re.compile(r'[\u4e00-\u9fff]')
RE_MD_IMAGE = re.compile(r'!\[([^\]]*)\]\([^)]+\)')
RE_MD_LINK = re.compile(r'\[([^\]]+)\]\([^)]+\)')
RE_MD_HTML_TAG = re.compile(r'<[^>]+>')
RE_MD_INLINE_CODE = re.compile(r'`([^`]+)`')
# Markdown 行内控制语法：其中的半角标点属于语法本身，符号标准化不能改成全角。
RE_MD_INLINE_SYNTAX = re.compile(
    r'`[^`\n]+`|!?\[[^\]\n]*\]\([^)\n]*\)|!?\[[^\]\n]*\]\[[^\]\n]*\]'
)
RE_MD_PLACEHOLDER = re.compile(r'\x00(\d+)\x00')
# 历史输出里图片语法曾被整体全角化，保留窄口径修复：只认「！[Image]（…）」。
RE_MD_FULLWIDTH_IMAGE = re.compile(r'！\[Image\]（([^）\n]*)）')
RE_MD_BOLD_ASTERISK = re.compile(r'\*\*(.*?)\*\*')
RE_MD_BOLD_UNDERSCORE = re.compile(r'__(.*?)__')
RE_MD_HEADER = re.compile(r'^\s*#+\s+(.*)')
# 标题井号后紧跟的空格是 Markdown 大纲语法，空格清理必须原样保留。
RE_MD_HEADING_PREFIX = re.compile(r'^\s*#+[ \t]+')
RE_MD_BLOCKQUOTE = re.compile(r'^\s*>\s?(.*)')
RE_MD_HORIZONTAL_RULE = re.compile(r'^\s*[-*_]{3,}\s*$')
RE_MD_UNORDERED_LIST_BLOCK = re.compile(r'^\s*[*+-]\s+')
RE_MD_UNORDERED_LIST = re.compile(r'^\s*[*+-]\s')
RE_MD_BULLET_WITH_CONTENT = re.compile(r'^(\s*[*+-]\s)(.*)')
RE_MD_EMPHASIS_ASTERISK = re.compile(r'(?<!\\)\*([^\s*][^*]*?)(?<!\\)\*')
# 文本源文件保持后缀输出时，这些行的 Markdown 结构本身必须原样保留。
# 注意：列表项只保护前缀，列表正文仍需参与断行合并，否则列表项内被 PDF/docling
# 拆开的续行永远无法还原（曾经把整行排除，导致只有十余处断行被合并）。
RE_MD_PROTECTED_LINE = re.compile(
    r'^\s*(?:#{1,6}\s|>\s?|\|>|```|~~~|---+\s*$|\*\*\*\s*$|___\s*$|<\/?[A-Za-z])'
)
RE_MD_LIST_PREFIX = re.compile(
    r'^(\s*(?:[-*+]\s+|\d{1,3}[.)]\s+|[（(][A-Za-z0-9]{1,3}[）)]\s*))'
)
# 列表项编号本体，用于识别「连续编号中唯一没有段尾标点的一项」
RE_LIST_NUMBER_TOKEN = re.compile(
    r'^\s*(?:[（(]\s*([0-9]{1,3}|[A-Za-z])\s*[）)]'
    r'|([0-9]{1,3})\s*[.．、)]'
    r'|([①②③④⑤⑥⑦⑧⑨⑩])'
    r'|([一二三四五六七八九十百]+)\s*[、.])'
)
RE_MD_BULLET_PREFIX = re.compile(r'^\s*[-*+]\s+')
# 兼容旧调用点：整行结构保护（含列表前缀），仅用于不需要处理列表正文的场景
RE_MD_STRUCTURAL_LINE = re.compile(
    r'^\s*(?:#{1,6}\s|>\s?|[-*+]\s|\d+[.)]\s|\|>|```|---+\s*$|\*\*\*\s*$|___\s*$|<\/?[A-Za-z])'
)
RE_MD_FENCE = re.compile(r'^\s*(?:```|~~~)')
# 排版清洗：PDF 粘贴常在中日文与相邻中文/英文/数字之间插入多余空格
_CJK_IDEOGRAPH = r'[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]'
_CJK_PUNCT = r"[，。、；：！？（）〔〕【】《》「」『』“”‘’—…·,.!?;:()\[\]{}<>/\\|@#$%^&*+=~`\-\"']"
_CJK_EDGE = rf'(?:{_CJK_IDEOGRAPH}|{_CJK_PUNCT})'
# 中文语境下标点两侧的多余空格也要去掉；纯英文行不参与，避免误伤 "Cor-Ban 27L"。
RE_CJK_SPACE_AFTER = re.compile(
    rf'(?<={_CJK_EDGE})[ \u3000]+(?={_CJK_IDEOGRAPH}|[A-Za-z0-9]|{_CJK_PUNCT})'
)
RE_CJK_SPACE_BEFORE = re.compile(
    rf'(?<=[A-Za-z0-9]|{_CJK_PUNCT})[ \u3000]+(?={_CJK_EDGE})'
)
# PDF 转换识别错误：外层自动编号转文字后又重复了一份相同序号（如「3. 3．标题」）
RE_DUPLICATE_NUMBER = re.compile(
    r'^([0-9]{1,3})([.．])[ \t]*([0-9]{1,3})([.．])(?![0-9.．])[ \t]*(?=\S)'
)
# docling 可能叠加多层编号（「3. 3． 3. 标题」），单遍剥离不收敛，限制迭代上限。
DUPLICATE_NUMBER_MAX_PASSES = 4
# docling 页面模型会把页面级连续序号叠到真实子标号前：「2. (a) 内容」「4. (b) 内容」
RE_LEADING_NUMBER_WITH_ALPHA = re.compile(
    r'^([0-9]{1,3})([.．])[ \t]+(?=[（(][A-Za-z][）)])'
)
# 不带前置数字的裸子标号（Word 列表符号或 docling 的「- (a)」「- 1)」）
RE_BARE_PAREN_LABEL = re.compile(
    r'^\s*(?:[-*+]\s+)?[（(]([A-Za-z0-9]{1,3})[）)](?=[\s\u4e00-\u9fff]|$)'
)
RE_BARE_NUM_PAREN_LABEL = re.compile(
    r'^\s*(?:[-*+]\s+)?[0-9]{1,3}[）)](?=[\s\u4e00-\u9fff]|$)'
)
# 同一个页面级计数器的纯数字前缀（无括号子标号）
RE_PAGE_NUMBER_PREFIX = re.compile(r'^([0-9]{1,3})([.．])[ \t]+(?=\S)')
RE_NOTE_LINE = re.compile(r'^\s*(?:警告|注意|说明|警示|提示)\s*[：:]')


def output_extension_for(config, source_name):
    """按“保持文件后缀”配置与源文件后缀决定输出后缀。"""
    source_ext = os.path.splitext(str(source_name))[1].lower()
    if config.get('keep_source_extension', False) and source_ext in ('.txt', '.md'):
        return source_ext
    return '.docx'
RE_CURRENCY_PREFIX = re.compile(r'^[¥￥$]')
RE_CURRENCY_SUFFIX = re.compile(r'(元|万元|亿元)$')
RE_NUMERIC_TABLE_TEXT = re.compile(r'^[-+]?(?:\d+(?:\.\d+)?|\.\d+)%?$')
CHINESE_NUM_PATTERN = r'[一二三四五六七八九十百千万零]+'
RE_TITLE_H1 = re.compile(r'^' + CHINESE_NUM_PATTERN + r'\s*、')
RE_TITLE_H2 = re.compile(r'^[（\(]' + CHINESE_NUM_PATTERN + r'[）\)]')
RE_HEADING_H1 = re.compile(r'^[一二三四五六七八九十百千万零]+\s*、')
RE_HEADING_H2 = re.compile(r'^[（\(][一二三四五六七八九十百千万零]+[）\)]')
# 单右括号单级编号（1) / a) / 1））与点号单级编号（1. / A.）同级；
# 双括号（(1) /（1））是更末级标号，保持四级，见 RE_HEADING_H4。
RE_HEADING_H3_PAREN = re.compile(r'^[0-9A-Za-z]+\s*[\)）](?=\s|[\u4e00-\u9fff]|$)')
RE_HEADING_H3 = re.compile(r'^[0-9A-Za-z]+\s*[\.．\)）](?=\s|[\u4e00-\u9fff]|$)')
RE_HEADING_H4 = re.compile(r'^[（\(]\d+[）\)]')
# 单独的字母序号：(a) / （a） / (A) 等，与 (1) 同属最末级标号
RE_HEADING_H4_ALPHA = re.compile(r'^[（\(][a-zA-Z][）\)]')
# 多级点分编号：每段可为数字或字母，至少一个点，末尾点可选。
# 兼容 “1.1 标题”“1.1. 标题”“A.B 标题”“1.A 标题”以及自动编号转文字后无尾巴的 “1.1标题”。
RE_DOTTED_NUMBER = re.compile(r'^([0-9A-Za-z]+(?:\.[0-9A-Za-z]+)+)\.?')
# 裸单级编号（无点）：1~3 位数字或单个字母，后接空白/中文/行尾；仅在点分编号为主干时启用。
RE_SINGLE_LABEL = re.compile(r'^([0-9]{1,3}|[A-Za-z])(?=[\s\u4e00-\u9fff]|$)')
RE_ATTACHMENT = re.compile(r'^附件\s*(\d+|[一二三四五六七八九十百千万零]+)?\s*[:：]?$')
RE_H2_INLINE_TITLE = re.compile(r'^[（\(](.+?)[）\)](.*)', re.DOTALL)
RE_HEADING_STYLE_NAME = re.compile(r'^(?:heading\s*\d+|标题\s*\d+)$', re.IGNORECASE)

# 可选增强：大纲级别上限（w:outlineLvl 的 9 表示正文级，不做标题处理）
MAX_HEADING_LEVEL = 9
MAX_TITLE_OUTLINE_LEVEL = MAX_HEADING_LEVEL - 1

# 标准样式统一：把各级字体字号写进 Word「标准样式」组里的内置样式，
# 再把这些样式套到对应段落上，同一个等级的段落共用一套样式。
BUILTIN_BODY_STYLE_NAME = 'Normal'
BUILTIN_BODY_STYLE_ALIASES = ('正文',)
BUILTIN_HEADING_STYLE_NAMES = {
    1: 'Heading 1',
    2: 'Heading 2',
    3: 'Heading 3',
    4: 'Heading 4',
}
BUILTIN_HEADING_STYLE_ALIASES = {
    1: ('标题 1', '标题1'),
    2: ('标题 2', '标题2'),
    3: ('标题 3', '标题3'),
    4: ('标题 4', '标题4'),
}
MAX_STYLE_HEADING_LEVEL = max(BUILTIN_HEADING_STYLE_NAMES)
# 内置样式自带的主题字体/主题色会盖掉显式字体，写样式前先清掉
STYLE_THEME_FONT_ATTRIBUTES = (
    'w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme', 'w:csTheme',
)
STYLE_THEME_COLOR_ATTRIBUTES = ('w:themeColor', 'w:themeShade', 'w:themeTint')
# 由样式统一提供的字符属性；段落上这些直接格式会被清掉，其余（加粗/斜体/下划线/颜色）保留
STYLE_DIRECT_RUN_TAGS = ('w:rFonts', 'w:sz', 'w:szCs')
HEADING_TEXT_MAX_LENGTH = 60
HEADING_TAIL_PUNCTUATION = '。；，、：!！?？…'
# 句中一旦出现这些断句标点，说明该段在讲完整句子，属于正文而不是标题
BODY_SENTENCE_PUNCTUATION = '。！？；!?'

# 章/节标题（可选增强）：第X章 / ATAxx章 记为 1 级，第X节 记为 2 级，其后编号标题整体顺延
RE_CHAPTER_TITLE = re.compile(
    r'^(?:第\s*[一二三四五六七八九十百千万零〇\d]+\s*章|ATA\s*\d+\s*章)', re.IGNORECASE
)
RE_SECTION_TITLE = re.compile(r'^第\s*[一二三四五六七八九十百千万零〇\d]+\s*节')
# 条标题：第X条 / 第XX条 / 第X.X条 / 第XX.XX条（点号支持半角/全角），
# 级别跟随章/节顺延，点分编号每多一段再深一级（无章无节时第X条为 1 级）。
_ARTICLE_NUM_SEGMENT = r'[一二三四五六七八九十百千万零〇\d]{1,6}'
RE_ARTICLE_TITLE = re.compile(
    rf'^第\s*{_ARTICLE_NUM_SEGMENT}(?:\s*[.．]\s*{_ARTICLE_NUM_SEGMENT})*\s*条'
)

# 前后行词语连续合并（PDF 粘贴常见）：句末/续句标点、jieba 跨行词探测、满行长度下限
MERGE_SENTENCE_END_PUNCTUATION = '。！？；!?…'
MERGE_SENTENCE_END_CLOSERS = '”’"）】》』」'
MERGE_CONNECTING_PUNCTUATION = '，、：,;:'
MERGE_JIEBA_PROBE_SIZE = 6
# 标题型行参与断行合并的最小长度：短标题 + 正文不构成断行，防止把标题与正文误并
MERGE_HEADING_MIN_LENGTH = 12
# 列表项被 PDF 断行后的短续行上限：「维护工作；」「使用标准；」这类词组尾部
MERGE_LIST_TAIL_MAX_LENGTH = 12
# 正文虚词断行：上一行以这些词收尾，或下一行以这些词开头，通常是同一句被拆开
MERGE_FUNCTION_WORDS = frozenset(
    '的 和 与 或 及 并 而 但 则 以 将 把 被 从 到 对 为 在 向 给 按 据 由 等'.split()
)


class LegacyConversionUnavailable(RuntimeError):
    """Raised when an old Word/WPS file is intentionally skipped."""


class SofficeConverter:
    """LibreOffice wrapper for converting legacy .doc/.wps files to .docx."""

    def __init__(self, soffice_path=None, timeout=120):
        self.soffice_path = soffice_path or self._find_soffice()
        self.timeout = int(timeout or 120)

    @property
    def available(self):
        return bool(self.soffice_path)

    @staticmethod
    def _find_soffice():
        for executable in ("soffice", "soffice.com", "libreoffice"):
            found = shutil.which(executable)
            if found:
                return found

        common_paths = []
        if sys.platform == "darwin":
            common_paths.extend(
                [
                    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                    "/opt/homebrew/bin/soffice",
                    "/usr/local/bin/soffice",
                ]
            )
        elif os.name == "nt":
            for base in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)")):
                if base:
                    common_paths.extend(
                        [
                            os.path.join(base, "LibreOffice", "program", "soffice.com"),
                            os.path.join(base, "LibreOffice", "program", "soffice.exe"),
                        ]
                    )
        else:
            common_paths.extend(
                [
                    "/usr/bin/soffice",
                    "/usr/local/bin/soffice",
                    "/usr/bin/libreoffice",
                    "/usr/local/bin/libreoffice",
                    "/snap/bin/libreoffice",
                    "/opt/libreoffice/program/soffice",
                    "/opt/libreoffice*/program/soffice",
                ]
            )

        for path in common_paths:
            if "*" in path:
                for candidate in sorted(Path("/").glob(path.lstrip("/"))):
                    if candidate.is_file():
                        return str(candidate)
                continue
            if path and os.path.isfile(path):
                return path
        return None

    def convert_to_docx(self, input_path, log=None):
        input_path = Path(input_path).expanduser().resolve()
        if not self.available:
            raise RuntimeError(
                "LibreOffice (soffice) 不可用，无法转换 "
                f"{input_path}。请安装 LibreOffice，或先将文档另存为 .docx 后再处理。"
            )

        work_dir = Path(tempfile.mkdtemp(prefix="wfp_soffice_"))
        out_dir = work_dir / "out"
        profile_dir = work_dir / "profile"
        out_dir.mkdir(parents=True, exist_ok=True)
        profile_dir.mkdir(parents=True, exist_ok=True)
        profile_uri = profile_dir.resolve().as_posix()
        if os.name == "nt":
            profile_uri = "/" + profile_uri

        cmd = [
            self.soffice_path,
            "--headless",
            "--norestore",
            f"-env:UserInstallation=file://{profile_uri}",
            "--convert-to",
            "docx",
            "--outdir",
            str(out_dir),
            str(input_path),
        ]
        if log:
            log(f"  > 正在使用 LibreOffice 转换为 .docx: {input_path.name}")

        creationflags = 0
        if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags = subprocess.CREATE_NO_WINDOW

        try:
            proc = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=self.timeout,
                creationflags=creationflags,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            shutil.rmtree(work_dir, ignore_errors=True)
            raise RuntimeError(f"LibreOffice 转换超时 ({self.timeout}s): {input_path}") from exc
        except FileNotFoundError as exc:
            shutil.rmtree(work_dir, ignore_errors=True)
            raise RuntimeError(f"找不到 soffice 可执行文件: {self.soffice_path}") from exc
        except Exception:
            shutil.rmtree(work_dir, ignore_errors=True)
            raise

        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout or "").strip()
            shutil.rmtree(work_dir, ignore_errors=True)
            raise RuntimeError(
                f"LibreOffice 无法将 {input_path} 转为 .docx。"
                f"请先手动转为 .docx 后再处理。详细错误: {detail}"
            )

        generated_files = sorted(out_dir.glob("*.docx"))
        if not generated_files:
            detail = (proc.stderr or proc.stdout or "").strip()
            shutil.rmtree(work_dir, ignore_errors=True)
            raise RuntimeError(
                f"LibreOffice 未生成 {input_path} 对应的 .docx。"
                f"请先手动转为 .docx 后再处理。详细信息: {detail}"
            )

        if log:
            log(f"  > LibreOffice 转换完成: {generated_files[0].name}")
        return generated_files[0], work_dir

def _initialize_com_for_thread(log_callback=None):
    if not (IS_WINDOWS and pythoncom is not None):
        return False
    try:
        pythoncom.CoInitialize()
        return True
    except Exception as e:
        if log_callback:
            log_callback(f"警告：后台线程初始化 COM 失败，将跳过 COM 自动化能力：{e}")
        return False

def _uninitialize_com_for_thread(initialized, log_callback=None):
    if not initialized or pythoncom is None:
        return
    try:
        pythoncom.CoUninitialize()
    except Exception as e:
        if log_callback:
            log_callback(f"警告：后台线程释放 COM 失败：{e}")

class WPSAppManager:
    def __init__(self, log_callback=None):
        self.log_callback = log_callback
        self.com_app = None

    def _log(self, message):
        if self.log_callback:
            self.log_callback(message)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()
        return False

    @staticmethod
    def _com_available():
        return IS_WINDOWS and win32com is not None

    @staticmethod
    def _com_unavailable_message(file_ext=None):
        if file_ext in ('.doc', '.wps'):
            return (
                f"当前环境无法使用 WPS/Word COM 直接转换 {file_ext} 文件。"
                "程序会改用 LibreOffice soffice 转换；如 soffice 不可用，请安装 LibreOffice，"
                "或先将文件另存为 .docx 后再处理。"
            )
        return (
            "当前环境无法使用 WPS/Word COM 自动化，无法执行 COM 预处理或自动接受修订；"
            "自动编号将改用 Python 转换。"
            "在 macOS/Kylin/Linux 等非 Windows 系统上这是预期行为。"
        )

    def get_app(self):
        if not self._com_available():
            raise RuntimeError(self._com_unavailable_message())

        if self.com_app is None:
            self._log("首次需要，正在启动WPS/Word应用（独立进程）...")
            try:
                self.com_app = win32com.client.DispatchEx('KWPS.Application')
                self._log("  > 已成功连接到WPS。")
            except Exception:
                try:
                    self.com_app = win32com.client.DispatchEx('Word.Application')
                    self._log("  > 已成功连接到Word。")
                except Exception as e:
                    raise RuntimeError(f"未能启动WPS或Word，请确保已安装。错误: {e}")

            try:
                self.com_app.Visible = False
            except Exception:
                pass
            try:
                self.com_app.DisplayAlerts = False
            except Exception:
                pass

        return self.com_app

    def quit(self):
        app, self.com_app = self.com_app, None
        if app is None:
            return
        self._log("所有任务完成，正在关闭WPS/Word应用...")
        try:
            # 残留文档会让 Quit() 弹窗或让进程无法退出，先全部静默关闭
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            try:
                while app.Documents.Count:
                    app.Documents(1).Close(SaveChanges=0)
            except Exception:
                pass
            app.Quit()
        except Exception as e:
            self._log(f"  > 警告：关闭应用时发生异常: {e}")
        finally:
            del app
            gc.collect()
            self._log("  > 应用已关闭。")

RE_ATTACHMENT_LINE = re.compile(
    r'^附件(?:[ \t\u3000]*(?P<number>\d+|[一二三四五六七八九十百千万零〇两]+))?'
    r'(?P<separator>[ \t\u3000]*[:：][ \t\u3000]*|[ \t\u3000]+)?(?P<title>.*)$', re.DOTALL
)
RE_ATTACHMENT_LIST_ITEM = re.compile(
    r'^(?:(?P<arabic>\d+)[.．、]|[（(](?P<arabic_paren>\d+)[）)]|'
    r'(?P<chinese>[一二三四五六七八九十]+)[、.．]|[（(](?P<chinese_paren>[一二三四五六七八九十]+)[）)])\s*\S'
)

BOLD_KEYS = ('title_bold', 'subtitle_bold', 'h1_bold', 'h2_bold',
             'attachment_bold', 'table_header_bold')

LEGACY_FORMAT_KEYS = ('table_smart_align', 'table_unified_borders')

FORMAT_CHOICES = {
    **{key: ('preserve', 'bold', 'normal') for key in BOLD_KEYS},
    'table_alignment': ('preserve', 'left', 'center', 'right'),
    'table_cell_alignment': ('preserve', 'smart', 'left', 'center', 'right'),
    'table_border_mode': ('preserve', 'unified'),
    'title_after_spacing_unit': ('pt', 'lines'),
    'attachment_title_after_spacing_unit': ('pt', 'lines'),
    'signature_layout': ('sealed', 'unsealed'),
    'signature_before_spacing_mode': ('auto', 'preserve', 'custom'),
    'signature_before_spacing_unit': ('pt', 'lines'),
}

def normalize_format_options(config):
    """Migrate legacy switches before merging defaults; never mutate the caller."""
    result = dict(config)
    for scope in ('title', 'attachment_title'):
        blank_key = 'enable_' + scope + '_after_blank_line'
        if blank_key in result:
            if result[blank_key] is None:
                result.pop(blank_key)
            elif not isinstance(result[blank_key], bool):
                raise ValueError(f'{blank_key} 必须为 true 或 false')
        if result.get(blank_key, False):
            result['enable_' + scope + '_after_spacing'] = False
    for key in ('enable_signature_formatting', 'signature_include_attachments', 'signature_normalize_date'):
        if key in result:
            if result[key] is None:
                result.pop(key)
            elif not isinstance(result[key], bool):
                raise ValueError(f'{key} 必须为 true 或 false')
    for old, new, enabled, disabled in (
        ('table_smart_align', 'table_cell_alignment', 'smart', 'preserve'),
        ('table_unified_borders', 'table_border_mode', 'unified', 'preserve'),
    ):
        if old in result:
            if new not in result:
                result[new] = enabled if result[old] else disabled
            result.pop(old)
    for key in BOLD_KEYS:
        if isinstance(result.get(key), bool):
            result[key] = 'bold' if result[key] else 'preserve'
    for key, choices in FORMAT_CHOICES.items():
        value = result.get(key)
        if value is None or value == '':
            result.pop(key, None)
        elif value not in choices:
            raise ValueError(f"无效配置 {key}={value!r}，可选值：{', '.join(choices)}")
    for key in ('title_after_spacing', 'attachment_title_after_spacing', 'signature_before_spacing'):
        if key in result and result[key] not in (None, ''):
            try:
                value = float(result[key])
            except (TypeError, ValueError):
                raise ValueError(f"{key} 必须为非负数") from None
            if not math.isfinite(value) or value < 0:
                raise ValueError(f"{key} 必须为非负数")
            result[key] = value
    return result


class WordProcessor:
    def __init__(
        self,
        config,
        log_callback=None,
        remove_blank_lines=True,
        blank_line_mode=None,
        com_manager=None,
        soffice_path=None,
        soffice_timeout=120,
    ):
        self.config = config
        self.temp_files = []
        # 章/节识别开启时，编号标题（一、/（一）/1./（1））需要整体顺延的大纲级数
        self._outline_offset = 0
        # 条标题（第X条）的大纲级别，按章/节是否出现决定
        self._article_level = 1
        # 点分编号「1.」的基准大纲级别：文档仅用点分多级编号且无传统标题时为 1（作为一级标题），
        # 否则为 3（保持公文三级标题约定）。由 _prepare_dotted_base 按文档结构决定。
        self._dotted_base_level = 3
        # 文档是否确实使用多级点分编号（1.1/1.1.1 等），决定是否识别“裸单级编号”（如自动编号转出的 1标题）
        self._has_dotted_hierarchy = False
        # 更深的点分编号（x.x / x.x.x）下游的单级编号，视为其下一级：
        # {段落 XML 元素 id: 大纲级别}，由 _prepare_single_number_restart 预扫描。
        # _single_number_levels 对应「1.」起头的连续编号，_paren_number_levels 对应「(1)」标号。
        # _single_number_owners 持有元素代理，避免代理被回收后 id() 变化导致查不到。
        self._single_number_levels = {}
        self._paren_number_levels = {}
        self._single_number_owners = []
        # 单独的末级标号（(1)/(a) 等无所属正文或下游标题）段落集合，按正文处理、不给等级标题
        self._non_heading_paras = set()
        # 逐段排版时登记的最终等级（key 为段落的 XML 元素 id），供排版结束后统一套用内置样式
        self._level_records = {}
        self.sys_temp_dir = tempfile.gettempdir()
        self.log_callback = log_callback
        self.com_manager = com_manager or WPSAppManager(log_callback)
        self._owns_com_manager = com_manager is None
        self.soffice_path = soffice_path
        self.soffice_timeout = soffice_timeout
        self.soffice_converter = None
        self.blank_line_mode = self._normalize_blank_line_mode(
            blank_line_mode or self.config.get('blank_line_mode'),
            remove_blank_lines=remove_blank_lines
        )
        self.remove_blank_lines = self.blank_line_mode == BLANK_LINE_MODE_DELETE_SINGLE

    def __del__(self):
        # 兜底：自建的 COM 管理器（测试、直接调用等）若未显式 quit，退出时释放
        # Word/WPS 进程，避免残留进程持续写诊断日志占满磁盘。共享 com_manager
        # 的场景（CLI/GUI/Web）由外层 WPSAppManager 统一关闭。
        try:
            if getattr(self, '_owns_com_manager', False):
                self.com_manager.quit()
        except Exception:
            pass

    def _log(self, message):
        if self.log_callback: self.log_callback(message)

    @staticmethod
    def _com_available():
        return WPSAppManager._com_available()

    @staticmethod
    def _com_unavailable_message(file_ext=None):
        return WPSAppManager._com_unavailable_message(file_ext)

    @staticmethod
    def _paragraph_has_ooxml(para, tag_name):
        return para._p.find('.//' + qn(tag_name)) is not None

    @classmethod
    def _has_field_codes(cls, para):
        return (
            cls._paragraph_has_ooxml(para, 'w:fldChar')
            or cls._paragraph_has_ooxml(para, 'w:instrText')
        )

    @classmethod
    def _has_drawing_or_pict(cls, para):
        return (
            cls._paragraph_has_ooxml(para, 'w:drawing')
            or cls._paragraph_has_ooxml(para, 'w:pict')
        )

    @classmethod
    def _has_embedded_object(cls, para):
        return cls._paragraph_has_ooxml(para, 'w:object')

    @staticmethod
    def _normalize_blank_line_mode(mode, remove_blank_lines=True):
        if mode in BLANK_LINE_MODE_OPTIONS:
            return mode
        if mode in ('preserve', 'none'):
            return BLANK_LINE_MODE_PRESERVE
        if mode in ('delete_single', 'remove_single'):
            return BLANK_LINE_MODE_DELETE_SINGLE
        if mode in ('keep_single', 'compress'):
            return BLANK_LINE_MODE_KEEP_SINGLE
        if isinstance(mode, bool):
            return BLANK_LINE_MODE_DELETE_SINGLE if mode else BLANK_LINE_MODE_KEEP_SINGLE
        return BLANK_LINE_MODE_DELETE_SINGLE if remove_blank_lines else BLANK_LINE_MODE_KEEP_SINGLE

    def _cleanup_temp_files(self):
        if not self.temp_files:
            return
        self._log("正在清理本轮临时文件...")
        for f in self.temp_files:
            try:
                if os.path.exists(f):
                    os.remove(f)
                    self._log(f"  > 临时文件 {os.path.basename(f)} 已删除")
            except OSError as e:
                self._log(f"  > 警告：删除临时文件 {f} 失败: {e}")
        self.temp_files.clear()

    def _make_temp_docx_path(self, prefix, base_name):
        safe_base_name = RE_SAFE_FILENAME_CHARS.sub('_', base_name).strip(' ._')
        safe_base_name = (safe_base_name or 'document')[:80]
        temp_name = f"~temp_{prefix}_{safe_base_name}_{os.getpid()}_{uuid.uuid4().hex[:8]}.docx"
        temp_path = os.path.join(self.sys_temp_dir, temp_name)
        self.temp_files.append(temp_path)
        return temp_path

    def _get_wps_app(self):
        return self.com_manager.get_app()

    def _get_soffice_converter(self):
        if self.soffice_converter is None:
            self.soffice_converter = SofficeConverter(self.soffice_path, self.soffice_timeout)
        return self.soffice_converter

    def _convert_legacy_with_com(self, input_path, temp_docx_path):
        app = self._get_wps_app()
        doc_com = None
        try:
            doc_com = app.Documents.Open(os.path.abspath(input_path), ReadOnly=1)
            doc_com.SaveAs2(os.path.abspath(temp_docx_path), FileFormat=12)
        finally:
            if doc_com is not None:
                try:
                    doc_com.Close()
                except Exception:
                    pass
        self._log("文件格式转换完成。")

    def _convert_legacy_with_soffice(self, input_path, temp_docx_path):
        converter = self._get_soffice_converter()
        if not converter.available:
            raise LegacyConversionUnavailable(
                f"{self._com_unavailable_message(os.path.splitext(input_path)[1].lower())} "
                "已跳过该旧格式文件，继续处理其他可支持文件。"
            )

        converted_path, work_dir = converter.convert_to_docx(input_path, self._log)
        try:
            shutil.copy2(converted_path, temp_docx_path)
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)
        self._log("LibreOffice 文件格式转换完成。")
        
    def quit_com_app(self):
        if self._owns_com_manager:
            self.com_manager.quit()

    # ------------------------------------------------------------------
    # Conservative Chinese punctuation normalization
    # ------------------------------------------------------------------
    @staticmethod
    def _has_chinese(text):
        return bool(RE_HAS_CHINESE.search(text or ''))

    @staticmethod
    def _is_digit_or_latin(ch):
        if not ch:
            return False
        code = ord(ch)
        return (
            ch.isdigit()
            or ('A' <= ch <= 'Z')
            or ('a' <= ch <= 'z')
            or (0xFF10 <= code <= 0xFF19)
            or (0xFF21 <= code <= 0xFF3A)
            or (0xFF41 <= code <= 0xFF5A)
        )

    @staticmethod
    def _prev_non_space(text, index):
        i = index - 1
        while i >= 0 and text[i].isspace():
            i -= 1
        return text[i] if i >= 0 else ''

    @staticmethod
    def _next_non_space(text, index):
        i = index + 1
        while i < len(text) and text[i].isspace():
            i += 1
        return text[i] if i < len(text) else ''

    @classmethod
    def _is_after_digit_or_latin(cls, text, index):
        return cls._is_digit_or_latin(cls._prev_non_space(text, index))

    @classmethod
    def _is_before_digit_or_latin(cls, text, index):
        return cls._is_digit_or_latin(cls._next_non_space(text, index))

    @classmethod
    def _iter_ellipsis_segments(cls, text):
        dot_chars = {'.', '．', '。'}
        i = 0
        while i < len(text):
            ch = text[i]
            if ch not in dot_chars:
                yield i, i + 1, ch
                i += 1
                continue

            j = i + 1
            while j < len(text) and text[j] in dot_chars:
                j += 1

            if j - i >= 3 and not cls._is_after_digit_or_latin(text, i):
                yield i, j, '……'
            else:
                yield i, j, text[i:j]
            i = j

    @classmethod
    def _normalize_ellipsis(cls, text):
        return ''.join(value for _, _, value in cls._iter_ellipsis_segments(text))

    @classmethod
    def _normalize_simple_punctuation(cls, text):
        replacements = {
            ',': '，',
            '.': '。',
            '．': '。',
            ';': '；',
            ':': '：',
            '?': '？',
            '!': '！',
        }
        chars = list(text)
        for i, ch in enumerate(chars):
            if ch not in replacements:
                continue
            if ch in {'.', '．', '。'}:
                prev_is_dot = i > 0 and text[i - 1] in {'.', '．', '。'}
                next_is_dot = i + 1 < len(text) and text[i + 1] in {'.', '．', '。'}
                if prev_is_dot or next_is_dot:
                    continue
            if cls._is_after_digit_or_latin(text, i):
                continue
            chars[i] = replacements[ch]
        return ''.join(chars)

    @classmethod
    def _normalize_bracket_pairs(cls, text):
        pair_sets = [
            ({'(': '（', '（': '（'}, {')': '）', '）': '）'}),
            ({'[': '［', '［': '［'}, {']': '］', '］': '］'}),
        ]
        result = text

        for open_chars, close_chars in pair_sets:
            chars = list(result)
            stack = []
            for i, ch in enumerate(chars):
                if ch in open_chars:
                    stack.append(i)
                elif ch in close_chars and stack:
                    open_index = stack.pop()
                    content = ''.join(chars[open_index + 1:i])
                    if not cls._has_chinese(content):
                        continue
                    if cls._is_after_digit_or_latin(result, open_index):
                        continue
                    if cls._is_before_digit_or_latin(result, i):
                        continue
                    chars[open_index] = open_chars[chars[open_index]]
                    chars[i] = close_chars[ch]
            result = ''.join(chars)

        return result

    @classmethod
    def _normalize_quote_pairs(cls, text, quote_chars, left_quote, right_quote, skip_inner_latin=False):
        chars = list(text)
        open_index = None

        for i, ch in enumerate(chars):
            if ch not in quote_chars:
                continue

            prev_is_latin = cls._is_after_digit_or_latin(text, i)
            next_is_latin = cls._is_before_digit_or_latin(text, i)
            if skip_inner_latin and prev_is_latin and next_is_latin:
                continue

            if open_index is None:
                if prev_is_latin:
                    continue
                open_index = i
                continue

            content = ''.join(chars[open_index + 1:i])
            should_normalize = (
                cls._has_chinese(content)
                and not cls._is_after_digit_or_latin(text, open_index)
                and not cls._is_before_digit_or_latin(text, i)
            )
            if should_normalize:
                chars[open_index] = left_quote
                chars[i] = right_quote
            open_index = None

        return ''.join(chars)

    @classmethod
    def _normalize_symbols_in_text(cls, text):
        if not text or not cls._has_chinese(text):
            return text

        result = cls._normalize_bracket_pairs(text)
        result = cls._normalize_ellipsis(result)
        result = cls._normalize_quote_pairs(
            result,
            {'"', '“', '”', '„', '‟', '「', '」'},
            '“',
            '”',
        )
        result = cls._normalize_quote_pairs(
            result,
            {"'", '‘', '’', '‚', '‛'},
            '‘',
            '’',
            skip_inner_latin=True,
        )
        result = cls._normalize_simple_punctuation(result)
        return result

    @staticmethod
    def _restore_markdown_syntax_punctuation(text):
        """修复历史上被全角化的 Markdown 图片语法，只动明确的 `！[Image]（…）`。"""
        return RE_MD_FULLWIDTH_IMAGE.sub(r'![Image](\1)', text)

    @staticmethod
    def _stash_markdown_syntax(text):
        """把行内控制语法换成等长语义的占位符，避免符号标准化破坏渲染。"""
        spans = []

        def stash(match):
            spans.append(match.group(0))
            return '\x00%d\x00' % (len(spans) - 1)

        return RE_MD_INLINE_SYNTAX.sub(stash, text), spans

    @staticmethod
    def _restore_markdown_syntax(text, spans):
        return RE_MD_PLACEHOLDER.sub(lambda match: spans[int(match.group(1))], text)

    def _normalize_paragraph_symbols(self, para):
        if self._has_field_codes(para) or self._paragraph_has_ooxml(para, 'w:fldSimple'):
            return False
        revisions = {qn('w:' + name) for name in ('ins', 'del', 'moveFrom', 'moveTo')}
        if any(element.tag in revisions for element in para._p.iter()):
            self._log('  > 符号标准化跳过含修订的段落，保留原内容。')
            return False

        # Map visible characters to their original text nodes. Non-text XML is
        # never reconstructed, even when text and an image share a single run.
        valid_runs = {run._r for run in self._iter_inline_runs(para)}
        boundaries = {qn('w:' + name) for name in
                      ('drawing', 'pict', 'object', 'bookmarkStart', 'bookmarkEnd',
                       'commentRangeStart', 'commentRangeEnd')}
        containers = {qn('w:hyperlink'), qn('w:sdt')}
        chars, mapping = [], []
        segment, previous_container = 0, None
        for element in para._p.iter():
            if element.tag in boundaries:
                segment += 1
            if element.getparent() not in valid_runs:
                continue
            value = self._inline_child_text(element)
            if not value:
                continue
            container = tuple(parent for parent in element.iterancestors()
                              if parent.tag in containers)
            if container != previous_container:
                segment += 1
                previous_container = container
            chars.extend(value)
            mapping.extend((element, offset, segment) if element.tag == qn('w:t') else None
                           for offset in range(len(value)))
        text = ''.join(chars)
        normalized = self._normalize_symbols_in_text(text)
        if normalized == text:
            return False
        edits = {}
        # Brackets, quotes and simple punctuation preserve character counts.
        # Ellipsis segments carry exact original ranges, so an existing "…"
        # beside "..." cannot attract an unrelated diff match or lose styling.
        position = 0
        for start, end, value in self._iter_ellipsis_segments(self._normalize_bracket_pairs(text)):
            replacement = normalized[position:position + len(value)]
            position += len(value)
            if replacement == text[start:end]:
                continue
            source = mapping[start:end]
            if not source or any(item is None for item in source):
                self._log('  > 符号标准化保留无法安全映射的文字片段。')
                continue
            if end - start == len(replacement):
                for item, char in zip(source, replacement):
                    node, offset, _ = item
                    edits.setdefault(node, []).append((offset, offset + 1, char))
            elif len({item[2] for item in source}) != 1:
                self._log('  > 符号标准化保留跨对象、书签或超链接边界的变长片段。')
            else:
                # A length-changing punctuation span may cross formatted runs,
                # but not a protected boundary. Keep surrounding text/styles.
                first = True
                for node, offset, _ in source:
                    edits.setdefault(node, []).append((offset, offset + 1,
                                                       replacement if first else ''))
                    first = False
        for node, changes in edits.items():
            value = node.text or ''
            for start, end, replacement in sorted(changes, reverse=True):
                value = value[:start] + replacement + value[end:]
            node.text = value
            if value and (value[0].isspace() or value[-1].isspace()):
                node.set(qn('xml:space'), 'preserve')
        return bool(edits)

    def _normalize_document_symbols(self, doc):
        changes = 0

        for para in doc.paragraphs:
            if self._normalize_paragraph_symbols(para):
                changes += 1

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if self._normalize_paragraph_symbols(para):
                            changes += 1

        return changes

    # ------------------------------------------------------------------
    # 排版清洗：多余空格与 PDF 双层序号（均在标题定级之前执行）
    # ------------------------------------------------------------------
    @staticmethod
    def _leading_list_prefix_length(text):
        """行首列表前缀（可叠加，如「- (1) 」）的长度，其中的空格必须保留。"""
        value = text or ''
        end = 0
        while True:
            match = RE_MD_LIST_PREFIX.match(value[end:])
            if not match or match.end() == 0:
                break
            end += match.end()
        return end

    @classmethod
    def _leading_protected_prefix_length(cls, text):
        """行首标题井号后的空格与列表前缀中的空格都必须原样保留。"""
        value = text or ''
        heading = RE_MD_HEADING_PREFIX.match(value)
        if heading:
            return heading.end()
        return cls._leading_list_prefix_length(value)

    @classmethod
    def _remove_cjk_spaces_in_text(cls, text):
        """全角空格先统一成半角，再去掉中文语境下汉字/字母/数字/标点之间的空格。

        纯英文行（无汉字）只把全角空格换成半角，不动原有半角空格，避免误伤
        "Cor-Ban 27L"、"version 1.2" 这类英文数字组合。行首列表前缀与标题井号
        后的空格保留，否则「2. (a) 内容」会丢列表结构、「## 标题」会丢 Markdown
        大纲。
        """
        value = (text or '').replace('\u3000', ' ')
        if not RE_HAS_CHINESE.search(value):
            return value
        keep = cls._leading_protected_prefix_length(value)
        prefix, body = value[:keep], value[keep:]
        return prefix + RE_CJK_SPACE_BEFORE.sub('', RE_CJK_SPACE_AFTER.sub('', body))

    @classmethod
    def _cjk_space_edits(cls, text):
        """生成 Word 段落的文字改写区间：全角空格→半角、多余空格→空。"""
        normalized = (text or '').replace('\u3000', ' ')
        keep = cls._leading_protected_prefix_length(normalized)
        edits = {}
        if RE_HAS_CHINESE.search(normalized):
            for pattern in (RE_CJK_SPACE_AFTER, RE_CJK_SPACE_BEFORE):
                for match in pattern.finditer(normalized, keep):
                    edits[match.start()] = (match.start(), match.end(), '')
        for index, char in enumerate(text or ''):
            if char == '\u3000' and index not in edits:
                edits[index] = (index, index + 1, ' ')
        return sorted(edits.values())

    @classmethod
    def _remove_duplicate_numbering_in_text_line(cls, line):
        """「3. 3．标题」→「3. 标题」；序号不一致时保持原样，避免误删真实层级。

        docling 可能叠加多层编号（「3. 3． 3. 标题」），反复剥离直到收敛。
        """
        result = line or ''
        for _ in range(DUPLICATE_NUMBER_MAX_PASSES):
            match = RE_DUPLICATE_NUMBER.match(result)
            if not match or int(match.group(1)) != int(match.group(3)):
                break
            updated = result[:match.end(2)] + ' ' + result[match.end(0):]
            if updated == result:
                break
            result = updated
        return result

    def _paragraph_text_nodes(self, para):
        """收集段落可见字符及其来源 w:t 节点，非文本对象位置记为 None。"""
        valid_runs = {run._r for run in self._iter_inline_runs(para)}
        chars, mapping = [], []
        for element in para._p.iter():
            if element.getparent() not in valid_runs:
                continue
            value = self._inline_child_text(element)
            if not value:
                continue
            chars.extend(value)
            mapping.extend(
                (element, offset) if element.tag == qn('w:t') else None
                for offset in range(len(value))
            )
        return ''.join(chars), mapping

    def _apply_paragraph_text_edits(self, para, edits):
        """按可见字符区间改写 w:t，跨 run 安全；区间含非文本节点时跳过该区间。"""
        _, mapping = self._paragraph_text_nodes(para)
        node_edits = {}
        for start, end, replacement in edits:
            source = mapping[start:end]
            if not source or any(item is None for item in source):
                continue
            for index, (node, offset) in enumerate(source):
                node_edits.setdefault(node, []).append(
                    (offset, offset + 1, replacement if index == 0 else '')
                )
        for node, changes in node_edits.items():
            value = node.text or ''
            for start, end, replacement in sorted(changes, reverse=True):
                value = value[:start] + replacement + value[end:]
            node.text = value
            if value and (value[0].isspace() or value[-1].isspace()):
                node.set(qn('xml:space'), 'preserve')
        return bool(node_edits)

    def _cleanable_paragraph(self, para):
        if self._has_field_codes(para) or self._paragraph_has_ooxml(para, 'w:fldSimple'):
            return False
        revisions = {qn('w:' + name) for name in ('ins', 'del', 'moveFrom', 'moveTo')}
        return not any(element.tag in revisions for element in para._p.iter())

    def _iter_document_paragraphs(self, doc):
        yield from doc.paragraphs
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from cell.paragraphs

    def _remove_cjk_spaces_in_document(self, doc):
        changes = 0
        for para in self._iter_document_paragraphs(doc):
            if not self._cleanable_paragraph(para):
                continue
            text, _ = self._paragraph_text_nodes(para)
            edits = self._cjk_space_edits(text)
            if edits and self._apply_paragraph_text_edits(para, edits):
                changes += 1
        return changes

    @staticmethod
    def _md_fence_line_indices(lines, source_ext):
        """Markdown 围栏代码块（含围栏行）的行号，清洗时必须跳过。"""
        if source_ext != '.md':
            return set()
        result, in_fence = set(), False
        for index, line in enumerate(lines):
            if RE_MD_FENCE.match(line):
                result.add(index)
                in_fence = not in_fence
            elif in_fence:
                result.add(index)
        return result

    @classmethod
    def _spurious_leading_number_indices(cls, lines, skip=()):
        """判定哪些「N. (a)」的前置数字是 docling 页面模型叠加的冗余序号。

        不靠两个序号是否相等，而看前后连续性：
        1. 同一段落里子标号在相邻字母上延续（a→b、b→c…），说明 (a)/(b) 才是真实子项；
        2. 最近的裸父项标号「(n)」与行首数字不一致，说明该数字来自页面级计数。
        """
        skipped = set(skip)
        candidates = []
        for index, line in enumerate(lines):
            if index in skipped:
                continue
            prefix = RE_LEADING_NUMBER_WITH_ALPHA.match(line)
            if not prefix:
                continue
            label = RE_BARE_PAREN_LABEL.match(line[prefix.end():])
            if not label or not label.group(1).isalpha():
                continue
            candidates.append((index, int(prefix.group(1)), label.group(1).lower()))

        by_letter = {}
        for index, _, letter in candidates:
            by_letter.setdefault(letter, []).append(index)

        result = set()
        for index, number, letter in candidates:
            if (chr(ord(letter) + 1) in by_letter
                    or chr(ord(letter) - 1) in by_letter):
                result.add(index)
                continue
            for previous in range(index - 1, max(-1, index - 12), -1):
                if previous in skipped:
                    continue
                parent = RE_BARE_PAREN_LABEL.match(lines[previous])
                if parent and parent.group(1).isdigit():
                    if int(parent.group(1)) != number:
                        result.add(index)
                    break

        # 字母子标号的证据只在本轮内沿用，避免纯数字续行被清掉后把后面的
        # 真实序号「4. 若水分…」「5. 若发现…」级联误删。
        alpha_result = set(result)
        numeric_cont = cls._numbered_continuation_indices(lines, skipped)
        result |= numeric_cont

        # 同一页面计数器也会叠到警告/说明行上。
        for index, line in enumerate(lines):
            if index in skipped or index in result:
                continue
            if RE_DUPLICATE_NUMBER.match(line):
                continue  # 重复序号模式已单独处理
            prefix = RE_PAGE_NUMBER_PREFIX.match(line)
            if not prefix:
                continue
            body = line[prefix.end():]
            if RE_NOTE_LINE.match(body):
                result.add(index)
                continue
            if not alpha_result:
                continue  # 无字母子标号证据时不凭裸父项猜数字，避免误删真实编号
            for previous in range(index - 1, max(-1, index - 4), -1):
                if previous in skipped or not lines[previous].strip():
                    continue
                if (previous in alpha_result
                        or RE_BARE_PAREN_LABEL.match(lines[previous])
                        or RE_BARE_NUM_PAREN_LABEL.match(lines[previous])
                        or RE_LEADING_NUMBER_WITH_ALPHA.match(lines[previous])):
                    result.add(index)
                break
        return result

    @classmethod
    def _numbered_continuation_indices(cls, lines, skipped=()):
        """行首序号叠在前行断行续文上的行号（纯数字列表场景）。

        单独成方法，供冗余序号清理和续行合并共用：这里确认过的续行必须直接
        接回上一行，否则 jieba 跨词判据未必能识别「小」+「水泡」这类拼接。
        """
        result = set()
        for index, line in enumerate(lines):
            if index in skipped or RE_DUPLICATE_NUMBER.match(line):
                continue
            prefix = RE_PAGE_NUMBER_PREFIX.match(line)
            if not prefix or RE_NOTE_LINE.match(line[prefix.end():]):
                continue
            for previous in range(index - 1, max(-1, index - 4), -1):
                if previous in skipped or not lines[previous].strip():
                    continue
                if cls._numbered_continuation_prefix(lines[previous], line):
                    result.add(index)
                break
        return result

    @staticmethod
    def _numbered_continuation_prefix(prev_line, curr_line):
        """判断当前行的行首数字是否是叠在前行断行续文上的页面序号，返回要删的前缀长度。

        典型场景：前行「3. …沉淀物或小」无标点收尾，下一行「4. 水泡，…；」
        以句末标点收尾，序号还向前跳跃（4→6、5→8 等）。这类数字来自页面级
        计数器，应删除后让断行合并把两段接回一行。
        """
        prev_prefix = RE_PAGE_NUMBER_PREFIX.match(prev_line)
        curr_prefix = RE_PAGE_NUMBER_PREFIX.match(curr_line)
        if not prev_prefix or not curr_prefix:
            return 0
        if int(curr_prefix.group(1)) <= int(prev_prefix.group(1)):
            return 0  # 假序号总是向前叠加；回退/持平是真实列表项接续
        prev_body = prev_line[prev_prefix.end():]
        curr_body = curr_line[curr_prefix.end():]
        if WordProcessor._sentence_ends_at(prev_body) or not WordProcessor._sentence_ends_at(curr_body):
            return 0
        prev_last = prev_body.strip().rstrip(MERGE_SENTENCE_END_CLOSERS).rstrip()[-1:]
        if prev_last in MERGE_CONNECTING_PUNCTUATION:
            return 0  # 逗号类收尾是明确续行，交给断行合并处理，不必动序号
        return curr_prefix.end()

    def _remove_duplicate_numbering_in_document(self, doc):
        paragraphs = list(self._iter_document_paragraphs(doc))
        texts = []
        for para in paragraphs:
            text, _ = self._paragraph_text_nodes(para)
            texts.append(text)

        changes = 0
        for para, text in zip(paragraphs, texts):
            if not self._cleanable_paragraph(para):
                continue
            current = text
            cleaned = False
            for _ in range(DUPLICATE_NUMBER_MAX_PASSES):
                match = RE_DUPLICATE_NUMBER.match(current.lstrip())
                if not match or int(match.group(1)) != int(match.group(3)):
                    break
                shift = len(current) - len(current.lstrip())
                if not self._apply_paragraph_text_edits(
                    para, [(shift + match.end(2), shift + match.end(0), ' ')]
                ):
                    break
                cleaned = True
                updated, _ = self._paragraph_text_nodes(para)
                if updated == current:
                    break
                current = updated
            if cleaned:
                changes += 1

        spurious = self._spurious_leading_number_indices(texts)
        numeric_cont = self._numbered_continuation_indices(texts)
        for index in sorted(spurious):
            para = paragraphs[index]
            if not self._cleanable_paragraph(para):
                continue
            prefix = (RE_LEADING_NUMBER_WITH_ALPHA.match(texts[index])
                      or RE_PAGE_NUMBER_PREFIX.match(texts[index]))
            if not prefix:
                continue
            if self._apply_paragraph_text_edits(para, [(0, prefix.end(), '')]):
                changes += 1
            if index in numeric_cont and index:
                previous = paragraphs[index - 1]
                if (previous is not None
                        and previous._p.getparent() is para._p.getparent()
                        and self._cleanable_paragraph(previous)
                        and self._paragraph_is_plain_runs(previous)
                        and self._paragraph_is_plain_runs(para)):
                    self._merge_paragraph_range(paragraphs, index - 1, index)
        return changes

    def _remove_duplicate_numbering_in_text(self, text, source_ext):
        lines = text.split('\n')
        # 围栏代码块整段原样保留：必须在逐行剥离序号之前确定跳过区间，
        # 否则代码行里的「3. 3．…」会被当成 PDF 冗余序号改写。
        skipped = self._md_fence_line_indices(lines, source_ext)
        lines = [
            line if index in skipped else self._remove_duplicate_numbering_in_text_line(line)
            for index, line in enumerate(lines)
        ]
        spurious = self._spurious_leading_number_indices(lines, skipped)
        numeric_cont = self._numbered_continuation_indices(lines, skipped)
        result = []
        for index, line in enumerate(lines):
            if index not in spurious:
                result.append(line)
                continue
            prefix = (RE_LEADING_NUMBER_WITH_ALPHA.match(line)
                      or RE_PAGE_NUMBER_PREFIX.match(line))
            if not prefix:
                result.append(line)
                continue
            body = line[prefix.end():]
            if (index in numeric_cont and index and (index - 1) not in skipped
                    and result and self._text_line_is_mergeable(result[-1], source_ext)):
                result[-1] = self._join_text_lines(result[-1], body)
            else:
                result.append(body)
        return '\n'.join(result)

    def _apply_text_cleanup(self, text, source_ext, transform):
        """逐行应用文本清洗，Markdown 围栏代码块原样保留。"""
        result, in_fence = [], False
        for line in text.split('\n'):
            if source_ext == '.md' and RE_MD_FENCE.match(line):
                in_fence = not in_fence
                result.append(line)
                continue
            result.append(line if in_fence else transform(line))
        return '\n'.join(result)

    def _normalize_text_symbols(self, text, source_ext):
        def normalize_line(line):
            if source_ext != '.md':
                return self._normalize_symbols_in_text(line)
            line = self._restore_markdown_syntax_punctuation(line)
            if RE_MD_STRUCTURAL_LINE.match(line):
                return line
            protected, spans = self._stash_markdown_syntax(line)
            if not spans:
                return self._normalize_symbols_in_text(line)
            return self._restore_markdown_syntax(
                self._normalize_symbols_in_text(protected), spans
            )

        return self._apply_text_cleanup(text, source_ext, normalize_line)

    # ------------------------------------------------------------------
    # Markdown cleanup
    # ------------------------------------------------------------------
    @staticmethod
    def _clean_markdown(text):
        """
        Clean Markdown content to plain text:
        1. Remove images, links, HTML, inline code markers
        2. Remove bold/italic markers (*, **, __)
        3. Remove heading markers (#)
        4. Remove blockquote markers (>)
        5. Remove horizontal rules (---)
        6. Preserve original ordered-list numbering from the source text
        """
        if not text:
            return ""

        # Global inline element replacements
        # Remove images: ![alt](url) -> alt
        text = RE_MD_IMAGE.sub(r'\1', text)
        # Remove links: [text](url) -> text
        text = RE_MD_LINK.sub(r'\1', text)
        # Remove HTML tags
        text = RE_MD_HTML_TAG.sub('', text)
        # Remove inline code: `code` -> code
        text = RE_MD_INLINE_CODE.sub(r'\1', text)
        # Remove bold/italic (** or __)
        text = RE_MD_BOLD_ASTERISK.sub(r'\1', text)
        text = RE_MD_BOLD_UNDERSCORE.sub(r'\1', text)

        lines = text.split('\n')
        new_lines = []

        for line in lines:
            cleaned_line = line

            # Remove heading markers: # Title -> Title
            header_match = RE_MD_HEADER.match(cleaned_line)
            if header_match:
                cleaned_line = header_match.group(1)

            # Remove blockquote markers: > Text -> Text
            blockquote_match = RE_MD_BLOCKQUOTE.match(cleaned_line)
            if blockquote_match:
                cleaned_line = blockquote_match.group(1)

            # Remove horizontal rules: ---, ***, ___
            if RE_MD_HORIZONTAL_RULE.match(cleaned_line):
                cleaned_line = ""

            # Ordered-list numbers are intentionally kept exactly as written.

            # Remove remaining * italic markers (careful not to break list markers)
            is_bullet = RE_MD_UNORDERED_LIST.match(cleaned_line)

            def remove_stars(m):
                return m.group(1)

            if is_bullet:
                bullet_match = RE_MD_BULLET_WITH_CONTENT.match(cleaned_line)
                if bullet_match:
                    marker = bullet_match.group(1)
                    content = bullet_match.group(2)
                    content = RE_MD_EMPHASIS_ASTERISK.sub(remove_stars, content)
                    cleaned_line = marker + content
            else:
                cleaned_line = RE_MD_EMPHASIS_ASTERISK.sub(remove_stars, cleaned_line)

            new_lines.append(cleaned_line)

        return '\n'.join(new_lines)

    # ------------------------------------------------------------------
    # Blank line removal for plain text sources
    # ------------------------------------------------------------------
    @staticmethod
    def _remove_blank_lines_from_text(text, keep_single_blank_lines=False):
        """
        Normalize blank lines from txt/md text:
        - 2+ consecutive blank lines: merge to 1
        - Single blank line: delete by default, or keep when requested
        """
        if not text:
            return text

        lines = text.split('\n')
        result = []
        blank_count = 0

        for line in lines:
            if line.strip() == '':
                blank_count += 1
            else:
                if blank_count >= 2 or (blank_count == 1 and keep_single_blank_lines):
                    result.append('')
                result.append(line)
                blank_count = 0

        if blank_count >= 2 or (blank_count == 1 and keep_single_blank_lines):
            result.append('')

        return '\n'.join(result)

    def _normalize_text_blank_lines(self, text):
        if self.blank_line_mode == BLANK_LINE_MODE_PRESERVE:
            return text
        return self._remove_blank_lines_from_text(
            text,
            keep_single_blank_lines=self.blank_line_mode == BLANK_LINE_MODE_KEEP_SINGLE
        )

    def _log_blank_line_mode(self, source_name):
        if self.blank_line_mode == BLANK_LINE_MODE_PRESERVE:
            self._log(f"  > 未改动 {source_name} 中的空行。")
        elif self.blank_line_mode == BLANK_LINE_MODE_KEEP_SINGLE:
            self._log(f"  > 已保留 {source_name} 中的单个空行，并将多个空行合并为 1 个。")
        else:
            self._log(f"  > 已删除 {source_name} 中的单个空行，并将多个空行合并为 1 个。")

    # ------------------------------------------------------------------
    # Text file reading
    # ------------------------------------------------------------------
    @staticmethod
    def _read_text_file(path):
        """Try multiple encodings to read text file"""
        with open(path, 'rb') as f:
            bom = f.read(2)
        # UTF-16 文本文件（记事本/WPS 另存常见），BOM 决定字节序；
        # UTF-8 必须先试 utf-8-sig，否则 BOM 会作为 U+FEFF 留在正文里
        # （破坏 Markdown 标题、行首列表等判据）。
        encodings = ['utf-16'] if bom in (b'\xff\xfe', b'\xfe\xff') else [
            'utf-8-sig', 'utf-8', 'gbk', 'gb18030'
        ]
        for enc in encodings:
            try:
                with open(path, 'r', encoding=enc) as f:
                    content = f.read()
                return content
            except UnicodeDecodeError:
                continue
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()

    def convert_to_docx(self, input_path):
        file_ext = os.path.splitext(input_path)[1].lower()
        is_from_txt = (file_ext in ('.txt', '.md'))
        base_name = os.path.splitext(os.path.basename(input_path))[0]

        if file_ext == '.docx':
            self._log("检测到 .docx 文件，正在创建安全的处理副本...")
            temp_docx_path = self._make_temp_docx_path("copy", base_name)
            shutil.copy2(input_path, temp_docx_path)
            self._log(f"  > 副本创建成功: {os.path.basename(temp_docx_path)}")
            return temp_docx_path, False

        temp_docx_path = self._make_temp_docx_path("converted", base_name)

        if file_ext == '.txt':
            self._log("检测到 .txt 文件，正在创建 .docx...")
            text_content = self._read_text_file(input_path)
            text_content = self._normalize_text_blank_lines(text_content)
            self._log_blank_line_mode("TXT")
            doc = Document()
            for line in text_content.split('\n'):
                doc.add_paragraph(line.strip())
            doc.save(temp_docx_path)
            self._log("TXT转换完成。")
            return temp_docx_path, is_from_txt
        elif file_ext == '.md':
            self._log("检测到 .md 文件，正在清理 Markdown 标记并创建 .docx...")
            raw_text = self._read_text_file(input_path)
            cleaned_text = self._clean_markdown(raw_text)
            cleaned_text = self._normalize_text_blank_lines(cleaned_text)
            self._log_blank_line_mode("Markdown 文本")
            doc = Document()
            for line in cleaned_text.split('\n'):
                doc.add_paragraph(line.strip())
            doc.save(temp_docx_path)
            self._log("Markdown 转换完成。")
            return temp_docx_path, is_from_txt
        elif file_ext in ['.wps', '.doc']:
            self._log(f"正在转换 {file_ext} 文件为 .docx...")
            if self._com_available():
                try:
                    self._convert_legacy_with_com(input_path, temp_docx_path)
                    return temp_docx_path, is_from_txt
                except Exception as exc:
                    self._log(f"  > WPS/Word 转换失败，尝试使用 LibreOffice 兜底: {exc}")
            else:
                self._log(f"  > {self._com_unavailable_message(file_ext)}")

            self._convert_legacy_with_soffice(input_path, temp_docx_path)
            return temp_docx_path, is_from_txt
        
        raise ValueError(f"不支持的文件格式: {file_ext}")

    def _preprocess_document(self, docx_path, is_from_txt=False):
        """Return a working copy with platform-appropriate numbering handling."""
        if is_from_txt:
            return docx_path
        # Windows keeps COM first. If Office is absent or preprocessing fails,
        # its existing Python formatting fallback now also converts numbering.
        if IS_WINDOWS:
            if self._preprocess_com_tasks(docx_path):
                return docx_path
            self._log("  > WPS/Word 预处理不可用，回退到 Python 自动编号转文字；修订不自动接受。")
        elif not self.config.get('convert_auto_numbering', True):
            self._log("  > 自动编号转文字已关闭，保留原自动编号。")
            return docx_path

        from wfp_numbering import convert_docx_numbering

        self._log("正在对副本识别自动编号并转换为文字（包含表格）...")
        output_path = self._make_temp_docx_path("numbering", Path(docx_path).stem)
        report = convert_docx_numbering(
            docx_path, output_path, strict=False, preserve_suffix=False,
        )
        if report.warnings:
            self._log(f"  > 自动编号部分转换：已转换 {report.converted} 个段落；"
                      "不支持的编号保留原状，继续排版。")
            for warning in report.warnings:
                self._log(f"  > 编号转换警告：{warning}")
        else:
            self._log(f"  > 自动编号转换完成，共转换 {report.converted} 个段落。")
        return output_path

    def _preprocess_com_tasks(self, docx_path):
        if not self._com_available():
            self._log(f"  > {self._com_unavailable_message()}")
            return False
        self._log("正在对副本执行预处理（接受所有修订、转换自动编号）...")
        doc_com = None
        try:
            app = self._get_wps_app()
            doc_com = app.Documents.Open(os.path.abspath(docx_path))
            
            doc_com.TrackRevisions = False
            self._log("  > 已关闭修订追踪。")
            
            if doc_com.Revisions.Count > 0:
                doc_com.AcceptAllRevisions()
                self._log("  > 已接受文档副本中的所有修订。")
            
            doc_com.Content.ListFormat.ConvertNumbersToText()
            self._log("  > 已将副本中的自动编号转换为文本。")
            
            if doc_com.Revisions.Count > 0:
                doc_com.AcceptAllRevisions()
                self._log("  > 已接受编号转换产生的修订。")
            
            doc_com.TrackRevisions = False
            
            doc_com.Save()
            self._log("预处理完成。")
            return True
        except Exception as e:
            self._log(f"警告：执行预处理任务时出错: {e}")
            return False
        finally:
            if doc_com is not None:
                try:
                    doc_com.Close()
                except Exception as e:
                    self._log(f"  > 警告：关闭预处理文档时发生异常: {e}")

    def _create_page_number(self, paragraph, text):
        font_name = self.config['page_number_font']
        font_size = self.config['page_number_size']
        self._set_run_font(paragraph.add_run('— '), font_name, font_size, set_color=True)
        run_field = paragraph.add_run()
        self._set_run_font(run_field, font_name, font_size, set_color=True)
        fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = text
        fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
        run_field._r.extend([fldChar1, instrText, fldChar2])
        self._set_run_font(paragraph.add_run(' —'), font_name, font_size, set_color=True)

    def _apply_page_setup(self, doc, is_from_txt=False):
        self._log("正在应用页面边距和页码设置...")
        
        # 判断是否需要强制设置A4纸
        # 逻辑：如果是纯文本来源（包括直接输入）或者 用户勾选了强制A4，则设置为A4
        should_set_a4 = is_from_txt or self.config.get('force_a4', False)

        for section in doc.sections:
            section.top_margin = Cm(self.config['margin_top'])
            section.bottom_margin = Cm(self.config['margin_bottom'])
            section.left_margin = Cm(self.config['margin_left'])
            section.right_margin = Cm(self.config['margin_right'])
            section.footer_distance = Cm(self.config['footer_distance'])

            # 设置纸张大小为A4 (仅在需要时)
            if should_set_a4:
                section.page_width = Cm(21)
                section.page_height = Cm(29.7)

            if self.config['page_number_align'] == '居中':
                p = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
                p.clear(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; self._create_page_number(p, 'PAGE')
            elif self.config['page_number_align'] == '奇偶分页':
                doc.settings.odd_and_even_pages_header_footer = True
                footer_odd = section.footer
                p_odd = footer_odd.paragraphs[0] if footer_odd.paragraphs else footer_odd.add_paragraph()
                p_odd.clear(); p_odd.alignment = WD_ALIGN_PARAGRAPH.RIGHT; self._create_page_number(p_odd, 'PAGE')
                
                footer_even = section.even_page_footer
                p_even = footer_even.paragraphs[0] if footer_even.paragraphs else footer_even.add_paragraph()
                p_even.clear(); p_even.alignment = WD_ALIGN_PARAGRAPH.LEFT; self._create_page_number(p_even, 'PAGE')
        
        if should_set_a4:
            self._log("  > 已将页面大小设置为 A4。")

    def _run_font_family(self, font_name):
        """按配置给出西文字体（数字、字母）：未启用自定义时与中文字体一致。"""
        if self.config.get('use_custom_english_font', False):
            en_font = self.config.get('english_font') or font_name
            return en_font
        return font_name

    def _set_run_font(self, run, font_name, size_pt, set_color=False):
        run.font.size = Pt(size_pt)
        if set_color: run.font.color.rgb = RGBColor(0, 0, 0)
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        for theme_attr in ('w:eastAsiaTheme', 'w:asciiTheme', 'w:hAnsiTheme', 'w:cstheme', 'w:csTheme'):
            rFonts.attrib.pop(qn(theme_attr), None)
        rFonts.set(qn('w:eastAsia'), font_name)
        # 根据配置决定西文字体（数字、字母）
        en_font = self._run_font_family(font_name)
        run.font.name = en_font
        rFonts.set(qn('w:ascii'), en_font)
        rFonts.set(qn('w:hAnsi'), en_font)

    def _apply_font_to_runs(self, para, font_name, size_pt, set_color=False):
        for run in self._iter_inline_runs(para):
            self._set_run_font(run, font_name, size_pt, set_color=set_color)

    @staticmethod
    def _iter_inline_runs(para):
        """Include hyperlink/content-control runs, excluding nested text boxes/deletions."""
        for element in para._p.iter(qn('w:r')):
            parent = element.getparent()
            while parent is not None and parent is not para._p:
                if parent.tag in (qn('w:p'), qn('w:del'), qn('w:moveFrom')):
                    break
                parent = parent.getparent()
            if parent is para._p:
                yield Run(element, para)

    @staticmethod
    def _inline_child_text(child):
        if child.tag == qn('w:t'):
            return child.text or ''
        if child.tag == qn('w:tab'):
            return '\t'
        if child.tag in (qn('w:br'), qn('w:cr')):
            return '\n'
        return ''

    @classmethod
    def _inline_run_text(cls, run):
        return ''.join(cls._inline_child_text(child) for child in run._r)

    @classmethod
    def _paragraph_text(cls, para):
        return ''.join(cls._inline_run_text(run) for run in cls._iter_inline_runs(para))

    @classmethod
    def _split_inline_run(cls, run, offset):
        """Split only text at a boundary; move other XML children without duplicating them."""
        right = OxmlElement('w:r')
        right.attrib.update(run._r.attrib)
        properties = run._r.find(qn('w:rPr'))
        if properties is not None:
            right.append(deepcopy(properties))
        remaining = offset
        for child in list(run._r):
            if child.tag == qn('w:rPr'):
                continue
            text = cls._inline_child_text(child)
            if remaining == 0:
                right.append(child)
            elif len(text) > remaining:
                tail = deepcopy(child)
                child.text, tail.text = text[:remaining], text[remaining:]
                child.set(qn('xml:space'), 'preserve')
                tail.set(qn('xml:space'), 'preserve')
                right.append(tail)
                remaining = 0
            else:
                remaining -= len(text)
        run._r.addnext(right)

    @classmethod
    def _split_runs_at(cls, para, boundary):
        consumed = 0
        for run in list(cls._iter_inline_runs(para)):
            length = len(cls._inline_run_text(run))
            if consumed < boundary < consumed + length:
                cls._split_inline_run(run, boundary - consumed)
                return
            consumed += length

    def _format_inline_heading(self, para, title_len, apply_color):
        self._split_runs_at(para, title_len)
        consumed = 0
        for run in self._iter_inline_runs(para):
            length = len(self._inline_run_text(run))
            if length:
                is_title = consumed < title_len
                prefix = 'h2' if is_title else 'body'
                self._set_run_font(run, self.config[prefix + '_font'],
                                   self.config[prefix + '_size'], set_color=apply_color)
                if is_title:
                    self._set_run_bold(run, self.config.get('h2_bold', 'preserve'))
            consumed += length

    def _normalize_heading_brackets(self, para):
        text = self._paragraph_text(para)
        match = RE_H2_INLINE_TITLE.match(text)
        if not match:
            return
        replacements = {0: '（', len(match.group(1)) + 1: '）'}
        consumed = 0
        for run in self._iter_inline_runs(para):
            for child in run._r:
                original = self._inline_child_text(child)
                if child.tag == qn('w:t'):
                    child.text = ''.join(replacements.get(consumed + i, ch)
                                         for i, ch in enumerate(original))
                consumed += len(original)

    def _get_paragraph_font_info(self, para):
        """获取段落主要字体和字号信息"""
        if not para.runs:
            return None, None
        
        # 获取第一个非空run的字体信息
        for run in self._iter_inline_runs(para):
            if self._inline_run_text(run).strip():
                font_name = run.font.name
                font_size = run.font.size.pt if run.font.size else None
                return font_name, font_size
        return None, None

    @staticmethod
    def _get_paragraph_alignment(para):
        try:
            return para.alignment
        except ValueError:
            p_pr = para._p.pPr
            jc = p_pr.jc if p_pr is not None else None
            raw_value = jc.get(qn("w:val")) if jc is not None else None
            return {
                "start": WD_ALIGN_PARAGRAPH.LEFT,
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "end": WD_ALIGN_PARAGRAPH.RIGHT,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "distribute": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "mediumKashida": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "highKashida": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "lowKashida": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "thaiDistribute": WD_ALIGN_PARAGRAPH.JUSTIFY,
            }.get(raw_value)

    def _strip_leading_whitespace(self, para):
        if not para.runs:
            return
        while para.runs:
            first_run = para.runs[0]
            r_elem = first_run._r
            children = [child for child in r_elem if child.tag != qn('w:rPr')]
            only_text = all(child.tag == qn('w:t') for child in children)
            if not only_text or first_run.text.strip():
                break
            para._p.remove(r_elem)

        if not para.runs:
            return
        first_run = para.runs[0]
        original_text = first_run.text
        stripped_text = original_text.lstrip()
        if original_text != stripped_text:
            first_run.text = stripped_text
            self._log("  > 已移除段落前的多余空格。")
    
    def _reset_pagination_properties(self, para):
        para.paragraph_format.widow_control = False
        para.paragraph_format.keep_with_next = False
        para.paragraph_format.keep_lines_together = False
        para.paragraph_format.page_break_before = False
        para.paragraph_format.keep_together = False

    def _get_outline_level(self, para):
        """
        读取段落的当前大纲级别
        返回: 0-8 表示级别1-9，9 表示正文级（显式覆盖样式），None 表示未设置
        """
        pPr = para._p.get_or_add_pPr()
        outlineLvl = pPr.find(qn('w:outlineLvl'))
        if outlineLvl is not None:
            val = outlineLvl.get(qn('w:val'))
            if val is not None:
                return int(val)
        return None

    def _set_outline_level(self, para, level):
        """
        直接设置段落的大纲级别，不通过样式，不影响字体字号等格式
        level: 1-9 的整数，表示大纲级别
        返回: 原有的大纲级别 (0-8) 或 None
        """
        if level < 1 or level > 9:
            self._log(f"  > 警告：大纲级别 {level} 超出范围 (1-9)，已跳过设置")
            return None
        
        # 读取原有大纲级别；正文级（9）视为未设置，避免日志出现 Lv10
        original_level = self._get_outline_level(para)
        if original_level is not None and original_level >= 9:
            original_level = None
        
        # 设置新的大纲级别 (Word内部用0-8表示1-9级)
        pPr = para._p.get_or_add_pPr()
        outlineLvl = pPr.find(qn('w:outlineLvl'))
        if outlineLvl is None:
            outlineLvl = OxmlElement('w:outlineLvl')
            pPr.append(outlineLvl)
        outlineLvl.set(qn('w:val'), str(level - 1))
        
        return original_level

    def _format_heading(self, para, level):
        """
        为段落设置大纲级别（仅设置大纲级别，不影响其他格式）
        """
        if not self.config['set_outline']:
            self._log(f"  > 大纲级别设置已禁用，跳过")
            return
        
        # 获取段落文本预览用于日志
        text_preview = para.text.strip()[:30].replace("\n", " ")
        
        original_level = self._set_outline_level(para, level)
        
        if original_level is not None:
            self._log(f"  > 大纲级别: Lv{original_level + 1} → Lv{level} (覆盖) - \"{text_preview}...\"")
        else:
            self._log(f"  > 大纲级别: 无 → Lv{level} (新设) - \"{text_preview}...\"")

    def _apply_text_indent_and_align(self, para):
        pf = para.paragraph_format
        indent_unit = self._normalize_indent_unit(self.config.get('paragraph_indent_unit', 'cm'))
        pf.first_line_indent = None
        pf.left_indent = None
        pf.right_indent = None

        ind = para._p.get_or_add_pPr().get_or_add_ind()
        ind.attrib.pop(qn('w:left'), None)
        ind.attrib.pop(qn('w:right'), None)
        ind.attrib.pop(qn('w:leftChars'), None)
        ind.attrib.pop(qn('w:rightChars'), None)
        ind.attrib.pop(qn('w:hanging'), None)
        ind.attrib.pop(qn('w:hangingChars'), None)
        ind.attrib.pop(qn('w:firstLine'), None)
        ind.attrib.pop(qn('w:firstLineChars'), None)

        if indent_unit == 'chars':
            left_chars = self._config_float(self.config, 'left_indent_chars', 0.0)
            right_chars = self._config_float(self.config, 'right_indent_chars', 0.0)
            ind.set(qn('w:leftChars'), self._chars_to_word_value(left_chars))
            ind.set(qn('w:rightChars'), self._chars_to_word_value(right_chars))
        else:
            pf.left_indent = Cm(self._config_float(self.config, 'left_indent_cm', 0.0))
            pf.right_indent = Cm(self._config_float(self.config, 'right_indent_cm', 0.0))

        first_line_chars = self._config_float(self.config, 'first_line_indent_chars', 2.0)
        if first_line_chars > 0:
            ind.set(qn("w:firstLineChars"), self._chars_to_word_value(first_line_chars))

        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _iter_block_items(self, parent):
        parent_elm = parent.element.body if isinstance(parent, _Document) else parent._tc
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P): yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl): yield Table(child, parent)

    def _get_or_add_table_pr(self, table):
        tbl = table._tbl
        tbl_pr = tbl.tblPr
        if tbl_pr is None:
            tbl_pr = OxmlElement('w:tblPr')
            tbl.insert(0, tbl_pr)
        return tbl_pr

    def _set_table_borders(self, table, size_pt=0.5, color="000000"):
        size = max(1, int(float(size_pt) * 8))
        tbl_pr = self._get_or_add_table_pr(table)
        borders = tbl_pr.find(qn('w:tblBorders'))
        if borders is None:
            borders = OxmlElement('w:tblBorders')
            tbl_pr.append(borders)
        else:
            for child in list(borders):
                borders.remove(child)

        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            elem = OxmlElement(f'w:{edge}')
            elem.set(qn('w:val'), 'single')
            elem.set(qn('w:sz'), str(size))
            elem.set(qn('w:space'), '0')
            elem.set(qn('w:color'), color)
            borders.append(elem)

    def _set_cell_borders(self, cell, size_pt=0.5, color="000000"):
        size = max(1, int(float(size_pt) * 8))
        tc = cell._tc
        tc_pr = tc.tcPr
        if tc_pr is None:
            tc_pr = OxmlElement('w:tcPr')
            tc.insert(0, tc_pr)

        borders = tc_pr.find(qn('w:tcBorders'))
        if borders is None:
            borders = OxmlElement('w:tcBorders')
            tc_pr.append(borders)
        else:
            for child in list(borders):
                borders.remove(child)

        for edge in ('top', 'left', 'bottom', 'right'):
            elem = OxmlElement(f'w:{edge}')
            elem.set(qn('w:val'), 'single')
            elem.set(qn('w:sz'), str(size))
            elem.set(qn('w:space'), '0')
            elem.set(qn('w:color'), color)
            borders.append(elem)

    def _set_table_cell_margins(self, table, top_cm=0.0, bottom_cm=0.0, left_cm=0.05, right_cm=0.05):
        tbl_pr = self._get_or_add_table_pr(table)
        cell_mar = tbl_pr.find(qn('w:tblCellMar'))
        if cell_mar is None:
            cell_mar = OxmlElement('w:tblCellMar')
            tbl_pr.append(cell_mar)

        def set_side(tag, cm_value):
            node = cell_mar.find(qn(f'w:{tag}'))
            if node is None:
                node = OxmlElement(f'w:{tag}')
                cell_mar.append(node)
            node.set(qn('w:type'), 'dxa')
            node.set(qn('w:w'), str(int(Cm(float(cm_value)).twips)))

        set_side('top', top_cm)
        set_side('bottom', bottom_cm)
        set_side('left', left_cm)
        set_side('right', right_cm)

    def _set_table_width_percent(self, table, percent=100):
        percent = max(1, min(100, int(float(percent))))
        tbl_pr = self._get_or_add_table_pr(table)
        tbl_w = tbl_pr.find(qn('w:tblW'))
        if tbl_w is None:
            tbl_w = OxmlElement('w:tblW')
            tbl_pr.append(tbl_w)
        tbl_w.set(qn('w:type'), 'pct')
        tbl_w.set(qn('w:w'), str(percent * 50))

    def _set_table_indent(self, table, indent_twips=0):
        tbl_pr = self._get_or_add_table_pr(table)
        tbl_ind = tbl_pr.find(qn('w:tblInd'))
        if tbl_ind is None:
            tbl_ind = OxmlElement('w:tblInd')
            tbl_pr.append(tbl_ind)
        tbl_ind.set(qn('w:type'), 'dxa')
        tbl_ind.set(qn('w:w'), str(int(indent_twips)))

    @staticmethod
    def _table_text_weight(text):
        weight = 0.0
        for ch in text:
            weight += 0.5 if ord(ch) < 128 else 1.0
        return weight

    @staticmethod
    def _normalize_table_pcts(weights, min_pct, max_pct):
        total = sum(weights) or 1.0
        pcts = [w / total * 100 for w in weights]
        for i, value in enumerate(pcts):
            if value < min_pct:
                pcts[i] = min_pct
            elif value > max_pct:
                pcts[i] = max_pct
        total = sum(pcts) or 1.0
        return [value / total * 100 for value in pcts]

    def _set_table_col_widths_by_content(self, table, min_pct=8, max_pct=45):
        if not table.rows:
            return
        col_count = max(len(row.cells) for row in table.rows)
        if col_count == 0:
            return

        min_pct = max(1.0, float(min_pct))
        max_pct = max(min_pct, float(max_pct))
        max_weights = [1.0] * col_count
        for row in table.rows:
            for col_idx, cell in enumerate(row.cells):
                text = ''.join(p.text for p in cell.paragraphs).strip()
                if text:
                    max_weights[col_idx] = max(max_weights[col_idx], self._table_text_weight(text))

        pcts = self._normalize_table_pcts(max_weights, min_pct, max_pct)
        tbl = table._tbl
        tbl_grid = tbl.tblGrid
        if tbl_grid is None:
            tbl_grid = OxmlElement('w:tblGrid')
            tbl.insert(0, tbl_grid)
        else:
            for child in list(tbl_grid):
                tbl_grid.remove(child)

        for pct in pcts:
            grid_col = OxmlElement('w:gridCol')
            grid_col.set(qn('w:w'), str(int(pct * 50)))
            tbl_grid.append(grid_col)

        for row in table.rows:
            for col_idx, cell in enumerate(row.cells):
                tc = cell._tc
                tc_pr = tc.tcPr
                if tc_pr is None:
                    tc_pr = OxmlElement('w:tcPr')
                    tc.insert(0, tc_pr)
                tc_w = tc_pr.find(qn('w:tcW'))
                if tc_w is None:
                    tc_w = OxmlElement('w:tcW')
                    tc_pr.append(tc_w)
                tc_w.set(qn('w:type'), 'pct')
                tc_w.set(qn('w:w'), str(int(pcts[col_idx] * 50)))

    @staticmethod
    def _is_numeric_table_text(text):
        text = (text or '').strip()
        if not text:
            return False
        text = text.replace(',', '').replace('，', '').replace('％', '%')
        text = RE_CURRENCY_PREFIX.sub('', text)
        text = RE_CURRENCY_SUFFIX.sub('', text)
        return RE_NUMERIC_TABLE_TEXT.match(text) is not None

    @staticmethod
    def _is_short_table_text(text, max_len=4):
        text = (text or '').strip()
        return 0 < len(text) <= int(max_len)

    @staticmethod
    def _config_float(config, key, default):
        value = config.get(key, default)
        try:
            if value == '':
                return default
            return float(value)
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _normalize_spacing_unit(value):
        value = str(value or 'pt').strip().lower()
        if value in ('multiple', '倍数', 'multi'):
            return 'multiple'
        return 'pt'

    @staticmethod
    def _normalize_indent_unit(value):
        value = str(value or 'cm').strip().lower()
        if value in ('chars', 'char', 'character', 'characters', '字符'):
            return 'chars'
        return 'cm'

    @staticmethod
    def _chars_to_word_value(value):
        return str(int(round(max(0.0, float(value)) * 100)))

    def _apply_line_spacing(self, para, pt_key, unit_key, multiple_key, default_pt, default_multiple=1.0):
        self._apply_line_spacing_to_pPr(
            para._p.get_or_add_pPr(),
            pt_key,
            unit_key,
            multiple_key,
            default_pt,
            default_multiple,
            paragraph_format=para.paragraph_format,
        )

    def _apply_line_spacing_to_pPr(
        self, p_pr, pt_key, unit_key, multiple_key, default_pt,
        default_multiple=1.0, paragraph_format=None,
    ):
        """把行距写进 pPr（段落或样式定义共用）。"""
        unit = self._normalize_spacing_unit(self.config.get(unit_key, 'pt'))
        spacing = p_pr.get_or_add_spacing()
        if unit == 'multiple':
            multiple = self._config_float(self.config, multiple_key, default_multiple)
            if multiple <= 0:
                multiple = default_multiple
            spacing.set(qn('w:line'), str(int(round(multiple * 240))))
            spacing.set(qn('w:lineRule'), 'auto')
            return

        line_spacing_pt = self._config_float(self.config, pt_key, default_pt)
        if line_spacing_pt <= 0:
            if paragraph_format is not None:
                paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            spacing.attrib.pop(qn('w:line'), None)
            return
        spacing.set(qn('w:line'), str(int(round(line_spacing_pt * 20))))
        spacing.set(qn('w:lineRule'), 'exact')

    @staticmethod
    def _apply_bold_to_runs(para, enabled):
        if not enabled:
            return
        for run in para.runs:
            run.font.bold = True

    @staticmethod
    def _apply_bold_to_first_chars(para, char_count, enabled):
        if not enabled:
            return
        consumed = 0
        for run in para.runs:
            run_len = len(run.text or '')
            if run_len and consumed < char_count:
                run.font.bold = True
            consumed += run_len

    @staticmethod
    def _table_header_indices(table):
        """返回表头行下标：优先 w:tblHeader（Word 重复标题行），否则回退首行。"""
        indices = []
        for index, row in enumerate(table.rows):
            tr_pr = row._tr.trPr
            node = tr_pr.find(qn('w:tblHeader')) if tr_pr is not None else None
            if node is not None and node.get(qn('w:val'), 'true') not in ('0', 'false', 'off'):
                indices.append(index)
        if not indices and table.rows:
            indices = [0]
        return indices

    @staticmethod
    def _table_header_signature(table, indices):
        """用表头行的单元格文字生成可比较的签名。"""
        return tuple(
            tuple(''.join(p.text for p in cell.paragraphs).strip() for cell in table.rows[i].cells)
            for i in indices
        )

    def _merge_repeated_table_headers(self, doc):
        """合并相邻且表头相同的表格，去掉后续表格的重复表头行。

        Word 开启“重复标题行”后，跨页长表格常被拆成多个连续表格，每个表格
        都带同一表头。这里把它们的正文行并入第一个表格，删除重复表格与重复
        表头；表头相同由 w:tblHeader 标记行（无标记时取首行）逐格文字比对。
        """
        body = doc.element.body
        pending = None
        merged_groups = merged_tables = 0

        for child in list(body.iterchildren()):
            if isinstance(child, CT_P):
                para = Paragraph(child, doc)
                if (
                    pending is not None
                    and not para.text.strip()
                    and not self._has_drawing_or_pict(para)
                    and not self._has_embedded_object(para)
                ):
                    continue
                pending = None
                continue
            if not isinstance(child, CT_Tbl):
                pending = None
                continue

            table = Table(child, doc)
            indices = self._table_header_indices(table)
            if not indices:
                pending = None
                continue
            signature = self._table_header_signature(table, indices)

            if (
                pending is not None
                and signature == pending[2]
                and len(table.columns) == len(pending[1].columns)
            ):
                header_set = set(indices)
                for row_idx, row in enumerate(table.rows):
                    if row_idx not in header_set:
                        pending[1]._tbl.append(deepcopy(row._tr))
                body.remove(child)
                merged_tables += 1
                continue

            if pending is not None and merged_tables:
                merged_groups += 1
            pending = (child, table, signature)

        if pending is not None and merged_tables:
            merged_groups += 1
        if merged_tables:
            self._log(
                f"相邻相同表头表格合并完成：合并 {merged_tables} 个表格、"
                f"{merged_groups} 组，已去除重复表头。"
            )
        return merged_tables

    def _format_tables(self, doc, apply_color=True):
        if not self.config.get('enable_table_formatting', False):
            self._log("表格自动调整未启用，跳过表格内容格式化。")
            return

        tables = list(doc.tables)
        if not tables:
            self._log("未发现表格，跳过表格内容格式化。")
            return

        table_font = self.config.get('table_font', self.config.get('body_font', '仿宋_GB2312'))
        table_header_font = self.config.get('table_header_font', table_font)
        table_size = self._config_float(self.config, 'table_size', self.config.get('body_size', 12))
        row_height_cm = self._config_float(self.config, 'table_row_height_cm', 0.7)
        border_size_pt = self._config_float(self.config, 'table_border_size_pt', 0.5)
        width_percent = self._config_float(self.config, 'table_width_percent', 100)
        col_min_pct = self._config_float(self.config, 'table_col_min_pct', 8)
        col_max_pct = self._config_float(self.config, 'table_col_max_pct', 45)
        short_text_len = self._config_float(self.config, 'table_short_text_len', 4)
        auto_col_width = self.config.get('table_auto_col_width', True)
        header_bold = self.config.get('table_header_bold', 'bold')
        cell_alignment = self.config.get('table_cell_alignment', 'preserve')
        table_alignment = self.config.get('table_alignment', 'preserve')
        unified_borders = self.config.get('table_border_mode', 'unified') == 'unified'
        alignments = {'left': WD_ALIGN_PARAGRAPH.LEFT, 'center': WD_ALIGN_PARAGRAPH.CENTER,
                      'right': WD_ALIGN_PARAGRAPH.RIGHT}
        table_alignments = {'left': WD_TABLE_ALIGNMENT.LEFT, 'center': WD_TABLE_ALIGNMENT.CENTER,
                            'right': WD_TABLE_ALIGNMENT.RIGHT}

        self._log(f"开始格式化表格内容（共 {len(tables)} 个）...")
        for table_idx, table in enumerate(tables, start=1):
            self._log(f"  > 表格 {table_idx}: 调整宽度、行高、字体和单元格格式")
            table.autofit = not auto_col_width
            self._set_table_width_percent(table, width_percent)
            self._set_table_cell_margins(table)
            if table_alignment != 'preserve':
                self._set_table_indent(table, 0)
                table.alignment = table_alignments[table_alignment]
            if self.config.get('table_remove_text_wrapping', False):
                tbl_pr = self._get_or_add_table_pr(table)
                for floating in list(tbl_pr.findall(qn('w:tblpPr'))):
                    tbl_pr.remove(floating)
            if unified_borders:
                self._set_table_borders(table, size_pt=border_size_pt)
            if auto_col_width:
                self._set_table_col_widths_by_content(table, min_pct=col_min_pct, max_pct=col_max_pct)

            serial_col_idx = None
            if table.rows:
                for col_idx, cell in enumerate(table.rows[0].cells):
                    head_text = ''.join(self._paragraph_text(p) for p in cell.paragraphs).strip()
                    if '序号' in head_text or head_text == '序':
                        serial_col_idx = col_idx
                        break

            for row_idx, row in enumerate(table.rows):
                if row_height_cm > 0:
                    row.height = Cm(row_height_cm)
                    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

                for col_idx, cell in enumerate(row.cells):
                    if unified_borders:
                        self._set_cell_borders(cell, size_pt=border_size_pt)

                    cell_text = ''.join(self._paragraph_text(p) for p in cell.paragraphs).strip()
                    for para in cell.paragraphs:
                        if self._paragraph_text(para).strip():
                            for run in self._iter_inline_runs(para):
                                font_name = table_header_font if row_idx == 0 else table_font
                                self._set_run_font(run, font_name, table_size, set_color=apply_color)
                                if row_idx == 0:
                                    self._set_run_bold(run, header_bold)

                        para.paragraph_format.first_line_indent = Pt(0)
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)
                        self._apply_line_spacing(
                            para,
                            'table_line_spacing',
                            'table_line_spacing_unit',
                            'table_line_spacing_multiple',
                            22,
                        )

                        if cell_alignment == 'smart':
                            if row_idx == 0:
                                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            elif '合计' in cell_text or '总计' in cell_text:
                                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            elif serial_col_idx is not None and col_idx == serial_col_idx:
                                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            elif self._is_numeric_table_text(cell_text):
                                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            elif self._is_short_table_text(cell_text, short_text_len):
                                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            else:
                                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        elif cell_alignment != 'preserve':
                            para.alignment = alignments[cell_alignment]
    
    def _find_title_and_subtitle_paragraphs(self, doc, is_from_txt, start_index=0, exclude_indices=None, stop_at_body=False, end_index=None):
        """
        查找题目和副标题段落的索引范围
        返回: (title_indices, subtitle_indices)
        title_indices: 题目行的索引列表
        subtitle_indices: 副标题行的索引列表
        end_index: 可选上界（如首个附件段落索引），限定题目识别范围，不含该索引
        """
        all_blocks = list(self._iter_block_items(doc))
        scan_end = len(all_blocks) if end_index is None else min(end_index, len(all_blocks))
        
        # 查找首个标题行
        first_title_idx = -1
        
        if is_from_txt:
            self._log("文档源自 TXT，采用智能规则查找题目...")
            for idx in range(start_index, scan_end):
                block = all_blocks[idx]
                if isinstance(block, Paragraph) and block.text.strip():
                    text_to_check = block.text.strip()
                    if RE_TITLE_H1.match(text_to_check) or RE_TITLE_H2.match(text_to_check):
                        self._log(f"  > 首个非空行 (块 {idx + 1}) 符合标题格式，认定本文档无独立题目。")
                        return [], []
                    else:
                        self._log(f"  > 在块 {idx + 1} 发现首个非空段落，认定为题目首行。")
                        first_title_idx = idx
                        break
        else:
            self._log("正在预扫描以确定居中题目位置...")
            for idx in range(start_index, scan_end):
                block = all_blocks[idx]
                if not isinstance(block, Paragraph) or not self._paragraph_text(block).strip():
                    continue
                para = block
                text_to_check = para.text.lstrip()
                if RE_TITLE_H1.match(text_to_check) or RE_TITLE_H2.match(text_to_check):
                    self._log("  > 发现一级/二级标题，在此之前未找到居中题目。")
                    return [], []
                if self._get_paragraph_alignment(para) == WD_ALIGN_PARAGRAPH.CENTER:
                    if exclude_indices and idx in exclude_indices:
                        self._log("  > 发现的居中段落已被处理过，不再视为题目。")
                        return [], []
                    self._log(f"  > 在块 {idx + 1} 发现潜在题目首行。")
                    first_title_idx = idx
                    break
                if stop_at_body:
                    self._log("  > 题目位置之前发现正文段落，视为无独立题目。")
                    return [], []
        
        if first_title_idx == -1:
            self._log("  > 扫描结束，未能找到题目。")
            return [], []
        
        # 获取首个标题行的字体字号信息
        first_title_para = all_blocks[first_title_idx]
        title_font, title_size = self._get_paragraph_font_info(first_title_para)
        
        # 向下查找连续的标题行
        title_indices = [first_title_idx]
        idx = first_title_idx + 1
        
        while idx < len(all_blocks) and (end_index is None or idx < end_index):
            block = all_blocks[idx]
            if not isinstance(block, Paragraph):
                break
            
            para = block
            text = self._paragraph_text(para).strip()
            
            # 遇到空行，停止标题识别
            if not text:
                self._log(f"  > 在块 {idx + 1} 遇到空行，标题识别结束。")
                break
            
            # 检查是否居中
            if self._get_paragraph_alignment(para) != WD_ALIGN_PARAGRAPH.CENTER:
                break
            
            # 检查字体字号是否与首行相同
            para_font, para_size = self._get_paragraph_font_info(para)
            if para_font == title_font and para_size == title_size:
                self._log(f"  > 块 {idx + 1} 也是标题行（居中且字体字号相同）。")
                title_indices.append(idx)
                idx += 1
            else:
                # 字体字号不同，可能是副标题的开始
                break
        
        self._log(f"  > 共识别到 {len(title_indices)} 行标题。")
        
        # 查找副标题
        subtitle_indices = []
        subtitle_start_idx = idx
        
        # 跳过空行
        while subtitle_start_idx < len(all_blocks):
            block = all_blocks[subtitle_start_idx]
            if isinstance(block, Paragraph) and self._paragraph_text(block).strip():
                break
            if isinstance(block, Paragraph) and self._attachment_blank(block):
                subtitle_start_idx += 1
            else:
                # 遇到非段落（如表格），停止
                break
        
        # 检查是否有副标题
        if subtitle_start_idx < len(all_blocks):
            block = all_blocks[subtitle_start_idx]
            if isinstance(block, Paragraph):
                para = block
                text = self._paragraph_text(para).strip()
                
                # 副标题必须居中
                if text and self._get_paragraph_alignment(para) == WD_ALIGN_PARAGRAPH.CENTER:
                    # 检查字体字号是否与标题不同
                    para_font, para_size = self._get_paragraph_font_info(para)
                    if para_font != title_font or para_size != title_size:
                        self._log(f"  > 在块 {subtitle_start_idx + 1} 发现副标题首行（居中且字体字号与标题不同）。")
                        subtitle_indices.append(subtitle_start_idx)
                        
                        # 查找连续的副标题行
                        subtitle_font, subtitle_size = para_font, para_size
                        idx = subtitle_start_idx + 1
                        
                        while idx < len(all_blocks):
                            block = all_blocks[idx]
                            if not isinstance(block, Paragraph):
                                break
                            
                            para = block
                            text = self._paragraph_text(para).strip()
                            
                            # 遇到空行，停止副标题识别
                            if not text:
                                self._log(f"  > 在块 {idx + 1} 遇到空行，副标题识别结束。")
                                break
                            
                            # 检查是否居中
                            if self._get_paragraph_alignment(para) != WD_ALIGN_PARAGRAPH.CENTER:
                                break
                            
                            # 检查字体字号是否与副标题首行相同
                            para_font, para_size = self._get_paragraph_font_info(para)
                            if para_font == subtitle_font and para_size == subtitle_size:
                                self._log(f"  > 块 {idx + 1} 也是副标题行（居中且字体字号相同）。")
                                subtitle_indices.append(idx)
                                idx += 1
                            else:
                                break
                        
                        self._log(f"  > 共识别到 {len(subtitle_indices)} 行副标题。")
        
        return title_indices, subtitle_indices

    # ------------------------------------------------------------------
    # 可选增强一：浮动式图片 / 嵌入对象 / 表格 → 插入式
    # ------------------------------------------------------------------
    VML_SHAPE_TAGS = {
        'arc', 'curve', 'group', 'image', 'line', 'oval', 'polyline',
        'rect', 'roundrect', 'shape',
    }
    VML_FLOATING_PROPS = {
        'position', 'z-index', 'left', 'top', 'margin-left', 'margin-top',
        'mso-position-horizontal', 'mso-position-vertical',
        'mso-position-horizontal-relative', 'mso-position-vertical-relative',
        'mso-left-percent', 'mso-top-percent',
    }

    @staticmethod
    def _local_tag(tag):
        """返回不含命名空间的标签名。"""
        if not isinstance(tag, str):
            return ''
        return tag.rsplit('}', 1)[-1] if '}' in tag else tag

    @classmethod
    def _vml_shape_to_inline(cls, shape):
        """把 VML 浮动图形改为随文字排布。返回是否修改。"""
        style = shape.get('style')
        if not style:
            return False
        kept, floating = [], False
        for declaration in style.split(';'):
            piece = declaration.strip()
            if not piece:
                continue
            prop = piece.split(':', 1)[0].strip().lower()
            if prop == 'position':
                if 'absolute' in piece.lower():
                    floating = True
                continue
            if prop in cls.VML_FLOATING_PROPS:
                floating = True
                continue
            kept.append(piece)
        if not floating:
            return False
        shape.set('style', ';'.join(kept))
        return True

    @classmethod
    def _graphic_extent(cls, graphic):
        """缺 wp:extent 时，从 graphic 子树里取 cx/cy 作为兜底尺寸。"""
        for node in graphic.iter():
            if cls._local_tag(node.tag) != 'ext':
                continue
            cx, cy = node.get('cx'), node.get('cy')
            if cx and cy:
                return cx, cy
        return '0', '0'

    @classmethod
    def _anchor_to_inline(cls, drawing, next_docpr_id):
        """把 drawing 中的 wp:anchor 改写为 wp:inline。返回是否转换成功。"""
        anchor = None
        for child in drawing:
            if cls._local_tag(child.tag) == 'anchor':
                anchor = child
                break
        if anchor is None:
            return False

        parts = {}
        for child in anchor:
            parts.setdefault(cls._local_tag(child.tag), child)
        graphic = parts.get('graphic')
        if graphic is None:
            return False

        inline = OxmlElement('wp:inline')
        for attr in ('distT', 'distB', 'distL', 'distR'):
            inline.set(attr, '0')

        if 'extent' in parts:
            inline.append(copy.deepcopy(parts['extent']))
        else:
            extent = OxmlElement('wp:extent')
            cx, cy = cls._graphic_extent(graphic)
            extent.set('cx', cx)
            extent.set('cy', cy)
            inline.append(extent)

        effect = parts.get('effectExtent')
        if effect is not None:
            inline.append(copy.deepcopy(effect))

        doc_pr = copy.deepcopy(parts['docPr']) if 'docPr' in parts else OxmlElement('wp:docPr')
        if not (doc_pr.get('id') or '').strip():
            doc_pr.set('id', str(next_docpr_id()))
        inline.append(doc_pr)

        cnv = parts.get('cNvGraphicFramePr')
        if cnv is not None:
            inline.append(copy.deepcopy(cnv))
        inline.append(copy.deepcopy(graphic))

        drawing.remove(anchor)
        drawing.append(inline)
        return True

    def _convert_floating_to_inline(self, doc):
        """把浮动排版的图片、嵌入对象和表格转换为插入式。"""
        start_id = 1
        for node in doc.element.body.iter():
            if self._local_tag(node.tag) == 'docPr':
                try:
                    start_id = max(start_id, int(node.get('id') or 0) + 1)
                except ValueError:
                    pass
        id_source = itertools.count(start_id)

        drawings = 0
        for drawing in doc.element.body.iter(qn('w:drawing')):
            if self._anchor_to_inline(drawing, lambda: next(id_source)):
                drawings += 1

        shapes = 0
        for pict in doc.element.body.iter(qn('w:pict')):
            for node in pict.iter():
                if self._local_tag(node.tag) in self.VML_SHAPE_TAGS and self._vml_shape_to_inline(node):
                    shapes += 1

        tables = 0
        for tbl in doc.element.body.iter(qn('w:tbl')):
            tbl_pr = tbl.find(qn('w:tblPr'))
            if tbl_pr is None:
                continue
            floating_pr = tbl_pr.find(qn('w:tblpPr'))
            if floating_pr is not None:
                tbl_pr.remove(floating_pr)
                tables += 1
            overlap = tbl_pr.find(qn('w:tblOverlap'))
            if overlap is not None:
                tbl_pr.remove(overlap)

        self._log(
            f"浮动对象转插入式：图片/嵌入对象 {drawings} 处，VML 图形 {shapes} 处，环绕表格 {tables} 个。"
        )
        return drawings + shapes + tables

    # ------------------------------------------------------------------
    # 可选增强二：含图片 / 嵌入对象的段落改为单倍行距
    # ------------------------------------------------------------------
    def _iter_document_paragraphs(self, doc):
        """遍历正文与表格（含嵌套表格）中的所有段落。"""
        def walk_cell(cell):
            for para in cell.paragraphs:
                yield para
            for nested in cell.tables:
                yield from walk_table(nested)

        def walk_table(table):
            for row in table.rows:
                for cell in row.cells:
                    yield from walk_cell(cell)

        for para in doc.paragraphs:
            yield para
        for table in doc.tables:
            yield from walk_table(table)

    @staticmethod
    def _force_single_line_spacing(para):
        """写为单倍行距（w:line=240 / w:lineRule=auto）。返回是否发生变化。"""
        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        changed = (
            spacing.get(qn('w:line')) not in (None, '240')
            or spacing.get(qn('w:lineRule')) not in (None, 'auto')
        )
        spacing.set(qn('w:line'), '240')
        spacing.set(qn('w:lineRule'), 'auto')
        return changed

    def _apply_media_single_line_spacing(self, doc):
        """含图片或嵌入对象的段落改为单倍行距，避免固定行距裁切对象。"""
        changed = 0
        for para in self._iter_document_paragraphs(doc):
            if not (self._has_drawing_or_pict(para) or self._has_embedded_object(para)):
                continue
            if self._force_single_line_spacing(para):
                changed += 1
        self._log(f"图片/嵌入对象段落改为单倍行距：{changed} 个段落。")
        return changed

    # ------------------------------------------------------------------
    # 可选增强三：相同格式标题等级合并 + 补齐跳级
    # ------------------------------------------------------------------
    @staticmethod
    def _flag_value(node):
        """读取 w:b / w:i 之类的开关值，返回 True/False/None。"""
        if node is None:
            return None
        raw = (node.get(qn('w:val')) or '').strip().lower()
        return raw not in ('0', 'false', 'off')

    @classmethod
    def _get_style_outline_level(cls, style):
        """读取段落样式的 outlineLvl（含 basedOn 链），返回 1-9 或 None。"""
        seen = 0
        while style is not None and seen < 6:
            element = getattr(style, 'element', None)
            if element is None:
                return None
            p_pr = element.find(qn('w:pPr'))
            if p_pr is not None:
                node = p_pr.find(qn('w:outlineLvl'))
                value = node.get(qn('w:val')) if node is not None else None
                if value is not None:
                    try:
                        level = int(value) + 1
                    except ValueError:
                        return None
                    return level if 1 <= level <= MAX_TITLE_OUTLINE_LEVEL else None
            style = getattr(style, 'base_style', None)
            seen += 1
        return None

    @classmethod
    def _get_heading_style_info(cls, para):
        """返回 (是否标题样式或样式自带级别, 样式自带等级)。"""
        try:
            style = para.style
        except Exception:
            return False, None
        if style is None:
            return False, None
        name = (getattr(style, 'name', '') or '').strip()
        style_id = (getattr(style, 'style_id', '') or '').strip()
        style_level = cls._get_style_outline_level(style)
        has_heading_style = (
            bool(RE_HEADING_STYLE_NAME.match(name))
            or bool(RE_HEADING_STYLE_NAME.match(style_id))
            or style_level is not None
        )
        return has_heading_style, style_level

    @classmethod
    def _paragraph_format_signature(cls, para):
        """按字体、字号、加粗、斜体生成格式指纹，信息不足时回退到段落样式的定义。"""
        r_pr = None
        for run in para.runs:
            if run.text.strip():
                r_pr = run._r.find(qn('w:rPr'))
                break
        if r_pr is None:
            p_pr = para._p.find(qn('w:pPr'))
            if p_pr is not None:
                r_pr = p_pr.find(qn('w:rPr'))
        if r_pr is None:
            return cls._style_format_signature(para)

        fonts = None
        r_fonts = r_pr.find(qn('w:rFonts'))
        if r_fonts is not None:
            fonts = tuple(
                r_fonts.get(qn(f'w:{attr}')) or '' for attr in ('ascii', 'eastAsia', 'hAnsi', 'cs')
            )
        size_node = r_pr.find(qn('w:sz')) if r_pr.find(qn('w:sz')) is not None else r_pr.find(qn('w:szCs'))
        size = size_node.get(qn('w:val')) if size_node is not None else None
        bold = cls._flag_value(r_pr.find(qn('w:b')))
        italic = cls._flag_value(r_pr.find(qn('w:i')))
        if fonts is None and size is None and bold is None and italic is None:
            # 字体字号已交给标准样式统一时，用样式定义里的格式作为指纹
            return cls._style_format_signature(para)
        return (fonts, size, bold, italic)

    @classmethod
    def _style_format_signature(cls, para):
        """取段落样式（含 basedOn 链）提供的格式指纹，样式也没有信息时返回 None。"""
        try:
            style = para.style
        except Exception:  # noqa: BLE001 - 损坏的 pStyle 指向，按无信息处理
            return None
        fonts = size = bold = italic = None
        seen = 0
        while style is not None and seen < 6:
            element = getattr(style, 'element', None)
            r_pr = element.find(qn('w:rPr')) if element is not None else None
            if r_pr is not None:
                if fonts is None:
                    r_fonts = r_pr.find(qn('w:rFonts'))
                    if r_fonts is not None:
                        fonts = tuple(
                            r_fonts.get(qn(f'w:{attr}')) or ''
                            for attr in ('ascii', 'eastAsia', 'hAnsi', 'cs')
                        )
                if size is None:
                    size_node = r_pr.find(qn('w:sz'))
                    if size_node is None:
                        size_node = r_pr.find(qn('w:szCs'))
                    if size_node is not None:
                        size = size_node.get(qn('w:val'))
                if bold is None:
                    bold = cls._flag_value(r_pr.find(qn('w:b')))
                if italic is None:
                    italic = cls._flag_value(r_pr.find(qn('w:i')))
            style = getattr(style, 'base_style', None)
            seen += 1
        if fonts is None and size is None and bold is None and italic is None:
            return None
        return (fonts, size, bold, italic)

    @classmethod
    def _dotted_number_level(cls, text, offset=0, base=3, allow_single=False):
        """按点分编号深度推断大纲级别：同一文档内同一层级的级别固定，且下一层 = 上一层 + 1。

        以 base 为第一层（单级编号）的基准层级：
        - 单级编号（`1.` / `A.`，或点分编号为主干时的裸 `1` / `A`）→ base；
        - 多级点分编号（`1.1` / `A.B` / `1.A` / `1.1.1` / `1.1.1.1`，末尾点可有可无）按段数逐级 +1：
          1.→base、1.1→base+1、1.1.1→base+2、1.1.1.1→base+3（更深封顶 MAX_TITLE_OUTLINE_LEVEL）。

        base 由文档结构决定（见 _prepare_dotted_base）：文档仅用点分多级编号且无传统标题（一、/（一）/章/节）
        时 base=1（点分编号作为主干，1. 即一级标题）；否则 base=3（保持公文三级标题约定）。
        编号段可以是字母或数字；多级编号要求段落整体像标题，避免把 "1.5 倍" 这类小数误判为标题。
        allow_single 为真（文档使用多级点分编号）时才识别裸单级编号（如自动编号转文字后的 “1标题”）。
        无法推断时返回 None；offset 为章/节带来的顺延级数。该识别优先于段落既有样式/格式。

        「x.x.x 下游从 1 重新起头的连续编号」需要文档顺序上下文，由
        _prepare_single_number_restart 单独判定为其下一级（见 _paragraph_dotted_level），
        本函数只做不依赖上下文的纯文本深度推断。
        """
        value = (text or '').lstrip()
        offset = int(offset or 0)
        multi = RE_DOTTED_NUMBER.match(value)
        if multi:
            if not cls._looks_like_heading_text(value):
                return None
            depth = multi.group(1).count('.') + 1
            if depth >= 2:
                # 多级点分编号层级固定：x.x→2、x.x.x→3、x.x.x.x→4、x.x.x.x.x→5。
                # 不受「1.」基准级别或章/节顺延影响，避免传统标题体系下 x.x 被推到四级。
                return min(depth, MAX_TITLE_OUTLINE_LEVEL)
            return min(base + offset, MAX_TITLE_OUTLINE_LEVEL)
        # 单级编号 + 点（1. / A.）：保持原有三级标题约定，不加标题样判据
        if RE_HEADING_H3.match(value):
            return min(base + offset, MAX_TITLE_OUTLINE_LEVEL)
        # 裸单级编号（无点，如 1标题 / A标题）：仅在文档确实使用多级点分编号时启用
        if allow_single and RE_SINGLE_LABEL.match(value):
            return min(base + offset, MAX_TITLE_OUTLINE_LEVEL)
        return None

    @classmethod
    def _dotted_number_depth(cls, text):
        """返回文本开头点分编号的段数（1.1 → 2、1.1.1 → 3），无多级点分编号时返回 0。"""
        m = RE_DOTTED_NUMBER.match((text or '').lstrip())
        return m.group(1).count('.') + 1 if m else 0

    @classmethod
    def _apply_dotted_level_floor(cls, text, level):
        """多级点分编号段落的大纲级别按深度固定：x.x→Lv2、x.x.x→Lv3、x.x.x.x→Lv4、x.x.x.x.x→Lv5。

        自带大纲级别/标题样式同样服从该固定层级，封顶 MAX_TITLE_OUTLINE_LEVEL；
        单级编号（1. / 1标题）与正文（None）不受影响。
        """
        if level is None:
            return None
        depth = cls._dotted_number_depth(text)
        if depth >= 2:
            return min(depth, MAX_TITLE_OUTLINE_LEVEL)
        return level

    def _dotted_style_heading_level(self, para, text):
        """以多级点分编号开头、文字不像标题、但自带大纲级别/标题样式的段落：
        视为标题，级别取自带级别与编号深度下限中的较深者（1.1 → 至少 Lv2）。"""
        if self._dotted_number_depth(text) < 2:
            return None
        outline_level = self._effective_outline_level(para)
        if outline_level is None:
            return None
        return self._apply_dotted_level_floor(text, outline_level)

    @classmethod
    def _numbered_heading_level(cls, text, offset=0, base=3, allow_single=False):
        """按编号特征推断的大纲级别，无法推断时返回 None；offset 为章/节带来的顺延级数，
        base 为点分编号「1.」的基准层级，allow_single 见 _dotted_number_level。

        点分多级编号（1./1.1./1.1.1./1.1.1.1.）按深度识别，优先于一、/（一）/1./（1）等单层编号。
        """
        value = (text or '').lstrip()
        dotted = cls._dotted_number_level(value, offset, base, allow_single)
        if dotted is not None:
            return dotted
        level = None
        if RE_HEADING_H1.match(value):
            level = 1
        elif RE_HEADING_H2.match(value):
            level = 2
        elif RE_HEADING_H3.match(value):
            level = 3
        elif RE_HEADING_H4.match(value) or RE_HEADING_H4_ALPHA.match(value):
            level = 4
        if level is None:
            return None
        return min(level + int(offset or 0), MAX_TITLE_OUTLINE_LEVEL)

    @staticmethod
    def _looks_like_heading_text(text):
        """保守判断：短、句末与句中都没有断句标点的段落，才参与无证据的同格式归并。"""
        value = (text or '').strip()
        if not value or len(value) > HEADING_TEXT_MAX_LENGTH:
            return False
        if value[-1] in HEADING_TAIL_PUNCTUATION:
            return False
        return not any(punct in value for punct in BODY_SENTENCE_PUNCTUATION)

    # ------------------------------------------------------------------
    # 可选增强四：章/节层级识别
    # ------------------------------------------------------------------
    def _prepare_chapter_section_offset(self, doc):
        """按文档实际层级决定章/节/条与编号标题的级数。

        文档有“第X章” → 章=1、节=2，条=3，编号标题顺延 2 级；
        只有“第X节” → 节=1，条=2，编号标题顺延 1 级；都没有则保持原有层级。
        文档含“第X条”时也按章/节结构处理（条标题需要其上级层级），否则只在
        开启“识别章/节层级”时生效。
        """
        self._outline_offset = 0
        self._article_level = 1
        detect_chapter = self.config.get('detect_chapter_section', False)
        detect_article = self.config.get('detect_article_title', True)
        if not detect_chapter and not detect_article:
            return

        has_chapter = has_section = has_article = False
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            # 条标题常与正文同段（“第一条 为了……”），只看前缀，不做标点过滤
            if detect_article and not has_article and RE_ARTICLE_TITLE.match(text):
                has_article = True
            if len(text) > HEADING_TEXT_MAX_LENGTH:
                continue
            if any(punct in text for punct in BODY_SENTENCE_PUNCTUATION):
                continue
            if RE_CHAPTER_TITLE.match(text):
                has_chapter = True
                continue
            if not has_section and RE_SECTION_TITLE.match(text):
                has_section = True

        if has_article:
            # 条标题只在章/节之下顺延一级：章+节=3 级，只有章或只有节=2 级，都没有=1 级
            if has_chapter and has_section:
                self._article_level = 3
            elif has_chapter or has_section:
                self._article_level = 2
            else:
                self._article_level = 1

        if not detect_chapter and not (detect_article and has_article):
            return

        self._outline_offset = 2 if has_chapter else (1 if has_section else 0)
        if self._outline_offset:
            kind = '章' if has_chapter else '节'
            self._log(
                f"检测到{kind}标题，章/节按 1/2 级处理，编号标题整体顺延 {self._outline_offset} 级。"
            )
        elif not has_article:
            self._log("未检测到“第X章/第X节”标题，章/节识别不生效。")
        if has_article:
            self._log(
                f"检测到“第X条”条标题，条标题按 Lv{self._article_level} 处理，条后正文另起一段。"
            )

    def _prepare_dotted_base(self, doc):
        """按文档结构决定点分编号「1.」的基准大纲级别。

        规则（结合结构判断，而非写死三级）：
        - 文档使用了传统标题（一、/（一）/第X章/第X节）时，点分编号仍作为三级标题体系之下的层级，
          「1.」基准为 3（1.→L3、1.1.→L4 …）；
        - 文档【未】使用任何传统标题、且出现了带点的多级编号（1.1. / 1.1.1. 等，深度≥2）时，
          点分编号本身构成文档的主干大纲，「1.」作为一级标题，基准为 1（1.→L1、1.1.→L2、1.1.1.→L3、1.1.1.1.→L4）；
        - 其余（无传统标题、且只有扁平的 1./2. 列表）保持基准 3，沿用公文三级标题约定。
        """
        has_traditional = False
        has_dotted_hierarchy = False
        for para in doc.paragraphs:
            t = (para.text or '').strip()
            if not t:
                continue
            if (RE_HEADING_H1.match(t) or RE_HEADING_H2.match(t)
                    or RE_CHAPTER_TITLE.match(t) or RE_SECTION_TITLE.match(t)):
                has_traditional = True
            m = RE_DOTTED_NUMBER.match(t)
            # 只有像标题的多级编号才算主干大纲，避免正文里 "1.5 倍" 之类小数误触发 base=1
            if m and m.group(1).count('.') >= 1 and self._looks_like_heading_text(t):
                has_dotted_hierarchy = True
        self._has_dotted_hierarchy = has_dotted_hierarchy
        self._dotted_base_level = 1 if (not has_traditional and has_dotted_hierarchy) else 3
        if self._dotted_base_level == 1:
            self._log("文档以多级点分编号为主干大纲，点分编号按 1./1.1./1.1.1./1.1.1.1. → L1~L4 识别。")

    def _single_number_label(self, text):
        """段落开头的单级标号：返回 (序号, 是否带括号)，不是单级标号时返回 None。

        「1.」/「1．」/裸「1内容」→ (1, False)；「(1)」/「（1）」→ (1, True)。
        多级点分编号（1.1 等）与字母标号（A. / (a)）不参与本规则。
        """
        value = (text or '').lstrip()
        if not value:
            return None
        parenthesized = re.match(r'^[（\(](\d+)[）\)]', value)
        if parenthesized:
            return int(parenthesized.group(1)), True
        if RE_DOTTED_NUMBER.match(value):
            return None
        if not RE_HEADING_H3.match(value) and not (self._has_dotted_hierarchy and RE_SINGLE_LABEL.match(value)):
            return None
        digits = re.match(r'\d+', value)
        return (int(digits.group(0)), False) if digits else None

    def _sub_level_of_context(self, depth, restart_level):
        """当前编号上下文的下一级：有起头的单级编号时取其后一级，否则按点分深度推算。"""
        if restart_level is not None:
            return min(restart_level + 1, MAX_TITLE_OUTLINE_LEVEL)
        if depth >= 2:
            # 多级点分编号层级固定为深度本身，其下游单级标号取下一级：
            # x.x(L2)→L3、x.x.x(L3)→L4，不再叠加「1.」基准级别。
            return min(depth + 1, MAX_TITLE_OUTLINE_LEVEL)
        return None

    def _prepare_single_number_restart(self, doc):
        """多级点分编号下游的单级标号，按“当前层级的下一级”处理。

        文档里更深的点分编号（x.x / x.x.x）之后出现的单级标号都不是新的一级标题，
        而是该深层标题的子项：
        - 「1.」/裸「1内容」：默认按基准级别识别（纯点分文档即一级标题）；只要出现在
          更深的点分编号之后、且序号从 1 重新起头并依次递增（1、2、3…），就改判为
          该点分标题的下一级（x.x→L3、x.x.x→L4、x.x.x.x→L5…），与文档基准级别无关；
          避免把「1. 总则 … 2. 引用文件」这类顶层编号一起降级；
        - 「(1)」/「（1）」：默认固定为四级，与上游层级无关；这里同样按“当前上下文的
          下一级”处理（1.1 之后 → L3、1.1.1 之后 → L4…），避免子项比父项还浅。

        文档开头的 1、2、3（前面没有更深的点分标题）不受影响；出现新的点分编号或
        传统标题（一、/（一））时重新统计上下文。
        """
        self._single_number_levels = {}
        self._paren_number_levels = {}
        self._single_number_owners = []
        if not self._has_dotted_hierarchy:
            return
        # 「1.」的下一级规则只依赖“上游存在更深的点分编号”，与基准级别（公文三级 / 点分主干）无关
        depth = 0              # 当前上下文中最近的点分编号段数
        restart_level = None   # 当前“从 1 重新起头”序列的大纲级别
        previous_value = None  # 该序列的上一序号
        for para in doc.paragraphs:
            text = (para.text or '').lstrip()
            if not text:
                continue
            label = self._single_number_label(text)
            if label is not None:
                value, parenthesized = label
                if parenthesized:
                    level = self._sub_level_of_context(depth, restart_level)
                    if level is not None:
                        self._paren_number_levels[id(para._p)] = level
                        self._single_number_owners.append(para._p)
                    continue
                level = None
                if restart_level is not None and previous_value is not None and value > previous_value:
                    level = restart_level
                elif value == 1 and depth >= 2:
                    restart_level = min(depth + 1, MAX_TITLE_OUTLINE_LEVEL)
                    level = restart_level
                else:
                    restart_level, previous_value = None, None
                if level is not None:
                    self._single_number_levels[id(para._p)] = level
                    self._single_number_owners.append(para._p)
                    previous_value = value
                continue
            if RE_HEADING_H1.match(text) or RE_HEADING_H2.match(text):
                # 传统标题开启新的层级上下文
                depth, restart_level, previous_value = 0, None, None
                continue
            para_depth = self._dotted_number_depth(text)
            if para_depth >= 2 and self._looks_like_heading_text(text):
                depth, restart_level, previous_value = para_depth, None, None
        details = []
        for name, mapping in (("「1.」起头连续编号", self._single_number_levels),
                              ("「(1)」标号", self._paren_number_levels)):
            if mapping:
                levels = "、".join(f"Lv{level}" for level in sorted(set(mapping.values())))
                details.append(f"{name} {len(mapping)} 段 → {levels}")
        if details:
            self._log("多级编号下游的单级编号按上一级的下一级处理：" + "；".join(details) + "。")

    def _paragraph_dotted_level(self, para, text):
        """段落按编号推断的大纲级别：「1 开头连续编号」的下一级映射优先于点分编号深度。"""
        if para is not None:
            override = self._single_number_levels.get(id(para._p))
            if override is not None:
                return override
        return self._dotted_number_level(
            text, self._outline_offset, self._dotted_base_level, self._has_dotted_hierarchy
        )

    def _paren_number_level_of(self, para):
        """「(1)」标号在多级编号下游的下一级级别；不适用时返回 None。"""
        if para is None:
            return None
        return self._paren_number_levels.get(id(para._p))

    def _chapter_section_level(self, text):
        """识别章/节标题并返回绝对大纲级别（1/2）；不适用时返回 None。"""
        if not self._outline_offset:
            return None
        value = (text or '').strip()
        if not value or len(value) > HEADING_TEXT_MAX_LENGTH:
            return None
        if any(punct in value for punct in BODY_SENTENCE_PUNCTUATION):
            return None
        if RE_CHAPTER_TITLE.match(value):
            return 1
        if RE_SECTION_TITLE.match(value):
            return 2 if self._outline_offset >= 2 else 1
        return None

    # ------------------------------------------------------------------
    # 条标题：第X条 / 第X.X条
    # ------------------------------------------------------------------
    @staticmethod
    def _article_number_depth(text):
        """「第X条」编号段数：第1条→1、第1.1条→2、第1.1.1条→3；不匹配时返回 0。"""
        match = RE_ARTICLE_TITLE.match((text or '').strip())
        if not match:
            return 0
        number = match.group(0)
        return number.count('.') + number.count('．') + 1

    def _article_title_level(self, text):
        """条标题的绝对大纲级别：章/节之下顺延，点分编号每多一段再深一级。"""
        if not self.config.get('detect_article_title', True):
            return None
        value = (text or '').strip()
        if not value or len(value) > HEADING_TEXT_MAX_LENGTH:
            return None
        if any(punct in value for punct in BODY_SENTENCE_PUNCTUATION):
            return None
        depth = self._article_number_depth(value)
        if not depth:
            return None
        return min(self._article_level + depth - 1, MAX_TITLE_OUTLINE_LEVEL)

    def _split_article_paragraphs(self, doc):
        """把「第X条 正文……」拆成两段：条标题独立成段，其余文字另起一段作正文。"""
        if not self.config.get('detect_article_title', True):
            return 0
        split_count = 0
        for para in list(doc.paragraphs):
            text = para.text
            stripped = text.lstrip()
            if not stripped:
                continue
            match = RE_ARTICLE_TITLE.match(stripped)
            if match is None:
                continue
            if self._has_drawing_or_pict(para) or self._has_embedded_object(para):
                continue
            cut = len(text) - len(stripped) + match.end()
            if not text[cut:].strip():
                continue
            new_para = self._split_paragraph_at(para, cut)
            if new_para is not None:
                self._strip_leading_whitespace(new_para)
                split_count += 1
        if split_count:
            self._log(f"条标题识别：{split_count} 个「第X条」段落已按条拆分（条后正文另起一段）。")
        return split_count

    def _split_paragraph_at(self, para, cut):
        """在第 cut 个字符处把段落一分为二，返回放在原段落之后的新段落。"""
        if cut <= 0 or cut >= len(para.text):
            return None
        new_p = OxmlElement('w:p')
        para._p.addnext(new_p)
        new_para = Paragraph(new_p, para._parent)
        moved, position = [], 0
        for run in list(para.runs):
            run_text = run.text
            start, end = position, position + len(run_text)
            position = end
            if end <= cut:
                continue
            if start >= cut:
                moved.append(run._r)
                continue
            keep = run_text[:cut - start]
            tail = run_text[cut - start:]
            run.text = keep
            tail_run = new_para.add_run(tail)
            r_pr = run._r.find(qn('w:rPr'))
            if r_pr is not None:
                tail_run._r.insert(0, copy.deepcopy(r_pr))
        for element in moved:
            new_p.append(element)
        return new_para

    def _paragraph_heading_rank(self, text_to_check, para=None):
        """按编号/章节特征返回段落的大纲级别，非标题返回 0。

        与逐段排版时 H1/H2/H3/H4/章/节 的识别顺序保持一致，仅用于判断
        “末级编号标题其后是否还有正文”，不影响任何格式设置。
        para 提供时，会采用 _prepare_single_number_restart 预扫描的下一级映射。
        """
        value = (text_to_check or '').lstrip()
        chapter = self._chapter_section_level(value)
        if chapter is not None:
            return chapter
        article = self._article_title_level(value)
        if article is not None:
            return article
        if RE_HEADING_H1.match(value):
            return 1 + self._outline_offset
        if RE_HEADING_H2.match(value):
            return 2 + self._outline_offset
        dotted = self._paragraph_dotted_level(para, value)
        if dotted is not None:
            return dotted
        if RE_HEADING_H3.match(value):
            return 3 + self._outline_offset
        if RE_HEADING_H4.match(value) or RE_HEADING_H4_ALPHA.match(value):
            return self._paren_number_level_of(para) or (4 + self._outline_offset)
        return 0

    def _is_demoted_deepest_heading(self, block_idx):
        """局部末级标题且局部范围内无正文时返回 True，该段落应降级为正文。

        依赖 format_document 预计算的 self._demoted_body。
        """
        return self._demoted_body[block_idx]

    def _format_demoted_deepest_as_body(self, para, protected_paragraphs, current_block_num, preview):
        """把被判定为“末级且无后续正文”的标题段落按正文处理。

        写入直接格式 w:outlineLvl=9（正文级）强制覆盖样式自带的大纲级别；
        标题等级归并（normalize_heading_levels）在主循环之前运行，可能已给该段
        设过大纲级别，这里同样覆盖为正文。加入 protected_paragraphs 可避免
        随后的跳级补齐再次处理该段。
        """
        self._log(f"段落 {current_block_num}: 末级标题且其后无正文，按正文处理 - \"{preview}...\"")
        self._strip_leading_whitespace(para)
        ppr = para._p.get_or_add_pPr()
        outline = ppr.find(qn('w:outlineLvl'))
        if outline is None:
            outline = OxmlElement('w:outlineLvl')
            ppr.append(outline)
        outline.set(qn('w:val'), '9')  # 9 = 正文级；直接格式覆盖样式
        self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=self._apply_color)
        self._apply_text_indent_and_align(para)
        self._reset_pagination_properties(para)
        protected_paragraphs.append(para._p)

    def _heading_font_for_level(self, level):
        """按最终大纲级别选择标题字体：1 级用一级标题字体、2 级用二级、其余按正文。"""
        if level <= 1:
            return (
                self.config['h1_font'],
                self.config['h1_size'],
                self.config.get('h1_bold', False),
            )
        if level == 2:
            return (
                self.config['h2_font'],
                self.config['h2_size'],
                self.config.get('h2_bold', False),
            )
        return self.config['body_font'], self.config['body_size'], False

    # ------------------------------------------------------------------
    # 可选增强五：前后行词语连续合并（PDF 粘贴还原）
    # ------------------------------------------------------------------
    def _is_heading_like_text(self, text):
        """判断段落文字是否形似标题/附件标识，词语连续合并时必须跳过它们。"""
        value = (text or '').strip()
        if not value:
            return False
        if self.config.get('detect_article_title', True) and RE_ARTICLE_TITLE.match(value):
            return True
        return bool(
            RE_HEADING_H1.match(value)
            or RE_HEADING_H2.match(value)
            or RE_HEADING_H3.match(value)
            or RE_HEADING_H4.match(value)
            or RE_CHAPTER_TITLE.match(value)
            or RE_SECTION_TITLE.match(value)
            or RE_ATTACHMENT.match(value)
        )

    @staticmethod
    def _paragraph_is_plain_runs(para):
        """段落内只有普通的 run（无超链接/域/内容控件），合并时才不会丢内容。"""
        plain_tags = {
            qn('w:pPr'), qn('w:r'), qn('w:bookmarkStart'), qn('w:bookmarkEnd'), qn('w:proofErr'),
        }
        return all(child.tag in plain_tags for child in para._p)

    @staticmethod
    def _sentence_ends_at(text):
        """取出去掉收尾引号/括号后的末字，判断该行是否以句号类标点收尾。"""
        value = (text or '').strip().rstrip(MERGE_SENTENCE_END_CLOSERS).rstrip()
        return bool(value and value[-1] in MERGE_SENTENCE_END_PUNCTUATION)

    @staticmethod
    def _jieba_word_crosses_boundary(prev_text, next_text):
        """判断两行拼接处是否恰好切断一个词（如“高压九级接”+“替供气”的“接替”）。

        把前行末若干字与后行首若干字拼成探测串交给 jieba 分词，若存在某个词
        的区间横跨拼接点，说明 PDF 在词中间断行，应还原为整段。
        """
        if jieba is None:
            return False
        tail = prev_text[-MERGE_JIEBA_PROBE_SIZE:]
        head = next_text[:MERGE_JIEBA_PROBE_SIZE]
        boundary = len(tail)
        try:
            for word, start, end in jieba.tokenize(tail + head):
                if start < boundary < end and len(word.strip()) >= 2:
                    return True
        except Exception:
            return False
        return False

    @classmethod
    def _number_token(cls, text):
        """取列表项的编号类型与序号；无编号返回 None。"""
        value = RE_MD_BULLET_PREFIX.sub('', text or '', count=1)
        match = RE_LIST_NUMBER_TOKEN.match(value)
        if not match:
            return None
        for kind, part in enumerate(match.groups()):
            if part is None:
                continue
            if kind == 0 and part.isalpha():
                return (kind, ord(part.lower()))
            return (kind, part)
        return None

    @staticmethod
    def _number_value(token):
        if not token:
            return None
        _, raw = token
        if isinstance(raw, int) or (isinstance(raw, str) and raw.isascii() and raw.isdigit()):
            return int(raw)
        small = {'①': 1, '②': 2, '③': 3, '④': 4, '⑤': 5,
                 '⑥': 6, '⑦': 7, '⑧': 8, '⑨': 9, '⑩': 10,
                 '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                 '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}
        return small.get(raw)

    @classmethod
    def _is_next_number(cls, prev, curr):
        if not prev or not curr or prev[0] != curr[0]:
            return False
        left, right = cls._number_value(prev), cls._number_value(curr)
        return left is not None and right is not None and right == left + 1

    @classmethod
    def _numbered_list_gap_indices(cls, texts):
        """连续编号列表中唯一缺少段尾标点的项：下一行就是它被截断的续文。

        同组相邻编号样式一致、序号连续，且前后两项都以句末标点收尾；当前项
        却没有段尾标点，而紧跟的非编号行有段尾标点。这正是「其它编号有段尾
        标点、唯独本项没有、下一行有」的中断特征，无需分词或语言模型。
        """
        values = [text.strip() for text in texts]
        items = [
            (index, cls._number_token(text), text)
            for index, text in enumerate(values)
            if RE_MD_LIST_PREFIX.match(text) and cls._number_token(text)
        ]
        result = set()
        for position in range(1, len(items) - 1):
            prev_index, prev_number, prev_text = items[position - 1]
            index, number, text = items[position]
            next_index, next_number, next_text = items[position + 1]
            if not (cls._is_next_number(prev_number, number)
                    and cls._is_next_number(number, next_number)):
                continue
            if (not cls._sentence_ends_at(prev_text)
                    or cls._sentence_ends_at(text)
                    or not cls._sentence_ends_at(next_text)):
                continue
            tail_index = index + 1
            if tail_index >= len(values) or tail_index == next_index:
                continue
            tail = values[tail_index]
            if not tail or RE_MD_LIST_PREFIX.match(tail) or not cls._sentence_ends_at(tail):
                continue
            result.add(index)
        return result

    def _is_continuous_pair(self, prev_text, next_text, prev_is_heading=False,
                            prev_is_list_item=False, prev_numbered_gap=False):
        """判断相邻两行是否应连续合并：先排除句末断点，再看续句标点/跨行词。

        prev_is_heading 表示前行形似标题。短标题后面跟正文是正常结构，不能合并；
        达到 MERGE_HEADING_MIN_LENGTH 的长标题型行更可能是 PDF 断行的前半段，
        因此只有这类行才允许用跨行词证据与下一行合并。

        prev_is_list_item 用于列表项尾部被 PDF 截断的场景。此时下一行若是很短
        且以句末标点收尾的词组（「维护工作；」「使用标准；」），即使 jieba 没
        发现跨行词也判定为续文；短标题式的列表项（如「1) 拆卸准备」）不会命中，
        因为续行被限定为短词组且前行须带冒号或足够长。

        prev_numbered_gap 表示前行是连续编号列表中唯一没有段尾标点的一项，
        这是最直接的截断证据，优先于标题长度等弱判据。
        """
        if self._sentence_ends_at(prev_text):
            return False
        if prev_numbered_gap:
            return True
        last_char = prev_text.strip().rstrip(MERGE_SENTENCE_END_CLOSERS).rstrip()[-1:]
        if last_char and last_char in MERGE_CONNECTING_PUNCTUATION:
            return True
        if prev_is_heading and len(prev_text.strip()) < MERGE_HEADING_MIN_LENGTH:
            return False
        body = prev_text.strip()
        tail = next_text.strip()
        if (not prev_is_list_item and body and tail
                and not RE_NOTE_LINE.match(tail)
                and (body[-1] in MERGE_FUNCTION_WORDS
                     or tail[0] in MERGE_FUNCTION_WORDS)):
            return True  # 正文虚词结尾/开头：PDF 断行把同一句拆成了两段
        if (prev_is_list_item
                and tail and len(tail) <= MERGE_LIST_TAIL_MAX_LENGTH
                and self._sentence_ends_at(tail)
                and not RE_NOTE_LINE.match(tail)
                and not self._is_heading_like_text(tail)
                and ('：' in prev_text or ':' in prev_text
                     or len(prev_text.strip()) >= MERGE_HEADING_MIN_LENGTH * 2)):
            return True
        return self._jieba_word_crosses_boundary(prev_text, next_text)

    def _merge_continuous_lines(self, doc):
        """把前后行词语连续的相邻段落合并成整段，还原 PDF 粘贴丢失的段落流畅性。

        逐对判断相邻行：前行以句号类标点收尾即为自然段边界；否则若前行以
        逗号类标点收尾、或 jieba 发现拼接处切断了词，则两行连续，合并。
        """
        raw_texts = [para.text.strip() for para in doc.paragraphs]
        numbered_gaps = self._numbered_list_gap_indices(raw_texts)
        entries = []
        for index, para in enumerate(doc.paragraphs):
            text = raw_texts[index]
            # 含图像的段落本身不参与文本合并，但也不再作为隔断：PDF 粘贴常把
            # 图像插在一句话中间，后续文字仍需与前半句合并；图像段落原样保留。
            if self._has_drawing_or_pict(para) or self._has_embedded_object(para):
                entries.append((para, text, False, True))
                continue
            if not text:
                entries.append(None)
                continue
            if (
                not self._paragraph_is_plain_runs(para)
                or self._get_paragraph_alignment(para)
                in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT)
            ):
                entries.append(None)
                continue
            # 标题型行不直接排除：长行可能是 PDF 断行的前半段，先把断行还原，
            # 之后标题识别才不会把续行误判为“标题的专属下文”。短标题仍被挡住。
            heading_like = self._is_heading_like_text(text)
            if heading_like and len(text) < MERGE_HEADING_MIN_LENGTH and index not in numbered_gaps:
                entries.append(None)
                continue
            entries.append((para, text, heading_like, False))

        paragraphs = [entry[0] if entry is not None else None for entry in entries]
        merged_paragraphs = merged_lines = 0
        index, total = 0, len(entries)
        while index < total:
            if entries[index] is None or entries[index][3]:
                index += 1
                continue

            end = index
            skip = set()  # 图像段落：跳过其文字判定，但保留在文档中
            scan = index + 1
            while scan < total:
                candidate = entries[scan]
                if candidate is None:
                    break
                if candidate[3]:
                    skip.add(scan)
                    scan += 1
                    continue
                if not self._is_continuous_pair(
                    entries[end][1], candidate[1],
                    prev_is_heading=entries[end][2],
                    prev_is_list_item=bool(RE_MD_LIST_PREFIX.match(entries[end][1])),
                    prev_numbered_gap=end in numbered_gaps,
                ):
                    break
                end = scan
                scan += 1

            if end > index:
                preview = entries[index][1][:30].replace("\n", " ")
                self._merge_paragraph_range(paragraphs, index, end, skip)
                merged_paragraphs += 1
                merged_lines += end - index - len(skip)
                self._log(
                    f"  > 合并 {end - index + 1 - len(skip)} 行为一段（前后行词语连续）: \"{preview}...\""
                )
            index = end + 1

        self._log(
            f"前后行词语连续合并完成：合并 {merged_paragraphs} 段、"
            f"消除 {merged_lines} 处断行。"
        )
        return merged_lines

    @staticmethod
    def _merge_paragraph_range(all_paragraphs, start, end, skip=()):
        """把 [start, end] 区间的段落内容合并进第一段，并删除其余段落。

        skip 中的下标对应的段落原样保留（图像段落），不会被并入也不会被删除。
        """
        if end <= start:
            return
        first = all_paragraphs[start]
        first_p = first._p
        for offset, para in enumerate(all_paragraphs[start + 1:end + 1], start=start + 1):
            if para is None or offset in skip:
                continue
            runs = list(para.runs)
            if first.runs:
                first.runs[-1].text = first.runs[-1].text.rstrip()
            if runs:
                runs[0].text = runs[0].text.lstrip()
            prev_text = ''.join(run.text for run in first.runs)
            next_text = ''.join(run.text for run in runs)
            if (
                prev_text and next_text
                and prev_text[-1].isascii() and prev_text[-1].isalnum()
                and next_text[0].isascii() and next_text[0].isalnum()
            ):
                first.add_run(' ')
            for run in runs:
                first_p.append(run._r)
            para._p.getparent().remove(para._p)

    @staticmethod
    def _pick_heading_group_level(counts):
        """从 {(证据等级, 大纲级别): 次数} 计数中选出组内最具代表性的级别。"""
        if not counts:
            return None
        (_, level), _ = sorted(
            counts.items(), key=lambda item: (-item[0][0], -item[1], item[0][1])
        )[0]
        return level

    @staticmethod
    def _plan_heading_levels(records, on_skip=None):
        """
        相同格式标题等级合并。每个 record 需包含：
        index / signature / level / evidence / joinable / numbered_level。
        证据优先级：3 有标题样式、2 有大纲级别、1 只有编号特征。
        自身带编号且与合并等级不一致时保留编号等级（编号优先）。

        无证据的段落只在同时满足以下两点时才跟着补齐，避免正文被整段提成标题：
        1) 同格式组内存在强证据（标题样式或已有大纲级别）；
           仅凭“一、”这类文本编号推断出的证据不足以扩散到其它段落；
        2) 有证据的段落不少于无证据的段落，即该格式没有被正文大量占用。
        返回 {record_index: 目标等级}，只包含需要改动的记录。
        """
        groups, order = {}, []
        for record in records:
            if not record.get('joinable') or record.get('signature') is None:
                continue
            key = record['signature']
            group = groups.get(key)
            if group is None:
                group = groups[key] = {
                    'members': [], 'strong': {}, 'weak': {}, 'plain': 0,
                }
                order.append(key)
            group['members'].append(record)
            evidence = record.get('evidence', 0)
            if evidence > 0 and record.get('level'):
                bucket = group['strong'] if evidence >= 2 else group['weak']
                token = (evidence, record['level'])
                bucket[token] = bucket.get(token, 0) + 1
            else:
                group['plain'] += 1

        plan = {}
        for key in order:
            group = groups[key]
            strong_level = WordProcessor._pick_heading_group_level(group['strong'])
            weak_level = WordProcessor._pick_heading_group_level(group['weak'])
            evidenced = len(group['members']) - group['plain']
            fill_allowed = strong_level is not None and group['plain'] <= evidenced
            fill_level = strong_level if fill_allowed else None

            if group['plain'] and not fill_allowed and on_skip is not None:
                reason = (
                    '同类段落只有编号特征、缺少标题样式或大纲级别'
                    if strong_level is None
                    else '同类段落多为无标题证据的正文'
                )
                on_skip(reason, group['plain'])

            for record in group['members']:
                numbered_level = record.get('numbered_level')
                if record.get('evidence', 0) > 0:
                    level = strong_level or weak_level
                    if level is None:
                        continue
                    if numbered_level and numbered_level != level:
                        continue
                else:
                    level = fill_level
                    if level is None:
                        continue
                level = WordProcessor._apply_dotted_level_floor(record.get('text'), level)
                if record.get('level') != level:
                    plan[record['index']] = level
        return plan

    @staticmethod
    def _plan_level_compaction(levels):
        """补齐中间缺失的大纲级别（只补缺口，不整体上移）。返回 {旧级别: 新级别}。"""
        mapping, expected = {}, None
        for level in sorted(set(levels)):
            if not 1 <= level <= MAX_TITLE_OUTLINE_LEVEL:
                continue
            mapping[level] = level if expected is None else expected
            expected = mapping[level] + 1
        return {old: new for old, new in mapping.items() if old != new}

    def _effective_outline_level(self, para):
        """读取段落当前生效的大纲级别（1-8）：先看直接设置，再看样式自带。"""
        raw_level = self._get_outline_level(para)
        if raw_level is not None:
            if raw_level >= 9:  # 9 = 正文级；显式覆盖样式自带的大纲级别
                return None
            level = raw_level + 1  # w:outlineLvl 以 0 表示一级
            if 1 <= level <= MAX_TITLE_OUTLINE_LEVEL:
                return level
        _, style_level = self._get_heading_style_info(para)
        return style_level

    def _build_heading_records(self, doc, exclude_elements=None):
        """收集参与同格式归并的段落记录，index 为记录在列表中的位置。"""
        excluded = set(exclude_elements or ())
        records = []
        for para in doc.paragraphs:
            if para._p in excluded:
                continue
            text = para.text.strip()
            if not text:
                continue
            if self._has_drawing_or_pict(para) or self._has_embedded_object(para):
                continue
            if para._p in self._non_heading_paras:
                continue  # 单独的 (1)/(a) 标号按正文处理，不参与标题等级归并
            signature = self._paragraph_format_signature(para)
            if signature is None:
                continue

            has_heading_style, style_level = self._get_heading_style_info(para)
            # 「1.」起头连续编号 /「(1)」标号位于更深点分编号下游时，按上一级的下一级识别
            numbered = self._single_number_levels.get(id(para._p)) or self._paren_number_level_of(para)
            if numbered is None:
                numbered = self._numbered_heading_level(text, self._outline_offset, self._dotted_base_level, self._has_dotted_hierarchy)
            numbered = numbered or self._chapter_section_level(text)
            outline_level = self._effective_outline_level(para)
            # 编号/层级识别优先于段落既有样式或大纲级别（格式），避免格式覆盖编号识别；
            # 多级点分编号（x.x）按编号深度给级别下限（1.1 → 至少 Lv2），自带级别更浅时上调
            level = self._apply_dotted_level_floor(text, numbered or outline_level)

            if has_heading_style and level:
                evidence = 3
            elif outline_level:
                evidence = 2
            elif numbered:
                evidence = 1
            else:
                evidence = 0

            records.append({
                'index': len(records),
                'paragraph': para,
                'text': text,
                'signature': signature,
                'level': level,
                'numbered_level': numbered,
                'evidence': evidence,
                'joinable': evidence > 0 or self._looks_like_heading_text(text),
            })
        return records

    def _normalize_heading_levels(self, doc, exclude_elements=None):
        """相同格式的段落并入同一标题等级，需要同时启用 set_outline。"""
        records = self._build_heading_records(doc, exclude_elements)

        def _report_skip(reason, count):
            self._log(f"  > 跳过 {count} 个同格式段落：{reason}，按正文保留。")

        plan = self._plan_heading_levels(records, on_skip=_report_skip)
        if not plan:
            self._log("标题等级合并优化：未发现需要合并的相同格式标题。")
            return 0

        assigned = unified = 0
        for position, level in sorted(plan.items()):
            record = records[position]
            old_level = self._effective_outline_level(record['paragraph'])
            self._set_outline_level(record['paragraph'], level)
            preview = record['text'][:30].replace("\n", " ")
            if old_level is None:
                assigned += 1
                self._log(f"  > 相同格式补齐大纲级别 Lv{level}: \"{preview}...\"")
            else:
                unified += 1
                self._log(f"  > 相同格式统一大纲级别 Lv{old_level} → Lv{level}: \"{preview}...\"")
        self._log(f"标题等级合并优化完成：补齐 {assigned} 处，统一 {unified} 处。")
        return len(plan)

    def _compact_outline_levels(self, doc, exclude_elements=None):
        """检查是否缺少中间的样式等级，有跳级时向上补齐。"""
        excluded = set(exclude_elements or ())
        targets = []
        for para in doc.paragraphs:
            if para._p in excluded:
                continue
            level = self._effective_outline_level(para)
            if level is None:
                continue
            targets.append((para, level))

        mapping = self._plan_level_compaction([level for _, level in targets])
        if not mapping:
            self._log("标题等级跳级检查：层级连续，无需调整。")
            return 0

        changed = 0
        for para, level in targets:
            new_level = mapping.get(level)
            if new_level is None or new_level == level:
                continue
            # 多级点分编号（x.x）补齐跳级后仍不得浅于编号深度（1.1 → 至少 Lv2）
            new_level = self._apply_dotted_level_floor(para.text, new_level)
            if new_level == level:
                continue
            self._set_outline_level(para, new_level)
            changed += 1
            preview = para.text.strip()[:30].replace("\n", " ")
            self._log(f"  > 补齐跳级 Lv{level} → Lv{new_level}: \"{preview}...\"")
        detail = "，".join(f"Lv{old}→Lv{new}" for old, new in sorted(mapping.items()))
        self._log(f"标题等级跳级检查完成：{detail}，共调整 {changed} 处。")
        return changed

    # ------------------------------------------------------------------
    # 标准样式统一：正文 / 标题 1~4
    # ------------------------------------------------------------------
    @staticmethod
    def _style_level_key(level):
        """把大纲级别折算成样式等级：0 表示正文，1~4 表示标题 1~4。"""
        if not level:
            return 0
        return min(int(level), MAX_STYLE_HEADING_LEVEL)

    def _record_paragraph_level(self, para, level=None, strip=True):
        """登记段落的最终等级，排版结束后统一套用内置样式。

        strip=False 表示该段的字符格式不是单一等级（如二级标题与正文同段拆分），
        此时只套样式、不清除直接字体字号，避免正文部分跟着标题样式变形。
        """
        self._level_records[id(para._p)] = (para, level, strip)

    def _get_or_add_paragraph_style(self, doc, name, aliases=()):
        """取内置段落样式（正文 / 标题 N），文档里没有时按内置样式补建。"""
        wanted = {name.lower()}
        wanted.update(alias.lower() for alias in aliases)
        for style in doc.styles:
            if (getattr(style, 'name', '') or '').strip().lower() in wanted:
                return style
        styles_element = doc.styles.element
        element = styles_element.get_by_id(name.replace(' ', ''))
        if element is not None and element.type == WD_STYLE_TYPE.PARAGRAPH:
            return StyleFactory(element)
        return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)

    def _get_level_style(self, doc, level):
        key = self._style_level_key(level)
        name = BUILTIN_BODY_STYLE_NAME if key == 0 else BUILTIN_HEADING_STYLE_NAMES[key]
        aliases = BUILTIN_BODY_STYLE_ALIASES if key == 0 else BUILTIN_HEADING_STYLE_ALIASES[key]
        return self._get_or_add_paragraph_style(doc, name, aliases)

    def _write_style_font(self, style, font_name, size_pt, bold):
        """把字体、字号、加粗写进样式定义（覆盖内置样式自带的主题字体/主题色）。"""
        element = style.element
        r_pr = element.get_or_add_rPr()
        r_fonts = r_pr.get_or_add_rFonts()
        for attribute in STYLE_THEME_FONT_ATTRIBUTES:
            r_fonts.attrib.pop(qn(attribute), None)
        en_font = self._run_font_family(font_name)
        r_fonts.set(qn('w:eastAsia'), font_name)
        r_fonts.set(qn('w:ascii'), en_font)
        r_fonts.set(qn('w:hAnsi'), en_font)
        r_fonts.attrib.pop(qn('w:cs'), None)

        style.font.size = Pt(size_pt)
        size_value = str(int(round(float(size_pt) * 2)))
        sz_node = r_pr.find(qn('w:sz'))
        if sz_node is not None:
            sz_cs = r_pr.find(qn('w:szCs'))
            if sz_cs is None:
                sz_cs = OxmlElement('w:szCs')
                sz_node.addnext(sz_cs)
            sz_cs.set(qn('w:val'), size_value)
        style.font.bold = bool(bold)
        style.font.italic = False
        # 内置标题样式自带“复杂文种加粗/斜体”，一并去掉，避免同一段落里字重不一致
        for tag in ('w:bCs', 'w:iCs'):
            node = r_pr.find(qn(tag))
            if node is not None:
                r_pr.remove(node)
        style.font.color.rgb = RGBColor(0, 0, 0)
        color = r_pr.find(qn('w:color'))
        if color is not None:
            for attribute in STYLE_THEME_COLOR_ATTRIBUTES:
                color.attrib.pop(qn(attribute), None)

    def _define_level_style(self, style, level):
        """按配置写全某一级的样式定义：字体字号加粗、行距、大纲级别。"""
        if level:
            font, size, bold = self._heading_font_for_level(level)
        else:
            font, size, bold = self.config['body_font'], self.config['body_size'], False
        self._write_style_font(style, font, size, bold)

        p_pr = style.element.get_or_add_pPr()
        # 内置标题样式自带“与下段同页/段中不分页”，排版主流程会逐段关掉，这里保持一致
        for tag in ('w:keepNext', 'w:keepLines'):
            node = p_pr.find(qn(tag))
            if node is not None:
                p_pr.remove(node)
        if level:
            self._apply_line_spacing_to_pPr(
                p_pr, 'line_spacing', 'line_spacing_unit', 'line_spacing_multiple', 28
            )
            # 标题段落一律段前段后 0，样式定义与逐段排版写入的直接格式保持一致
            spacing = p_pr.get_or_add_spacing()
            for tag in ('w:before', 'w:after', 'w:beforeLines', 'w:afterLines'):
                spacing.set(qn(tag), '0')
            spacing.set(qn('w:beforeAutospacing'), '0')
            spacing.set(qn('w:afterAutospacing'), '0')
        outline = p_pr.find(qn('w:outlineLvl'))
        if not level or not self.config.get('set_outline', True):
            if outline is not None:
                p_pr.remove(outline)
            return
        if outline is None:
            outline = OxmlElement('w:outlineLvl')
            anchor = None
            for tag in ('w:rPr', 'w:sectPr', 'w:pPrChange'):
                anchor = p_pr.find(qn(tag))
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(outline)
            else:
                p_pr.append(outline)
        outline.set(qn('w:val'), str(level - 1))

    @staticmethod
    def _assign_paragraph_style(para, style):
        """把段落改成指定样式（只写 pStyle，不动其它直接格式）。"""
        try:
            current = para.style
        except Exception:  # noqa: BLE001 - 损坏的 pStyle 指向，按未设置处理
            current = None
        if current is not None and getattr(current, 'style_id', None) == style.style_id:
            return
        para.style = style

    @staticmethod
    def _strip_direct_run_fonts(para):
        """清掉段落里由样式统一提供的直接字体/字号，保留加粗、斜体、下划线、颜色。"""
        for run in para.runs:
            r_pr = run._r.find(qn('w:rPr'))
            if r_pr is None:
                continue
            for tag in STYLE_DIRECT_RUN_TAGS:
                child = r_pr.find(qn(tag))
                if child is not None:
                    r_pr.remove(child)

    @classmethod
    def _style_font_size_value(cls, style):
        """读取样式（含 basedOn 链）上的 w:sz（半磅字符串），取不到返回 None。"""
        seen = 0
        while style is not None and seen < 6:
            element = getattr(style, 'element', None)
            r_pr = element.find(qn('w:rPr')) if element is not None else None
            if r_pr is not None:
                size = r_pr.find(qn('w:sz'))
                if size is not None and size.get(qn('w:val')):
                    return size.get(qn('w:val'))
            style = getattr(style, 'base_style', None)
            seen += 1
        return None

    @classmethod
    def _style_font_attributes(cls, style):
        """读取样式（含 basedOn 链）上的 rFonts 属性（含主题字体），取不到返回 None。"""
        seen = 0
        while style is not None and seen < 6:
            element = getattr(style, 'element', None)
            r_pr = element.find(qn('w:rPr')) if element is not None else None
            if r_pr is not None:
                r_fonts = r_pr.find(qn('w:rFonts'))
                if r_fonts is not None and len(r_fonts.attrib):
                    return dict(r_fonts.attrib)
            style = getattr(style, 'base_style', None)
            seen += 1
        return None

    @staticmethod
    def _default_font_size_value(doc):
        """读取 styles.xml 中 docDefaults 的默认字号（半磅字符串）。"""
        r_pr = WordProcessor._default_run_properties(doc)
        if r_pr is None:
            return None
        size = r_pr.find(qn('w:sz'))
        return size.get(qn('w:val')) if size is not None else None

    @staticmethod
    def _default_font_attributes(doc):
        """读取 styles.xml 中 docDefaults 的默认字体属性（可能是主题字体）。"""
        r_pr = WordProcessor._default_run_properties(doc)
        if r_pr is None:
            return None
        r_fonts = r_pr.find(qn('w:rFonts'))
        if r_fonts is None or not len(r_fonts.attrib):
            return None
        return dict(r_fonts.attrib)

    @staticmethod
    def _default_run_properties(doc):
        """取出 docDefaults 里的 rPr 元素。"""
        doc_defaults = doc.styles.element.find(qn('w:docDefaults'))
        if doc_defaults is None:
            return None
        r_pr_default = doc_defaults.find(qn('w:rPrDefault'))
        if r_pr_default is None:
            return None
        return r_pr_default.find(qn('w:rPr'))

    @staticmethod
    def _set_paragraph_mark_size(para, size_value):
        """给段落标记写直接字号，用于保持空行（空段落）高度不变。"""
        p_pr = para._p.get_or_add_pPr()
        mark_r_pr = p_pr.find(qn('w:rPr'))
        if mark_r_pr is None:
            mark_r_pr = OxmlElement('w:rPr')
            anchor = None
            for tag in ('w:sectPr', 'w:pPrChange'):
                anchor = p_pr.find(qn(tag))
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(mark_r_pr)
            else:
                p_pr.append(mark_r_pr)
        WordProcessor._write_size_nodes(mark_r_pr, size_value)

    @staticmethod
    def _write_size_nodes(r_pr, size_value):
        """把字号写进 rPr（w:sz / w:szCs），按 schema 顺序插入。"""
        if hasattr(r_pr, 'get_or_add_sz'):
            size_node = r_pr.get_or_add_sz()
        else:
            size_node = r_pr.find(qn('w:sz'))
            if size_node is None:
                size_node = OxmlElement('w:sz')
                r_pr.append(size_node)
        size_node.set(qn('w:val'), str(size_value))
        sz_cs = r_pr.find(qn('w:szCs'))
        if sz_cs is None:
            sz_cs = OxmlElement('w:szCs')
            size_node.addnext(sz_cs)
        sz_cs.set(qn('w:val'), str(size_value))

    def _freeze_skipped_paragraph_sizes(self, doc, recorded_ids):
        """改写「正文」样式前，把未参与排版的段落的生效字体字号固化成直接格式。

        表格、页眉页脚、空行等段落不参与逐段排版，只会继承「正文」样式。不固化的话，
        改完样式后这些内容会跟着「正文」一起变（例如表格文字突然变成正文字号），
        而工具对表格的承诺是“不启用表格自动调整就不动表格”。
        """
        new_size = str(int(round(float(self.config['body_size']) * 2)))
        new_font = self.config['body_font']
        for para in self._iter_untouched_paragraphs(doc):
            if id(para._p) in recorded_ids:
                continue
            try:
                current_style = para.style
            except Exception:  # noqa: BLE001 - 损坏的 pStyle 指向，按默认样式处理
                current_style = None
            old_size = self._style_font_size_value(current_style) or self._default_font_size_value(doc)
            old_fonts = (
                self._style_font_attributes(current_style)
                or self._default_font_attributes(doc)
            )
            size_changed = bool(old_size) and old_size != new_size
            font_changed = bool(old_fonts) and self._fonts_would_change(old_fonts, new_font)
            if not size_changed and not font_changed:
                continue
            if not para.text.strip():
                if size_changed:
                    self._set_paragraph_mark_size(para, old_size)
                continue
            for run in para.runs:
                r_pr = run._r.get_or_add_rPr()
                r_fonts = r_pr.find(qn('w:rFonts'))
                if font_changed and (
                    r_fonts is None
                    or not (r_fonts.get(qn('w:eastAsia')) or r_fonts.get(qn('w:eastAsiaTheme')))
                ):
                    if r_fonts is None:
                        r_fonts = r_pr.get_or_add_rFonts()
                    for attribute, value in old_fonts.items():
                        r_fonts.set(attribute, value)
                if size_changed and r_pr.find(qn('w:sz')) is None:
                    self._set_run_direct_size(r_pr, old_size)

    @staticmethod
    def _fonts_would_change(old_fonts, new_font):
        """未排版段落的原字体是否会因为改写「正文」样式而变化。"""
        themed = any(
            attribute in old_fonts
            for attribute in (qn('w:eastAsiaTheme'), qn('w:asciiTheme'), qn('w:hAnsiTheme'))
        )
        if themed:
            return True
        current = old_fonts.get(qn('w:eastAsia')) or old_fonts.get(qn('w:ascii'))
        return bool(current) and current != new_font

    def _iter_untouched_paragraphs(self, doc):
        """遍历所有不参与排版的段落：表格单元格（含嵌套）、页眉页脚。"""
        for table in doc.tables:
            yield from self._iter_table_paragraphs(table)
        for section in doc.sections:
            for container in (
                section.header, section.footer,
                section.first_page_header, section.first_page_footer,
                section.even_page_header, section.even_page_footer,
            ):
                # 只看已定义的页眉页脚：访问未定义的会顺手创建页眉/页脚部件
                if container.is_linked_to_previous:
                    continue
                for para in container.paragraphs:
                    yield para
        for para in doc.paragraphs:
            if not para.text.strip():
                yield para

    def _iter_table_paragraphs(self, table):
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para
                for nested in cell.tables:
                    yield from self._iter_table_paragraphs(nested)

    @staticmethod
    def _set_run_direct_size(r_pr, size_value):
        """给 run 写直接字号（w:sz / w:szCs），保持其原有大小不随样式变化。"""
        WordProcessor._write_size_nodes(r_pr, size_value)

    def _apply_level_styles(self, doc):
        """把「正文 / 标题 1~4」样式定义写对，并让同级段落统一套用同一套样式。

        等级以排版结束时的最终结果为准（含等级合并、跳级补齐、末级降级），
        字体、字号、加粗、行距写进样式定义，段落上只保留原文的加粗/斜体/下划线/颜色等
        直接格式，这样同一个等级的段落共用一套标准样式，后续在 Word 里改样式即可整体同步。
        """
        records = self._level_records
        self._level_records = {}
        if not records:
            return 0

        resolved = []
        for para, level, strip in records.values():
            # 编号/层级识别得到的登记等级（数字推导，优先于格式）已写入大纲级别，
            # 此处以登记等级为准；仅当登记等级缺失（正文）时，才回退到既有样式的大纲级别。
            if level is None:
                fallback = self._effective_outline_level(para)
                if fallback is not None:
                    # 多级点分编号（x.x）按编号深度给级别下限，如 1.1 至少 Lv2
                    level = self._apply_dotted_level_floor(para.text, fallback)
            resolved.append((para, level, strip))
        if not resolved:
            return 0

        self._log("\n正在用标准样式统一同一个等级的段落（正文 / 标题 1~4）...")
        # 改写「正文」样式会影响所有继承它的未排版段落（表格、页眉页脚、空行），先固化它们的字体字号
        self._freeze_skipped_paragraph_sizes(doc, set(records))

        styles = {}
        for _, level, _ in resolved:
            key = self._style_level_key(level)
            if key in styles:
                continue
            style = self._get_level_style(doc, level)
            self._define_level_style(style, level)
            styles[key] = style
            self._log(f"  > 样式「{style.name}」：{level and f'大纲 Lv{level}' or '正文'}")

        counts = Counter()
        for para, level, strip in resolved:
            style = styles.get(self._style_level_key(level))
            if style is None:
                continue
            self._assign_paragraph_style(para, style)
            if strip:
                self._strip_direct_run_fonts(para)
            counts[self._style_level_key(level)] += 1

        detail = "，".join(
            f"{'正文' if key == 0 else f'标题 {key}'} {counts[key]} 段"
            for key in sorted(counts)
        )
        self._log(f"标准样式统一完成：{detail}。")
        return sum(counts.values())

    def _clear_unused_custom_styles(self, doc):
        """删除文档中未被引用的自定义（非内置）样式，返回被删除的样式名列表。

        仅清除「自定义样式」，内置样式（正文 / 标题 1~4 等）始终保留；被引用的样式及其
        基于链（basedOn）上的祖先样式一并保留，避免误删导致文档损坏。页眉页脚段落、
        正文段落、表格、run 上的直接样式引用都会被纳入“已使用”统计。
        """
        used_ids = set()
        used_names = set()

        def _note(style):
            if style is None:
                return
            try:
                sid = style.style_id
            except Exception:  # noqa: BLE001
                return
            if sid:
                used_ids.add(sid)
            name = getattr(style, 'name', None)
            if name:
                used_names.add(name)
            # 把 basedOn 链上的祖先样式也标记为已使用，避免误删被继承的自定义样式
            try:
                base = getattr(style, 'base_style', None)
            except Exception:  # noqa: BLE001
                base = None
            if base is not None:
                _note(base)

        for para in doc.paragraphs:
            _note(self._safe_style(para))
            for run in para.runs:
                _note(self._safe_style(run))
        for table in doc.tables:
            _note(self._safe_style(table))
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        _note(self._safe_style(para))
                        for run in para.runs:
                            _note(self._safe_style(run))

        removed = []
        for style in list(doc.styles):
            try:
                if style.builtin:
                    continue
            except Exception:  # noqa: BLE001
                continue
            try:
                sid = style.style_id
                name = getattr(style, 'name', None) or ''
            except Exception:  # noqa: BLE001
                continue
            if sid in used_ids or name in used_names:
                continue
            try:
                style_element = style._element
                parent = style_element.getparent()
                if parent is None:
                    continue
                parent.remove(style_element)
                removed.append(name)
            except Exception:  # noqa: BLE001
                self._log(f"  > 清除自定义样式「{name}」失败，已跳过。")
        return removed

    @staticmethod
    def _safe_style(target):
        """安全读取段落/run/表格的样式对象，损坏或不可用时返回 None。"""
        try:
            return target.style
        except Exception:  # noqa: BLE001 - 损坏的 pStyle / rStyle 指向
            return None

    def _remove_headers_footers(self, doc):
        """删除所有节的页眉页脚内容（段落与表格），返回删除的元素数量。

        仅处理本节独立定义的页眉页脚（is_linked_to_previous 为 False），与上一节共享的
        页眉页脚不在此处删除，避免误改上游节的内容。
        """
        removed = 0
        for section in doc.sections:
            for container in (
                section.header, section.footer,
                section.first_page_header, section.first_page_footer,
                section.even_page_header, section.even_page_footer,
            ):
                if container.is_linked_to_previous:
                    continue
                try:
                    element = container._element
                except Exception:  # noqa: BLE001
                    continue
                for child in list(element):
                    if child.tag in (qn('w:p'), qn('w:tbl')):
                        element.remove(child)
                        removed += 1
        return removed

    def _classify_attachments(self, blocks, is_from_txt):
        """Use layout and neighboring items; ambiguous inline names stay in the list."""
        candidates = {i: self._parse_attachment_line(p) for i, p in enumerate(blocks)
                      if isinstance(p, Paragraph)}
        candidates = {i: info for i, info in candidates.items() if info is not None}
        roles = {}

        def neighbor(index, step):
            index += step
            while 0 <= index < len(blocks) and self._attachment_blank(blocks[index]):
                index += step
            return index if 0 <= index < len(blocks) else None

        for index, info in candidates.items():
            para = blocks[index]
            following = neighbor(index, 1)
            previous = neighbor(index, -1)
            next_para = blocks[following] if following is not None and isinstance(blocks[following], Paragraph) else None
            next_text = self._paragraph_text(next_para).strip() if next_para is not None else ''
            next_item = bool(RE_ATTACHMENT_LIST_ITEM.match(next_text))
            centered_title = (next_para is not None and following not in candidates and not next_item
                              and self._get_paragraph_alignment(next_para) == WD_ALIGN_PARAGRAPH.CENTER
                              and not self._has_drawing_or_pict(next_para))
            page_start = self._attachment_has_page_start(blocks, index)
            centered = self._get_paragraph_alignment(para) == WD_ALIGN_PARAGRAPH.CENTER
            adjacent_labels = any(j in candidates and candidates[j]['title']
                                  and not self._attachment_has_page_start(blocks, j)
                                  for j in (previous, following))
            in_list = previous is not None and roles.get(blocks[previous]._p if isinstance(blocks[previous], Paragraph) else None, {}).get('kind') in ('list', 'list_item')
            heading_next = bool(RE_HEADING_H1.match(next_text) or RE_HEADING_H2.match(next_text))
            document_start = all(self._attachment_blank(block) for block in blocks[:index])
            strong_body = page_start or centered or centered_title
            explicit_list = not info['number'] and (RE_ATTACHMENT_LIST_ITEM.match(info['title']) or next_item)
            if explicit_list:
                kind = 'list'
            elif strong_body:
                kind = 'body'
            elif in_list or (info['title'] and adjacent_labels):
                kind = 'list'
            elif document_start and info['number']:
                kind = 'body'
            elif not info['title'] and (info['number'] or not info['colon']) and not (not info['number'] and next_item):
                kind = 'body'
            elif info['number'] and info['title'] and heading_next:
                kind = 'body'
            else:
                kind = 'list'
                if info['number'] and info['title']:
                    self._log(f"附件识别：\"{self._paragraph_text(para).strip()[:40]}\" 缺少附件正文线索，按清单保留，不强制分页。")
            info['kind'] = kind
            roles[para._p] = info
            if kind == 'list':
                enumeration = RE_ATTACHMENT_LIST_ITEM.match(info['title'])
                if info['title'] and enumeration is None:
                    continue
                enumeration_kind = enumeration.lastgroup if enumeration is not None else None
                j = index + 1
                while j < len(blocks):
                    block = blocks[j]
                    if self._attachment_blank(block):
                        j += 1
                        continue
                    if not isinstance(block, Paragraph) or j in candidates or self._attachment_has_page_start(blocks, j):
                        break
                    enumeration = RE_ATTACHMENT_LIST_ITEM.match(self._paragraph_text(block).strip())
                    if enumeration is None or (enumeration_kind and enumeration.lastgroup != enumeration_kind):
                        break
                    enumeration_kind = enumeration.lastgroup
                    roles[block._p] = {'kind': 'list_item'}
                    j += 1
        return roles

    @classmethod
    def _parse_attachment_line(cls, para):
        original = cls._paragraph_text(para)
        text = original.lstrip()
        match = RE_ATTACHMENT_LINE.fullmatch(text.rstrip())
        if not match:
            return None
        title = match.group('title').strip()
        separator = match.group('separator') or ''
        # A boundary is required: “附件1所列资料” is prose, not a label.
        if title and (not separator or title.endswith(('。', '；', ';', '！', '？', '!', '?'))
                      or re.match(r'^(?:所列|所示|所附|中的|中所|详见)', title)):
            return None
        return {'number': match.group('number'), 'title': title, 'colon': ':' in separator or '：' in separator,
                'prefix_end': len(original) - len(text) + match.start('title')}

    def _format_attachment_body(self, doc, blocks, index, info, roles, processed, apply_color, is_from_txt):
        para = blocks[index]
        self._log(f"附件正文标识：{self._paragraph_text(para).strip()[:40]}")
        inline_title = bool(info['title'] and info.get('title_element') is None)
        if inline_title:
            self._format_attachment_label_runs(para, info['prefix_end'], apply_color, inline_title=True)
        else:
            self._strip_leading_whitespace(para)
            self._apply_font_to_runs(para, self.config['attachment_font'], self.config['attachment_size'], set_color=apply_color)
            self._apply_bold_to_runs(para, self.config.get('attachment_bold', 'preserve'))
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.left_indent = Pt(0)
        para.paragraph_format.first_line_indent = None
        para._p.get_or_add_pPr().get_or_add_ind().set(qn('w:firstLineChars'), '0')
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        self._apply_line_spacing(para, 'line_spacing', 'line_spacing_unit', 'line_spacing_multiple', 28)
        original_page_break = para.paragraph_format.page_break_before
        preceding_page_break = self._attachment_has_page_start(blocks, index, include_current=False)
        self._reset_pagination_properties(para)
        para.paragraph_format.page_break_before = bool(original_page_break or not preceding_page_break)
        self._format_heading(para, 1)
        end = next((j for j in range(index + 1, len(blocks))
                    if isinstance(blocks[j], Paragraph) and blocks[j]._p in roles), len(blocks))
        first = index + 1
        while first < end and self._attachment_blank(blocks[first]):
            first += 1
        titles, subtitles = [], []
        if not inline_title and first < end and isinstance(blocks[first], Paragraph):
            candidate = blocks[first]
            text = self._paragraph_text(candidate).strip()
            is_heading = any(pattern.match(text) for pattern in (RE_HEADING_H1, RE_HEADING_H2, RE_HEADING_H3, RE_HEADING_H4))
            safe_title = info.get('title_element') is candidate._p or not (self._has_drawing_or_pict(candidate) or self._has_embedded_object(candidate))
            if text and not is_heading and safe_title and (is_from_txt or self._get_paragraph_alignment(candidate) == WD_ALIGN_PARAGRAPH.CENTER):
                titles, subtitles = self._find_title_and_subtitle_paragraphs(doc, is_from_txt, first, end)
        for j in titles:
            self._format_title_paragraph(blocks[j], apply_color, attachment=True)
        for j in subtitles:
            self._format_title_paragraph(blocks[j], apply_color, subtitle=True, attachment=True)
        processed.update(titles + subtitles)
        group = subtitles or titles
        if group:
            self._apply_title_after_spacing(blocks[group[-1]], attachment=True)
        elif inline_title:
            self._apply_title_after_spacing(para, attachment=True)

    @classmethod
    def _attachment_blank(cls, block):
        return (isinstance(block, Paragraph) and not cls._paragraph_text(block).strip()
                and not cls._has_drawing_or_pict(block) and not cls._has_embedded_object(block)
                and not cls._has_field_codes(block)
                and block._p.find('./' + qn('w:pPr') + '/' + qn('w:sectPr')) is None)

    @classmethod
    def _attachment_has_page_start(cls, blocks, index, include_current=True):
        para = blocks[index]
        if include_current and para.paragraph_format.page_break_before:
            return True
        for previous in reversed(blocks[:index]):
            if not isinstance(previous, Paragraph):
                return False
            section = previous._p.find('./' + qn('w:pPr') + '/' + qn('w:sectPr'))
            if section is not None:
                section_type = section.find(qn('w:type'))
                return section_type is None or section_type.get(qn('w:val')) not in ('continuous', 'nextColumn')
            for run in reversed(list(cls._iter_inline_runs(previous))):
                for child in reversed(list(run._r)):
                    if child.tag == qn('w:rPr') or (child.tag == qn('w:t') and not (child.text or '').strip()):
                        continue
                    return child.tag == qn('w:br') and child.get(qn('w:type')) == 'page'
            if not cls._attachment_blank(previous):
                return False
        return False

    def _format_attachment_label_runs(self, para, boundary, apply_color, inline_title=False):
        self._split_runs_at(para, boundary)
        consumed = 0
        for run in self._iter_inline_runs(para):
            length = len(self._inline_run_text(run))
            prefix = 'attachment' if consumed < boundary else ('title' if inline_title else None)
            if prefix is not None and length:
                self._set_run_font(run, self.config[prefix + '_font'], self.config[prefix + '_size'], set_color=apply_color)
                self._set_run_bold(run, self.config.get(prefix + '_bold', 'preserve'))
            consumed += length

    @staticmethod
    def _set_run_bold(run, mode):
        if mode is True:
            mode = 'bold'
        if mode in (False, None, 'preserve'):
            return
        run.font.bold = mode == 'bold'
        run.font.cs_bold = mode == 'bold'

    def _format_attachment_list(self, para, info, apply_color):
        self._log(f"附件清单：{self._paragraph_text(para).strip()[:40]}（不强制分页）")
        self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
        if info['kind'] == 'list':
            self._format_attachment_label_runs(para, info['prefix_end'], apply_color)
        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        spacing.set(qn('w:beforeAutospacing'), '0')
        spacing.set(qn('w:afterAutospacing'), '0')
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        self._apply_line_spacing(para, 'line_spacing', 'line_spacing_unit', 'line_spacing_multiple', 28)
        original_page_break = para.paragraph_format.page_break_before
        self._reset_pagination_properties(para)
        para.paragraph_format.page_break_before = original_page_break

    def _apply_title_after_spacing(self, para, attachment=False):
        scope = 'attachment_title' if attachment else 'title'
        if self.config.get('enable_' + scope + '_after_blank_line', False):
            spacing = para._p.get_or_add_pPr().get_or_add_spacing()
            for name in ('after', 'afterLines', 'afterAutospacing'):
                spacing.set(qn('w:' + name), '0')
            # Defer insertion so attachment/title indices remain stable during
            # the formatting passes, including later attachment title scans.
            if not hasattr(self, '_pending_title_blank_lines'):
                self._pending_title_blank_lines = []
            self._pending_title_blank_lines.append(para)
            return
        if not self.config.get('enable_' + scope + '_after_spacing', False):
            return
        value = self._config_float(self.config, scope + '_after_spacing', 0)
        unit = self.config.get(scope + '_after_spacing_unit', 'pt')
        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        spacing.set(qn('w:afterAutospacing'), '0')
        # Explicit zero also overrides line spacing inherited from a style.
        spacing.set(qn('w:after'), '0' if unit == 'lines' else str(int(round(value * 20))))
        spacing.set(qn('w:afterLines'), str(int(round(value * 100))) if unit == 'lines' else '0')
        contextual = para._p.get_or_add_pPr().find(qn('w:contextualSpacing'))
        if contextual is None:
            contextual = OxmlElement('w:contextualSpacing')
            para._p.get_or_add_pPr().insert_element_before(
                contextual, 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection',
                'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle',
                'w:rPr', 'w:sectPr', 'w:pPrChange')
        contextual.set(qn('w:val'), '0')

    def _insert_title_blank_lines(self):
        seen = set()
        for para in getattr(self, '_pending_title_blank_lines', []):
            if para._p in seen:
                continue
            seen.add(para._p)
            if (para._p.find('./' + qn('w:pPr') + '/' + qn('w:sectPr')) is not None
                    or re.search(r'\n[ \t\u3000]*$', self._paragraph_text(para))):
                self._log('标题后已有换行/分页或分节边界，保留原结构，不追加空行。')
                continue
            following = para._p.getnext()
            markers = {qn('w:' + name) for name in ('bookmarkStart', 'bookmarkEnd',
                        'commentRangeStart', 'commentRangeEnd')}
            while following is not None and following.tag in markers:
                following = following.getnext()
            if following is not None and following.tag == qn('w:p'):
                existing = Paragraph(following, para._parent)
                number = following.find('./' + qn('w:pPr') + '/' + qn('w:numPr') + '/' + qn('w:numId'))
                if (self._attachment_blank(existing)
                        and not self._paragraph_has_ooxml(existing, 'w:fldSimple')
                        and (number is None or number.get(qn('w:val')) == '0')):
                    self._log('标题后已有空行，保留原空行，不重复添加。')
                    continue
            element = OxmlElement('w:p')
            para._p.addnext(element)
            blank = Paragraph(element, para._parent)
            pf = blank.paragraph_format
            pf.left_indent = pf.right_indent = pf.first_line_indent = Pt(0)
            pf.space_before = pf.space_after = Pt(0)
            pf.page_break_before = pf.keep_with_next = pf.keep_together = False
            spacing = element.get_or_add_pPr().get_or_add_spacing()
            for name in ('beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing'):
                spacing.set(qn('w:' + name), '0')
            number = element.get_or_add_pPr().get_or_add_numPr().get_or_add_numId()
            number.val = 0
            self._apply_line_spacing(blank, 'line_spacing', 'line_spacing_unit', 'line_spacing_multiple', 28)
            run = blank.add_run()
            self._set_run_font(run, self.config['body_font'], self.config['body_size'])
            element.get_or_add_pPr().append(deepcopy(run._r.rPr))
            self._log('已在标题组末尾添加一个空行（采用正文行距）。')
        self._pending_title_blank_lines = []


    @staticmethod
    def _join_text_lines(prev, nxt):
        """文本输出时拼接两行：ASCII  alphanumeric 交界补空格，与 docx 合并行为一致。"""
        left = prev.rstrip()
        right = nxt.lstrip()
        if (
            left and right
            and left[-1].isascii() and left[-1].isalnum()
            and right[0].isascii() and right[0].isalnum()
        ):
            return left + ' ' + right
        return left + right

    @staticmethod
    def _md_list_parts(line, source_ext):
        """拆分 Markdown 列表前缀与正文；非列表行前缀为空。"""
        value = line or ''
        if source_ext != '.md':
            return '', value.strip()
        match = RE_MD_LIST_PREFIX.match(value)
        if not match:
            return '', value.strip()
        return match.group(1), value[match.end():].strip()

    def _text_line_is_mergeable(self, line, source_ext):
        """空行与 Markdown 标题/表格/围栏等结构行是硬边界；列表项只保护前缀。"""
        value = (line or '').strip()
        if not value:
            return False
        if source_ext == '.md' and RE_MD_PROTECTED_LINE.match(line):
            return False
        prefix, body = self._md_list_parts(line, source_ext)
        if prefix:
            return bool(body)
        return not self._is_heading_like_text(value)

    def _merge_text_lines(self, text, source_ext):
        """在纯文本上执行与 docx 相同的“前后行词语连续合并”。"""
        lines = text.split('\n')
        # 围栏代码块里的行（含围栏本身）不参与合并，也不作为合并的桥梁。
        skipped = self._md_fence_line_indices(lines, source_ext)
        entries = [
            (*self._md_list_parts(line, source_ext),
             self._text_line_is_mergeable(line, source_ext) and index not in skipped)
            for index, line in enumerate(lines)
        ]
        numbered_gaps = self._numbered_list_gap_indices(lines)
        result = []
        index, total = 0, len(lines)
        merged_lines = 0
        while index < total:
            if not entries[index][2]:
                result.append(lines[index])
                index += 1
                continue

            end = index
            while end + 1 < total:
                next_prefix, next_body, next_mergeable = entries[end + 1]
                if not next_mergeable or next_prefix:
                    break  # 下一行是独立列表项或围栏行，不与前一项合并
                if not self._is_continuous_pair(
                    entries[end][1], next_body,
                    prev_is_list_item=bool(entries[end][0]),
                    prev_numbered_gap=end in numbered_gaps,
                ):
                    break
                end += 1

            if end > index:
                combined = lines[index]
                for nxt in range(index + 1, end + 1):
                    combined = self._join_text_lines(combined, entries[nxt][1])
                result.append(combined)
                merged_lines += end - index
            else:
                result.append(lines[index])
            index = end + 1

        self._log(f"文本输出：消除 {merged_lines} 处断行。")
        return '\n'.join(result)

    def _format_text_document(self, input_path, output_path, source_ext):
        """保持 txt/md 后缀输出：应用全部非格式化清洗，跳过 Word 专属排版。"""
        text = self._read_text_file(input_path)
        applied = []
        # 顺序：空格清理 → 冗余序号清理 → 断行合并 → 符号标准化。
        # 空格先清掉可让后续判据基于干净文本；冗余序号必须在断行合并之前，
        # 否则 docling 叠在续行前的页面序号会被当作独立列表项而拒绝合并。
        if self.config.get('remove_cjk_spaces', False):
            text = self._apply_text_cleanup(text, source_ext, self._remove_cjk_spaces_in_text)
            applied.append("中英文/数字间多余空格清理")
        if self.config.get('remove_duplicate_numbering', False):
            text = self._remove_duplicate_numbering_in_text(text, source_ext)
            applied.append("冗余序号清理")
        if self.config.get('merge_continuous_lines', False):
            text = self._merge_text_lines(text, source_ext)
            applied.append("前后行词语连续合并")
        if self.config.get('normalize_punctuation', False):
            text = self._normalize_text_symbols(text, source_ext)
            applied.append("符号标准化")
        if not applied:
            self._log("文本输出：未开启任何清洗选项，内容保持原样。")

        with open(output_path, 'w', encoding='utf-8', newline='') as handle:
            handle.write(text)

        self._log(
            f"保持源文件后缀输出（{source_ext}）：已应用"
            + ("、".join(applied) if applied else "无清洗项")
            + "；字体/样式/大纲/表格/页眉页脚等 Word 专属排版选项已跳过。"
        )

    def format_document(self, input_path, output_path):
        source_ext = os.path.splitext(str(input_path))[1].lower()
        if self.config.get('keep_source_extension', False) and source_ext in ('.txt', '.md'):
            if os.path.splitext(str(output_path))[1].lower() != source_ext:
                output_path = os.path.splitext(str(output_path))[0] + source_ext
            self._format_text_document(input_path, output_path, source_ext)
            return

        processing_path, is_from_txt = self.convert_to_docx(input_path)
        processing_path = self._preprocess_document(processing_path, is_from_txt)
        
        doc = Document(processing_path)

        apply_color = not is_from_txt
        self._apply_color = apply_color

        # 清洗顺序：空格清理 → 冗余序号清理 → 断行合并 → 表格合并 → 符号标准化，
        # 全部先于标题证据预扫描与版式排版。
        # 空格先清掉可让后续判据基于干净文本；冗余序号必须在断行合并之前，
        # 否则 docling 叠在续行前的页面序号会被当作独立列表项而拒绝合并。
        if self.config.get('remove_cjk_spaces', False):
            self._log("正在去除中文与中文/英文/数字之间的多余空格...")
            space_changes = self._remove_cjk_spaces_in_document(doc)
            self._log(f"多余空格清理完成，共处理 {space_changes} 个段落。")

        if self.config.get('remove_duplicate_numbering', False):
            self._log("正在按前后连续性清除 PDF/docling 叠加的冗余序号...")
            dup_changes = self._remove_duplicate_numbering_in_document(doc)
            self._log(f"冗余序号清理完成，共处理 {dup_changes} 个段落。")

        if self.config.get('merge_continuous_lines', False):
            self._log("正在合并前后行词语连续的断行（PDF 粘贴导致的换行）...")
            self._merge_continuous_lines(doc)

        if self.config.get('merge_repeated_table_headers', False):
            self._log("正在合并相邻且表头相同的表格（Word 重复标题行）...")
            self._merge_repeated_table_headers(doc)

        if self.config.get('normalize_punctuation', False):
            symbol_changes = self._normalize_document_symbols(doc)
            self._log(f"符号标准化完成，共修复 {symbol_changes} 个段落/表格单元格。")

        # 以上清洗全部完成后才做标题证据预扫描与定级，避免重复序号/空格误判层级
        self._prepare_chapter_section_offset(doc)
        self._prepare_dotted_base(doc)
        self._prepare_single_number_restart(doc)

        if self.config.get('convert_floating_objects', False):
            self._log("正在把浮动式对象转换为插入式...")
            self._convert_floating_to_inline(doc)

        # 条标题：先把「第X条 正文……」拆成“条标题 + 正文”两段，再进入逐段排版
        self._level_records = {}
        self._non_heading_paras = set()
        if self.config.get('detect_article_title', True):
            self._log("正在识别条标题（第X条）...")
            self._split_article_paragraphs(doc)

        all_blocks = list(self._iter_block_items(doc))
        processed_indices = set()

        # 上游 2.7.9 附件分类/签名（在 fork 清理之后、公文版式之前执行）
        attachment_roles = {}
        signature_enabled = self.config.get('enable_signature_formatting', False)
        if self.config.get('enable_attachment_formatting', False) or signature_enabled:
            attachment_roles = self._classify_attachments(all_blocks, is_from_txt)

        # 预计算各段落的编号/章节标题层级，并对每个标题判断“局部末级且局部无正文”：
        # 一个标题的局部范围是它之后到下一个同级或更高级标题为止；若该范围内
        # 既没有更深的下级标题、也没有正文内容，说明它其实只是正文（不设大纲级别）。
        heading_ranks = [
            self._paragraph_heading_rank(block.text, block) if isinstance(block, Paragraph) else 0
            for block in all_blocks
        ]
        total = len(all_blocks)
        is_content = [False] * total  # 非标题的实体内容（正文段落、图片或表格）
        for _i, _block in enumerate(all_blocks):
            if heading_ranks[_i] != 0:
                continue
            if isinstance(_block, Paragraph):
                is_content[_i] = (
                    bool(_block.text.strip())
                    or self._has_drawing_or_pict(_block)
                    or self._has_embedded_object(_block)
                )
            else:
                is_content[_i] = True  # 表格等非段落内容也算作该标题下的内容

        # 单独的末级标号（(1)/(a) 等）：若其后（跳过空段）既无所属正文、也没有真正的下游标题，
        # 说明它只是列表标号而非结构标题，按正文处理、不给等级标题。
        def _is_h4_label(text):
            t = (text or '').lstrip()
            return bool(RE_HEADING_H4.match(t) or RE_HEADING_H4_ALPHA.match(t))

        _i = 0
        while _i < total:
            _blk = all_blocks[_i]
            if not (isinstance(_blk, Paragraph) and _is_h4_label(_blk.text)):
                _i += 1
                continue
            _run_end = _i
            while _run_end < total:
                _b = all_blocks[_run_end]
                if isinstance(_b, Paragraph) and _is_h4_label(_b.text):
                    _run_end += 1
                else:
                    break
            # 该组之后的下一个“非空”块（跳过空段落）
            _after = _run_end
            while _after < total:
                _b = all_blocks[_after]
                if isinstance(_b, Paragraph):
                    if _b.text.strip() or self._has_drawing_or_pict(_b) or self._has_embedded_object(_b):
                        break
                else:
                    break  # 表格等非段落内容即视为正文
                _after += 1
            _has_support = False
            if _after < total:
                _b = all_blocks[_after]
                if not isinstance(_b, Paragraph):
                    _has_support = True  # 表格等算作正文
                else:
                    _btext = _b.text
                    # 非 (1)/(a) 同类、且非空：正文或真正的下游标题都算“有支撑”
                    if not _is_h4_label(_btext):
                        _has_support = bool(_btext.strip()) or self._has_drawing_or_pict(_b) or self._has_embedded_object(_b)
            if not _has_support:
                for _k in range(_i, _run_end):
                    if isinstance(all_blocks[_k], Paragraph):
                        self._non_heading_paras.add(all_blocks[_k]._p)
                        heading_ranks[_k] = 0
                        is_content[_k] = bool(all_blocks[_k].text.strip())
            _i = _run_end

        # 前缀和：统计任意区间内的“更深标题数”和“正文内容数”。
        prefix_heading = [0] * (total + 1)
        prefix_content = [0] * (total + 1)
        for _i in range(total):
            prefix_heading[_i + 1] = prefix_heading[_i] + (1 if heading_ranks[_i] > 0 else 0)
            prefix_content[_i + 1] = prefix_content[_i] + (1 if is_content[_i] else 0)
        # 每个标题右侧最近的“同级或更高级标题”，即其局部范围的终点。
        scope_end = [total] * total
        _stack = []
        for _i in range(total - 1, -1, -1):
            _rank = heading_ranks[_i]
            if _rank <= 0:
                continue
            while _stack and heading_ranks[_stack[-1]] > _rank:
                _stack.pop()
            scope_end[_i] = _stack[-1] if _stack else total
            _stack.append(_i)
        # 只有文档中最深的标题级别才可能“其实只是正文”，上级标题即使局部没有正文也保留
        # （它们是结构层级，例如紧接着下一个同级标题的分章标题）。
        deepest_rank = max(heading_ranks) if heading_ranks else 0
        self._demoted_body = [False] * total
        for _i in range(total):
            if heading_ranks[_i] <= 0 or heading_ranks[_i] < deepest_rank:
                continue
            _end = scope_end[_i]
            # 局部范围内更深的下级标题数（非 0 说明它是父级标题，应保留）。
            deeper = prefix_heading[_end] - prefix_heading[_i + 1]
            # 局部范围内的正文内容数。
            content = prefix_content[_end] - prefix_content[_i + 1]
            if deeper == 0 and content == 0:
                self._demoted_body[_i] = True

        # 连续相邻的同级序号子标题：若前面各项都因“没有专属下文”降级为正文，
        # 那么末项之后的正文属于整组序号段落的总体说明，而不是末项的专属下文，
        # 末项不应单独被标题化。
        _run_start = 0
        while _run_start < total:
            _rank = heading_ranks[_run_start]
            if (
                _rank <= 0
                or not isinstance(all_blocks[_run_start], Paragraph)
                or not _is_h4_label(all_blocks[_run_start].text)
            ):
                _run_start += 1
                continue
            _run_end = _run_start + 1
            while (
                _run_end < total
                and heading_ranks[_run_end] == _rank
                and isinstance(all_blocks[_run_end], Paragraph)
                and _is_h4_label(all_blocks[_run_end].text)
            ):
                _run_end += 1
            _last = _run_end - 1
            if _run_end - _run_start >= 2 and not self._demoted_body[_last]:
                _end = scope_end[_last]
                _has_deeper = prefix_heading[_end] - prefix_heading[_last + 1]
                _has_content = prefix_content[_end] - prefix_content[_last + 1]
                if (
                    _has_deeper == 0
                    and _has_content > 0
                    and all(self._demoted_body[k] for k in range(_run_start, _last))
                ):
                    self._demoted_body[_last] = True
            _run_start = _run_end

        if not is_from_txt:
            self._log("正在扫描图表标题...")
            for idx, block in enumerate(all_blocks):
                is_pic_para = isinstance(block, Paragraph) and self._has_drawing_or_pict(block)
                is_table = isinstance(block, Table)
                
                if not (is_pic_para or is_table): continue
                
                for direction in [-1, 1]:
                    caption_found = False
                    for i in range(idx + direction, -1 if direction == -1 else len(all_blocks), direction):
                        if i in processed_indices: continue
                        potential_caption = all_blocks[i]
                        if not isinstance(potential_caption, Paragraph): break 
                        text = potential_caption.text.strip()
                        if text: 
                            if self._get_paragraph_alignment(potential_caption) == WD_ALIGN_PARAGRAPH.CENTER and (text.startswith("图") or text.startswith("表")):
                                detected_type = "图" if text.startswith("图") else "表"
                                self._log(f"  > 发现 {detected_type} 的标题: \"{text[:30]}...\" (在段落 {i+1})")
                                config_font_key = f'{("figure" if detected_type == "图" else "table")}_caption_font'
                                config_size_key = f'{("figure" if detected_type == "图" else "table")}_caption_size'
                                config_font = self.config[config_font_key]
                                config_size = self.config[config_size_key]
                                self._apply_font_to_runs(potential_caption, config_font, config_size, set_color=apply_color)
                                processed_indices.add(i)
                                caption_found = True
                            break 
                    if caption_found: break 

        # 查找主标题和副标题
        main_end = next((i for i, p in enumerate(all_blocks)
                         if self.config.get('enable_attachment_formatting', False)
                         and isinstance(p, Paragraph) and p._p in attachment_roles), None)
        title_indices, subtitle_indices = self._find_title_and_subtitle_paragraphs(doc, is_from_txt, end_index=main_end)
        signature_formatter, signature_roles = None, {}
        if signature_enabled:
            from wfp_signature import SignatureFormatter
            signature_formatter = SignatureFormatter(self)
            excluded = {all_blocks[i]._p for i in set(title_indices + subtitle_indices) | processed_indices
                        if isinstance(all_blocks[i], Paragraph)}
            signature_roles = signature_formatter.classify(all_blocks, attachment_roles, excluded)
            self._log('落款识别完成，确认 {} 组；未确认的段落沿用现有排版。'.format(len(signature_roles) // 2))
        
        # 将标题和副标题索引加入已处理集合
        for idx in title_indices:
            processed_indices.add(idx)
        for idx in subtitle_indices:
            processed_indices.add(idx)

        # 已处理段落不参与后续的大纲级别归纳
        protected_paragraphs = [
            all_blocks[idx]._p for idx in processed_indices
            if isinstance(all_blocks[idx], Paragraph)
        ]

        self._log("预扫描完成，开始逐段格式化...")
        if self.config['set_outline']:
            self._log("【大纲级别设置已启用】")
        else:
            self._log("【大纲级别设置已禁用】")

        outline_enabled = self.config.get('set_outline', True)
        if self.config.get('normalize_heading_levels', False):
            if outline_enabled:
                self._log("\n正在合并相同格式的标题等级...")
                self._normalize_heading_levels(doc, exclude_elements=protected_paragraphs)
            else:
                self._log("\n【标题等级合并优化】需要先启用“自动设置大纲级别”，已跳过。")
            
        # 格式化主标题
        if title_indices:
            self._log(f"\n开始格式化主标题（共 {len(title_indices)} 行）...")
            for idx in title_indices:
                para = all_blocks[idx]
                self._log(f"段落 {idx + 1}: 主标题行 - \"{para.text[:30]}...\"")
                self._strip_leading_whitespace(para)
                self._apply_font_to_runs(para, self.config['title_font'], self.config['title_size'], set_color=apply_color)
                self._apply_bold_to_runs(para, self.config.get('title_bold', False))
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.first_line_indent = None
                
                # 设置标题行间距
                spacing = para._p.get_or_add_pPr().get_or_add_spacing()
                spacing.set(qn('w:beforeAutospacing'), '0')
                spacing.set(qn('w:afterAutospacing'), '0')
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
                self._apply_line_spacing(
                    para,
                    'title_line_spacing',
                    'title_line_spacing_unit',
                    'title_line_spacing_multiple',
                    33,
                )
                
                self._reset_pagination_properties(para)
        
        # 格式化副标题
        if subtitle_indices:
            self._log(f"\n开始格式化副标题（共 {len(subtitle_indices)} 行）...")
            for idx in subtitle_indices:
                para = all_blocks[idx]
                self._log(f"段落 {idx + 1}: 副标题行 - \"{para.text[:30]}...\"")
                self._strip_leading_whitespace(para)
                self._apply_font_to_runs(para, self.config['subtitle_font'], self.config['subtitle_size'], set_color=apply_color)
                self._apply_bold_to_runs(para, self.config.get('subtitle_bold', 'preserve'))
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.first_line_indent = None
                
                # 设置副标题行间距
                spacing = para._p.get_or_add_pPr().get_or_add_spacing()
                spacing.set(qn('w:beforeAutospacing'), '0')
                spacing.set(qn('w:afterAutospacing'), '0')
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
                self._apply_line_spacing(
                    para,
                    'subtitle_line_spacing',
                    'subtitle_line_spacing_unit',
                    'subtitle_line_spacing_multiple',
                    33,
                )
                
                self._reset_pagination_properties(para)

        main_group = subtitle_indices or title_indices
        if main_group:
            self._apply_title_after_spacing(all_blocks[main_group[-1]])

        block_idx = 0
        while block_idx < len(all_blocks):
            block = all_blocks[block_idx]
            
            if block_idx in processed_indices:
                if block_idx not in title_indices and block_idx not in subtitle_indices:
                    self._log(f"块 {block_idx + 1}: 已作为图表/附件标题处理 - 跳过")
                block_idx += 1
                continue

            current_block_num = block_idx + 1
            if isinstance(block, Table): 
                self._log(f"块 {current_block_num}: 表格 - 跳过"); block_idx += 1; continue
            
            para = block
            if para._p in signature_roles:
                # Retain original paragraph properties until the signature pass, after page setup.
                block_idx += 1
                continue
            if not self._paragraph_text(para).strip():
                self._log(f"段落 {current_block_num}: 空白 - 跳过"); block_idx += 1; continue

            attachment = attachment_roles.get(para._p) if self.config.get('enable_attachment_formatting', False) else None
            if attachment is not None:
                if attachment['kind'] == 'body':
                    self._format_attachment_body(doc, all_blocks, block_idx, attachment, attachment_roles,
                                                 processed_indices, apply_color, is_from_txt)
                else:
                    self._format_attachment_list(para, attachment, apply_color)
                block_idx += 1
                continue
            
            is_pic = self._has_drawing_or_pict(para)
            is_embedded_obj = self._has_embedded_object(para)
            if is_pic or is_embedded_obj:
                log_msg = "图片" if is_pic else "嵌入对象"
                self._log(f"段落 {current_block_num}: {log_msg} - 仅格式化文字")
                
                text_to_check = self._paragraph_text(para).lstrip()
                para_text_preview = text_to_check[:30].replace("\n", " ")

                if RE_HEADING_H1.match(text_to_check):
                    self._log(f"  > 文字识别为一级标题: \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['h1_font'], self.config['h1_size'], set_color=apply_color)
                    self._apply_bold_to_runs(para, self.config.get('h1_bold', False))
                elif RE_HEADING_H2.match(text_to_check):
                    self._log(f"  > 文字识别为二级标题: \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['h2_font'], self.config['h2_size'], set_color=apply_color)
                    self._apply_bold_to_runs(para, self.config.get('h2_bold', False))
                elif (dotted_level := self._paragraph_dotted_level(para, text_to_check)) is not None:
                    font, size, bold = self._heading_font_for_level(dotted_level)
                    self._log(f"  > 文字识别为多级编号标题(Lv{dotted_level}): \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, font, size, set_color=apply_color)
                    self._apply_bold_to_runs(para, bold)
                elif RE_HEADING_H3.match(text_to_check):
                    self._log(f"  > 文字识别为三级标题: \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                elif RE_HEADING_H4.match(text_to_check) or RE_HEADING_H4_ALPHA.match(text_to_check):
                    self._log(f"  > 文字识别为四级标题: \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                elif text_to_check:
                    self._log(f"  > 文字识别为正文: \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)

                block_idx += 1
                continue

            # 默认按正文登记等级；标题分支会改写为对应等级，附件标识等特殊段落撤销登记
            self._record_paragraph_level(para)

            original_text, text_to_check = para.text, para.text.lstrip()
            text_to_check_stripped = para.text.strip()
            leading_space_count = len(original_text) - len(text_to_check)
            para_text_preview = text_to_check[:30].replace("\n", " ")
            chapter_level = self._chapter_section_level(text_to_check)
            article_level = self._article_title_level(text_to_check)
            
            spacing = para._p.get_or_add_pPr().get_or_add_spacing()
            spacing.set(qn('w:beforeAutospacing'), '0'); spacing.set(qn('w:afterAutospacing'), '0')
            para.paragraph_format.space_before, para.paragraph_format.space_after = Pt(0), Pt(0)
            self._apply_line_spacing(
                para,
                'line_spacing',
                'line_spacing_unit',
                'line_spacing_multiple',
                28,
            )

            is_attachment_enabled = self.config.get('enable_attachment_formatting', False)
            is_attachment_candidate = False
            if is_from_txt:
                if RE_ATTACHMENT.match(text_to_check_stripped): is_attachment_candidate = True
            elif self._get_paragraph_alignment(para) in [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY, None] and RE_ATTACHMENT.match(text_to_check_stripped):
                is_attachment_candidate = True

            if is_attachment_enabled and is_attachment_candidate:
                self._log(f"段落 {current_block_num}: 附件标识 - \"{para_text_preview}...\"")
                # 附件标识有自己的字体和分页规则，不并入「正文/标题」标准样式
                self._level_records.pop(id(para._p), None)
                self._strip_leading_whitespace(para)
                self._apply_font_to_runs(para, self.config['attachment_font'], self.config['attachment_size'], set_color=apply_color)
                self._reset_pagination_properties(para)
                para.paragraph_format.page_break_before = True
                para.paragraph_format.left_indent = Pt(0)
                para.paragraph_format.first_line_indent = None
                
                ind = para._p.get_or_add_pPr().get_or_add_ind()
                ind.set(qn("w:firstLineChars"), "0")
                
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                self._format_heading(para, 1)

                # 查找并格式化附件的标题和副标题
                search_idx = block_idx + 1
                
                # 查找附件的标题和副标题：限定在附件标识之后紧邻的段落，
                # 跳过已处理段落（如图表标题），避免把远处居中段误认成附件标题
                att_title_indices, att_subtitle_indices = self._find_title_and_subtitle_paragraphs(
                    doc, is_from_txt, search_idx,
                    exclude_indices=processed_indices,
                    stop_at_body=not is_from_txt,
                )
                
                # 将附件的标题和副标题加入已处理集合
                for idx in att_title_indices:
                    processed_indices.add(idx)
                for idx in att_subtitle_indices:
                    processed_indices.add(idx)
                
                # 格式化附件的标题
                if att_title_indices:
                    self._log(f"  > 识别到附件标题（共 {len(att_title_indices)} 行）")
                    for idx in att_title_indices:
                        para_title = all_blocks[idx]
                        self._log(f"    段落 {idx + 1}: 附件标题行 - \"{para_title.text.strip()[:30]}...\"")
                        self._strip_leading_whitespace(para_title)
                        self._apply_font_to_runs(para_title, self.config['title_font'], self.config['title_size'], set_color=apply_color)
                        self._apply_bold_to_runs(para_title, self.config.get('title_bold', False))
                        para_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        para_title.paragraph_format.first_line_indent = None
                        
                        # 设置标题行间距
                        spacing = para_title._p.get_or_add_pPr().get_or_add_spacing()
                        spacing.set(qn('w:beforeAutospacing'), '0')
                        spacing.set(qn('w:afterAutospacing'), '0')
                        para_title.paragraph_format.space_before = Pt(0)
                        para_title.paragraph_format.space_after = Pt(0)
                        self._apply_line_spacing(
                            para_title,
                            'title_line_spacing',
                            'title_line_spacing_unit',
                            'title_line_spacing_multiple',
                            33,
                        )
                        
                        self._reset_pagination_properties(para_title)
                        self._format_heading(para_title, 1)
                
                # 格式化附件的副标题
                if att_subtitle_indices:
                    self._log(f"  > 识别到附件副标题（共 {len(att_subtitle_indices)} 行）")
                    for idx in att_subtitle_indices:
                        para_subtitle = all_blocks[idx]
                        self._log(f"    段落 {idx + 1}: 附件副标题行 - \"{para_subtitle.text.strip()[:30]}...\"")
                        self._strip_leading_whitespace(para_subtitle)
                        self._apply_font_to_runs(para_subtitle, self.config['subtitle_font'], self.config['subtitle_size'], set_color=apply_color)
                        para_subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        para_subtitle.paragraph_format.first_line_indent = None
                        
                        # 设置副标题行间距
                        spacing = para_subtitle._p.get_or_add_pPr().get_or_add_spacing()
                        spacing.set(qn('w:beforeAutospacing'), '0')
                        spacing.set(qn('w:afterAutospacing'), '0')
                        para_subtitle.paragraph_format.space_before = Pt(0)
                        para_subtitle.paragraph_format.space_after = Pt(0)
                        self._apply_line_spacing(
                            para_subtitle,
                            'subtitle_line_spacing',
                            'subtitle_line_spacing_unit',
                            'subtitle_line_spacing_multiple',
                            33,
                        )
                        
                        self._reset_pagination_properties(para_subtitle)
                
                # 计算下一个要处理的块索引。没有附件标题/副标题时，
                # 只跳过附件标识本段，避免漏处理紧随其后的正文段落。
                handled_indices = att_title_indices + att_subtitle_indices
                if handled_indices:
                    next_idx = max(handled_indices) + 1
                else:
                    next_idx = block_idx + 1
                
                block_idx = next_idx
                continue
            
            elif chapter_level is not None:
                if self._is_demoted_deepest_heading(block_idx):
                    self._format_demoted_deepest_as_body(para, protected_paragraphs, current_block_num, para_text_preview)
                else:
                    kind = "章" if chapter_level == 1 else "节"
                    self._log(f"段落 {current_block_num}: {kind}标题（大纲 Lv{chapter_level}）- \"{para_text_preview}...\"")
                    self._strip_leading_whitespace(para)
                    self._format_heading(para, chapter_level)
                    font, size, bold = self._heading_font_for_level(chapter_level)
                    self._apply_font_to_runs(para, font, size, set_color=apply_color)
                    self._apply_bold_to_runs(para, bold)
                    self._apply_text_indent_and_align(para)
                    self._reset_pagination_properties(para)
                    self._record_paragraph_level(para, chapter_level)

            elif article_level is not None:
                if self._is_demoted_deepest_heading(block_idx):
                    self._format_demoted_deepest_as_body(para, protected_paragraphs, current_block_num, para_text_preview)
                else:
                    font, size, bold = self._heading_font_for_level(article_level)
                    self._log(f"段落 {current_block_num}: 条标题（大纲 Lv{article_level}）- \"{para_text_preview}...\"")
                    self._strip_leading_whitespace(para)
                    self._format_heading(para, article_level)
                    self._apply_font_to_runs(para, font, size, set_color=apply_color)
                    self._apply_bold_to_runs(para, bold)
                    self._apply_text_indent_and_align(para)
                    self._reset_pagination_properties(para)
                    self._record_paragraph_level(para, article_level)

            elif RE_HEADING_H1.match(text_to_check):
                level = 1 + self._outline_offset
                if self._is_demoted_deepest_heading(block_idx):
                    self._format_demoted_deepest_as_body(para, protected_paragraphs, current_block_num, para_text_preview)
                else:
                    font, size, bold = self._heading_font_for_level(level)
                    self._log(f"段落 {current_block_num}: 一级标题（大纲 Lv{level}）- \"{para_text_preview}...\"")
                    self._strip_leading_whitespace(para)
                    self._format_heading(para, level)
                    self._apply_font_to_runs(para, font, size, set_color=apply_color)
                    self._apply_bold_to_runs(para, bold)
                    self._apply_text_indent_and_align(para)
                    self._reset_pagination_properties(para)
                    self._record_paragraph_level(para, level)

            elif RE_HEADING_H2.match(text_to_check):
                level = 2 + self._outline_offset
                if self._is_demoted_deepest_heading(block_idx):
                    self._format_demoted_deepest_as_body(para, protected_paragraphs, current_block_num, para_text_preview)
                else:
                    head_font, head_size, head_bold = self._heading_font_for_level(level)
                    self._log(f"段落 {current_block_num}: 二级标题（大纲 Lv{level}）- \"{para_text_preview}...\"")
                    self._strip_leading_whitespace(para)

                    parts = para.text.split('。', 1)

                    if len(parts) == 2 and parts[1].strip():
                        self._log("  > 检测到二级标题与正文在同一段落，执行段内格式拆分。")
                        title_len = len(parts[0]) + 1

                        original_runs = []
                        for r in para.runs:
                            original_runs.append({
                                'text': r.text, 'bold': r.bold, 'italic': r.italic,
                                'underline': r.underline, 'font_color': r.font.color.rgb
                            })

                        para.clear()

                        char_count = 0
                        for run_info in original_runs:
                            run_text = run_info['text']
                            run_end_pos = char_count + len(run_text)

                            title_run, body_run, new_run = None, None, None

                            if run_end_pos <= title_len:
                                new_run = para.add_run(run_text)
                                self._set_run_font(new_run, head_font, head_size, set_color=apply_color)

                            elif char_count >= title_len:
                                new_run = para.add_run(run_text)
                                self._set_run_font(new_run, self.config['body_font'], self.config['body_size'], set_color=apply_color)

                            else:
                                split_index = title_len - char_count
                                title_part = run_text[:split_index]
                                body_part = run_text[split_index:]

                                if title_part:
                                    title_run = para.add_run(title_part)
                                    self._set_run_font(title_run, head_font, head_size, set_color=apply_color)
                                if body_part:
                                    body_run = para.add_run(body_part)
                                    self._set_run_font(body_run, self.config['body_font'], self.config['body_size'], set_color=apply_color)

                            runs_to_format = [r for r in [title_run, body_run] if r] or ([new_run] if new_run else [])
                            for r in runs_to_format:
                                if r:
                                    r.bold = run_info['bold']; r.italic = run_info['italic']
                                    r.underline = run_info['underline']
                                    if run_info['font_color']: r.font.color.rgb = run_info['font_color']

                            char_count = run_end_pos

                        self._format_heading(para, level)
                        self._apply_bold_to_first_chars(para, title_len, head_bold)
                        self._apply_text_indent_and_align(para)
                        self._reset_pagination_properties(para)
                        # 标题与正文同段：保留段内各自的直接字体字号，只套样式
                        self._record_paragraph_level(para, level, strip=False)

                    else:
                        match = RE_H2_INLINE_TITLE.match(text_to_check)
                        if match and not (text_to_check.startswith('（') and text_to_check.strip().endswith('）')):
                            self._log("  > 已将二级标题的括号统一为中文括号。")
                            for r in para.runs: r.text = r.text.replace('(', '（', 1).replace(')', '）', 1)
                        self._format_heading(para, level)
                        self._apply_font_to_runs(para, head_font, head_size, set_color=apply_color)
                        self._apply_bold_to_runs(para, head_bold)
                        self._apply_text_indent_and_align(para)
                        self._reset_pagination_properties(para)
                        self._record_paragraph_level(para, level)
                    
            elif ((dotted_level := self._paragraph_dotted_level(para, text_to_check)) is not None
                    or (dotted_level := self._dotted_style_heading_level(para, text_to_check)) is not None):
                level = dotted_level
                if self._is_demoted_deepest_heading(block_idx):
                    self._format_demoted_deepest_as_body(para, protected_paragraphs, current_block_num, para_text_preview)
                else:
                    font, size, bold = self._heading_font_for_level(level)
                    self._log(f"段落 {current_block_num}: 多级编号标题（大纲 Lv{level}）- \"{para_text_preview}...\"")
                    self._strip_leading_whitespace(para)
                    self._format_heading(para, level)
                    self._apply_font_to_runs(para, font, size, set_color=apply_color)
                    self._apply_bold_to_runs(para, bold)
                    self._apply_text_indent_and_align(para)
                    self._reset_pagination_properties(para)
                    self._record_paragraph_level(para, level)
                

            elif RE_HEADING_H4.match(text_to_check) or RE_HEADING_H4_ALPHA.match(text_to_check):
                if para._p in self._non_heading_paras:
                    # 单独的 (1)/(a) 标号，无所属正文或下游标题：按正文处理，不给等级标题
                    self._log(f"段落 {current_block_num}: 单独末级标号按正文 - \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                    self._apply_text_indent_and_align(para)
                    self._reset_pagination_properties(para)
                else:
                    # 位于多级编号下游时取“当前上下文的下一级”，否则仍按四级标题处理
                    level = self._paren_number_level_of(para) or (4 + self._outline_offset)
                    if self._is_demoted_deepest_heading(block_idx):
                        self._format_demoted_deepest_as_body(para, protected_paragraphs, current_block_num, para_text_preview)
                    else:
                        font, size, bold = self._heading_font_for_level(level)
                        self._log(f"段落 {current_block_num}: 四级标题（大纲 Lv{level}）- \"{para_text_preview}...\"")
                        self._strip_leading_whitespace(para)
                        self._format_heading(para, level)
                        self._apply_font_to_runs(para, font, size, set_color=apply_color)
                        self._apply_bold_to_runs(para, bold)
                        self._apply_text_indent_and_align(para)
                        self._reset_pagination_properties(para)
                        self._record_paragraph_level(para, level)
                
            elif not is_from_txt:
                para_alignment = self._get_paragraph_alignment(para)
                if para_alignment in [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT]:
                    align_text = "居中" if para_alignment == WD_ALIGN_PARAGRAPH.CENTER else "右对齐"
                    self._log(f"段落 {current_block_num}: {align_text}正文 - 保留原对齐")
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                    self._reset_pagination_properties(para)
                elif leading_space_count > 5:
                    self._log(f"段落 {current_block_num}: 正文 (保留前导空格) - \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    self._reset_pagination_properties(para)
                elif (para.paragraph_format.first_line_indent is None or para.paragraph_format.first_line_indent.pt == 0) and leading_space_count == 0:
                    self._log(f"段落 {current_block_num}: 正文 (保留0缩进) - \"{para_text_preview}...\"")
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    self._reset_pagination_properties(para)
                else:
                    self._log(f"段落 {current_block_num}: 正文 (应用标准缩进) - \"{para_text_preview}...\"")
                    self._strip_leading_whitespace(para)
                    self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                    self._apply_text_indent_and_align(para)
                    self._reset_pagination_properties(para)
            else:
                self._log(f"段落 {current_block_num}: 正文 (源自TXT，强制缩进) - \"{para_text_preview}...\"")
                self._strip_leading_whitespace(para)
                self._apply_font_to_runs(para, self.config['body_font'], self.config['body_size'], set_color=apply_color)
                self._apply_text_indent_and_align(para)
                self._reset_pagination_properties(para)
            
            block_idx += 1
        
        self._format_tables(doc, apply_color=apply_color)

        if self.config.get('media_single_line_spacing', False):
            self._apply_media_single_line_spacing(doc)

        if self.config.get('normalize_heading_levels', False) and outline_enabled:
            self._log("\n正在检查标题等级是否跳级...")
            self._compact_outline_levels(doc, exclude_elements=protected_paragraphs)

        if self.config.get('use_builtin_styles', True):
            self._apply_level_styles(doc)
        else:
            self._level_records = {}

        self._apply_page_setup(doc, is_from_txt=is_from_txt)

        if self.config.get('remove_headers_footers', False):
            removed_hf = self._remove_headers_footers(doc)
            if removed_hf:
                self._log(f"删除页眉页脚内容：{removed_hf} 处。")

        if self.config.get('clear_unused_styles', True):
            removed = self._clear_unused_custom_styles(doc)
            if removed:
                self._log(f"清除未使用的自定义样式：{len(removed)} 个（{', '.join(removed)}）。")

        self._log("正在保存最终文档...")
        doc.save(output_path)
