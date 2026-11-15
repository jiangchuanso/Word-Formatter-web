# -*- coding: utf-8 -*-
"""Shared configuration defaults for Word Formatter Pro."""

from wfp_core import BLANK_LINE_MODE_DELETE_SINGLE, DEFAULT_BLANK_LINE_MODE

FONT_SIZE_MAP = {
    '一号 (26pt)': 26, '小一 (24pt)': 24, '二号 (22pt)': 22, '小二 (18pt)': 18,
    '三号 (16pt)': 16, '小三 (15pt)': 15, '四号 (14pt)': 14, '小四 (12pt)': 12,
    '五号 (10.5pt)': 10.5, '小五 (9pt)': 9,
}

DEFAULT_CONFIG = {
    'convert_auto_numbering': True,  # Non-Windows option; Windows uses COM or Python fallback.
    'page_number_align': '奇偶分页', 'footer_distance': 2.5, 'line_spacing': 28,
    'margin_top': 3.7, 'margin_bottom': 3.5, 'margin_left': 2.8, 'margin_right': 2.6,
    'title_font': '方正小标宋简体', 'h1_font': '黑体', 'h2_font': '楷体_GB2312',
    'body_font': '仿宋_GB2312', 'page_number_font': '宋体',
    'table_caption_font': '黑体', 'figure_caption_font': '黑体', 'attachment_font': '黑体',
    'subtitle_font': '楷体_GB2312',
    'title_size': 22, 'h1_size': 16, 'h2_size': 16, 'body_size': 16, 'page_number_size': 14,
    'table_caption_size': 14, 'figure_caption_size': 14, 'attachment_size': 16,
    'subtitle_size': 16, 'title_line_spacing': 33, 'subtitle_line_spacing': 33,
    'line_spacing_unit': 'pt', 'line_spacing_multiple': 1.0,
    'title_line_spacing_unit': 'pt', 'title_line_spacing_multiple': 1.0,
    'subtitle_line_spacing_unit': 'pt', 'subtitle_line_spacing_multiple': 1.0,
    'left_indent_cm': 0.0, 'right_indent_cm': 0.0,
    'paragraph_indent_unit': 'cm',
    'left_indent_chars': 0.0, 'right_indent_chars': 0.0, 'first_line_indent_chars': 2.0,
    'title_bold': False, 'h1_bold': False, 'h2_bold': False,
    'set_outline': True, 'enable_attachment_formatting': True,
    'use_builtin_styles': True,  # 把各级格式写进内置标准样式（正文/标题 1~4）并用样式统一同级文字
    'detect_article_title': True,  # 识别“第X条”为条标题，条后正文另起一段
    'force_a4': False, 'use_custom_english_font': False, 'english_font': 'Times New Roman',
    'blank_line_mode': DEFAULT_BLANK_LINE_MODE, 'normalize_punctuation': False,
    # 可选增强项，默认全部关闭，避免改变既有排版结果
    'convert_floating_objects': False,  # 浮动式图片/嵌入对象/表格改为插入式
    'media_single_line_spacing': False,  # 含图片或嵌入对象的段落改为单倍行距
    'normalize_heading_levels': False,  # 合并相同格式标题等级并补齐跳级
    'detect_chapter_section': False,  # 识别“第X章/第X节”为 1/2 级标题，其后编号标题整体顺延
    'merge_broken_lines': False,  # 合并等长断行（PDF 粘贴导致的换行）
    'enable_table_formatting': False, 'table_header_font': '仿宋_GB2312',
    'table_font': '仿宋_GB2312', 'table_size': 12, 'table_line_spacing': 22,
    'table_line_spacing_unit': 'pt', 'table_line_spacing_multiple': 1.0,
    'table_row_height_cm': 0.7, 'table_auto_col_width': True, 'table_width_percent': 100,
    'table_header_bold': True, 'table_smart_align': False,
    'table_unified_borders': True, 'table_border_size_pt': 0.5,
    'table_col_min_pct': 8, 'table_col_max_pct': 45, 'table_short_text_len': 4,
    'clear_unused_styles': True,  # 排版结束后清除文档中未被引用的自定义（非内置）样式
    'remove_headers_footers': False,  # 排版清洗：删除页眉页脚内容
}

# 功能分类快捷配置：一键套用某一类功能及其推荐参数（Web / GUI 共用）
# 两类开关各司其职：format = 公文格式排版（版式与体例），clean = 排版清洗（净化还原源文档）。
# “识别章/节层级”在两类中都开启：公文排版用于确定层级字体，排版清洗用于辅助判断标题等级。
PRESET_GROUPS = {
    'format': {
        'label': '公文格式排版',
        'description': '页边距与 A4、标题/副标题/层级字体字号行距、大纲层级、附件与表格、标准样式统一',
        'values': {
            'force_a4': True,
            'margin_top': 3.7, 'margin_bottom': 3.5,
            'margin_left': 2.8, 'margin_right': 2.6,
            'footer_distance': 2.5, 'page_number_align': '奇偶分页',
            'title_font': '方正小标宋简体', 'title_size': 22, 'title_line_spacing': 33,
            'subtitle_font': '楷体_GB2312', 'subtitle_size': 16, 'subtitle_line_spacing': 33,
            'h1_font': '黑体', 'h1_size': 16,
            'h2_font': '楷体_GB2312', 'h2_size': 16,
            'body_font': '仿宋_GB2312', 'body_size': 16, 'line_spacing': 28,
            'table_caption_font': '黑体', 'table_caption_size': 14,
            'figure_caption_font': '黑体', 'figure_caption_size': 14,
            'attachment_font': '黑体', 'attachment_size': 16,
            'set_outline': True,
            'use_builtin_styles': True,
            'enable_attachment_formatting': True,
            'enable_table_formatting': True,
            'table_auto_col_width': True,
            'table_unified_borders': True,
            'table_header_bold': True,
            'detect_chapter_section': True,
        },
    },
    'clean': {
        'label': '排版清洗',
        'description': '符号标准化、断行合并、浮动对象、图片行距、标题等级合并、章/节层级辅助判断、清除未使用样式、删除页眉页脚、空行处理',
        'values': {
            'normalize_punctuation': True,
            'merge_broken_lines': True,
            'convert_floating_objects': True,
            'media_single_line_spacing': True,
            'normalize_heading_levels': True,
            'detect_chapter_section': True,  # 清洗时也用章/节层级辅助判断标题等级
            'clear_unused_styles': True,
            'remove_headers_footers': True,  # 清洗时删除页眉页脚
            'blank_line_mode': BLANK_LINE_MODE_DELETE_SINGLE,
        },
    },
}

def apply_preset(config, name):
    """把指定快捷配置组（format / clean）的推荐值合并进 config，返回新字典。"""
    preset = PRESET_GROUPS.get(name)
    if preset is None:
        raise KeyError(f"未知的快捷配置组: {name}")
    merged = dict(config)
    merged.update(preset['values'])
    return merged


PRESET_FONT_OPTIONS = {
    'title': ['方正小标宋简体', '方正小标宋_GBK', '华文中宋'],
    'h1': ['黑体', '方正黑体_GBK', '方正黑体简体', '华文黑体'],
    'h2': ['楷体_GB2312', '方正楷体_GBK', '楷体', '方正楷体简体', '华文楷体'],
    'body': ['仿宋_GB2312', '方正仿宋_GBK', '仿宋', '方正仿宋简体', '华文仿宋'],
    'page_number': ['宋体', 'Times New Roman'],
    'table_caption': ['黑体', '宋体', '仿宋_GB2312'],
    'figure_caption': ['黑体', '宋体', '仿宋_GB2312'],
    'attachment': ['黑体', '宋体', '仿宋_GB2312'],
    'subtitle': ['楷体_GB2312', '方正楷体_GBK', '楷体', '方正楷体简体', '华文楷体'],
    'table': ['仿宋_GB2312', '宋体', '黑体', '楷体_GB2312', '方正仿宋_GBK', '仿宋'],
    'english': ['Times New Roman'],
}
