# -*- coding: utf-8 -*-
"""Word Formatter Pro - Web 服务端。

本文件不修改原项目的任何文件，仅通过 sys.path 注入项目根目录，
复用 wfp_core.WordProcessor 完成与本地版完全一致的后端排版逻辑。
Office / LibreOffice 安装在服务器端（仅处理 .doc/.wps 旧格式时由 Windows 版 WPS 调用，
或在 Linux/macOS 下由 LibreOffice 调用）。

排版任务以「异步 job」执行：上传完成后立即返回 job_id，前端轮询进度；
同一批文件在一个工作线程里串行处理并复用同一个 WPS/Word COM 实例，
避免多请求并发时 COM 冲突，也避免每个文件都重新启动 Office 进程。
"""

from __future__ import annotations

import asyncio
import json
import logging
import math
import os
import re
import shutil
import sys
import tempfile
import threading
import time
import uuid
from contextlib import asynccontextmanager
from pathlib import Path
from urllib.parse import quote

# ----- 将项目根目录加入 sys.path，以便 import wfp_core / wfp_config -----
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile  # noqa: E402
from fastapi.middleware.cors import CORSMiddleware  # noqa: E402
from fastapi.responses import FileResponse, HTMLResponse  # noqa: E402
from fastapi.staticfiles import StaticFiles  # noqa: E402

from wfp_config import (  # noqa: E402
    DEFAULT_CONFIG,
    FONT_SIZE_MAP,
    PRESET_FONT_OPTIONS,
    PRESET_GROUPS,
)
from wfp_core import (  # noqa: E402
    BLANK_LINE_MODE_OPTIONS,
    RE_SAFE_FILENAME_CHARS,
    SUPPORTED_FILE_EXTENSIONS,
    LegacyConversionUnavailable,
    WPSAppManager,
    WordProcessor,
    _initialize_com_for_thread,
    _uninitialize_com_for_thread,
)

from wfp_version import APP_TITLE, __version__  # noqa: E402

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("wfp-web")

STATIC_DIR = Path(__file__).resolve().parent / "static"
UPLOAD_DIR = Path(__file__).resolve().parent / "uploads"
OUTPUT_DIR = Path(__file__).resolve().parent / "outputs"
for d in (UPLOAD_DIR, OUTPUT_DIR):
    d.mkdir(parents=True, exist_ok=True)

# 可通过环境变量调整的服务端限制
MAX_UPLOAD_BYTES = int(os.environ.get("WFP_MAX_UPLOAD_MB", "100")) * 1024 * 1024
MAX_FILES_PER_JOB = int(os.environ.get("WFP_MAX_FILES", "50"))
MAX_TEXT_CHARS = int(os.environ.get("WFP_MAX_TEXT_CHARS", "500000"))
JOB_TTL_SEC = int(os.environ.get("WFP_JOB_TTL_HOURS", "24")) * 3600
CLEANUP_INTERVAL_SEC = 3600

# job 状态保存在进程内；服务重启后 outputs 文件仍可下载，但进度状态会丢失
JOBS: dict[str, dict] = {}
# 用可重入锁：部分状态更新路径会在持锁时追加日志
JOBS_LOCK = threading.RLock()
# WPS/Word COM 与 pythoncom 不能跨线程并发使用，批次之间也必须串行
FORMAT_LOCK = threading.Lock()


class UploadTooLarge(Exception):
    """单个上传文件超过服务端限制。"""


# --------------------------------------------------------------------------- #
# 清理：过期的上传/输出目录与内存中的 job 状态
# --------------------------------------------------------------------------- #
def _cleanup_stale_jobs(max_age_sec: int = JOB_TTL_SEC):
    cutoff = time.time() - max_age_sec
    for base in (UPLOAD_DIR, OUTPUT_DIR):
        for child in base.iterdir():
            try:
                if child.is_dir() and child.stat().st_mtime < cutoff:
                    shutil.rmtree(child, ignore_errors=True)
            except OSError:
                pass
    with JOBS_LOCK:
        stale = [
            job_id for job_id, job in JOBS.items()
            if job["created"] < cutoff and job["status"] in ("done", "error", "cancelled")
        ]
        for job_id in stale:
            JOBS.pop(job_id, None)


async def _cleanup_loop():
    while True:
        await asyncio.sleep(CLEANUP_INTERVAL_SEC)
        try:
            await asyncio.to_thread(_cleanup_stale_jobs)
        except Exception:  # noqa: BLE001
            logger.exception("清理临时 job 目录失败")


@asynccontextmanager
async def lifespan(_app: FastAPI):
    task = asyncio.create_task(_cleanup_loop())
    try:
        yield
    finally:
        task.cancel()
        try:
            await task
        except asyncio.CancelledError:
            pass


app = FastAPI(title=f"{APP_TITLE} Web", version=__version__, lifespan=lifespan)

# 默认同源访问即可；确需跨域时用 WFP_CORS_ORIGINS 显式放行（逗号分隔）
_cors_origins = [o.strip() for o in os.environ.get("WFP_CORS_ORIGINS", "").split(",") if o.strip()]
if _cors_origins:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=_cors_origins,
        allow_methods=["*"],
        allow_headers=["*"],
    )


# --------------------------------------------------------------------------- #
# 前端参数说明：字号映射、字体预设、默认值、空行模式、服务端限制等
# --------------------------------------------------------------------------- #
def build_frontend_schema():
    return {
        "app_title": APP_TITLE,
        "version": __version__,
        "font_size_map": FONT_SIZE_MAP,                     # 显示文案 -> 浮点磅值
        "font_size_keys": list(FONT_SIZE_MAP.keys()),
        "preset_font_options": PRESET_FONT_OPTIONS,          # 各分组预设字体列表
        "preset_groups": PRESET_GROUPS,                      # 快捷配置：公文格式排版 / 排版清洗
        "default_config": DEFAULT_CONFIG,
        "blank_line_mode_options": BLANK_LINE_MODE_OPTIONS,
        "page_number_align_options": ["奇偶分页", "居中"],
        "supported_extensions": list(SUPPORTED_FILE_EXTENSIONS),
        "units": {
            "spacing": [{"value": "pt", "label": "磅值"}, {"value": "multiple", "label": "倍数"}],
            "indent": [{"value": "cm", "label": "厘米"}, {"value": "chars", "label": "字符"}],
        },
        "limits": {
            "max_upload_bytes": MAX_UPLOAD_BYTES,
            "max_files": MAX_FILES_PER_JOB,
            "max_text_chars": MAX_TEXT_CHARS,
        },
    }


@app.get("/api/schema")
async def schema():
    return build_frontend_schema()


# --------------------------------------------------------------------------- #
# 配置转换：严格按 DEFAULT_CONFIG 的类型转换，避免表单字符串变成“真值”或错类型
# --------------------------------------------------------------------------- #
BOOL_TRUE = {"true", "1", "yes", "on", "是", "开"}
BOOL_FALSE = {"false", "0", "no", "off", "否", "关"}
SIZE_KEYS = {k for k in DEFAULT_CONFIG if k.endswith("_size")}
SPACING_UNIT_KEYS = {"line_spacing_unit"} | {
    k for k in DEFAULT_CONFIG if k.endswith("_line_spacing_unit")
}
INDENT_UNIT_KEYS = {"paragraph_indent_unit"}
ENUM_CHOICES = {
    "page_number_align": {"奇偶分页", "居中"},
    "blank_line_mode": set(BLANK_LINE_MODE_OPTIONS),
}


def _as_bool(value, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in BOOL_TRUE:
        return True
    if text in BOOL_FALSE:
        return False
    return default


def _as_number(value, default):
    if isinstance(value, bool):
        return default
    try:
        number = float(value)
    except (TypeError, ValueError):
        return default
    return number if math.isfinite(number) else default


def _as_font_size(value, default):
    if isinstance(value, str):
        text = value.strip()
        if text in FONT_SIZE_MAP:
            return float(FONT_SIZE_MAP[text])
        # 允许 "16.5" / "16pt" / "三号 (16pt)" 这类输入
        match = re.search(r"(\d+(?:\.\d+)?)\s*pt\)\s*$", text, re.IGNORECASE)
        if match:
            value = match.group(1)
        else:
            match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*pt", text, re.IGNORECASE)
            if match:
                value = match.group(1)
    size = _as_number(value, default)
    return size if 0 < size <= 200 else default


def _normalize_unit(value, kind: str) -> str:
    text = str(value or "").strip().lower()
    if kind == "spacing":
        return "multiple" if text in {"multiple", "倍数", "倍", "multi"} else "pt"
    return "chars" if text in {"chars", "char", "character", "characters", "字符"} else "cm"


def _coerce_config(raw: dict) -> dict:
    """把前端提交的表单值转换为 WordProcessor 期望的类型。

    - 只接受 DEFAULT_CONFIG 中存在的键，忽略模板里的未知字段；
    - 按默认值类型强制转换（bool / 数字 / 字符串），空值或非法值回退内置默认；
    - 字号支持「三号 (16pt)」文案、纯数字与 "16pt"；
    - 单位与枚举做白名单归一，避免拼写错误静默生效。
    """
    cfg = dict(DEFAULT_CONFIG)
    if not isinstance(raw, dict):
        return cfg
    for key, default in DEFAULT_CONFIG.items():
        if key not in raw:
            continue
        value = raw[key]
        if isinstance(value, str):
            value = value.strip()
        if value is None or value == "":
            continue
        if isinstance(default, bool):
            cfg[key] = _as_bool(value, default)
        elif key in SIZE_KEYS:
            cfg[key] = _as_font_size(value, default)
        elif isinstance(default, (int, float)):
            cfg[key] = _as_number(value, default)
        elif key in SPACING_UNIT_KEYS:
            cfg[key] = _normalize_unit(value, "spacing")
        elif key in INDENT_UNIT_KEYS:
            cfg[key] = _normalize_unit(value, "indent")
        elif key in ENUM_CHOICES:
            cfg[key] = str(value) if str(value) in ENUM_CHOICES[key] else default
        else:
            cfg[key] = str(value)
    return cfg


def _parse_config_form(config: str) -> dict:
    try:
        raw = json.loads(config)
    except (ValueError, TypeError):
        raise HTTPException(status_code=400, detail="参数格式错误：config 不是合法 JSON")
    if not isinstance(raw, dict):
        raise HTTPException(status_code=400, detail="参数格式错误：config 应为对象")
    return raw


# --------------------------------------------------------------------------- #
# 文件名与下载链接
# --------------------------------------------------------------------------- #
def _safe_output_stem(original_name: str) -> str:
    """去掉 multipart 可能携带的路径、Windows 非法字符与控制字符。"""
    name = (original_name or "").replace("\\", "/").split("/")[-1]
    stem = os.path.splitext(name)[0]
    stem = RE_SAFE_FILENAME_CHARS.sub("_", stem).strip(" ._")
    stem = stem.strip(".")
    return (stem or "document")[:120]


def _unique_output_path(out_dir: Path, stem: str) -> Path:
    candidate = out_dir / f"{stem}.docx"
    counter = 2
    while candidate.exists():
        candidate = out_dir / f"{stem}_{counter}.docx"
        counter += 1
    return candidate


def _download_url(job_id: str, filename: str) -> str:
    # 中文/空格/# 等字符必须编码，否则下载链接会被浏览器截断
    return f"/api/download/{job_id}/{quote(filename, safe='')}"


# --------------------------------------------------------------------------- #
# job 状态读写
# --------------------------------------------------------------------------- #
def _new_job(job_id: str, total: int, skipped: list[str]):
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "pending",
            "cancelled": False,
            "created": time.time(),
            "total": total,
            "processed": 0,
            "current": "",
            "results": [],
            "log": [],
            "skipped_unsupported": skipped,
            "error": None,
        }


def _job_update(job_id: str, **fields):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job:
            job.update(fields)


def _job_log(job_id: str, message: str):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job and len(job["log"]) < 2000:
            job["log"].append(str(message))


def _job_append_result(job_id: str, result: dict, processed: int):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job:
            job["results"].append(result)
            job["processed"] = processed
            job["current"] = ""


def _job_cancelled(job_id: str) -> bool:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        return bool(job and job["cancelled"])


def _job_snapshot(job_id: str, after: int = 0, log_after: int = 0):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return None
        return {
            "job_id": job_id,
            "status": job["status"],
            "processed": job["processed"],
            "total": job["total"],
            "current": job["current"],
            "results": [dict(r) for r in job["results"][max(0, after):]],
            "result_count": len(job["results"]),
            "log_lines": job["log"][max(0, log_after):],
            "log_count": len(job["log"]),
            "skipped_unsupported": list(job["skipped_unsupported"]),
            "error": job["error"],
        }


# --------------------------------------------------------------------------- #
# 排版执行：一个批次只初始化一次 COM、只启动一次 WPS/Word
# --------------------------------------------------------------------------- #
def _format_document(input_path: str, output_path: str, config: dict, com_mgr, log_cb):
    processor = WordProcessor(config, log_cb, com_manager=com_mgr)
    try:
        processor.format_document(input_path, output_path)
    finally:
        # 清理转换/复制过程中产生的系统临时副本，避免临时目录无限累积
        processor._cleanup_temp_files()


def _format_one(job_id: str, original_name: str, in_path: str, out_dir: Path,
                config: dict, com_mgr) -> dict:
    out_path = _unique_output_path(out_dir, _safe_output_stem(original_name))
    log_lines: list[str] = []

    def log_cb(msg):
        if len(log_lines) < 2000:
            log_lines.append(str(msg))

    base = {"file": original_name}
    try:
        _format_document(in_path, str(out_path), config, com_mgr, log_cb)
        return {
            **base,
            "status": "success",
            "filename": out_path.name,
            "url": _download_url(job_id, out_path.name),
            "log": "\n".join(log_lines),
        }
    except LegacyConversionUnavailable as exc:
        return {**base, "status": "skipped", "reason": str(exc), "log": "\n".join(log_lines)}
    except Exception as exc:  # noqa: BLE001
        logger.exception("排版失败: %s", original_name)
        # 失败时可能留下未写完的 .docx，不能让它出现在结果目录里
        try:
            out_path.unlink(missing_ok=True)
        except OSError:
            pass
        return {
            **base,
            "status": "failed",
            "reason": str(exc),
            "log": "\n".join(log_lines) + (f"\n❌ {exc}" if log_lines else f"❌ {exc}"),
        }


def _process_job(job_id: str, upload_dir: Path, out_dir: Path,
                 saved_paths: list[tuple[str, str]], config: dict):
    """在线程池的单个工作线程里串行处理整批文件。"""
    try:
        with FORMAT_LOCK:
            if _job_cancelled(job_id):
                return
            _job_update(job_id, status="running")
            job_log = lambda msg: _job_log(job_id, msg)  # noqa: E731
            com_initialized = _initialize_com_for_thread(job_log)
            try:
                with WPSAppManager(job_log) as com_mgr:
                    for index, (original_name, in_path) in enumerate(saved_paths, start=1):
                        if _job_cancelled(job_id):
                            _job_log(job_id, "任务已取消，停止处理剩余文件。")
                            break
                        _job_update(job_id, current=original_name)
                        _job_log(job_id, f"[{index}/{len(saved_paths)}] 开始处理：{original_name}")
                        result = _format_one(job_id, original_name, in_path, out_dir, config, com_mgr)
                        _job_append_result(job_id, result, index)
            finally:
                _uninitialize_com_for_thread(com_initialized, job_log)
        _job_update(job_id, current="")
    except Exception as exc:  # noqa: BLE001
        logger.exception("批处理任务异常: %s", job_id)
        _job_update(job_id, status="error", error=str(exc))
        _job_log(job_id, f"❌ 任务异常：{exc}")
    finally:
        # 原始上传副本只在本轮排版中使用，处理完立即释放磁盘
        shutil.rmtree(upload_dir, ignore_errors=True)
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if job:
                if job["cancelled"]:
                    job["status"] = "cancelled"
                elif job["status"] in ("pending", "running"):
                    job["status"] = "done"
                job["current"] = ""


def _job_future_callback(future):
    exc = future.exception()
    if exc:
        logger.error("排版后台任务异常终止: %s", exc)


# --------------------------------------------------------------------------- #
# 路由：首页（按 mtime 缓存，避免每次请求都读盘）
# --------------------------------------------------------------------------- #
_INDEX_CACHE = {"mtime": -1.0, "text": ""}


@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = STATIC_DIR / "index.html"
    if not html_path.exists():
        raise HTTPException(status_code=500, detail="前端页面缺失：web/static/index.html")
    try:
        mtime = html_path.stat().st_mtime
        if _INDEX_CACHE["mtime"] != mtime:
            _INDEX_CACHE["text"] = html_path.read_text(encoding="utf-8")
            _INDEX_CACHE["mtime"] = mtime
    except OSError as exc:
        raise HTTPException(status_code=500, detail=f"前端页面读取失败：{exc}")
    return HTMLResponse(_INDEX_CACHE["text"], headers={"Cache-Control": "no-cache"})


# --------------------------------------------------------------------------- #
# 上传保存：分块写盘，避免大文件整体驻留内存
# --------------------------------------------------------------------------- #
async def _write_upload(uf: UploadFile, dest: Path) -> int:
    total = 0
    with dest.open("wb") as handle:
        while True:
            chunk = await uf.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_UPLOAD_BYTES:
                raise UploadTooLarge()
            handle.write(chunk)
    return total


async def _save_uploads(files: list[UploadFile], upload_dir: Path):
    saved: list[tuple[str, str]] = []
    skipped: list[str] = []
    for uf in files:
        original_name = uf.filename or "unnamed"
        ext = os.path.splitext(original_name)[1].lower()
        if ext not in SUPPORTED_FILE_EXTENSIONS:
            skipped.append(original_name)
            continue
        dest = upload_dir / f"{uuid.uuid4().hex}{ext}"
        try:
            size = await _write_upload(uf, dest)
        except UploadTooLarge:
            dest.unlink(missing_ok=True)
            skipped.append(f"{original_name}（超过单文件 {MAX_UPLOAD_BYTES // (1024 * 1024)}MB 限制）")
        except OSError as exc:
            dest.unlink(missing_ok=True)
            skipped.append(f"{original_name}（写入失败：{exc}）")
        else:
            if size == 0:
                dest.unlink(missing_ok=True)
                skipped.append(f"{original_name}（空文件）")
            else:
                saved.append((original_name, str(dest)))
        finally:
            await uf.close()
    return saved, skipped


@app.post("/api/format")
async def format_files(
    files: list[UploadFile] = File(...),
    config: str = Form(...),
):
    """创建异步排版任务，立即返回 job_id；用 /api/jobs/{id} 轮询进度。"""
    coerced = _coerce_config(_parse_config_form(config))
    if not files:
        raise HTTPException(status_code=400, detail="未接收到任何文件")
    if len(files) > MAX_FILES_PER_JOB:
        raise HTTPException(
            status_code=400,
            detail=f"一次最多处理 {MAX_FILES_PER_JOB} 个文件，请分批提交",
        )

    job_id = uuid.uuid4().hex
    upload_dir = UPLOAD_DIR / job_id
    out_dir = OUTPUT_DIR / job_id
    upload_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    saved_paths, skipped = await _save_uploads(files, upload_dir)
    if not saved_paths:
        shutil.rmtree(upload_dir, ignore_errors=True)
        shutil.rmtree(out_dir, ignore_errors=True)
        detail = "没有可处理的文件（支持的格式：" + ", ".join(SUPPORTED_FILE_EXTENSIONS) + "）"
        if skipped:
            detail += "；已跳过：" + "、".join(skipped)
        raise HTTPException(status_code=400, detail=detail)

    _new_job(job_id, len(saved_paths), skipped)
    loop = asyncio.get_running_loop()
    future = loop.run_in_executor(None, _process_job, job_id, upload_dir, out_dir, saved_paths, coerced)
    future.add_done_callback(_job_future_callback)
    return {"job_id": job_id, "total": len(saved_paths), "skipped_unsupported": skipped}


@app.get("/api/jobs/{job_id}")
async def get_job(job_id: str, after: int = 0, log_after: int = 0):
    """轮询任务进度。after / log_after 只返回增量结果与日志，减少轮询流量。"""
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        raise HTTPException(status_code=400, detail="非法任务 ID")
    snapshot = _job_snapshot(job_id, after, log_after)
    if snapshot is None:
        raise HTTPException(status_code=404, detail="任务不存在或状态已过期（服务可能已重启）")
    return snapshot


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str):
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        raise HTTPException(status_code=400, detail="非法任务 ID")
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="任务不存在或状态已过期")
        if job["status"] in ("done", "error", "cancelled"):
            return {"cancelled": False, "status": job["status"]}
        job["cancelled"] = True
        status = job["status"]
    _job_log(job_id, "收到取消请求，将在当前文件处理完成后停止。")
    return {"cancelled": True, "status": status}


# --------------------------------------------------------------------------- #
# 路由：直接输入文本排版（单个文档，保持同步返回）
# --------------------------------------------------------------------------- #
def _run_format_text(text: str, output_path: str, config: dict):
    log_lines: list[str] = []

    def log_cb(msg):
        if len(log_lines) < 2000:
            log_lines.append(str(msg))

    with tempfile.NamedTemporaryFile("w", suffix=".txt", encoding="utf-8", delete=False) as handle:
        handle.write(text)
        tmp_txt = handle.name

    com_initialized = _initialize_com_for_thread(log_cb)
    try:
        with WPSAppManager(log_cb) as com_mgr:
            _format_document(tmp_txt, output_path, config, com_mgr, log_cb)
    finally:
        _uninitialize_com_for_thread(com_initialized, log_cb)
        try:
            os.remove(tmp_txt)
        except OSError:
            pass
    return "\n".join(log_lines)


@app.post("/api/format-text")
async def format_text(
    text: str = Form(...),
    config: str = Form(...),
    output_name: str = Form("排版结果"),
):
    if not text.strip():
        raise HTTPException(status_code=400, detail="文本内容为空")
    if len(text) > MAX_TEXT_CHARS:
        raise HTTPException(
            status_code=400,
            detail=f"文本过长（上限 {MAX_TEXT_CHARS} 字符），请缩短后重试",
        )
    coerced = _coerce_config(_parse_config_form(config))
    # 直接输入文本：强制 A4（与本地版一致）
    coerced["force_a4"] = True

    job_id = uuid.uuid4().hex
    out_dir = OUTPUT_DIR / job_id
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = _unique_output_path(out_dir, _safe_output_stem(output_name))
    filename = out_path.name

    loop = asyncio.get_running_loop()
    try:
        log = await loop.run_in_executor(
            None, _run_format_text, text, str(out_path), coerced
        )
    except Exception as exc:  # noqa: BLE001
        logger.exception("文本排版失败")
        shutil.rmtree(out_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"排版失败：{exc}")

    return {
        "job_id": job_id,
        "log": log,
        "results": [{
            "file": filename,
            "filename": out_path.name,
            "status": "success",
            "url": _download_url(job_id, out_path.name),
            "log": log,
        }],
    }


# --------------------------------------------------------------------------- #
# 路由：下载结果
# --------------------------------------------------------------------------- #
@app.get("/api/download/{job_id}/{filename}")
async def download(job_id: str, filename: str):
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        raise HTTPException(status_code=400, detail="非法任务 ID")
    if (
        not filename
        or filename != os.path.basename(filename)
        or ".." in filename
        or RE_SAFE_FILENAME_CHARS.search(filename)
        or not filename.lower().endswith(".docx")
    ):
        raise HTTPException(status_code=400, detail="非法文件名")
    file_path = OUTPUT_DIR / job_id / filename
    if not file_path.is_file():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    return FileResponse(
        str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# --------------------------------------------------------------------------- #
# 路由：健康检查
# --------------------------------------------------------------------------- #
@app.get("/api/health")
async def health():
    with JOBS_LOCK:
        active = sum(1 for job in JOBS.values() if job["status"] in ("pending", "running"))
    return {"status": "ok", "version": __version__, "active_jobs": active}


# --------------------------------------------------------------------------- #
# 静态资源
# --------------------------------------------------------------------------- #
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("WFP_HOST", "0.0.0.0")
    # 与 start.bat / start.sh 的默认端口保持一致
    port = int(os.environ.get("WFP_PORT", "4618"))
    logger.info("启动 %s v%s，监听 http://%s:%d", APP_TITLE, __version__, host, port)
    uvicorn.run(app, host=host, port=port)
