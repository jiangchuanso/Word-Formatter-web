# -*- coding: utf-8 -*-
"""Word Formatter Pro - Web 服务端。

本文件不修改原项目的任何文件，仅通过 sys.path 注入项目根目录，
复用 wfp_core.WordProcessor 完成与本地版完全一致的后端排版逻辑。
Office / LibreOffice 安装在服务器端（仅处理 .doc/.wps 旧格式时由 Windows 版 WPS 调用，
或在 Linux/macOS 下由 LibreOffice 调用）。
"""

import asyncio
import logging
import os
import shutil
import sys
import tempfile
import time
import uuid
from pathlib import Path

# ----- 将项目根目录加入 sys.path，以便 import wfp_core / wfp_config -----
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile  # noqa: E402
from fastapi.responses import FileResponse, HTMLResponse  # noqa: E402
from fastapi.staticfiles import StaticFiles  # noqa: E402
from fastapi.middleware.cors import CORSMiddleware  # noqa: E402

from wfp_config import DEFAULT_CONFIG, FONT_SIZE_MAP, PRESET_FONT_OPTIONS  # noqa: E402
from wfp_core import (  # noqa: E402
    BLANK_LINE_MODE_OPTIONS,
    LegacyConversionUnavailable,
    SUPPORTED_FILE_EXTENSIONS,
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


def _cleanup_stale_jobs(max_age_hours: int = 24):
    """清理超过指定时长的临时 job 目录，避免 uploads/outputs 无限累积。"""
    import time

    cutoff = time.time() - max_age_hours * 3600
    for base in (UPLOAD_DIR, OUTPUT_DIR):
        for child in base.iterdir():
            if child.is_dir():
                try:
                    if child.stat().st_mtime < cutoff:
                        shutil.rmtree(child, ignore_errors=True)
                except OSError:
                    pass


app = FastAPI(title=f"{APP_TITLE} Web", version=__version__)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# --------------------------------------------------------------------------- #
# 前端参数说明：把字号映射、字体预设、默认值、空行模式等暴露给前端，用于渲染表单
# --------------------------------------------------------------------------- #
def build_frontend_schema():
    return {
        "app_title": APP_TITLE,
        "version": __version__,
        "font_size_map": FONT_SIZE_MAP,                     # 显示文案 -> 浮点磅值
        "font_size_keys": list(FONT_SIZE_MAP.keys()),
        "preset_font_options": PRESET_FONT_OPTIONS,          # 各分组预设字体列表
        "default_config": DEFAULT_CONFIG,
        "blank_line_mode_options": BLANK_LINE_MODE_OPTIONS,
        "supported_extensions": list(SUPPORTED_FILE_EXTENSIONS),
    }


_LAST_CLEANUP_TS = 0.0
_CLEANUP_INTERVAL_SEC = 3600


@app.get("/api/schema")
async def schema():
    global _LAST_CLEANUP_TS
    now = time.time()
    if now - _LAST_CLEANUP_TS > _CLEANUP_INTERVAL_SEC:
        _LAST_CLEANUP_TS = now
        try:
            _cleanup_stale_jobs()
        except Exception:  # noqa: BLE001
            logger.exception("清理临时 job 目录失败")
    return build_frontend_schema()


# --------------------------------------------------------------------------- #
# 真正执行排版的核心逻辑（与本地版 GUI worker 等价，但运行在线程池里）
# --------------------------------------------------------------------------- #
def _coerce_config(raw: dict) -> dict:
    """把前端提交的字符串表单值转换为 WordProcessor 期望的类型。

    规则参考本地版 wfp_gui.collect_config：
    - *_size 字段：若在字号映射里则用映射的浮点值，否则尝试 float；
    - 含小数点的纯数字 -> float，否则 int；
    - 复选框（开关）已经是 bool。
    """
    cfg = {}
    for key, value in raw.items():
        if key in DEFAULT_CONFIG and isinstance(value, str):
            value = value.strip()
            if value == "":
                cfg[key] = DEFAULT_CONFIG[key]
                continue
        if isinstance(value, bool):
            cfg[key] = value
            continue
        if isinstance(value, str) and "_size" in key:
            if value in FONT_SIZE_MAP:
                cfg[key] = FONT_SIZE_MAP[value]
            else:
                try:
                    cfg[key] = float(value)
                except (ValueError, TypeError):
                    cfg[key] = 16
            continue
        if isinstance(value, str):
            try:
                cfg[key] = float(value) if ("." in value) else int(value)
            except (ValueError, TypeError):
                cfg[key] = value
        else:
            cfg[key] = value
    # 兜底：补齐缺省键
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    return cfg


def _run_format(input_path: str, output_path: str, config: dict):
    """同步执行排版（在后台线程中调用）。"""
    log_lines = []

    def log_cb(msg):
        log_lines.append(str(msg))

    com_initialized = _initialize_com_for_thread(log_cb)
    processor = None
    try:
        with WPSAppManager(log_cb) as com_mgr:
            processor = WordProcessor(config, log_cb, com_manager=com_mgr)
            processor.format_document(input_path, output_path)
    finally:
        # 清理转换/复制过程中产生的系统临时副本，避免临时目录无限累积
        if processor is not None:
            processor._cleanup_temp_files()
        _uninitialize_com_for_thread(com_initialized, log_cb)
    return log_lines


# --------------------------------------------------------------------------- #
# 路由：首页
# --------------------------------------------------------------------------- #
@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = STATIC_DIR / "index.html"
    if not html_path.exists():
        raise HTTPException(status_code=500, detail="前端页面缺失：web/static/index.html")
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


# --------------------------------------------------------------------------- #
# 路由：上传文件并排版（支持单文件与多文件批量）
# --------------------------------------------------------------------------- #
@app.post("/api/format")
async def format_files(
    files: list[UploadFile] = File(...),
    config: str = Form(...),  # JSON 字符串，前端将参数表单序列化后传入
):
    import json

    try:
        raw_config = json.loads(config)
    except Exception:
        raise HTTPException(status_code=400, detail="参数格式错误")

    coerced = _coerce_config(raw_config)

    if not files:
        raise HTTPException(status_code=400, detail="未接收到任何文件")

    job_id = uuid.uuid4().hex
    job_upload_dir = UPLOAD_DIR / job_id
    job_output_dir = OUTPUT_DIR / job_id
    job_upload_dir.mkdir(parents=True, exist_ok=True)
    job_output_dir.mkdir(parents=True, exist_ok=True)

    # 保存上传文件
    saved_paths = []
    skipped = []
    for uf in files:
        ext = os.path.splitext(uf.filename or "")[1].lower()
        if ext not in SUPPORTED_FILE_EXTENSIONS:
            skipped.append(uf.filename or "(未知文件名)")
            continue
        safe_name = f"{uuid.uuid4().hex}{ext}"
        dest = job_upload_dir / safe_name
        data = await uf.read()
        if not data:
            skipped.append(uf.filename or "(空文件)")
            continue
        dest.write_bytes(data)
        saved_paths.append((uf.filename or safe_name, str(dest)))

    if not saved_paths:
        raise HTTPException(
            status_code=400,
            detail="没有可处理的文件（支持的格式："
            + ", ".join(SUPPORTED_FILE_EXTENSIONS)
            + "）",
        )

    loop = asyncio.get_running_loop()
    results = []
    # 串行执行：WPS/Word COM 应用实例与 pythoncom 线程初始化并非线程安全，
    # 并行调用会导致 "应用程序正忙" 或 COM 初始化冲突，故逐个文件处理。
    for original_name, in_path in saved_paths:
        base = os.path.splitext(original_name)[0]
        out_path = job_output_dir / f"{base}_formatted.docx"
        try:
            logs = await loop.run_in_executor(
                None, _run_format, in_path, str(out_path), coerced
            )
            results.append({
                "file": original_name,
                "status": "success",
                "url": f"/api/download/{job_id}/{out_path.name}",
                "log": "\n".join(logs),
            })
        except LegacyConversionUnavailable as e:
            results.append({"file": original_name, "status": "skipped", "reason": str(e)})
        except Exception as e:  # noqa: BLE001
            logger.exception("排版失败: %s", original_name)
            results.append({"file": original_name, "status": "failed", "reason": str(e)})

    return {
        "job_id": job_id,
        "results": results,
        "skipped_unsupported": skipped,
    }


# --------------------------------------------------------------------------- #
# 路由：直接输入文本排版
# --------------------------------------------------------------------------- #
@app.post("/api/format-text")
async def format_text(
    text: str = Form(...),
    config: str = Form(...),
):
    import json

    if not text.strip():
        raise HTTPException(status_code=400, detail="文本内容为空")

    try:
        raw_config = json.loads(config)
    except Exception:
        raise HTTPException(status_code=400, detail="参数格式错误")

    coerced = _coerce_config(raw_config)
    # 直接输入文本：强制 A4（与本地版一致）
    coerced["force_a4"] = True

    job_id = uuid.uuid4().hex
    job_output_dir = OUTPUT_DIR / job_id
    job_output_dir.mkdir(parents=True, exist_ok=True)

    out_path = job_output_dir / "formatted_document.docx"
    loop = asyncio.get_running_loop()
    try:
        logs = await loop.run_in_executor(
            None, _run_format_text, text, str(out_path), coerced
        )
    except Exception as e:  # noqa: BLE001
        logger.exception("文本排版失败")
        raise HTTPException(status_code=500, detail=f"排版失败：{e}")

    return {
        "job_id": job_id,
        "url": f"/api/download/{job_id}/{out_path.name}",
        "log": "\n".join(logs),
    }


def _run_format_text(text: str, output_path: str, config: dict):
    import tempfile

    log_lines = []

    def log_cb(msg):
        log_lines.append(str(msg))

    fd, tmp_txt = tempfile.mkstemp(suffix=".txt", text=True)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(text)

    com_initialized = _initialize_com_for_thread(log_cb)
    processor = None
    try:
        with WPSAppManager(log_cb) as com_mgr:
            processor = WordProcessor(config, log_cb, com_manager=com_mgr)
            processor.format_document(tmp_txt, output_path)
    finally:
        # 清理排版过程产生的系统临时副本
        if processor is not None:
            processor._cleanup_temp_files()
        _uninitialize_com_for_thread(com_initialized, log_cb)
        try:
            os.remove(tmp_txt)
        except OSError:
            pass
    return log_lines


# --------------------------------------------------------------------------- #
# 路由：下载结果
# --------------------------------------------------------------------------- #
@app.get("/api/download/{job_id}/{filename}")
async def download(job_id: str, filename: str):
    # 防止路径穿越
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(status_code=400, detail="非法文件名")
    out_dir = OUTPUT_DIR / job_id
    file_path = out_dir / filename
    if not file_path.exists():
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
    return {"status": "ok", "version": __version__}


# --------------------------------------------------------------------------- #
# 静态资源（如果存在）
# --------------------------------------------------------------------------- #
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("WFP_HOST", "0.0.0.0")
    port = int(os.environ.get("WFP_PORT", "8000"))
    logger.info("启动 %s v%s，监听 http://%s:%d", APP_TITLE, __version__, host, port)
    uvicorn.run(app, host=host, port=port)
