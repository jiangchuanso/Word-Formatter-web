# -*- coding: utf-8 -*-
"""Web 版回归测试：标准库 unittest，不依赖 pytest / httpx。

会启动一个本机 uvicorn 子进程，用 urllib 构造 multipart 请求，
覆盖参数转换、表单字段完整性、异步任务进度、下载安全与取消等路径。

运行：
    python -m unittest web.test_app -v
或：
    python web/test_app.py
"""

from __future__ import annotations

import io
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import unittest
import urllib.error
import urllib.request
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "web"))

import app as webapp  # noqa: E402
from wfp_config import DEFAULT_CONFIG  # noqa: E402

BASE = ""
JOB_IDS: list[str] = []
_SERVER = None


def _free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def _request(path_or_url: str, data=None, headers=None, method=None, timeout=60):
    url = path_or_url if path_or_url.startswith("http") else BASE + path_or_url
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, resp.read(), dict(resp.headers)
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read(), dict(exc.headers)


def _json(path_or_url: str, **kwargs):
    status, body, headers = _request(path_or_url, **kwargs)
    try:
        payload = json.loads(body)
    except (ValueError, TypeError):
        payload = None
    return status, payload, headers


def _post_multipart(path, fields, files=()):
    boundary = "----wfp" + uuid.uuid4().hex
    buf = io.BytesIO()
    for key, value in fields.items():
        buf.write(f'--{boundary}\r\nContent-Disposition: form-data; name="{key}"\r\n\r\n'.encode())
        buf.write(str(value).encode("utf-8"))
        buf.write(b"\r\n")
    for name, filename, content in files:
        buf.write(
            f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"; '
            f'filename="{filename}"\r\nContent-Type: application/octet-stream\r\n\r\n'.encode()
        )
        buf.write(content)
        buf.write(b"\r\n")
    buf.write(f"--{boundary}--\r\n".encode())
    return _request(
        path,
        data=buf.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )


def _wait_ready(url: str, timeout: float = 25):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1) as resp:
                if resp.status == 200:
                    return
        except Exception:  # noqa: BLE001
            time.sleep(0.2)
    raise RuntimeError("uvicorn 启动超时")


def _wait_job(job_id: str, timeout: float = 60):
    deadline = time.time() + timeout
    snapshot = None
    while time.time() < deadline:
        status, body, _ = _json(f"/api/jobs/{job_id}?after=0&log_after=0")
        assert status == 200, body
        snapshot = body
        if snapshot["status"] in ("done", "error", "cancelled"):
            return snapshot
        time.sleep(0.2)
    raise AssertionError(f"任务未在 {timeout}s 内结束：{snapshot}")


class ConfigCoercionTests(unittest.TestCase):
    """不依赖服务端的纯函数测试。"""

    def test_bool_strings_and_unknown_keys(self):
        cfg = webapp._coerce_config({
            "force_a4": "false",
            "use_builtin_styles": "true",
            "convert_auto_numbering": "0",
            "set_outline": "nonsense",
            "not_a_real_key": "should-be-ignored",
        })
        self.assertIs(cfg["force_a4"], False)
        self.assertIs(cfg["use_builtin_styles"], True)
        self.assertIs(cfg["convert_auto_numbering"], False)
        self.assertIs(cfg["set_outline"], DEFAULT_CONFIG["set_outline"])
        self.assertNotIn("not_a_real_key", cfg)

    def test_font_size_and_numeric_values(self):
        cfg = webapp._coerce_config({
            "body_size": "三号 (16pt)",
            "title_size": "15.5",
            "page_number_size": "14pt",
            "line_spacing": "28.5",
            "margin_top": "not-a-number",
            "table_border_size_pt": "0.75",
        })
        self.assertEqual(cfg["body_size"], 16)
        self.assertEqual(cfg["title_size"], 15.5)
        self.assertEqual(cfg["page_number_size"], 14)
        self.assertEqual(cfg["line_spacing"], 28.5)
        self.assertEqual(cfg["margin_top"], DEFAULT_CONFIG["margin_top"])
        self.assertEqual(cfg["table_border_size_pt"], 0.75)

    def test_units_and_enums(self):
        cfg = webapp._coerce_config({
            "line_spacing_unit": "倍数",
            "table_line_spacing_unit": "multiple",
            "paragraph_indent_unit": "字符",
            "blank_line_mode": "invalid-mode",
            "page_number_align": "居中",
        })
        self.assertEqual(cfg["line_spacing_unit"], "multiple")
        self.assertEqual(cfg["table_line_spacing_unit"], "multiple")
        self.assertEqual(cfg["paragraph_indent_unit"], "chars")
        self.assertEqual(cfg["blank_line_mode"], DEFAULT_CONFIG["blank_line_mode"])
        self.assertEqual(cfg["page_number_align"], "居中")

    def test_empty_values_fall_back_to_defaults(self):
        cfg = webapp._coerce_config({"body_font": "  ", "body_size": "", "force_a4": True})
        self.assertEqual(cfg["body_font"], DEFAULT_CONFIG["body_font"])
        self.assertEqual(cfg["body_size"], DEFAULT_CONFIG["body_size"])
        self.assertIs(cfg["force_a4"], True)

    def test_safe_names_and_download_url(self):
        self.assertEqual(webapp._safe_output_stem("C:/docs/2026:report?.txt"), "2026_report")
        self.assertEqual(webapp._safe_output_stem("../.."), "document")
        # 路径分量被剥离，只保留最后一段文件名
        self.assertEqual(webapp._safe_output_stem("..\\..\\autorun.docx"), "autorun")
        url = webapp._download_url("a" * 32, "报告 draft #1.docx")
        self.assertNotIn("#", url)
        self.assertNotIn(" ", url)
        self.assertIn("%23", url)

    def test_form_covers_all_default_config_keys(self):
        html = (ROOT / "web" / "static" / "index.html").read_text(encoding="utf-8")
        tokens = set(re.findall(r"'([a-z][a-z0-9_]{2,})'", html))
        missing = sorted(set(DEFAULT_CONFIG) - tokens)
        self.assertFalse(missing, f"表单缺少配置项：{missing}")
        self.assertNotIn("unused_placeholder", html)


class ApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        global BASE, _SERVER
        port = _free_port()
        BASE = f"http://127.0.0.1:{port}"
        env = os.environ.copy()
        env["PYTHONUTF8"] = "1"
        _SERVER = subprocess.Popen(
            [sys.executable, "-m", "uvicorn", "web.app:app", "--host", "127.0.0.1",
             "--port", str(port), "--log-level", "warning"],
            cwd=str(ROOT), env=env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        )
        try:
            _wait_ready(BASE + "/api/health")
        except Exception:
            cls._stop_server()
            raise

    @classmethod
    def _stop_server(cls):
        global _SERVER
        if _SERVER is not None:
            _SERVER.terminate()
            try:
                _SERVER.wait(timeout=5)
            except subprocess.TimeoutExpired:
                _SERVER.kill()
                _SERVER.wait(timeout=5)
            if _SERVER.stdout:
                _SERVER.stdout.close()
            _SERVER = None

    @classmethod
    def tearDownClass(cls):
        cls._stop_server()
        for job_id in JOB_IDS:
            for base in (ROOT / "web" / "outputs", ROOT / "web" / "uploads"):
                shutil.rmtree(base / job_id, ignore_errors=True)

    def test_01_schema_health_and_index(self):
        status, payload, _ = _json("/api/schema")
        self.assertEqual(status, 200)
        self.assertIn("line_spacing_unit", payload["default_config"])
        self.assertIn("max_upload_bytes", payload["limits"])
        self.assertEqual(payload["page_number_align_options"], ["奇偶分页", "居中"])

        status, payload, _ = _json("/api/health")
        self.assertEqual(status, 200)
        self.assertEqual(payload["status"], "ok")

        status, body, _ = _request("/")
        self.assertEqual(status, 200)
        self.assertIn(b'id="runBtn"', body)

    def test_02_format_text_and_download(self):
        text = "在线排版测试\n\n一、第一节\n这是正文内容。\n"
        status, body, _ = _post_multipart(
            "/api/format-text",
            {"config": json.dumps({}), "text": text, "output_name": "排版结果"},
        )
        self.assertEqual(status, 200, body)
        data = json.loads(body)
        JOB_IDS.append(data["job_id"])
        self.assertEqual(len(data["results"]), 1)
        result = data["results"][0]
        self.assertEqual(result["status"], "success")
        self.assertTrue(result["url"].startswith("/api/download/"))

        status, content, headers = _request(result["url"])
        self.assertEqual(status, 200)
        self.assertTrue(content.startswith(b"PK"), "下载内容应为 docx（ZIP）")
        from docx import Document
        doc = Document(io.BytesIO(content))
        self.assertTrue(any("第一节" in p.text for p in doc.paragraphs))

    def test_03_batch_job_progress_and_special_filename(self):
        content = "批量排版测试\n\n（一）小节标题\n正文段落。\n".encode("utf-8")
        status, body, _ = _post_multipart(
            "/api/format",
            {"config": json.dumps({"force_a4": "false", "body_size": "三号 (16pt)"})},
            [
                ("files", "report draft #1.txt", content),
                ("files", "notes.txt", content),
            ],
        )
        self.assertEqual(status, 200, body)
        created = json.loads(body)
        JOB_IDS.append(created["job_id"])
        self.assertEqual(created["total"], 2)

        job = _wait_job(created["job_id"])
        self.assertEqual(job["status"], "done", job.get("error"))
        self.assertEqual(job["processed"], 2)
        self.assertEqual(len(job["results"]), 2)
        self.assertTrue(all(r["status"] == "success" for r in job["results"]), job["results"])

        # 文件名含空格和 # 时，链接必须编码且可正常下载
        special = next(r for r in job["results"] if "report" in r["file"])
        self.assertIn("%23", special["url"])
        status, content, _ = _request(special["url"])
        self.assertEqual(status, 200)
        self.assertTrue(content.startswith(b"PK"))
        self.assertNotIn("/", special["filename"])

        # 增量轮询：已消费的结果与日志不应重复返回
        status, payload, _ = _json(
            f"/api/jobs/{created['job_id']}?after={job['result_count']}&log_after={job['log_count']}"
        )
        self.assertEqual(status, 200)
        self.assertEqual(payload["results"], [])
        self.assertEqual(payload["log_lines"], [])

    def test_04_unsupported_and_empty_upload(self):
        status, body, _ = _post_multipart(
            "/api/format", {"config": "{}"}, [("files", "virus.exe", b"MZ")],
        )
        self.assertEqual(status, 400)
        self.assertIn("没有可处理的文件", json.loads(body)["detail"])

        status, body, _ = _post_multipart(
            "/api/format", {"config": "{}"}, [("files", "empty.txt", b"")],
        )
        self.assertEqual(status, 400)

    def test_05_download_guards(self):
        cases = [
            "/api/download/not-a-job-id/a.docx",
            "/api/download/" + "a" * 32 + "/..docx",
            "/api/download/" + "a" * 32 + "/%2e%2e%2fsecret.docx",
            "/api/download/" + "a" * 32 + "/missing.docx",
        ]
        for path in cases:
            status, _, _ = _request(path)
            self.assertIn(status, (400, 404), path)

    def test_06_cancel_and_invalid_config(self):
        status, body, _ = _post_multipart(
            "/api/format", {"config": "not-json"}, [("files", "a.txt", b"hello")],
        )
        self.assertEqual(status, 400)

        status, body, _ = _post_multipart(
            "/api/format", {"config": "{}"}, [("files", "a.txt", b"hello")],
        )
        self.assertEqual(status, 200, body)
        job_id = json.loads(body)["job_id"]
        JOB_IDS.append(job_id)
        status, payload, _ = _json(f"/api/jobs/{job_id}/cancel", data=b"", method="POST")
        self.assertEqual(status, 200)
        job = _wait_job(job_id)
        self.assertIn(job["status"], ("cancelled", "done"))

        status, _, _ = _json("/api/jobs/" + "0" * 32)
        self.assertEqual(status, 404)
        status, _, _ = _json("/api/jobs/invalid-id")
        self.assertEqual(status, 400)


if __name__ == "__main__":
    unittest.main(verbosity=2)
