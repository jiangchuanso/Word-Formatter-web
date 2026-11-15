# Word Formatter Pro · Web 在线版

把本地版排版引擎改造为在线网站。**原项目文件（`wfp_*.py` 等）不做任何修改**，
所有在线功能都放在本 `web/` 目录中，方便继续 `fork` / 同步上游。

客户端通过浏览器上传 `.docx / .doc / .wps / .txt / .md` 文件、或直接输入文本，
并调整页面格式参数（页边距、各级标题字体字号、行距单位、缩进单位、表格样式、附件、全局选项等），
服务端复用同一套引擎 `wfp_core.WordProcessor` 完成排版，返回格式化后的 `.docx`。

> 服务器端需安装 Office / WPS（Windows，处理 `.doc/.wps`）或 LibreOffice（Linux/macOS）。
> 处理旧格式时需本机有对应软件；`.docx / .txt / .md` 不需要 Office。

---

## 目录结构

```
web/
├── app.py            # FastAPI 后端，复用项目根的 wfp_core / wfp_config
├── requirements.txt  # 在线版依赖（fastapi / uvicorn / python-docx / lxml ...）
├── test_app.py       # Web 层回归测试（标准库 unittest，无需额外依赖）
├── static/
│   └── index.html    # 前端单页（上传 + 参数表单 + 进度 + 下载）
└── README.md         # 本文件
```

> 运行后会自动生成 `web/uploads/`、`web/outputs/` 目录存放临时文件。
> 上传副本在处理完成后立即删除；输出文件与任务状态超过 24 小时自动清理。

---

## 快速开始

### Windows

双击项目根目录的 `start.bat`：

1. 自动创建 `venv` 虚拟环境（如不存在）；
2. 检查依赖，**仅当缺失时**才执行 `pip install`（避免每次启动联网等待）；
3. 启动网站（默认 `http://localhost:4618`）。

浏览器打开 `http://localhost:4618` 即可使用。同一局域网内其他设备用
`http://<服务器IP>:4618` 访问。

### Linux / macOS

```bash
chmod +x start.sh
./start.sh
```

### 手动启动（源码环境）

```bash
pip install -r web/requirements.txt
python -m uvicorn web.app:app --host 0.0.0.0 --port 4618
# 或
python web/app.py
```

---

## 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `WFP_HOST` | `0.0.0.0` | 监听地址，设为 `127.0.0.1` 仅本机可访问 |
| `WFP_PORT` | `4618` | 监听端口（`start.bat` / `start.sh` / `app.py` 三处一致） |
| `WFP_MAX_UPLOAD_MB` | `100` | 单个上传文件大小上限（MB） |
| `WFP_MAX_FILES` | `50` | 单次提交的文件数量上限 |
| `WFP_MAX_TEXT_CHARS` | `500000` | 直接输入文本的长度上限（字符） |
| `WFP_JOB_TTL_HOURS` | `24` | 输出文件与任务状态的保留时长 |
| `WFP_CORS_ORIGINS` | 空 | 逗号分隔的跨域白名单；默认不加 CORS 中间件（同源访问） |
| `WFP_SKIP_DEPS` | `0` | 设为 `1` 可让启动脚本完全跳过依赖检查 |

---

## 使用流程

1. 选择 **上传文件** 或 **直接输入文本**（文件可拖拽到页面任意位置；文本模式可自定义输出文件名）。
2. 在「参数设置」中调整参数：分组可点击折叠、支持关键字搜索；字体与字号既可下拉选择，也可手动输入。
3. 点击 **开始排版**：批量任务立即返回任务号，页面轮询显示逐文件进度、状态统计与处理日志，可随时 **取消**。
4. 完成后在「结果」区逐项 **下载** 与原文件同名的 `.docx`（重名自动追加序号，文件名中的空格 / `#` 等特殊字符会正确编码）。
5. 参数可保存到浏览器本地、导出为 JSON、从 JSON 导入，或一键恢复内置默认。

支持**批量上传**：同一批次串行排版并复用同一个 WPS/Word COM 实例，避免并发 COM 冲突和反复启动 Office 进程。

---

## API 速览

| 方法 | 路径 | 说明 |
|------|------|------|
| GET  | `/api/schema`     | 返回字体映射、默认值、单位选项、空行模式、服务端限制，用于渲染表单 |
| POST | `/api/format`     | `files`(多文件) + `config`(JSON 字符串)；创建异步任务，返回 `job_id` 与 `total` |
| GET  | `/api/jobs/{job_id}` | 轮询进度；`after` / `log_after` 返回增量结果与日志，避免重复传输 |
| POST | `/api/jobs/{job_id}/cancel` | 请求取消，当前文件处理完后停止剩余文件 |
| POST | `/api/format-text`| `text` + `config` + 可选 `output_name`；单文档同步排版并返回下载链接 |
| GET  | `/api/download/{job_id}/{filename}` | 下载结果（校验任务 ID 与文件名，防路径穿越） |
| GET  | `/api/health`     | 健康检查与当前活跃任务数 |

`config` 为参数表单序列化后的 JSON 字符串，字段与 `wfp_config.DEFAULT_CONFIG` 一致。
服务端会按默认值类型强制转换（布尔、数值、字号文案、单位、枚举），空值或非法值回退内置默认，
未知字段一律忽略，避免模板文件或手写 JSON 中的 `"false"` 等字符串被当成真值。

---

## 回归测试

```bash
python -m unittest web.test_app -v
# 或
python web/test_app.py
```

测试会自动启动一个本机 uvicorn 子进程，覆盖配置转换、表单字段完整性、
异步任务进度与取消、特殊文件名下载、非法路径与空文件等路径，无需 pytest / httpx。

---

## 复用于本地版引擎

后端 `app.py` 通过 `sys.path` 注入项目根目录，直接 `from wfp_core import WordProcessor`，
因此本地版的每一次改进（识别规则、字体预设等）都会自动在 Web 版生效，无需维护两份逻辑。
