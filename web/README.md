# Word Formatter Pro · Web 在线版

把本地版排版引擎改造为在线网站。**原项目文件（`wfp_*.py` 等）不做任何修改**，
所有在线功能都放在本 `web/` 目录中，方便你继续 `fork` / 同步上游。

客户端通过浏览器上传 `.docx / .doc / .wps / .txt / .md` 文件、或直接输入文本，
并调整页面格式参数（页边距、各级标题字体字号、表格样式、附件、全局选项等），
服务端复用同一套引擎 `wfp_core.WordProcessor` 完成排版，返回格式化后的 `.docx`。

> 服务器端需安装 Office / WPS（Windows，处理 `.doc/.wps`）或 LibreOffice（Linux/macOS）。
> 处理旧格式时需本机有对应软件；`.docx / .txt / .md` 不需要 Office。

---

## 目录结构

```
web/
├── app.py            # FastAPI 后端，复用项目根的 wfp_core / wfp_config
├── requirements.txt  # 在线版依赖（fastapi / uvicorn / python-docx ...）
├── static/
│   └── index.html    # 前端单页（上传 + 参数表单 + 下载）
└── README.md         # 本文件
```

> 运行后会自动生成 `web/uploads/`、`web/outputs/` 目录存放临时文件。

---

## 快速开始

### Windows

双击项目根目录的 `start.bat`：

1. 自动创建 `venv` 虚拟环境（如不存在）；
2. 安装依赖；
3. 启动网站（默认 `http://localhost:8000`）。

浏览器打开 `http://localhost:8000` 即可使用。同一局域网内其他设备用
`http://<服务器IP>:8000` 访问。

### Linux / macOS

```bash
chmod +x start.sh
./start.sh
```

---

## 自定义监听地址 / 端口

默认 `0.0.0.0:8000`。可用环境变量覆盖：

```bash
# Windows (在 start.bat 里 set，或在启动前执行)
set WFP_HOST=127.0.0.1
set WFP_PORT=9000

# Linux / macOS
WFP_HOST=127.0.0.1 WFP_PORT=9000 ./start.sh
```

---

## 使用流程

1. 选择 **上传文件** 或 **直接输入文本**。
2. 在「参数设置」中调整页面格式（所有参数与本地版 GUI 一一对应）。
3. 点击 **开始排版**，右侧实时显示处理日志与进度。
4. 完成后在「结果」区点击 **下载** 获取 `_formatted.docx`。

支持**批量上传**：每个文件独立排版，分别生成下载链接。

---

## API 速览

| 方法 | 路径 | 说明 |
|------|------|------|
| GET  | `/api/schema`     | 返回字体映射、默认值、空行模式等，用于渲染表单 |
| POST | `/api/format`     | `files`(多文件) + `config`(JSON) 批量排版 |
| POST | `/api/format-text`| `text` + `config`(JSON) 文本排版 |
| GET  | `/api/download/{job_id}/{filename}` | 下载结果 |
| GET  | `/api/health`     | 健康检查 |

`config` 为参数表单序列化后的 JSON 字符串，字段与 `wfp_config.DEFAULT_CONFIG` 一致。

---

## 复用于本地版引擎

后端 `app.py` 通过 `sys.path` 注入项目根目录，直接 `from wfp_core import WordProcessor`，
因此本地版的每一次改进（识别规则、字体预设等）都会自动在 Web 版生效，无需维护两份逻辑。
