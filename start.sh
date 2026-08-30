#!/usr/bin/env bash
# =========================================================================
#  Word Formatter Pro - Web 版启动脚本 (Linux / macOS)
#  功能：创建/激活 venv 虚拟环境，安装依赖，启动 FastAPI 网站。
#  原项目文件不会被修改；所有在线功能位于 web/ 目录。
#  所需 LibreOffice（处理 .doc/.wps 旧格式）须装在服务器本机。
# =========================================================================

set -e

# 设置终端窗口标题（兼容 xterm / screen 等常见终端）
printf '\033]0;Word Formatter Pro - Web 服务\007'

# 强制 UTF-8，避免部分环境下读取依赖清单时报错
export PYTHONUTF8=1

# 项目根目录（本文件所在目录）
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$ROOT/.venv"
WEB="$ROOT/web"

# 服务监听端口，可通过环境变量 WFP_PORT 覆盖
: "${WFP_PORT:=4615}"
export WFP_PORT

# 优先使用 python3
PYTHON_BIN="$(command -v python3 || command -v python)"
if [ -z "$PYTHON_BIN" ]; then
    echo "未找到 Python，请先安装 Python 3.9+。"
    exit 1
fi

echo "[1/4] 检查虚拟环境 ..."
if [ ! -d "$VENV" ]; then
    echo "  未检测到虚拟环境，正在创建 venv ..."
    "$PYTHON_BIN" -m venv "$VENV"
fi

echo "[2/4] 激活虚拟环境 ..."
echo "  source .venv/bin/activate 激活虚拟环境，后续 python/pip 均指向 .venv"
# shellcheck disable=SC1091
. "$VENV/bin/activate"

echo "[3/4] 安装/更新依赖..."
python -m pip install -r "$WEB/requirements.txt"

echo "[4/4] 启动网站 ..."
echo "--------------------------------------------------------------"
echo "  访问地址:  http://localhost:${WFP_PORT}   (本机)"
echo "            http://<服务器IP>:${WFP_PORT}   (局域网/其他设备)"
echo "  关闭网站:  在下方窗口按 Ctrl+C"
echo "--------------------------------------------------------------"
exec python "$WEB/app.py"
