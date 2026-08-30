@echo off
title Word Formatter Pro - Web 服务
echo =========================================================================
echo  Word Formatter Pro - Web 版启动脚本 (Windows)
echo  功能：创建/激活 venv 虚拟环境，安装依赖，启动 FastAPI 网站。
echo  原项目文件不会被修改；所有在线功能位于 web/ 目录。
echo  所需 Office / WPS（处理 .doc/.wps）或 LibreOffice（Linux/macOS）须装在服务器本机。
echo =========================================================================
setlocal enabledelayedexpansion
echo 项目根目录（本文件所在目录）
set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"
set "WEB=%ROOT%web"
if not defined WFP_PORT set "WFP_PORT=4615"
echo 端口: %WFP_PORT%（可通过环境变量 WFP_PORT 覆盖）
if not exist "%VENV%" (
    echo [1/4] 未检测到虚拟环境，正在创建 venv ...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo 创建虚拟环境失败，请确认已安装 Python 3.9+。
        pause
        exit /b 1
    )
)
echo [2/4] 激活虚拟环境 ...
echo 调用 .venv\Scripts\activate.bat 激活虚拟环境，后续 python/pip 均指向 .venv
call "%VENV%\Scripts\activate.bat"
if errorlevel 1 (
    echo 虚拟环境激活失败，请检查 .venv 是否完整。
    pause
    exit /b 1
)
echo [3/4] 安装/更新依赖...
python -m pip install -r "%WEB%\requirements.txt"
if errorlevel 1 (
    echo 依赖安装失败，请检查网络或 requirements.txt。
    pause
    exit /b 1
)
echo [4/4] 启动网站 ...
echo --------------------------------------------------------------
echo  访问地址:  http://localhost:%WFP_PORT%   (本机)
echo            http://<服务器IP>:%WFP_PORT%   (局域网/其他设备)
echo  关闭网站:  在下方窗口按 Ctrl+C
echo --------------------------------------------------------------
python "%WEB%\app.py"
if errorlevel 1 (
    echo 网站启动失败，请查看上方错误信息。
    pause
)
endlocal