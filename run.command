#!/bin/bash
# 定时关机助手 —— macOS 启动脚本（双击本文件即可运行）
cd "$(dirname "$0")" || exit 1

if command -v python3 >/dev/null 2>&1; then
    PY=python3
elif command -v python >/dev/null 2>&1; then
    PY=python
else
    echo "未找到 Python，请先安装：https://www.python.org/downloads/macos/"
    echo "（注意：系统自带的 /usr/bin/python3 不含 tkinter）"
    read -r -p "按回车键退出..."
    exit 1
fi

if ! "$PY" -c "import tkinter" >/dev/null 2>&1; then
    echo "当前 Python 缺少 tkinter，无法显示界面。"
    echo "请安装带 Tk 的 Python：brew install python-tk@3.13  或下载 python.org 官方版本。"
    read -r -p "按回车键退出..."
    exit 1
fi

exec "$PY" shutdown_timer.py
