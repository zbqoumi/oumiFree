#!/bin/bash
# 欧米freeGPT注册机 启动脚本 (macOS / Linux)
set -e

cd "$(dirname "$0")"

echo "=== 欧米freeGPT注册机 环境检查 ==="

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未找到 python3，请先安装 Python 3.11+"
    exit 1
fi

PYVER=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "Python 版本: $PYVER"

# 创建虚拟环境
if [ ! -d ".venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv .venv
fi

source .venv/bin/activate

# 安装依赖
echo "安装依赖..."
pip install -r requirements.txt

# 安装 Playwright 浏览器
echo "安装 Playwright Chromium..."
python3 -m playwright install chromium

echo ""
echo "=== 启动欧米freeGPT注册机 ==="
python3 start_blackcat.py
