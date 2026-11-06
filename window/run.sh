#!/bin/bash
cd "$(dirname "$0")" || exit 1

# 启动顺序：便携运行时 → app/.venv → 系统 python3
if [ -x "app/runtime/bin/python3" ]; then
    exec app/runtime/bin/python3 app/main.py
fi
if [ -x "app/.venv/bin/python" ]; then
    exec app/.venv/bin/python app/main.py
fi
for py in python3 python; do
    if command -v "$py" >/dev/null 2>&1; then
        exec "$py" app/main.py
    fi
done

echo "未找到 Python 运行环境。请先安装 Python 3.8+，并在本目录执行："
echo "    pip install -r requirements.txt"
exit 1
