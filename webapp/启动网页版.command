#!/bin/bash
cd "$(dirname "$0")"
export PATH="$HOME/Library/Python/3.9/bin:/usr/local/bin:$PATH"
pip3 install -r requirements.txt --user -q
echo "启动后请用浏览器打开: http://127.0.0.1:8787"
python3 app.py
