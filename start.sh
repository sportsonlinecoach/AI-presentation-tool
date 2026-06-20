#!/bin/bash
# AI-presentation-tool サーバー起動スクリプト
# Automator アプリ / 手動実行 どちらからでも使える

cd "$(dirname "$0")"

# すでに起動中なら知らせるだけで二重起動しない
if lsof -ti tcp:5060 >/dev/null 2>&1; then
    osascript -e 'display notification "すでに起動中です (localhost:5060)" with title "AI presentation tool"'
    open "http://localhost:5060"
    exit 0
fi

# 起動（少し待ってからブラウザを開く）
( sleep 3 && open "http://localhost:5060" ) &
.venv/bin/python run.py
