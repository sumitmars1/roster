#!/bin/sh
set -eu
cd "$(dirname "$0")"
PORT="${ROSTER_PORT:-8878}"
if [ -f .server.pid ] && kill -0 "$(cat .server.pid)" 2>/dev/null; then
  echo "已在运行：http://127.0.0.1:$PORT"
  exit 0
fi
if lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
  echo "端口 $PORT 已被占用，请设置 ROSTER_PORT 或先停止占用进程。"
  exit 1
fi
nohup .venv/bin/uvicorn app:app --host 127.0.0.1 --port "$PORT" > server.log 2>&1 &
sleep 2
PID=$(lsof -nP -iTCP:"$PORT" -sTCP:LISTEN -t | head -1)
if [ -z "$PID" ]; then
  echo "启动失败，日志：$(pwd)/server.log"
  tail -20 server.log
  exit 1
fi
echo "$PID" > .server.pid
echo "已启动：http://127.0.0.1:$PORT"
echo "日志：$(pwd)/server.log"
