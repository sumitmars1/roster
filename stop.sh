#!/bin/sh
set -eu
cd "$(dirname "$0")"
if [ -f .server.pid ]; then
  PID=$(cat .server.pid)
  kill "$PID" 2>/dev/null || true
  rm -f .server.pid
  echo "已停止 $PID"
else
  echo "没有 .server.pid"
fi
