#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$APP_DIR/.venv"
PORT="${PORT:-5000}"
HOST="${HOST:-0.0.0.0}"

cd "$APP_DIR"

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 is required but not installed."
  exit 1
fi

if git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  if git remote get-url origin >/dev/null 2>&1; then
    if [ -n "$(git status --porcelain)" ]; then
      echo "Local changes detected; skipping git pull to avoid conflicts."
      echo "Commit/stash changes or run pull manually, then rerun."
    else
      git pull --rebase origin "$(git rev-parse --abbrev-ref HEAD)"
    fi
  fi
fi

if ! command -v tesseract >/dev/null 2>&1; then
  if command -v apt-get >/dev/null 2>&1; then
    apt-get update
    apt-get install -y tesseract-ocr
  else
    echo "tesseract is not installed. Install it manually and rerun this script."
    exit 1
  fi
fi

if [ ! -d "$VENV_DIR" ]; then
  python3 -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"
pip install --upgrade pip
pip install -r requirements.txt

exec python -m flask --app app run --host "$HOST" --port "$PORT" --no-debugger --no-reload
