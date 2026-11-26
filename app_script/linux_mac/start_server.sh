#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$PROJECT_ROOT"

VENV_PYTHON="$PROJECT_ROOT/.venv/bin/python"

if [ ! -x "$VENV_PYTHON" ]; then
  bash "$SCRIPT_DIR/install_requirements.sh"
fi

if [ ! -f "$PROJECT_ROOT/.env" ] && [ -f "$PROJECT_ROOT/.env.example" ]; then
  cp "$PROJECT_ROOT/.env.example" "$PROJECT_ROOT/.env"
  echo "Created .env from .env.example. Remember to set GEMINI_API_KEY in .env!"
fi

echo "Starting Adaptive Batch Image-to-Document Extractor at http://127.0.0.1:8000 ..."
if command -v xdg-open >/dev/null 2>&1; then
  xdg-open "http://127.0.0.1:8000" >/dev/null 2>&1 &
elif command -v open >/dev/null 2>&1; then
  open "http://127.0.0.1:8000" >/dev/null 2>&1 &
fi

exec "$VENV_PYTHON" -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
