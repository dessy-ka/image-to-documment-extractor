#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$PROJECT_ROOT"

VENV_PYTHON="$PROJECT_ROOT/.venv/bin/python"
REQ_FILE="$PROJECT_ROOT/requirements.txt"

if [ ! -f "$REQ_FILE" ]; then
  echo "Error: requirements.txt not found at $REQ_FILE" >&2
  exit 1
fi

if [ ! -x "$VENV_PYTHON" ]; then
  echo "Creating virtual environment (.venv)..."
  python3 -m venv .venv
fi

echo "Upgrading pip and installing dependencies from requirements.txt..."
"$VENV_PYTHON" -m pip install --upgrade pip
"$VENV_PYTHON" -m pip install -r "$REQ_FILE"

if [ ! -f "$PROJECT_ROOT/.env" ] && [ -f "$PROJECT_ROOT/.env.example" ]; then
  cp "$PROJECT_ROOT/.env.example" "$PROJECT_ROOT/.env"
  echo "Created .env from .env.example. Remember to set GEMINI_API_KEY in .env!"
fi

echo "Dependencies installed successfully."
