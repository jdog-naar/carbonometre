#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$ROOT_DIR/carbo"
REQ_FILE="$ROOT_DIR/requirements.txt"
APP_FILE="$ROOT_DIR/app.py"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Erreur: python3 introuvable dans le PATH." >&2
  exit 1
fi

if [ ! -f "$REQ_FILE" ]; then
  echo "Erreur: requirements.txt introuvable dans $ROOT_DIR" >&2
  exit 1
fi

if [ ! -f "$APP_FILE" ]; then
  echo "Erreur: app.py introuvable dans $ROOT_DIR" >&2
  exit 1
fi

if [ ! -d "$VENV_DIR" ]; then
  echo "[launch] Creation du venv: $VENV_DIR"
  python3 -m venv "$VENV_DIR"
fi

# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

echo "[launch] Mise a jour de pip"
python -m pip install -U pip

echo "[launch] Installation des dependances"
python -m pip install -r "$REQ_FILE"

echo "[launch] Demarrage de l'app Streamlit"
echo "[launch] Ouvre: http://localhost:8501"
exec python -m streamlit run "$APP_FILE"
