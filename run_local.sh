#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate

# Fast local run: skip installs unless explicitly requested.
if [ "${ANC_INSTALL_DEPS:-0}" = "1" ]; then
  pip install -r requirements.txt >/dev/null
fi

streamlit run code/geo_placer_web.py
