#!/usr/bin/env bash
set -e
PY=python3
command -v python3 >/dev/null 2>&1 || PY=python
[ -d ".venv" ] || $PY -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip >/dev/null
pip install openpyxl python-docx pyyaml >/dev/null
python apps/parser_orders/src/gui/parser_gui.py
