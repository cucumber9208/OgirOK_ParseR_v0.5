@echo off
setlocal ENABLEDELAYEDEXPANSION
where python >nul 2>nul
if %errorlevel% neq 0 (
  where py >nul 2>nul && set PY=py -3
) else (
  set PY=python
)
if not exist .venv (
  %PY% -m venv .venv
)
call .venv\Scripts\activate
python -m pip install --upgrade pip >nul
pip install openpyxl python-docx pyyaml >nul
python apps\parser_orders\src\gui\parser_gui.py
