@echo off
rem Run from source (Python 3.12 required). Normally use ExcelAuto.exe instead.
cd /d "%~dp0"
py -3 -m pip install -q -r requirements.txt
py -3 main.py
