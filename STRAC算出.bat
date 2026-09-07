@echo off
cd /d "%~dp0"
start "" http://localhost:8765/strac_tool.html
python -m http.server 8765
