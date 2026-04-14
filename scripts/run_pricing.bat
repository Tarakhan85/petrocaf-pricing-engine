@echo off
cd /d %~dp0\..
call .venv\Scripts\activate
python scripts\run_pricing.py --config config\settings.json
pause
