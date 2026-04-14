@echo off
cd /d %~dp0\..
call .venv\Scripts\activate
python scripts\validate_inputs.py --config config\settings.json
pause
