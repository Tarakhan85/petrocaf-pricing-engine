@echo off
cd /d %~dp0\..
call .venv\Scripts\activate
python scripts\build_sqlite.py --master-folder data\master --db-path data\output\petrocaf_master.db
pause
