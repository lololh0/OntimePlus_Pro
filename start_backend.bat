@echo off
cd /d "%~dp0"
pip install -r requirements.txt
python -m uvicorn feedback_backend:app --reload
pause
