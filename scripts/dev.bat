@echo off
setlocal
cd /d "%~dp0\.."

start "quick-data-vite" cmd /k "cd frontend && npm run dev"
timeout /t 2 /nobreak >nul
.venv312\Scripts\python.exe -m app --dev --debug
