@echo off
setlocal
cd /d "%~dp0\.."

echo [build] compilando frontend...
cd frontend
call npm run build
if errorlevel 1 exit /b 1
cd ..

echo [build] copiando para app\ui ...
if exist app\ui rmdir /s /q app\ui
mkdir app\ui
xcopy /s /e /y frontend\dist\* app\ui\ >nul

echo [build] pronto.
