@echo off
setlocal
set "ROOT=%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed or not in PATH.
  echo Please install Node.js 18 or newer from https://nodejs.org/
  pause
  exit /b 1
)

where pnpm >nul 2>nul
if errorlevel 1 (
  echo pnpm is not installed or not in PATH.
  echo Please run: npm install -g pnpm
  pause
  exit /b 1
)

if not exist "%ROOT%node_modules" (
  echo Installing dependencies...
  cd /d "%ROOT%" && pnpm install
  if errorlevel 1 (
    echo Dependency installation failed.
    pause
    exit /b 1
  )
)

start "AI-Figure-Toolbox" /min cmd /k "cd /d "%ROOT%" && pnpm dev -- --host 127.0.0.1 --port 5173"
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:5173"
