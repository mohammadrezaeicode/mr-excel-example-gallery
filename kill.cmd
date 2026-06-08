@echo off
setlocal enabledelayedexpansion

if "%~1"=="" (
  echo Usage: %~nx0 ^<port^>
  echo Example: %~nx0 5000
  exit /b 1
)

set "PORT=%~1"
set "found=0"

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%PORT%" ^| findstr "LISTENING"') do (
  set "found=1"
  echo Killing PID %%a listening on port %PORT%...
  taskkill /PID %%a /F >nul 2>&1
  if errorlevel 1 (
    echo Failed to kill PID %%a
  ) else (
    echo PID %%a terminated.
  )
)

if "%found%"=="0" (
  echo No process found using port %PORT%.
)

endlocal
