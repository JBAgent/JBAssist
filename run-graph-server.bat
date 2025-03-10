@echo off
echo Starting MS Graph MCP Server...

:: Check if Node.js is installed
where node >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Node.js is not installed or not in PATH.
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

:: Check if npm is installed
where npm >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo npm is not installed or not in PATH.
    pause
    exit /b 1
)

:: Install dependencies if needed
if not exist node_modules (
    echo Installing dependencies...
    call npm install
)

:: Build if needed
if not exist build (
    echo Building project...
    call npm run build
)

:: Run the server
call npm start

:: If error occurs, pause
if %ERRORLEVEL% neq 0 (
    echo An error occurred.
    pause
)
