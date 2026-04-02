@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "VENV_PY=%SCRIPT_DIR%.venv\Scripts\python.exe"

if /I "%~1"=="/?" goto :usage
if /I "%~1"=="-h" goto :usage
if /I "%~1"=="--help" goto :usage

pushd "%SCRIPT_DIR%" >nul 2>&1 || (
    echo Failed to enter project folder:
    echo %SCRIPT_DIR%
    exit /b 1
)

if not exist "%VENV_PY%" (
    echo Virtual environment was not found:
    echo   %VENV_PY%
    echo.
    echo Run install.cmd first.
    popd >nul
    exit /b 1
)

if "%~1"=="" (
    echo Starting xl-agent-core...
    "%VENV_PY%" -m xl_agent_core --help
    set "EXIT_CODE=%errorlevel%"
    popd >nul
    exit /b %EXIT_CODE%
)

if /I "%~1"=="mcp" (
    shift /1
    goto :mcp
)
if /I "%~1"=="shell" goto :shell
if /I "%~1"=="test" goto :test

echo Running xl-agent-core %*
"%VENV_PY%" -m xl_agent_core %*
set "EXIT_CODE=%errorlevel%"
popd >nul
exit /b %EXIT_CODE%

:mcp
echo Starting xl-agent-core MCP server...
"%VENV_PY%" -m xl_agent_core.mcp.server %1 %2 %3 %4 %5 %6 %7 %8 %9
set "EXIT_CODE=%errorlevel%"
popd >nul
exit /b %EXIT_CODE%

:shell
if not exist "%SCRIPT_DIR%.venv\Scripts\activate.bat" (
    echo Activation script was not found:
    echo   %SCRIPT_DIR%.venv\Scripts\activate.bat
    popd >nul
    exit /b 1
)
echo Opening a project shell with the virtual environment activated...
call "%SCRIPT_DIR%.venv\Scripts\activate.bat"
cmd /k
set "EXIT_CODE=%errorlevel%"
popd >nul
exit /b %EXIT_CODE%

:test
echo Running smoke tests...
"%VENV_PY%" -m unittest -q
set "EXIT_CODE=%errorlevel%"
popd >nul
exit /b %EXIT_CODE%

:usage
echo Usage:
echo   start.cmd
echo   start.cmd mcp [server arguments]
echo   start.cmd shell
echo   start.cmd test
echo   start.cmd [xl arguments]
echo.
echo Examples:
echo   start.cmd mcp
echo   start.cmd mcp --transport streamable-http --port 8000
echo   start.cmd probe workbook.xlsx
echo   start.cmd read sheets workbook.xlsx
echo   start.cmd verify diff in.xlsx out.xlsx
exit /b 0
