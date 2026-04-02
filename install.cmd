@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%.venv"
set "RECREATE=0"
set "PYTHON_EXE="
set "PYTHON_ARGS="

if /I "%~1"=="/?" goto :usage
if /I "%~1"=="-h" goto :usage
if /I "%~1"=="--help" goto :usage

:parse_args
if "%~1"=="" goto :args_done
if /I "%~1"=="--recreate" (
    set "RECREATE=1"
    shift /1
    goto :parse_args
)
if /I "%~1"=="--python" (
    if "%~2"=="" (
        echo Missing value for --python.
        exit /b 1
    )
    set "PYTHON_EXE=%~2"
    set "PYTHON_ARGS="
    shift /1
    shift /1
    goto :parse_args
)
echo Unknown option: %~1
echo.
goto :usage

:args_done
pushd "%SCRIPT_DIR%" >nul 2>&1 || (
    echo Failed to enter project folder:
    echo %SCRIPT_DIR%
    exit /b 1
)

if not defined PYTHON_EXE call :detect_python
if errorlevel 1 goto :fail

call :validate_python "%PYTHON_EXE%" "%PYTHON_ARGS%"
if errorlevel 1 (
    echo Python 3.10 or newer is required.
    goto :fail
)

if "%RECREATE%"=="1" if exist "%VENV_DIR%\Scripts\python.exe" (
    echo Removing existing virtual environment...
    rmdir /s /q "%VENV_DIR%"
    if exist "%VENV_DIR%\Scripts\python.exe" (
        echo Failed to remove existing virtual environment:
        echo %VENV_DIR%
        goto :fail
    )
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Creating virtual environment...
    call :run_python -m venv "%VENV_DIR%"
    if errorlevel 1 goto :fail
) else (
    echo Reusing existing virtual environment:
    echo %VENV_DIR%
)

set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo Virtual environment Python was not found:
    echo %VENV_PY%
    goto :fail
)

echo.
echo Upgrading packaging tools...
"%VENV_PY%" -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto :fail

echo.
echo Installing xl-agent-core in editable mode...
"%VENV_PY%" -m pip install -e .
if errorlevel 1 goto :fail

echo.
echo Verifying installation...
"%VENV_PY%" -m xl_agent_core --help >nul
if errorlevel 1 goto :fail
"%VENV_PY%" -m xl_agent_core.mcp.server --help >nul
if errorlevel 1 goto :fail

echo.
echo Installation completed successfully.
echo Virtual environment:
echo   %VENV_DIR%
echo.
echo Next steps:
echo   %VENV_DIR%\Scripts\activate
echo   xl --help
echo   xl-mcp --help
echo   python -m unittest -q
popd >nul
exit /b 0

:detect_python
where py >nul 2>&1
if not errorlevel 1 (
    call :validate_python "py" "-3"
    if not errorlevel 1 (
        set "PYTHON_EXE=py"
        set "PYTHON_ARGS=-3"
        exit /b 0
    )
)

where python >nul 2>&1
if not errorlevel 1 (
    call :validate_python "python" ""
    if not errorlevel 1 (
        set "PYTHON_EXE=python"
        set "PYTHON_ARGS="
        exit /b 0
    )
)

echo Could not find a usable Python 3 interpreter.
echo Install Python 3.10 or newer, or rerun with:
echo   install.cmd --python "C:\Path\To\python.exe"
exit /b 1

:validate_python
if "%~2"=="" (
    "%~1" -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
) else (
    "%~1" %~2 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
)
exit /b %errorlevel%

:run_python
if defined PYTHON_ARGS (
    "%PYTHON_EXE%" %PYTHON_ARGS% %*
) else (
    "%PYTHON_EXE%" %*
)
exit /b %errorlevel%

:fail
echo.
echo Installation failed.
popd >nul
exit /b 1

:usage
echo Usage:
echo   install.cmd
echo   install.cmd --recreate
echo   install.cmd --python "C:\Path\To\python.exe"
echo.
echo Options:
echo   --recreate   Recreate the .venv folder before installing.
echo   --python     Use a specific Python executable.
exit /b 0
