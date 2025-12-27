@echo off
REM Batch file to set up Python environment, install dependencies, and run update_excel.py

REM Check if .venv directory exists
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment. Ensure Python is installed.
        pause
        exit /b 1
    )
)

REM Activate the virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo Failed to activate virtual environment.
    pause
    exit /b 1
)

REM Install requirements
echo Installing requirements...
pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install requirements.
    pause
    exit /b 1
)

REM Run the Python script
echo Running update_excel.py...
python update_excel.py
if errorlevel 1 (
    echo Script execution failed.
    pause
    exit /b 1
)

echo Done.
pause