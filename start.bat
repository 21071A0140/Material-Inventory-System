@echo off
echo.
echo   +------------------------------------------+
echo   ^|  MATERIAL INVENTORY AUTOMATION SYSTEM    ^|
echo   +------------------------------------------+
echo.

REM Check if venv exists
if not exist venv (
    echo   Setting up virtual environment...
    python -m venv venv
)

REM Activate and install
call venv\Scripts\activate
echo   Installing dependencies...
pip install -r requirements.txt -q

REM Check API key
if "%ANTHROPIC_API_KEY%"=="" (
    echo.
    echo   WARNING: ANTHROPIC_API_KEY not set.
    echo   Set it with: set ANTHROPIC_API_KEY=your_key_here
    echo.
)

echo   Starting server at http://localhost:8000
echo   Open index.html in your browser to use the app.
echo.

REM Open browser
timeout /t 2 >nul
start index.html

REM Start server
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
