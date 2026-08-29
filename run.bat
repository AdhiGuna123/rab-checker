@echo off
echo ========================================
echo   LOCAL RAB MATHEMATICAL CHECKER
echo ========================================
echo.
echo Starting application...
echo.
echo Application will open in your browser at: http://localhost:8501
echo.
echo Press Ctrl+C to stop the application.
echo ========================================
echo.

set PYTHON_PATH=C:\Users\DESIGN1\AppData\Local\Programs\Python\Python311\python.exe
set STREAMLIT_PATH=C:\Users\DESIGN1\AppData\Local\Programs\Python\Python311\Scripts\streamlit.exe

"%STREAMLIT_PATH%" run app.py

pause
