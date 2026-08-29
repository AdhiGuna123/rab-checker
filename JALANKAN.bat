@echo off
title RAB MATHEMATICAL CHECKER
color 0A

echo.
echo  ╔═══════════════════════════════════════════════════════════╗
echo  ║       LOCAL RAB MATHEMATICAL CHECKER                     ║
echo  ║       Aplikasi Pengecek RAB Excel                        ║
echo  ╚═══════════════════════════════════════════════════════════╝
echo.
echo  Menyiapkan aplikasi...
echo.

:: Cek Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python tidak ditemukan!
    echo  Silakan install Python dari: https://python.org
    echo  Centang "Add Python to PATH" saat install.
    echo.
    pause
    exit /b
)

:: Cek dependencies
echo  Memeriksa dependencies...
pip show streamlit >nul 2>&1
if errorlevel 1 (
    echo  Menginstall dependencies...
    pip install -r "%~dp0requirements.txt"
)

echo.
echo  ═══════════════════════════════════════════════════════════
echo   Aplikasi akan terbuka di browser Anda.
echo   Alamat: http://localhost:8501
echo   Tekan Ctrl+C untuk menutup aplikasi.
echo  ═══════════════════════════════════════════════════════════
echo.

:: Jalankan Streamlit
cd /d "%~dp0"
streamlit run app.py

pause
