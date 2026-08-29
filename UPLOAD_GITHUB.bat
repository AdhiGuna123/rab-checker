@echo off
title UPLOAD TO GITHUB
color 0F

echo.
echo  ╔═══════════════════════════════════════════════════════════════╗
echo  ║           UPLOAD TO GITHUB - OTOMATIS                       ║
echo  ╚═══════════════════════════════════════════════════════════════╝
echo.

:: Cek Git
git --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Git tidak ditemukan!
    echo  Silakan install Git dari: https://git-scm.com
    pause
    exit /b
)

echo  ═══════════════════════════════════════════════════════════════
echo   PASTIKAN ANDA SUDAH BUAT REPOSITORY DI GITHUB!
echo   Nama Repository: rab-checker
echo  ═══════════════════════════════════════════════════════════════
echo.

set /p GITHUB_USERNAME="Masukkan username GitHub Anda: "
set /p REPO_NAME="Masukkan nama repository (rab-checker): "

if "%REPO_NAME%"=="" set REPO_NAME=rab-checker

echo.
echo  Menginisialisasi Git...
cd /d "%~dp0"
git init
git remote add origin https://github.com/%GITHUB_USERNAME%/%REPO_NAME%.git

echo.
echo  Menambahkan file...
git add app.py
git add excel_reader.py
git add checker.py
git add report.py
git add requirements.txt
git add README.md
git add .gitignore

echo.
echo  Commit changes...
git commit -m "Initial commit - RAB Mathematical Checker"

echo.
echo  Push ke GitHub...
git branch -M main
git push -u origin main

echo.
echo  ╔═══════════════════════════════════════════════════════════════╗
echo  ║  UPLOAD BERHASIL!                                            ║
echo  ║                                                               ║
echo  ║  Langkah selanjutnya:                                         ║
echo  ║  1. Buka: https://share.streamlit.io                         ║
echo  ║  2. Login dengan akun GitHub                                  ║
echo  ║  3. Klik "New app"                                            ║
echo  ║  4. Pilih repository: %REPO_NAME%                             ║
echo  ║  5. Main file: app.py                                         ║
echo  ║  6. Klik "Deploy"                                             ║
echo  ║                                                               ║
echo  ║  Kirim link ke teman Anda!                                    ║
echo  ╚═══════════════════════════════════════════════════════════════╝
echo.

pause
