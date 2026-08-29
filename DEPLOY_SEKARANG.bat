@echo off
title DEPLOY RAB CHECKER TO GITHUB
color 0F

echo.
echo  ╔═══════════════════════════════════════════════════════════════╗
echo  ║        DEPLOY RAB CHECKER - OTOMATIS                        ║
echo  ╚═══════════════════════════════════════════════════════════════╝
echo.

:: Cek gh
gh --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] GitHub CLI tidak ditemukan!
    echo  Install dari: https://cli.github.com
    pause
    exit /b
)

:: Cek login
gh auth status >nul 2>&1
if errorlevel 1 (
    echo  Anda belum login ke GitHub.
    echo  Menjalankan login...
    echo.
    gh auth login
)

echo.
echo  ═══════════════════════════════════════════════════════════════
echo   Membuat repository dan mengupload file...
echo  ═══════════════════════════════════════════════════════════════
echo.

cd /d "%~dp0"

:: Init git
echo  [1/5] Inisialisasi Git...
git init
git config user.email "user@example.com"
git config user "User"

:: Add files
echo  [2/5] Menambahkan file...
git add app.py excel_reader.py checker.py report.py requirements.txt README.md .gitignore

:: Commit
echo  [3/5] Commit changes...
git commit -m "Initial commit - RAB Mathematical Checker"

:: Create repo
echo  [4/5] Membuat repository...
gh repo create rab-checker --public --source=. --push --description "Aplikasi Pengecek RAB Excel"

:: Push
echo  [5/5] Push ke GitHub...
git branch -M main
git push -u origin main

echo.
echo  ╔═══════════════════════════════════════════════════════════════╗
echo  ║  UPLOAD BERHASIL!                                            ║
echo  ║                                                               ║
echo  ║  Selanjutnya:                                                 ║
echo  ║  1. Buka: https://share.streamlit.io                         ║
echo  ║  2. Login pakai akun GitHub                                   ║
echo  ║  3. Klik "New app"                                            ║
echo  ║  4. Pilih: rab-checker / main / app.py                        ║
echo  ║  5. Klik "Deploy"                                             ║
echo  ║                                                               ║
echo  ║  Kirim link ke teman Anda!                                    ║
echo  ╚═══════════════════════════════════════════════════════════════╝
echo.

pause
