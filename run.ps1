# LOCAL RAB MATHEMATICAL CHECKER
# PowerShell Launch Script

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   LOCAL RAB MATHEMATICAL CHECKER" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$PythonPath = "C:\Users\DESIGN1\AppData\Local\Programs\Python\Python311\python.exe"
$StreamlitPath = "C:\Users\DESIGN1\AppData\Local\Programs\Python\Python311\Scripts\streamlit.exe"

Write-Host "Starting application..." -ForegroundColor Yellow
Write-Host ""
Write-Host "Application will open in your browser at: http://localhost:8501" -ForegroundColor Green
Write-Host ""
Write-Host "Press Ctrl+C to stop the application." -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

& $StreamlitPath run app.py

Read-Host "Press Enter to exit"
