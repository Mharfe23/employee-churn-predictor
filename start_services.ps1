Write-Host "Starting Employee Churn Predictor Services..." -ForegroundColor Green
Write-Host ""

# Check if Python is available
try {
    python --version | Out-Null
    Write-Host "✓ Python found" -ForegroundColor Green
} catch {
    Write-Host "✗ Python not found. Please install Python and add it to PATH." -ForegroundColor Red
    exit 1
}

# Check if Node.js is available
try {
    node --version | Out-Null
    Write-Host "✓ Node.js found" -ForegroundColor Green
} catch {
    Write-Host "✗ Node.js not found. Please install Node.js and add it to PATH." -ForegroundColor Red
    exit 1
}

# Check if npm is available
try {
    npm --version | Out-Null
    Write-Host "✓ npm found" -ForegroundColor Green
} catch {
    Write-Host "✗ npm not found. Please install npm." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Starting Flask Server..." -ForegroundColor Yellow
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PSScriptRoot\flask_server'; python app.py" -WindowStyle Normal

Write-Host "Waiting 5 seconds for Flask server to start..." -ForegroundColor Yellow
Start-Sleep -Seconds 5

Write-Host "Starting Excel Add-in Development Server..." -ForegroundColor Yellow
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PSScriptRoot\employe_ml_excel_addin'; npm start" -WindowStyle Normal

Write-Host ""
Write-Host "Services are starting..." -ForegroundColor Green
Write-Host "- Flask Server: http://localhost:5000" -ForegroundColor Cyan
Write-Host "- Excel Add-in: https://localhost:3001" -ForegroundColor Cyan
Write-Host ""
Write-Host "Please wait for both services to fully start before using the Excel add-in." -ForegroundColor Yellow
Write-Host "Press any key to exit this launcher..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
