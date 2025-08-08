Write-Host "=== Employee Churn Predictor Troubleshooting ===" -ForegroundColor Green
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if ($isAdmin) {
    Write-Host "✓ Running as Administrator" -ForegroundColor Green
} else {
    Write-Host "✗ NOT running as Administrator - Please run as Administrator" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as administrator'" -ForegroundColor Yellow
    exit 1
}

# Check Node.js
try {
    $nodeVersion = node --version
    Write-Host "✓ Node.js found: $nodeVersion" -ForegroundColor Green
} catch {
    Write-Host "✗ Node.js not found" -ForegroundColor Red
    Write-Host "Download from: https://nodejs.org/" -ForegroundColor Yellow
    exit 1
}

# Check npm
try {
    $npmVersion = npm --version
    Write-Host "✓ npm found: $npmVersion" -ForegroundColor Green
} catch {
    Write-Host "✗ npm not found" -ForegroundColor Red
    exit 1
}

# Check Python
try {
    $pythonVersion = python --version
    Write-Host "✓ Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "✗ Python not found" -ForegroundColor Red
    Write-Host "Download from: https://python.org/" -ForegroundColor Yellow
    exit 1
}

# Check if ports are in use
Write-Host ""
Write-Host "Checking port availability..." -ForegroundColor Yellow

$port3001 = Get-NetTCPConnection -LocalPort 3001 -ErrorAction SilentlyContinue
if ($port3001) {
    Write-Host "✗ Port 3001 is in use by: $($port3001.ProcessName)" -ForegroundColor Red
    Write-Host "Please close the application using port 3001" -ForegroundColor Yellow
} else {
    Write-Host "✓ Port 3001 is available" -ForegroundColor Green
}

$port5000 = Get-NetTCPConnection -LocalPort 5000 -ErrorAction SilentlyContinue
if ($port5000) {
    Write-Host "✗ Port 5000 is in use by: $($port5000.ProcessName)" -ForegroundColor Red
    Write-Host "Please close the application using port 5000" -ForegroundColor Yellow
} else {
    Write-Host "✓ Port 5000 is available" -ForegroundColor Green
}

# Check if directories exist
Write-Host ""
Write-Host "Checking project structure..." -ForegroundColor Yellow

$flaskDir = "flask_server"
$addinDir = "employe_ml_excel_addin"

if (Test-Path $flaskDir) {
    Write-Host "✓ Flask server directory found" -ForegroundColor Green
} else {
    Write-Host "✗ Flask server directory not found" -ForegroundColor Red
}

if (Test-Path $addinDir) {
    Write-Host "✓ Excel add-in directory found" -ForegroundColor Green
} else {
    Write-Host "✗ Excel add-in directory not found" -ForegroundColor Red
}

# Check if node_modules exists
if (Test-Path "$addinDir\node_modules") {
    Write-Host "✓ Node.js dependencies installed" -ForegroundColor Green
} else {
    Write-Host "✗ Node.js dependencies not installed" -ForegroundColor Red
    Write-Host "Run: cd $addinDir && npm install" -ForegroundColor Yellow
}

# Check if virtual environment exists
if (Test-Path "$flaskDir\venv") {
    Write-Host "✓ Python virtual environment found" -ForegroundColor Green
} else {
    Write-Host "✗ Python virtual environment not found" -ForegroundColor Red
    Write-Host "Run: cd $flaskDir && python -m venv venv" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Troubleshooting Complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "If all checks passed, try running:" -ForegroundColor Yellow
Write-Host ".\start_services.ps1" -ForegroundColor Cyan
Write-Host ""
Write-Host "If issues persist, try manual setup:" -ForegroundColor Yellow
Write-Host "1. cd flask_server && python app.py" -ForegroundColor Cyan
Write-Host "2. cd employe_ml_excel_addin && npm start" -ForegroundColor Cyan
