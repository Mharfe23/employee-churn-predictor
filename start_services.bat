@echo off
echo Starting Employee Churn Predictor Services...
echo.

echo Starting Flask Server...
start "Flask Server" cmd /k "cd flask_server && python app.py"

echo Waiting 3 seconds for Flask server to start...
timeout /t 3 /nobreak > nul

echo Starting Excel Add-in Development Server...
start "Excel Add-in" cmd /k "cd employe_ml_excel_addin && npm start"

echo.
echo Services are starting...
echo - Flask Server: http://localhost:5000
echo - Excel Add-in: https://localhost:3001
echo.
echo Please wait for both services to fully start before using the Excel add-in.
pause
