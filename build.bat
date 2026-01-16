@echo off
:: Ensure the virtual environment exists
if not exist venv (
    echo Error: venv not found. Please run setup_offline.bat first.
    pause
    exit /b
)

:: Create the executable
echo Building executable...
venv\Scripts\pyinstaller.exe --onefile --name "AzureTaskMonitor" app.py

echo.
echo -------------------------------------------------------
echo Build complete!
echo Your executable is located in the 'dist' folder:
echo dist\AzureTaskMonitor.exe
echo -------------------------------------------------------
pause