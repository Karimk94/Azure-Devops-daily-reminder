@echo off
echo Creating virtual environment...
python -m venv venv

echo Upgrading pip...
venv\Scripts\python.exe -m pip install --upgrade pip

echo Installing requirements...
venv\Scripts\pip.exe install -r requirements.txt

echo Done! You can now activate the environment using: venv\Scripts\activate.bat
pause