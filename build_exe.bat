@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Creating EXE...
pip install pyinstaller

pyinstaller --onefile --windowed salary_slip_mailer.py

echo EXE Created Successfully!
pause
