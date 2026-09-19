@echo off
setlocal

py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m pip install pyinstaller==6.22.3

pyinstaller --noconfirm --clean TMS_Report_Generator.spec
if errorlevel 1 exit /b 1

echo.
echo DONE: dist\TMS_Report_Generator\
echo Copy/distribute the whole folder, not only the EXE.
pause
