:: Bertin F. 26-05-2024

@echo off
Title create CTG.exe
::color 17

mkdir %TEMP%\CTG_exe
@echo ceate %TEMP%\CTG_exe successfully

:: create a venv with python 3.9.7
:: adapted from https://stackoverflow.com/questions/45833736/how-to-store-python-version-in-a-variable-inside-bat-file-in-an-easy-way?noredirect=1
for /F "tokens=* USEBACKQ" %%F in (`python --version`) do (set var=%%F)
echo create a virtual environment with %var%
cd %TEMP%\CTG_exe
python -m venv venv

:: activate the venv
set VIRTUAL_ENV=%TEMP%\CTG_exe\venv
call %VIRTUAL_ENV%\Scripts\activate.bat

:: install packages
::pip install git+https://github.com/Bertin-fap/ctg.git#egg=ctg
pip install %userprofile%\pyvenv\ctg
::pip install ctg
pip install auto-py-to-exe

:: set the default directories
:: ICON contains the icon file with the format.ico
:: PGM contain the application lauch python program

set "ICON=%TEMP%/CTG_exe/venv/Lib/site-packages/ctg/ctgfuncts/CTG_RefFiles/logoctg4.ico"
set "DATA=%TEMP%/CTG_exe/venv/Lib/site-packages/ctg;ctg/"
set "PGM=%TEMP%/CTG_exe/venv/Lib/site-packages/ctg/ctgfuncts/CTG_RefFiles/CTG_METER.py"

:: make the executable 
pyinstaller --noconfirm --onefile --console^
 --icon "%ICON%"^
 --add-data "%DATA%"^
 "%PGM%"
 
:: remove the directories build
set "BUILD=%TEMP%\CTG_exe\build"
rmdir /s /q %BUILD%

:: rename the directory dist to aaaa_mm_jj BiblioMeter 
:: adapted from http://stackoverflow.com/a/10945887/1810071
for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
for /f %%x in ('wmic path win32_localtime get /format:list ^| findstr "="') do set %%x
set fmonth=00%Month%
set fday=00%Day%
set dirname="%Year%-%fmonth:~-2%-%fday:~-2% CTG_Meter"
rename dist %dirname%
 
set "new_file_name=%dirname%.exe"
ren %TEMP%\CTG_exe\%dirname%\CTG_METER.exe %new_file_name%
echo %new_file_name% is located in %TEMP%\CTG_exe\%dirname% 

:: Copy Exe
set input_file=%TEMP%\CTG_exe\%dirname%\%dirname%.exe%
set /p "rep=do you want to copy this file in a folder (y/n) : "
if NOT %rep%==y GOTO FIN
set /p "rep1=do you use %userprofile%\CTG (y/n) : "
if NOT %rep%==y GOTO A
copy  %input_file% %userprofile%\CTG
GOTO FIN
A: set /p "new_dir=enter the full path of the folder : "
set output_file=%new_dir%\%dirname%.exe%
copy  %input_file% %output_file%
:FIN

pause