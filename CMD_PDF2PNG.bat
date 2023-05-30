@echo off

REM path
call %userprofile%\scoop\apps\miniconda3\current\Scripts\activate.bat

REM execute
python CMD_PDF2PNG.py

REM pause
exit
