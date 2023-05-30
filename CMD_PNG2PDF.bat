@echo off

REM path
call %userprofile%\scoop\apps\miniconda3\current\Scripts\activate.bat

REM execute
python CMD_PNG2PDF.py

REM pause
exit
