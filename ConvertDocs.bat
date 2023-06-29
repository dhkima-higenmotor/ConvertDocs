@echo off

REM path
call %userprofile%\scoop\apps\miniconda3\current\Scripts\activate.bat

REM execute
python ConvertDocs.py

REM pause
exit
