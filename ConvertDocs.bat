@echo off

REM path
set root=%USERPROFILE%\miniforge3
call %root%\Scripts\activate.bat %root%
call conda activate base

REM execute
python ConvertDocs.py

REM pause
exit