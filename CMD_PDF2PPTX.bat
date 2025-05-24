@echo off

REM path
set root=%USERPROFILE%\miniforge3
call %root%\Scripts\activate.bat %root%
call conda activate base

REM execute
python CMD_PDF2PPTX.py

REM pause
exit
