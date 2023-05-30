from tkinter import filedialog
import os
from pathlib import Path

INPUT = filedialog.askopenfilename(initialdir="D:\WORK", title="Select file", filetypes=(("PPTX files", "*.pptx"), ("all files", "*.*")))
DIRPATH = Path(INPUT).resolve().parent

if INPUT != "":
   userprofile = os.getenv("USERPROFILE")
   COMMAND = f"{userprofile}\\scoop\\apps\\libreoffice\\current\\LibreOffice\\program\\soffice.exe --headless --convert-to pdf {INPUT}"
   os.chdir(DIRPATH)
   os.system(COMMAND)
