from tkinter import filedialog
import os

INPUT = filedialog.askopenfilename(initialdir="D:\WORK", title="Select file", filetypes=(("PPTX files", "*.pptx"), ("all files", "*.*")))

if INPUT != "":
    COMMAND = f"C:\\Users\\dhkima\\scoop\\apps\\libreoffice\\current\\LibreOffice\\program\\soffice.exe --headless --convert-to pdf {INPUT}"
    print(COMMAND)
    os.system(COMMAND)

