from tkinter import filedialog
import os

INPUT = filedialog.askopenfilename(initialdir="D:\WORK", title="Select file", filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))

if INPUT != "":
    COMMAND = f"pdf2pptx {INPUT} -r 300"
    print(COMMAND)
    os.system(COMMAND)

