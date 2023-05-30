import img2pdf
from tkinter import filedialog
#from pathlib import Path
import os

def images2PdfFile(src_dir : '', tgt_dir : '', pdfname):
    img_files = [os.path.join(src_dir, nm) for nm in os.listdir(src_dir)]
    with open(f"{tgt_dir}\\{pdfname}.pdf", "wb") as pdf_file:
        pdf_file.write(img2pdf.convert(img_files))

DIR = filedialog.askdirectory(initialdir="D:\WORK", title="Select directory")
print(f"DIR = {DIR}")
pdfname = os.path.basename(os.path.normpath(DIR))
images2PdfFile(DIR, DIR, pdfname)
