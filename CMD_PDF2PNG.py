import fitz  # PyMuPDF
from tkinter import filedialog
from pathlib import Path

INPUT = filedialog.askopenfilename(initialdir="D:\WORK", title="Select file", filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))

if INPUT != "":
    print(f"INPUT = {INPUT}")
    DIR = Path(INPUT).stem
    DIRPATH = Path(INPUT).resolve().parent
    print(f"DIRPATH = {DIRPATH}")
    Path(f"{DIRPATH}\\{DIR}").mkdir(parents=True, exist_ok=True)
    doc = fitz.open(INPUT)
    for i, page in enumerate(doc):
        img = page.get_pixmap(dpi=600)
        img.save(f"{DIRPATH}\\{DIR}\\{i}.png")


