import fitz  # PyMuPDF
from tkinter import filedialog
from pathlib import Path

INPUT = filedialog.askopenfilename(initialdir="D:\WORK", title="Select file", filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))

if INPUT != "":
    print(f"INPUT = {INPUT}")
    DIR = Path(INPUT).stem
    Path(f".\\{DIR}").mkdir(parents=True, exist_ok=True)
    doc = fitz.open(INPUT)
    for i, page in enumerate(doc):
        img = page.get_pixmap(dpi=600)
        img.save(f".\\{DIR}\\{i}.png")


