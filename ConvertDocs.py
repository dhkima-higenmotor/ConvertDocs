import tkinter as tk
import subprocess
import sys
import shutil

def run_pdf_to_png():
    subprocess.run([sys.executable,'CMD_PDF2PNG.py'],capture_output=True,text=True,check=True)

def run_pdf_to_pptx():
    subprocess.run([sys.executable,'CMD_PDF2PPTX.py'],capture_output=True,text=True,check=True)

def run_png_to_pdf():
    subprocess.run([sys.executable,'CMD_PNG2PDF.py'],capture_output=True,text=True,check=True)

def run_pptx_to_pdf():
    subprocess.run([sys.executable,'CMD_PPTX2PDF.py'],capture_output=True,text=True,check=True)

def run_png_to_pptx():
    subprocess.run([sys.executable,'CMD_PNG2PPTX.py'],capture_output=True,text=True,check=True)

def run_pptx_to_png():
    subprocess.run([sys.executable,'CMD_PPTX2PNG.py'],capture_output=True,text=True,check=True)

def run_pdf_arranger():
    dfarranger_path = shutil.which('pdfarranger')
    subprocess.run(dfarranger_path,capture_output=True,text=True,check=True)

# Create the main window
root = tk.Tk()
root.title("ConvertDoc")

# Set font
font_style = ("D2Coding", 12)

# Create and place buttons
btn_pdf_png = tk.Button(root, text="PDF_PNG", font=font_style, command=run_pdf_to_png)
btn_pdf_png.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

btn_pdf_pptx = tk.Button(root, text="PDF_PPTX", font=font_style, command=run_pdf_to_pptx)
btn_pdf_pptx.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

btn_png_pdf = tk.Button(root, text="PNG_PDF", font=font_style, command=run_png_to_pdf)
btn_png_pdf.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

btn_pptx_pdf = tk.Button(root, text="PPTX_PDF", font=font_style, command=run_pptx_to_pdf)
btn_pptx_pdf.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

btn_png_pptx = tk.Button(root, text="PNG_PPTX", font=font_style, command=run_png_to_pptx)
btn_png_pptx.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

btn_pptx_png = tk.Button(root, text="PPTX_PNG", font=font_style, command=run_pptx_to_png)
btn_pptx_png.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

btn_pdf_arranger = tk.Button(root, text="PdfArranger", font=font_style, command=run_pdf_arranger)
btn_pdf_arranger.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

btn_exit = tk.Button(root, text="Exit", font=font_style, command=root.quit)
btn_exit.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

# Configure grid to expand buttons
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Start the GUI event loop
root.mainloop()
