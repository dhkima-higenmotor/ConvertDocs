import img2pdf
from tkinter import filedialog, messagebox
import tkinter as tk
import os
from pathlib import Path

def images_to_pdf(src_dir, tgt_path):
    """Converts all images in the source directory to a single PDF file."""
    # Get all file paths in the source directory and sort them by name
    all_files = [os.path.join(src_dir, fname) for fname in os.listdir(src_dir)]
    all_files.sort() # Sort files by name
    
    # Filter out non-image files
    allowed_extensions = (".png", ".jpg", ".jpeg", ".bmp", ".gif")
    img_files = [f for f in all_files if f.lower().endswith(allowed_extensions)]

    if not img_files:
        messagebox.showwarning("No Images Found", "The selected directory contains no supported image files.")
        return

    try:
        a4inpt = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
        layout_fun = img2pdf.get_layout_fun(a4inpt)
        with open(tgt_path, "wb") as pdf_file:
            # Convert images to PDF with A4 landscape paper size
            pdf_file.write(img2pdf.convert(img_files, playout_fun=layout_fun))
        messagebox.showinfo("Success", f"PDF created successfully at:\n{tgt_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during PDF conversion: {e}")

def main():
    """Main function to handle directory selection and PDF creation."""
    # Set the initial directory to the user's home directory for better cross-platform compatibility
    initial_dir = str(Path.home())
    
    src_dir = filedialog.askdirectory(initialdir=initial_dir, title="Select a directory with images")

    if not src_dir:
        # User cancelled the dialog
        return

    # The PDF will be named after the selected folder and saved in the same location
    pdf_name = f"{os.path.basename(src_dir)}.pdf"
    tgt_path = os.path.join(src_dir, pdf_name)

    images_to_pdf(src_dir, tgt_path)

if __name__ == "__main__":
    main()
