import os
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF

def pdf_to_png():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not file_path:
        return  # User cancelled the dialog

    try:
        pdf_document = fitz.open(file_path)
        
        # Get the directory and base name of the PDF file
        pdf_dir = os.path.dirname(file_path)
        pdf_name_without_ext = os.path.splitext(os.path.basename(file_path))[0]
        
        # Create the output subdirectory
        output_dir = os.path.join(pdf_dir, pdf_name_without_ext)
        os.makedirs(output_dir, exist_ok=True)

        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)  # Load page
            pix = page.get_pixmap()  # Render page to an image (pixmap)
            
            # Define output PNG file path
            output_png_path = os.path.join(output_dir, f"{page_num + 1:03d}.png")
            
            # Save the pixmap as PNG
            pix.save(output_png_path)

        pdf_document.close()
        messagebox.showinfo("Success", f"PDF converted to PNGs successfully!\nImages saved in: {output_dir}")

    except fitz.FileNotFoundError:
        messagebox.showerror("Error", "Selected PDF file not found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    pdf_to_png()