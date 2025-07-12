import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

def find_libreoffice_path():
    """Attempts to find the LibreOffice executable (soffice) path."""
    # Common paths for Windows
    if sys.platform == "win32":
        program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
        userprofile = os.environ.get("userprofile")
        
        windows_paths = [
            Path(program_files) / "LibreOffice" / "program" / "soffice.exe",
            Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe",
            Path(program_files) / "LibreOffice 7" / "program" / "soffice.exe", # Specific version
            Path(program_files_x86) / "LibreOffice 7" / "program" / "soffice.exe",
            Path(program_files) / "LibreOffice 6" / "program" / "soffice.exe", # Specific version
            Path(program_files_x86) / "LibreOffice 6" / "program" / "soffice.exe",
            Path(userprofile) / "scoop" / "apps" / "libreoffice" / "current" / "LibreOffice" / "program" / "soffice.exe",
        ]
        for path in windows_paths:
            if path.is_file():
                return str(path)

    # Common paths for Linux
    elif sys.platform.startswith("linux") or sys.platform == "darwin": # macOS also uses similar paths
        linux_paths = [
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/opt/libreoffice/program/soffice",
        ]
        for path in linux_paths:
            if os.path.isfile(path) and os.access(path, os.X_OK):
                return path
    
    # Check if soffice is in PATH
    try:
        # This will raise FileNotFoundError if not in PATH
        return subprocess.check_output(["which", "soffice"]).decode("utf-8").strip()
    except (subprocess.CalledProcessError, FileNotFoundError):
        pass

    return None

def pptx_to_pdf():
    root = tk.Tk()
    root.withdraw() # Hide the main window

    libreoffice_path = find_libreoffice_path()
    if not libreoffice_path:
        messagebox.showerror("Error", "LibreOffice executable (soffice) not found.\nPlease ensure LibreOffice is installed and accessible in your system's PATH, or check common installation directories.")
        return

    file_path = filedialog.askopenfilename(
        title="Select a PPTX file",
        filetypes=[("PowerPoint files", "*.pptx")]
    )

    if not file_path:
        return # User cancelled the dialog

    try:
        input_file_path = Path(file_path)
        output_dir = input_file_path.parent
        output_file_name = input_file_path.stem + ".pdf"
        output_pdf_path = output_dir / output_file_name

        # LibreOffice command to convert PPTX to PDF
        # --headless: run without GUI
        # --convert-to pdf: specify output format
        # --outdir: specify output directory
        command = [
            libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            str(input_file_path),
            "--outdir", str(output_dir)
        ]

        # Execute the command
        result = subprocess.run(command, capture_output=True, text=True, check=True)

        if output_pdf_path.is_file():
            messagebox.showinfo("Success", f"PPTX converted to PDF successfully!\nSaved as: {output_pdf_path}")
        else:
            messagebox.showerror("Error", f"PDF conversion failed.\nLibreOffice output: {result.stderr}\nCheck if the file was created: {output_pdf_path}")

    except FileNotFoundError:
        messagebox.showerror("Error", "LibreOffice command not found. Please ensure LibreOffice is installed.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"LibreOffice conversion failed with error:\n{e.stderr}\n{e.stdout}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    pptx_to_pdf()