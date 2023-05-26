import PySimpleGUI as sg
import os

INPUT = sg.popup_get_file("Select xlsx file",  title="File selector")
os.system(f"pdf2pptx {INPUT}")

