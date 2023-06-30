import PySimpleGUI as sg
import subprocess

# Theme
sg.theme('DarkBlue3')
sg.set_options(font=["D2Coding", 12])

# GUI Layout
layout = [
    [
        sg.Button("PDF_PNG"),
        sg.Button("PDF_PPTX")
    ],
    [
        sg.Button("PNG_PDF"),
        sg.Button("PPTX_PDF")
    ],
    [
        sg.Button("PdfArranger")
    ],
    [
        sg.Button('Exit')
    ]
]

window = sg.Window('ConvertDoc', layout)

while True:
    event, values = window.read() 
    if event==sg.WIN_CLOSED or event=='Exit':
        break
    if event == 'PDF_PNG':
        subprocess.call([r'D:\github\ConvertDocs\CMD_PDF2PNG.bat'])
    if event == 'PDF_PPTX':
        subprocess.call([r'D:\github\ConvertDocs\CMD_PDF2PPTX.bat'])
    if event == 'PNG_PDF':
        subprocess.call([r'D:\github\ConvertDocs\CMD_PNG2PDF.bat'])
    if event == 'PPTX_PDF':
        subprocess.call([r'D:\github\ConvertDocs\CMD_PPTX2PDF.bat'])
    if event == 'PdfArranger':
        #subprocess.check_output(r'cd /d C:\Users\dhkima\scoop\apps\pdfarranger\current & pdfarranger.exe', shell=True)
        subprocess.check_output(r'cd /d %userprofile%\scoop\apps\pdfarranger\current & pdfarranger.exe', shell=True)

window.close()
