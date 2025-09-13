import os
import pyttsx3
import pypdf
# import webbrowser
import docx2txt
import openpyxl
from tkinter.filedialog import *

player = pyttsx3.init()
def speak(text):
    if text:
        # print(text)
        player.say(text)
        player.runAndWait()
        
book = askopenfilename()

if not book:
    print("No file selected.")
    exit()
os.startfile(book)
ext = os.path.splitext(book)[1]

if ext == '.pdf':
    # webbrowser.open_new(book)
    pdfreader = pypdf.PdfReader(book)
    pages = pdfreader.get_num_pages()

    for num in range(0, pages):
        page = pdfreader.get_page(num)
        text = page.extract_text(extraction_mode="layout")
        speak(text)
elif ext == '.txt':
    with open(book, 'r', encoding='utf-8') as f:
        text = f.read()
        speak(text)
elif ext == '.docx':
    text = docx2txt.process(book)
    speak(text)  
elif ext == '.xlsx':
    wb = openpyxl.load_workbook(book)
    text = ""
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            line = ' '.join(str(cell) for cell in row if cell is not None)
            text += line + '\n'
    speak(text)
else:
    print(f"Unsupported file type: {ext}")