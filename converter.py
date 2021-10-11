import tkinter
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askdirectory
from PIL import Image
import os, sys, shutil, os.path
from pdf2image import convert_from_path, convert_from_bytes

def img_to_pdf():
    source_file_path = Image.open(askopenfilename())
    pdf = source_file_path.convert('RGB')
    files = [('All Files', '*.*'),
             ('Portable document file', '*.pdf')]
    save_file_path = asksaveasfilename(filetypes = files)
    exit_file_path = pdf.save(save_file_path)

def pdf_to_img():
    images = convert_from_path(askopenfilename(), poppler_path=r'C:\Program Files\poppler-0.68.0\bin')
    for img in images:
        img.save(asksaveasfilename(), 'JPEG')
