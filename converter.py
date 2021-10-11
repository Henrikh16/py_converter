import tkinter
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askdirectory
from PIL import Image
import os, sys, shutil, os.path
from pdf2image import convert_from_path, convert_from_bytes
from docx2pdf import convert
import os
import comtypes.client
import time


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

def doc_to_pdf():
    wdFormatPDF = 17
    in_file=askopenfilename()
    out_file=asksaveasfilename()
    in_file2=askopenfilename()
    out_file2=asksaveasfilename()
    print (in_file)
    print (out_file)
    print (in_file2)
    print (out_file2)
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = True
    time.sleep(3)
    doc=word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Visible = False
    doc = word.Documents.Open(in_file2)
    doc.SaveAs(out_file2, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
