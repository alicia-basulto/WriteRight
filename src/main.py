from docx import Document
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import PyPDF2
import customtkinter as ctk
import fitz
import language_tool_python


#function to process PDF file
def process_pdf(archivo):
    with open(archivo, 'rb') as pdf_file:
        doc = fitz.open(archivo)
        texto_completo = ""

    for pagina_numero in range(len(doc)):
        pagina = doc[pagina_numero]
        texto_pagina = pagina.get_text()
        texto_completo += " ".join(texto_pagina.split()) + " "

    return texto_completo

#function to process DOCX file
def process_docx(file):
    doc = Document(file)
    texto_completo = ""

    for paragraph in doc.paragraphs:
        texto_completo += paragraph.text + "\n"

    return texto_completo

#function to manage "Select file" button
def seleccionar_archivo():
    file = filedialog.askopenfilename(filetypes=[("Archivos PDF y DOCX", "*.pdf;*.docx")])
    
    if file:
        if file.endswith('.pdf'):
            extractText = process_pdf(file)
        elif file.endswith('.docx'):
            extractText = process_docx(file)
        else:
            extractText = "Unsupported file format"

        
        writingText = extractText.replace("\n"," ")
    
        tool = language_tool_python.LanguageTool('en-US')
        #Checks the grammar and spelling
        correcciones = tool.check(extractText)
        #Create a secondary screen to show the corrections
        ventana_correcciones = tk.Tk()
        ventana_correcciones.title("Correcciones de Gramática")
        cuadro_correcciones = Text(ventana_correcciones, wrap=tk.WORD, width=400, height=100)
        cuadro_correcciones.pack()
        #Create a vertical scroll bar
        scrollbar_vertical = Scrollbar(ventana_correcciones, command=cuadro_correcciones.yview)
        scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)
        cuadro_correcciones.config(yscrollcommand=scrollbar_vertical.set)
        #Create a horizontal scroll bar
        scrollbar_horizontal = Scrollbar(ventana_correcciones, command=cuadro_correcciones.xview, orient=tk.HORIZONTAL)
        scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)
        cuadro_correcciones.config(xscrollcommand=scrollbar_horizontal.set)     
        #Show the corrections in the new box
        for correccion in correcciones:
            cuadro_correcciones.insert(tk.END, correccion.context +" error: " + correccion.message + "\n")


        ventana_correcciones.mainloop()
       

#Create interface
window = ctk.CTk()
window.title("Write Right: PDF & DOCX lector")
window.geometry("300x400")

#Select file button
selectFileBtn = ctk.CTkButton(master=window, text="Select a file", command=seleccionar_archivo)
selectFileBtn.place(relx=0.5,rely=0.5,anchor=CENTER)


window.mainloop()
