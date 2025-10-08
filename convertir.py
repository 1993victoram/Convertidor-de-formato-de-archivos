"""
LEER ANTES DE USAR ESTE PROGRAMA

Este es un programa de aplicación con Python para convertir archivos entre Word (.docx) y PDF.
Usa Microsoft Word para conversiones con imágenes y pdf2docx para PDF a Word.

Requisitos:
    pip install pywin32 pdf2docx

Notas:
- Tener Microsoft Word instalado en Windows con licencia activada, en caso contrario no funcionará.
- En ningún caso este programa es un sustituto de Microsoft Word. 
- Solo es una herramienta para agilizar la elaboración de trabajos en cada formato.
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import os
import sys
import win32com.client
from pdf2docx import Converter

def mostrar_en_explorador(ruta):
    """Abre el explorador de Windows y resalta el archivo."""
    try:
        os.system(f'explorer /select,"{ruta}"')
    except Exception as e:
        messagebox.showwarning("⚠ Aviso", f"No se pudo abrir el explorador:\n{e}")

def word_a_pdf():
    """Convierte un archivo Word (.docx) a PDF conservando imágenes (usa Microsoft Word)."""
    try:
        word_file = filedialog.askopenfilename(
            title="Selecciona un archivo Word",
            filetypes=[("Archivos Word", "*.docx")]
        )
        if not word_file:
            return

        pdf_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            title="Guardar PDF como"
        )
        if not pdf_file:
            return
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(word_file))
        doc.SaveAs(os.path.abspath(pdf_file), FileFormat=17)  # 17 = PDF
        doc.Close()
        word.Quit()

        messagebox.showinfo("✅ Conversión exitosa", f"Word → PDF:\n\nArchivo guardado en:\n{pdf_file}")
        os.startfile(pdf_file)  # Abrir el archivo convertido

    except Exception as e:
        messagebox.showerror("❌ Error", f"No se pudo convertir a PDF:\n{e}")

def pdf_a_word():
    """Convierte un archivo PDF a Word (.docx)."""
    try:
        pdf_file = filedialog.askopenfilename(
            title="Selecciona un archivo PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if not pdf_file:
            return

        word_file = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Archivos Word", "*.docx")],
            title="Guardar Word como"
        )
        if not word_file:
            return

        cv = Converter(pdf_file)
        cv.convert(word_file, start=0, end=None)
        cv.close()

        messagebox.showinfo("✅ Conversión exitosa", f"PDF → Word:\n\nArchivo guardado en:\n{word_file}")
        os.startfile(word_file)  # Abrir el archivo convertido

    except Exception as e:
        messagebox.showerror("❌ Error", f"No se pudo convertir a Word:\n{e}")

def resource_path(relative_path):
    """Devuelve la ruta absoluta para acceder a recursos en .exe o .py."""
    try:
        base_path = sys._MEIPASS  # PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# === Interfaz gráfica ===
convertidor = tk.Tk()
convertidor.title("Convertidor de Word ↔ PDF con imágenes")
convertidor.geometry("600x500")
convertidor.resizable(False, False)
convertidor.config(background="#3d57af")

# Icono de la ventana
try:
    icono = resource_path("icono.ico")
    convertidor.iconbitmap(icono)
except Exception as e:
    messagebox.showwarning("⚠ Icono no encontrado", f"No se pudo cargar el icono:\n{e}")

tk.Label(
    convertidor, 
    text="Seleccione el tipo de archivo a convertir", 
    bg="#3d57af", fg="#000000", 
    font=("Arial 16 bold")
).place(x=50, y=50)

frame_convertidor = tk.Frame(convertidor, bg="#5c7dbb", width=500, height=300)
frame_convertidor.place(x=50, y=100)

# Word → PDF
label_word = tk.Label(frame_convertidor, text="Archivo Word (.docx)", bg="#5c7dbb", fg="#000000", font=("Arial 14"))
label_word.place(x=50, y=45)

boton_word = tk.Button(
    frame_convertidor, 
    text="Convertir Word a PDF", 
    bg="#ffffff", fg="#000000", 
    font=("Arial 12"), 
    command=word_a_pdf
)
boton_word.place(x=260, y=40, width=190, height=40)

# PDF → Word
label_pdf = tk.Label(frame_convertidor, text="Archivo PDF (.pdf)", bg="#5c7dbb", fg="#000000", font=("Arial 14"))
label_pdf.place(x=50, y=155)

boton_pdf = tk.Button(
    frame_convertidor, 
    text="Convertir PDF a Word", 
    bg="#ffffff", fg="#000000", 
    font=("Arial 12"), 
    command=pdf_a_word
)
boton_pdf.place(x=260, y=150, width=190, height=40)

convertidor.mainloop()
