import os
import threading
from tkinter import *
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx2pdf import convert as convert_to_pdf


# ------------------ Conversion Functions ------------------ #

def pdf_to_word(pdf_file, word_file, start_page=0, end_page=None):
    try:
        cv = Converter(pdf_file)
        cv.convert(word_file, start=start_page, end=end_page)
        cv.close()
        return True, f"Successfully converted {pdf_file} to {word_file}"
    except Exception as e:
        return False, str(e)


def word_to_pdf(word_file, output_dir=None):
    try:
        if output_dir:
            convert_to_pdf(word_file, output_dir)
        else:
            convert_to_pdf(word_file)
        return True, f"Successfully converted {word_file} to PDF."
    except Exception as e:
        return False, str(e)


# ------------------ GUI Functions ------------------ #

def browse_pdf():
    path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_path_var.set(path)


def browse_docx_output():
    path = filedialog.asksaveasfilename(defaultextension=".docx",
                                        filetypes=[("Word Files", "*.docx")])
    word_output_var.set(path)


def browse_docx_input():
    path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    word_path_var.set(path)


def browse_output_dir():
    path = filedialog.askdirectory()
    output_dir_var.set(path)


def start_pdf_to_word():
    pdf_file = pdf_path_var.get()
    word_file = word_output_var.get()
    start_page = start_var.get()
    end_page = end_var.get()

    if not pdf_file or not word_file:
        messagebox.showerror("Error", "Please select both PDF and output Word file.")
        return

    # Convert empty entries to None or int
    start_page = int(start_page) if start_page.isdigit() else 0
    end_page = int(end_page) if end_page.isdigit() else None

    status_label.config(text="Converting... Please wait.")

    def task():
        success, msg = pdf_to_word(pdf_file, word_file, start_page, end_page)
        status_label.config(text=msg)
        if success:
            messagebox.showinfo("Success", msg)
        else:
            messagebox.showerror("Error", msg)

    threading.Thread(target=task).start()


def start_word_to_pdf():
    word_file = word_path_var.get()
    out_dir = output_dir_var.get() if output_dir_var.get().strip() else None

    if not word_file:
        messagebox.showerror("Error", "Please select a Word file.")
        return

    status_label.config(text="Converting... Please wait.")

    def task():
        success, msg = word_to_pdf(word_file, out_dir)
        status_label.config(text=msg)
        if success:
            messagebox.showinfo("Success", msg)
        else:
            messagebox.showerror("Error", msg)

    threading.Thread(target=task).start()


# ------------------ Building Tkinter UI ------------------ #

root = Tk()
root.title("PDF <-> Word Converter")
root.geometry("550x400")
root.resizable(False, False)

title_label = Label(root, text="PDF ↔ DOCX Converter", font=("Arial", 18, "bold"))
title_label.pack(pady=10)

notebook_frame = Frame(root)
notebook_frame.pack(fill=BOTH, expand=True)

# Variables
pdf_path_var = StringVar()
word_output_var = StringVar()
word_path_var = StringVar()
output_dir_var = StringVar()
start_var = StringVar()
end_var = StringVar()

# ---------------- PDF → Word Frame ---------------- #

pdf_frame = LabelFrame(root, text="PDF → Word (.docx)", padx=10, pady=10, font=("Arial", 12))
pdf_frame.pack(fill="x", padx=10, pady=10)

Label(pdf_frame, text="PDF File:").grid(row=0, column=0, sticky="w")
Entry(pdf_frame, textvariable=pdf_path_var, width=45).grid(row=0, column=1)
Button(pdf_frame, text="Browse", command=browse_pdf).grid(row=0, column=2)

Label(pdf_frame, text="Output Word File:").grid(row=1, column=0, sticky="w")
Entry(pdf_frame, textvariable=word_output_var, width=45).grid(row=1, column=1)
Button(pdf_frame, text="Browse", command=browse_docx_output).grid(row=1, column=2)

Label(pdf_frame, text="Start Page:").grid(row=2, column=0, sticky="w")
Entry(pdf_frame, textvariable=start_var, width=10).grid(row=2, column=1, sticky="w")

Label(pdf_frame, text="End Page:").grid(row=3, column=0, sticky="w")
Entry(pdf_frame, textvariable=end_var, width=10).grid(row=3, column=1, sticky="w")

Button(pdf_frame, text="Convert PDF to WORD", bg="#4CAF50", fg="white",
       command=start_pdf_to_word).grid(row=4, column=1, pady=10)

# ---------------- Word → PDF Frame ---------------- #

word_frame = LabelFrame(root, text="Word (.docx) → PDF", padx=10, pady=10, font=("Arial", 12))
word_frame.pack(fill="x", padx=10, pady=10)

Label(word_frame, text="Word File:").grid(row=0, column=0, sticky="w")
Entry(word_frame, textvariable=word_path_var, width=45).grid(row=0, column=1)
Button(word_frame, text="Browse", command=browse_docx_input).grid(row=0, column=2)

Label(word_frame, text="Output Directory:").grid(row=1, column=0, sticky="w")
Entry(word_frame, textvariable=output_dir_var, width=45).grid(row=1, column=1)
Button(word_frame, text="Browse", command=browse_output_dir).grid(row=1, column=2)

Button(word_frame, text="Convert WORD to PDF", bg="#2196F3", fg="white",
       command=start_word_to_pdf).grid(row=2, column=1, pady=10)

# Status Label
status_label = Label(root, text="", fg="blue", font=("Arial", 12))
status_label.pack(pady=10)

root.mainloop()
