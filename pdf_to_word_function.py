import pdfplumber
from pdf2docx import Converter
from docx import Document
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to convert PDF to Word for a specified range of pages or the whole document
def pdf_to_word(pdf_path, start_page, end_page, output_path, convert_whole_document):
    if convert_whole_document:
        # Convert the whole document
        cv = Converter(pdf_path)
        cv.convert(output_path)
        cv.close()
    else:
        # Convert the specified page range to a Word document
        cv = Converter(pdf_path)
        cv.convert(output_path, start=start_page-1, end=end_page)
        cv.close()

# Function to open file dialog and get the PDF path
def browse_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    entry_pdf_path.delete(0, tk.END)
    entry_pdf_path.insert(0, pdf_path)

# Function to open save dialog and get the output file path
def browse_output():
    output_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    entry_output_path.delete(0, tk.END)
    entry_output_path.insert(0, output_path)

# Function to handle the conversion process
def convert_pdf_to_word():
    pdf_path = entry_pdf_path.get()
    if not os.path.isfile(pdf_path):
        messagebox.showerror("Error", "Please select a valid PDF file.")
        return

    convert_whole_document = var_convert_whole_document.get()

    if not convert_whole_document:
        try:
            start_page = int(entry_start_page.get())
            end_page = int(entry_end_page.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter valid start and end page numbers.")
            return
    else:
        # Inform the user that the page range will be ignored
        messagebox.showinfo("Info", "The whole document will be converted. The specified page range will be ignored.")
        start_page = 1
        end_page = None

    output_path = entry_output_path.get()
    if not output_path:
        messagebox.showerror("Error", "Please specify an output file path.")
        return

    pdf_to_word(pdf_path, start_page, end_page, output_path, convert_whole_document)

# Create the main window
window = tk.Tk()
window.title("PDF to Word Converter")

# Create and place the PDF file input
label_pdf_path = tk.Label(window, text="Select PDF file:")
label_pdf_path.grid(row=0, column=0, padx=10, pady=10)

entry_pdf_path = tk.Entry(window, width=50)
entry_pdf_path.grid(row=0, column=1, padx=10, pady=10)

button_browse = tk.Button(window, text="Browse", command=browse_pdf)
button_browse.grid(row=0, column=2, padx=10, pady=10)

# Create and place the start and end page inputs
label_start_page = tk.Label(window, text="Start Page:")
label_start_page.grid(row=1, column=0, padx=10, pady=10)

entry_start_page = tk.Entry(window, width=10)
entry_start_page.grid(row=1, column=1, padx=10, pady=10)

label_end_page = tk.Label(window, text="End Page:")
label_end_page.grid(row=2, column=0, padx=10, pady=10)

entry_end_page = tk.Entry(window, width=10)
entry_end_page.grid(row=2, column=1, padx=10, pady=10)

# Create and place the output file location input
label_output_path = tk.Label(window, text="Save As (Word file):")
label_output_path.grid(row=3, column=0, padx=10, pady=10)

entry_output_path = tk.Entry(window, width=50)
entry_output_path.grid(row=3, column=1, padx=10, pady=10)

button_browse_output = tk.Button(window, text="Browse", command=browse_output)
button_browse_output.grid(row=3, column=2, padx=10, pady=10)

# Create and place the checkbox for converting the whole document
var_convert_whole_document = tk.BooleanVar()
checkbox_convert_whole_document = tk.Checkbutton(window, text="Convert Whole Document", variable=var_convert_whole_document)
checkbox_convert_whole_document.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Create and place the Convert button
button_convert = tk.Button(window, text="Convert", command=convert_pdf_to_word)
button_convert.grid(row=5, column=0, columnspan=3, pady=20)

# Run the application
window.mainloop()
