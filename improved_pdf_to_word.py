import pdfplumber
from pdf2docx import Converter
from docx import Document
import os
import tkinter as tk
from tkinter import filedialog, messagebox

class ProgressTracker:
    def __init__(self, progress_callback):
        self.progress_callback = progress_callback
        self.current_page = 0
        self.total_pages = 0

    def update_progress(self, current, total):
        self.current_page = current
        self.total_pages = total
        progress = (current / total) * 100
        self.progress_callback(progress)

def pdf_to_word(pdf_path, start_page, end_page, output_path, convert_whole_document):
    try:
        def update_progress_label(progress):
            progress_label.config(text=f"Progress: {progress:.2f}%")
            window.update_idletasks()

        progress_tracker = ProgressTracker(update_progress_label)
        
        if convert_whole_document:
            cv = Converter(pdf_path)
            total_pages = len(cv.pdf.pages)
            
            def progress_callback(current, total):
                progress_tracker.update_progress(current, total)
            
            cv.convert(output_path, progress_callback=progress_callback)
            cv.close()
        else:
            cv = Converter(pdf_path)
            total_pages = end_page - start_page + 1
            current_page = 0
            
            for i in range(start_page-1, end_page):
                cv.convert(output_path, start=i, end=i+1)
                current_page += 1
                progress_tracker.update_progress(current_page, total_pages)
            cv.close()
            
        messagebox.showinfo("Success", "PDF successfully converted to Word document.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")

# Function to open file dialog and get the PDF path
def browse_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    entry_pdf_path.delete(0, tk.END)
    entry_pdf_path.insert(0, pdf_path)

# Function to open file dialog and get the output path
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

    # Get total pages first
    try:
        with pdfplumber.open(pdf_path) as pdf:
            num_pages = len(pdf.pages)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading PDF file: {e}")
        return

    convert_whole_document = var_convert_whole_document.get()
    
    if convert_whole_document:
        start_page = 1
        end_page = num_pages
        messagebox.showinfo("Info", "Converting the whole document.")
    else:
        try:
            start_page = int(entry_start_page.get())
            end_page = int(entry_end_page.get())
            
            # Validate page range
            if start_page < 1 or end_page > num_pages or start_page > end_page:
                messagebox.showerror("Error", f"Please enter a valid page range (1-{num_pages}).")
                return
        except ValueError:
            messagebox.showerror("Error", "Please enter valid start and end page numbers.")
            return

    output_path = entry_output_path.get()
    if not output_path:
        messagebox.showerror("Error", "Please specify an output file path.")
        return

    # Show progress label and reset progress
    progress_label.config(text="Progress: 0%")
    progress_label.grid(row=6, column=0, columnspan=3, pady=10)
    window.update_idletasks()

    # Start the conversion process
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

# Create and place the progress label
progress_label = tk.Label(window, text="Progress: 0%")
progress_label.grid(row=6, column=0, columnspan=3, pady=10)

# Run the application
window.mainloop()