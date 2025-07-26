import os
import pytesseract
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog, scrolledtext, Canvas
from tkinter import font as tkFont
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import pandas as pd
import cv2
import numpy as np
from tabulate import tabulate
from pdf2image import convert_from_path
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image, ImageTk
import fitz  # PyMuPDF

# Set the path to the Tesseract executable (Windows-specific)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Create the main window
root = Tk()
root.title("Advanced PDF Manipulation Tool")
root.geometry("1000x700")
root.config(bg="#F0F0F0")

# UI Enhancements
font_style = tkFont.Font(family="Arial", size=12)
button_font = tkFont.Font(family="Arial", size=12, weight="bold")

# Header frame
header_frame = Frame(root, bg="#2E8B57", height=80)
header_frame.pack(fill="x")
header_label = Label(header_frame, text="Advanced PDF Manipulation Tool", font=("Arial", 18, "bold"), fg="white", bg="#2E8B57")
header_label.pack(padx=20, pady=20)

# Sidebar frame for buttons
sidebar_frame = Frame(root, width=200, bg="#2E8B57", height=600, relief="sunken", bd=2)
sidebar_frame.pack(side=LEFT, fill=Y)

# Label to display the recently edited PDF
recent_pdf_label = Label(root, text="Recently Edited PDF: None", font=("Arial", 12), bg="#F0F0F0", fg="#333")
recent_pdf_label.pack(pady=10)

def update_recent_pdf(filename):
    """Update the label with the name of the recently edited PDF."""
    recent_pdf_label.config(text=f"Recently Edited PDF: {os.path.basename(filename)}")

def display_text(text):
    """Display extracted text in a new window."""
    text_window = Toplevel(root)
    text_window.title("Extracted Text")
    text_window.geometry("600x400")
    text_area = scrolledtext.ScrolledText(text_window, wrap=WORD, font=("Arial", 12))
    text_area.pack(expand=True, fill=BOTH, padx=10, pady=10)
    text_area.insert(INSERT, text)

def extract_text():
    """Extract text from a PDF (non-OCR) and save it to a text file."""
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")], initialdir=os.getcwd())
    if not file:
        return
    
    try:
        pdf_reader = PdfReader(file)
        extracted_text = []
        
        for page in pdf_reader.pages:
            text = page.extract_text()
            if text:
                extracted_text.append(text)
        
        result = "\n".join(extracted_text) if extracted_text else "No text found in the PDF."
        output = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if output:
            with open(output, "w", encoding="utf-8") as txt_file:
                txt_file.write(result)
        display_text(result)
        update_recent_pdf(file)
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process PDF: {str(e)}")

def extract_text_ocr():
    """Extract text from a PDF using OCR and save it to a PDF file."""
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file:
        return
    images = convert_from_path(file)
    text = "\n".join(pytesseract.image_to_string(img) for img in images)
    display_text(text if text else "No text found using OCR.")
    update_recent_pdf(file)
    
    output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    if output_pdf:
        c = canvas.Canvas(output_pdf, pagesize=letter)
        c.setFont("Courier", 10)
        y_position = 750
        
        for line in text.split("\n"):
            c.drawString(50, y_position, line)
            y_position -= 15
            if y_position < 50:
                c.showPage()
                c.setFont("Courier", 10)
                y_position = 750
        
        c.save()
        messagebox.showinfo("Success", f"OCR Text saved to {output_pdf}")

def split_pdf():
    """Split a PDF into specified page ranges."""
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file:
        return
    
    pdf_reader = PdfReader(file)
    total_pages = len(pdf_reader.pages)
    page_input = simpledialog.askstring("Split Pages", "Enter page ranges to split (e.g., 2-5,8-10):")
    
    if not page_input:
        return
    
    page_ranges = []
    for part in page_input.split(","):
        if "-" in part:
            start, end = map(int, part.split("-"))
            page_ranges.append((start, end))
        else:
            page = int(part)
            page_ranges.append((page, page))
    
    for i, (start, end) in enumerate(page_ranges):
        if start < 1 or end > total_pages:
            messagebox.showerror("Error", f"Page range {start}-{end} is out of bounds.")
            return
        
        pdf_writer = PdfWriter()
        for page in range(start-1, end):
            pdf_writer.add_page(pdf_reader.pages[page])
        
        output = filedialog.asksaveasfilename(
            defaultextension=f"_part{i+1}.pdf", 
            filetypes=[("PDF Files", "*.pdf")]
        )
        
        if output:
            with open(output, "wb") as out:
                pdf_writer.write(out)
    
    messagebox.showinfo("Success", "PDF split successfully into specified ranges!")

def merge_pdfs():
    """Merge multiple PDFs into a single PDF."""
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")], title="Select PDFs to Merge")
    if not files:
        return
    pdf_writer = PdfWriter()
    for file in files:
        pdf_reader = PdfReader(file)
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)
    output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="Save Merged PDF")
    if output:
        with open(output, "wb") as out:
            pdf_writer.write(out)
        update_recent_pdf(output)
        messagebox.showinfo("Success", "PDFs merged successfully!")

def images_to_pdf():
    """Convert selected images into a single PDF."""
    files = filedialog.askopenfilenames(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")], title="Select Images")
    if not files:
        return
    images = [Image.open(file).convert("RGB") for file in files]
    output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="Save as PDF")
    if output:
        images[0].save(output, save_all=True, append_images=images[1:])
        update_recent_pdf(output)
        messagebox.showinfo("Success", "Images converted to PDF successfully!")

def extract_images():
    """Extract images from a PDF and save them to a folder."""
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file:
        return

    try:
        pdf_document = fitz.open(file)
        image_folder = filedialog.askdirectory(title="Select Folder to Save Images")
        if not image_folder:
            return

        img_count = 0
        for page_num, page in enumerate(pdf_document, start=1):
            for img_index, img in enumerate(page.get_images(full=True), start=1):
                xref = img[0]
                base_image = pdf_document.extract_image(xref)
                image_bytes = base_image["image"]
                img_extension = base_image["ext"]
                img_filename = os.path.join(image_folder, f"page_{page_num}_img_{img_index}.{img_extension}")

                with open(img_filename, "wb") as img_file:
                    img_file.write(image_bytes)
                
                img_count += 1

        messagebox.showinfo("Success", f"Extracted {img_count} images from the PDF!")
        update_recent_pdf(file)
    
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract images: {str(e)}")

def extract_tables():
    """Extract tables from a PDF, including color info and OCR text."""
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file:
        return
    
    try:
        extracted_tables = []
        with pdfplumber.open(file) as pdf:
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table)
                    extracted_tables.append(df)
        
        if not extracted_tables:
            messagebox.showinfo("No Tables Found", "No tables detected in the PDF.")
            return
        
        images = convert_from_path(file)
        for img in images:
            img_cv = np.array(img)
            hsv = cv2.cvtColor(img_cv, cv2.COLOR_BGR2HSV)
            dominant_color = cv2.mean(hsv)[:3]
            color_str = f"Background Color (HSV): {dominant_color}"
        
        def preprocess_image(img):
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            return binary
        
        ocr_text = "\n".join([pytesseract.image_to_string(preprocess_image(np.array(img))) for img in images])
        
        table_window = Toplevel()
        table_window.title("Extracted Tables with Colors and OCR")
        table_window.geometry("900x600")
        text_area = scrolledtext.ScrolledText(table_window, wrap=NONE, font=("Courier", 12))
        text_area.pack(expand=True, fill=BOTH, padx=10, pady=10)
        
        for df in extracted_tables:
            table_str = tabulate(df, headers='keys', tablefmt='grid', showindex=False)
            text_area.insert(INSERT, table_str + "\n\n" + color_str + "\n\n")
        
        text_area.insert(INSERT, "OCR Extracted Text:\n" + ocr_text)
        
        output_csv = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if output_csv:
            combined_df = pd.concat(extracted_tables, ignore_index=True)
            combined_df.to_csv(output_csv, index=False, header=False)
            messagebox.showinfo("Success", f"Tables saved to {output_csv}")
        
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf:
            c = canvas.Canvas(output_pdf, pagesize=letter)
            c.setFont("Courier", 10)
            y_position = 750
            
            for df in extracted_tables:
                table_str = tabulate(df, headers='keys', tablefmt='grid', showindex=False)
                for line in table_str.split("\n"):
                    c.drawString(50, y_position, line)
                    y_position -= 15
                    if y_position < 50:
                        c.showPage()
                        c.setFont("Courier", 10)
                        y_position = 750
            
            c.save()
            messagebox.showinfo("Success", f"Tables saved to {output_pdf}")
    
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract tables: {str(e)}")

def encrypt_pdf():
    """Encrypt a PDF with a user-provided password."""
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file:
        return

    try:
        pdf_reader = PdfReader(file)
        pdf_writer = PdfWriter()

        # Copy all pages to the writer
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

        # Prompt user for a password
        password = simpledialog.askstring("Password", "Enter password for PDF encryption:", show="*")
        if not password:
            messagebox.showwarning("Warning", "No password provided. Encryption cancelled.")
            return

        # Encrypt the PDF with the password
        pdf_writer.encrypt(user_password=password)

        # Save the encrypted PDF
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="Save Encrypted PDF")
        if output:
            with open(output, "wb") as out:
                pdf_writer.write(out)
            update_recent_pdf(output)
            messagebox.showinfo("Success", "PDF encrypted successfully!")
    
    except Exception as e:
        messagebox.showerror("Error", f"Failed to encrypt PDF: {str(e)}")
def decrypt_pdf():
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not file:
        return
    try:
        pdf_reader = PdfReader(file)
        if not pdf_reader.is_encrypted:
            messagebox.showinfo("Info", "PDF is not encrypted.")
            return
        password = simpledialog.askstring("Password", "Enter PDF password:", show="*")
        if not password or pdf_reader.decrypt(password) == 0:
            messagebox.showerror("Error", "Incorrect password or decryption failed.")
            return
        pdf_writer = PdfWriter()
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output:
            with open(output, "wb") as out:
                pdf_writer.write(out)
            update_recent_pdf(output)
            messagebox.showinfo("Success", "PDF decrypted successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to decrypt PDF: {str(e)}")

# Create UI buttons
buttons = [
    ("Extract Text", extract_text),
    ("Extract Text (OCR)", extract_text_ocr),
    ("Split PDF", split_pdf),
    ("Merge PDFs", merge_pdfs),
    ("Images to PDF", images_to_pdf),
    ("Extract Images", extract_images),
    ("Extract Tables", extract_tables),
    ("Encrypt PDF", encrypt_pdf),
    ("Decrypt PDF", decrypt_pdf),
] 

for i, (text, cmd) in enumerate(buttons):
    btn = Button(sidebar_frame, text=text, font=button_font, bg="#4CAF50", fg="white", relief="flat", width=20, height=2, command=cmd)
    btn.grid(row=i, column=0, pady=10, padx=10)

# Start the main event loop
root.mainloop()
