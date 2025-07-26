# 📄 Advanced PDF Manipulation Tool

[![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python)](https://www.python.org/)
[![Tkinter GUI](https://img.shields.io/badge/Tkinter-GUI-green)](https://docs.python.org/3/library/tkinter.html)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> A feature-rich desktop application built with Python and Tkinter for advanced PDF manipulation, including OCR, table/image extraction, encryption, and more.

---

## 🧰 Features

- **Extract Text**: Extract plain text from PDFs
- **OCR Text Extraction**: Use Tesseract OCR to extract text from scanned PDFs
- **Split PDFs**: Extract specified page ranges into separate PDFs
- **Merge PDFs**: Combine multiple PDF files into one
- **Images to PDF**: Convert images into a single PDF
- **Extract Images**: Pull embedded images from PDFs and save them
- **Extract Tables**: Extract tables with `pdfplumber`, save as CSV/PDF
- **Encrypt/Decrypt PDFs**: Secure PDFs with passwords
- **User-Friendly Interface**: Intuitive GUI with sidebar controls and status display

---

## 💻 Technologies Used

- Python 3.11+
- Tkinter for GUI
- PyPDF2, pdfplumber, pytesseract
- OpenCV, NumPy, pandas
- Pillow, reportlab, pdf2image, PyMuPDF

---
## Installation

1.  **Clone the repository (or download the `App.py` file):**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install Python dependencies:**
    ```bash
    pip install PyPDF2 pdfplumber pandas opencv-python numpy tabulate pdf2image reportlab Pillow PyMuPDF pytesseract
    ```

3.  **Install Tesseract OCR Engine:**
    *   **Windows**: Download the installer from [Tesseract-OCR GitHub](https://tesseract-ocr.github.io/tessdoc/Downloads.html). During installation, note the installation path (e.g., `C:\Program Files\Tesseract-OCR`).
    *   **macOS**:
        ```bash
        brew install tesseract
        ```
    *   **Linux (Debian/Ubuntu)**:
        ```bash
        sudo apt-get install tesseract-ocr
        ```

4.  **Configure Tesseract Path in the Script:**
    Open `App.py` and update the `pytesseract.pytesseract.tesseract_cmd` variable to point to your Tesseract executable.
    For example, on Windows:
    ```python
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    ```
    On Linux/macOS, it might be:
    ```python
    pytesseract.pytesseract.tesseract_cmd = r'/usr/local/bin/tesseract' # Or wherever tesseract is installed
    ```

## 👩🏻‍💻Usage

To run the application, simply execute the Python script:

```bash
python App.py
