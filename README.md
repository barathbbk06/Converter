# Local Word ↔ PDF Converter for Windows

A reliable, fast, and **100% offline** document converter designed for Windows 10 and 11. It allows you to convert Word documents to PDF and PDF files to Word documents without sending any data to external servers, cloud services, or third-party APIs.

---

## 🌟 Key Features

- **Bidirectional Conversion:**
  - **Word to PDF (`.docx`, `.doc` → `.pdf`)**: Native fidelity using local Microsoft Word COM automation (`pywin32`) or headless LibreOffice fallback.
  - **PDF to Word (`.pdf` → `.docx`)**: Extracts editable text, headings, tables, formatting, and images using `pdf2docx`.
- **100% Local & Private:** No files, text, or telemetry leave your computer.
- **Easy File Selection:**
  - Native Windows file picker dialogs for input documents and output folders.
  - Full path typing and pasting supported.
  - Automatic conversion direction detection based on file extension.
- **Safe Output & Duplicate Handling:**
  - Automatically calculates output filenames.
  - If a file with the same name already exists, prompts you to **Overwrite**, **Create Unique Name** (e.g., `document (1).pdf`), or **Cancel**.
  - Never alters or deletes the original input file.
- **High-DPI Desktop Interface:**
  - Built with Python `Tkinter` and `ttk`.
  - Non-blocking background worker threads (the GUI remains responsive during large file conversions).
  - One-click actions to **Open Converted File**, **Open Destination Folder**, and **Copy Output Path**.
- **CLI & Scripting Support:** Also includes a command-line interface (`converter.py`) for automated workflows.

---

## 📂 Project Structure

```text
pdf to word/
│
├── main.py              # Tkinter desktop graphical interface
├── converter.py         # Core local conversion backend (also usable via CLI)
├── requirements.txt     # Python package requirements
├── run.bat              # One-click Windows batch launcher
├── test_converter.py    # Automated test suite
└── README.md            # Detailed instructions and documentation
```

---

## 📋 Prerequisites & Software Requirements

### 1. Python (Windows 10 / 11)
- Python 3.10 or newer (tested on Python 3.14).
- Ensure **"Add python.exe to PATH"** was checked during Python installation.

### 2. Office Software for Word-to-PDF Conversion
To convert `.docx` to `.pdf` locally with high fidelity, ensure **one** of the following is installed on your computer:
- **Microsoft Word (Office 2013, 2016, 2019, 2021, or Microsoft 365)** *(Recommended)*:
  Uses Microsoft Word's built-in PDF export engine via Windows COM automation.
- **LibreOffice (Free & Open Source)** *(Alternative)*:
  If Microsoft Word is not installed, the tool will automatically detect LibreOffice (`soffice.exe`) in standard locations (`C:\Program Files\LibreOffice\program\soffice.exe`) or PATH and run in headless mode.

> **Note on PDF-to-Word:** Does not require Microsoft Office or LibreOffice. It runs entirely via Python's local `pdf2docx` engine.

---

## 🚀 Quick Start (Installation & Execution)

### Step 1: Check Python
Open Command Prompt or PowerShell and type:
```cmd
python --version
```
If Python is not recognized, download and install it from [python.org](https://www.python.org/downloads/) (make sure to check **Add Python to PATH**).

### Step 2: Install Python Dependencies
Open PowerShell or Command Prompt in this folder:
```cmd
pip install -r requirements.txt
```
This installs:
- `pywin32` (Windows COM bridge for Word automation)
- `pdf2docx` (Local PDF to DOCX extractor)
- `python-docx` (Document formatting support)

### Step 3: Run the Application
You have two simple ways to launch the app:

#### Option A: Double-Click `run.bat` (Recommended)
Simply double-click `run.bat` in Windows File Explorer. The batch script will:
1. Verify that Python is available.
2. Verify dependencies and automatically install missing packages if needed.
3. Launch the converter window.

#### Option B: Run via Command Prompt / PowerShell
```cmd
python main.py
```

---

## 🖥️ Using the Graphical Interface

1. **Select Conversion Type:**
   - Choose `Word to PDF (.docx → .pdf)` or `PDF to Word (.pdf → .docx)` from the dropdown.
   - *(Note: Selecting an input file will also auto-detect and switch the mode automatically).*
2. **Choose Input File:**
   - Click **Browse Input File...** to select your file, or type/paste the absolute path.
3. **Choose Destination Folder:**
   - Click **Choose Destination...** to pick where to save the converted file, or click **Use Input File's Folder**.
4. **Click "Convert Document Now":**
   - The progress bar will activate.
   - If a file with the target name already exists, a dialog asks whether to **Overwrite**, **Create Unique Name**, or **Cancel**.
5. **Access Your Converted File:**
   - Once completed, the full path appears in the result box.
   - Click **Open Converted File** to open it immediately in your default application.
   - Click **Open Destination Folder** to highlight the file in Windows File Explorer.
   - Click **Copy Path** to copy the path to your clipboard.

---

## ⚡ Command-Line Usage (Optional)

The `converter.py` module can also be invoked directly from PowerShell or Command Prompt for batch scripts:

```cmd
# Convert Word to PDF:
python converter.py "C:\path\to\document.docx"

# Convert PDF to Word:
python converter.py "C:\path\to\document.pdf"

# Specify a custom destination directory or output file:
python converter.py "C:\path\to\document.docx" -o "D:\ConvertedDocs"

# Force overwrite if file exists:
python converter.py "C:\path\to\document.docx" -f
```

---

## 🧪 Running the Automated Test Suite

To verify all conversion functions and collision handlers without opening the GUI, run:

```cmd
python test_converter.py
```

This automated test:
1. Detects Microsoft Word COM and LibreOffice availability.
2. Creates an authentic `.docx` test document with formatted text and tables.
3. Converts `.docx` → `.pdf` and verifies file creation.
4. Validates duplicate filename collision handling.
5. Converts `.pdf` → `.docx` and verifies table/text extraction.
6. Validates error handling for missing files and invalid extensions.
7. Automatically cleans up test artifacts.

---

## 📦 Packaging into a Standalone `.exe` (PyInstaller)

If you would like to distribute or run this application as a single portable Windows executable (`.exe`) without needing Python installed on other PCs:

1. Install PyInstaller:
   ```cmd
   pip install pyinstaller
   ```
2. Build the executable:
   ```cmd
   pyinstaller --noconsole --onefile --name "WordPdfConverter" main.py
   ```
3. Your standalone executable will be generated inside the `dist\` folder:
   ```cmd
   dist\WordPdfConverter.exe
   ```

---

## 🔍 Important Technical Notes & Limitations

- **Scanned PDFs vs. Searchable PDFs:**
  - PDFs with selectable text, tables, and images convert cleanly into editable Microsoft Word documents.
  - If a PDF was created by a physical paper scanner (image-only PDF with no text layer), `pdf2docx` will preserve the scanned page images inside the Word document. Extracting editable text from raw scanned images requires Optical Character Recognition (OCR), such as Tesseract OCR.
- **Complex Layouts:**
  - Multi-column newsletters, magazine spreads, or heavily nested drawing objects are approximated to standard Word flow formatting.
- **Original Files Protected:**
  - The script opens documents in read-only mode and never deletes, renames, or modifies your original input files.
- **Protected/Encrypted Documents:**
  - Password-protected Word documents or encrypted PDFs must have their passwords removed before conversion.
