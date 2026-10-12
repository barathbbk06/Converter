"""
Local Word <-> PDF Conversion Engine
------------------------------------
Provides reliable, 100% offline conversion between Word (.docx/.doc) and PDF (.pdf).

Features:
- Word to PDF: Microsoft Word COM automation (pywin32) with LibreOffice headless fallback.
- PDF to Word: Local PDF structure extraction via pdf2docx.
- Path handling: Safe handling for spaces, Unicode characters, and long paths.
- Conflict resolution: Functions for generating unique filenames.
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Tuple, Optional


class ConversionError(Exception):
    """Custom exception for conversion failures with user-friendly messages."""
    pass


def get_unique_filepath(target_path: str) -> str:
    """
    If target_path exists, generates a unique filename by appending (1), (2), etc.
    Example: 'report.pdf' -> 'report (1).pdf'
    """
    target_path = os.path.abspath(target_path)
    if not os.path.exists(target_path):
        return target_path

    directory, filename = os.path.split(target_path)
    base_name, ext = os.path.splitext(filename)

    counter = 1
    while True:
        candidate_name = f"{base_name} ({counter}){ext}"
        candidate_path = os.path.join(directory, candidate_name)
        if not os.path.exists(candidate_path):
            return candidate_path
        counter += 1


def is_ms_word_available() -> Tuple[bool, str]:
    """
    Checks if Microsoft Word can be automated via COM (pywin32).
    Returns (True, version_string) or (False, reason).
    """
    if sys.platform != "win32":
        return False, "Microsoft Word COM is only available on Windows."

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False, "pywin32 package is not installed."

    try:
        pythoncom.CoInitialize()
        try:
            word = win32com.client.DispatchEx("Word.Application")
            version = str(getattr(word, "Version", "Unknown"))
            word.Quit()
            return True, f"Microsoft Word {version}"
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        return False, f"Word COM automation failed: {e}"


def get_libreoffice_path() -> Optional[str]:
    """
    Finds the path to LibreOffice's soffice/libreoffice executable if installed.
    Works on Windows, Linux, and macOS.
    """
    # 1. Check PATH
    for bin_name in ["soffice", "libreoffice"]:
        which_path = shutil.which(bin_name)
        if which_path and os.path.isfile(which_path):
            return which_path

    # 2. Check standard Linux/Unix paths
    linux_paths = [
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/usr/lib/libreoffice/program/soffice"
    ]
    for lp in linux_paths:
        if os.path.isfile(lp):
            return lp

    # 3. Check standard Windows installation directories
    if sys.platform == "win32":
        candidate_paths = [
            os.path.expandvars(r"%ProgramFiles%\LibreOffice\program\soffice.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\LibreOffice\program\soffice.exe"),
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in candidate_paths:
            if os.path.isfile(p):
                return p

    return None


def convert_word_to_pdf_com(input_path: str, output_path: str) -> None:
    """
    Converts a Word document (.docx or .doc) to PDF using Microsoft Word COM.
    """
    import pythoncom
    import win32com.client

    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    # 17 represents wdFormatPDF
    wd_format_pdf = 17

    pythoncom.CoInitialize()
    word_app = None
    doc = None
    try:
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0  # wdAlertsNone

        doc = word_app.Documents.Open(input_path, ReadOnly=True, ConfirmConversions=False)
        doc.SaveAs(output_path, FileFormat=wd_format_pdf)
    except Exception as e:
        raise ConversionError(f"Microsoft Word COM conversion failed: {e}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)  # wdDoNotSaveChanges
            except Exception:
                pass
        if word_app is not None:
            try:
                word_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def convert_word_to_pdf_libreoffice(input_path: str, output_path: str, soffice_path: str) -> None:
    """
    Converts a Word document to PDF using LibreOffice in headless mode.
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    dest_dir = os.path.dirname(output_path)

    os.makedirs(dest_dir, exist_ok=True)

    cmd = [
        soffice_path,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        dest_dir,
        input_path,
    ]

    try:
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
    except subprocess.CalledProcessError as e:
        raise ConversionError(f"LibreOffice conversion process error (exit code {e.returncode}): {e.stderr}")
    except Exception as e:
        raise ConversionError(f"Failed to execute LibreOffice: {e}")

    # LibreOffice outputs with default input base name in dest_dir
    input_base_name = Path(input_path).stem
    default_lo_output = os.path.join(dest_dir, f"{input_base_name}.pdf")

    if os.path.exists(default_lo_output) and os.path.abspath(default_lo_output) != output_path:
        # If output_path is different (e.g., unique name like 'doc (1).pdf'), rename it
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(default_lo_output, output_path)

    if not os.path.exists(output_path):
        raise ConversionError("LibreOffice finished but the output PDF was not found.")


def convert_word_to_pdf(input_path: str, output_path: str) -> str:
    """
    Converts Word (.docx / .doc) to PDF (.pdf).
    Tries Microsoft Word COM first; falls back to LibreOffice headless if Word is unavailable.
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    if not os.path.isfile(input_path):
        raise ConversionError(f"Input file does not exist: {input_path}")

    ext = Path(input_path).suffix.lower()
    if ext not in [".docx", ".doc"]:
        raise ConversionError(f"Unsupported input file extension '{ext}'. Expected .docx or .doc")

    # Create destination directory if needed
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 1. Try Microsoft Word COM automation
    word_available, word_info = is_ms_word_available()
    if word_available:
        try:
            convert_word_to_pdf_com(input_path, output_path)
            if os.path.isfile(output_path) and os.path.getsize(output_path) > 0:
                return output_path
        except Exception as e:
            # If COM failed, try LibreOffice fallback
            lo_path = get_libreoffice_path()
            if lo_path:
                convert_word_to_pdf_libreoffice(input_path, output_path, lo_path)
                return output_path
            raise e

    # 2. Try LibreOffice headless
    lo_path = get_libreoffice_path()
    if lo_path:
        convert_word_to_pdf_libreoffice(input_path, output_path, lo_path)
        return output_path

    # Neither available
    raise ConversionError(
        "Neither Microsoft Word nor LibreOffice is available on this system.\n\n"
        "To convert Word documents to PDF locally, please ensure either:\n"
        "1. Microsoft Office (Word) is installed with pywin32, OR\n"
        "2. LibreOffice (free & open-source) is installed."
    )


def convert_pdf_to_word(input_path: str, output_path: str) -> str:
    """
    Converts PDF (.pdf) to Word (.docx) using pdf2docx.
    Preserves text, paragraphs, tables, images, and formatting locally.
    """
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    if not os.path.isfile(input_path):
        raise ConversionError(f"Input file does not exist: {input_path}")

    ext = Path(input_path).suffix.lower()
    if ext != ".pdf":
        raise ConversionError(f"Unsupported input file extension '{ext}'. Expected .pdf")

    # Create destination directory if needed
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    try:
        from pdf2docx import Converter
    except ImportError:
        raise ConversionError(
            "The 'pdf2docx' package is not installed.\n"
            "Please run: pip install pdf2docx"
        )

    cv = None
    try:
        cv = Converter(input_path)
        cv.convert(output_path)
    except Exception as e:
        raise ConversionError(f"PDF to Word conversion failed: {e}")
    finally:
        if cv is not None:
            try:
                cv.close()
            except Exception:
                pass

    if not os.path.isfile(output_path) or os.path.getsize(output_path) == 0:
        raise ConversionError("Conversion completed but output Word document was not created.")

    return output_path


def get_default_output_path(input_path: str, destination_dir: str, target_format: str) -> str:
    """
    Generates standard output file path based on input file name and target format.
    target_format: 'pdf' or 'docx'
    """
    base_stem = Path(input_path).stem
    ext = f".{target_format.lower().lstrip('.')}"
    filename = f"{base_stem}{ext}"
    return os.path.join(destination_dir, filename)


if __name__ == "__main__":
    import argparse

    # Ensure UTF-8 output on Windows console
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="Local Word <-> PDF Converter CLI (100% Offline)")
    parser.add_argument("input", help="Path to input file (.docx, .doc, or .pdf)")
    parser.add_argument("-o", "--output", help="Optional output file path or destination directory")
    parser.add_argument("-f", "--force", action="store_true", help="Overwrite output file if it exists")

    args = parser.parse_args()
    input_file = os.path.abspath(args.input)

    if not os.path.isfile(input_file):
        print(f"Error: Input file does not exist: {input_file}", file=sys.stderr)
        sys.exit(1)

    in_ext = Path(input_file).suffix.lower()
    if in_ext in [".docx", ".doc"]:
        target_ext = "pdf"
        convert_func = convert_word_to_pdf
    elif in_ext == ".pdf":
        target_ext = "docx"
        convert_func = convert_pdf_to_word
    else:
        print(f"Error: Unsupported input file extension '{in_ext}'. Expected .docx, .doc, or .pdf", file=sys.stderr)
        sys.exit(1)

    if args.output:
        out_cand = os.path.abspath(args.output)
        if os.path.isdir(out_cand):
            out_file = get_default_output_path(input_file, out_cand, target_ext)
        else:
            out_file = out_cand
    else:
        out_file = get_default_output_path(input_file, os.path.dirname(input_file), target_ext)

    if os.path.exists(out_file) and not args.force:
        unique_file = get_unique_filepath(out_file)
        print(f"Notice: '{os.path.basename(out_file)}' exists. Saving as unique file '{os.path.basename(unique_file)}'.")
        out_file = unique_file

    print(f"Converting: {input_file} -> {out_file}")
    try:
        res = convert_func(input_file, out_file)
        print(f"Success! Output file: {res}")
    except Exception as err:
        print(f"Conversion Error: {err}", file=sys.stderr)
        sys.exit(1)

