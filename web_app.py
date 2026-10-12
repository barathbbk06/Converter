"""
Web Application Server for Local & Cloud Word <-> PDF Converter
---------------------------------------------------------------
Flask-based web server supporting both local desktop web browsing
and 100% containerized cloud hosting (Render, Railway, Docker, etc.).
"""

import os
import sys
import time
import uuid
import shutil
import threading
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename

# Import core conversion engine
from converter import (
    convert_word_to_pdf,
    convert_pdf_to_word,
    is_ms_word_available,
    get_libreoffice_path,
    ConversionError
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB maximum upload limit

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "temp_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


def cleanup_old_files():
    """Background cleanup worker to delete converted files older than 30 minutes."""
    while True:
        try:
            now = time.time()
            if os.path.exists(UPLOAD_DIR):
                for task_id in os.listdir(UPLOAD_DIR):
                    task_path = os.path.join(UPLOAD_DIR, task_id)
                    if os.path.isdir(task_path):
                        # If folder is older than 30 minutes (1800s), remove it
                        if now - os.path.getmtime(task_path) > 1800:
                            shutil.rmtree(task_path, ignore_errors=True)
        except Exception:
            pass
        time.sleep(300)  # Check every 5 minutes


# Start background cleaner thread
cleanup_thread = threading.Thread(target=cleanup_old_files, daemon=True)
cleanup_thread.start()


@app.route("/")
def index():
    """Render the main web interface."""
    return render_template("index.html")


@app.route("/api/status", methods=["GET"])
def api_status():
    """Returns detected system engines (MS Word, LibreOffice, pdf2docx)."""
    word_available, word_info = is_ms_word_available()
    lo_path = get_libreoffice_path()

    return jsonify({
        "status": "online",
        "platform": sys.platform,
        "word_com_available": word_available,
        "word_com_info": word_info,
        "libreoffice_available": bool(lo_path),
        "libreoffice_path": lo_path if lo_path else "Not detected",
        "pdf2docx_available": True
    })


@app.route("/api/convert", methods=["POST"])
def api_convert():
    """
    Accepts uploaded file and converts between Word and PDF.
    Expects multipart/form-data:
      - file: the uploaded document
      - mode: 'word-to-pdf' or 'pdf-to-word'
    """
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file uploaded."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"success": False, "error": "No file selected."}), 400

    mode = request.form.get("mode", "word-to-pdf").strip().lower()
    orig_filename = secure_filename(file.filename)
    if not orig_filename:
        orig_filename = "document"

    in_ext = Path(orig_filename).suffix.lower()

    # Validation
    if mode == "word-to-pdf" and in_ext not in [".docx", ".doc"]:
        return jsonify({
            "success": False,
            "error": f"Invalid format '{in_ext}'. Word to PDF mode requires a .docx or .doc file."
        }), 400

    if mode == "pdf-to-word" and in_ext != ".pdf":
        return jsonify({
            "success": False,
            "error": f"Invalid format '{in_ext}'. PDF to Word mode requires a .pdf file."
        }), 400

    # Create task folder
    task_id = str(uuid.uuid4())
    task_folder = os.path.join(UPLOAD_DIR, task_id)
    os.makedirs(task_folder, exist_ok=True)

    input_path = os.path.join(task_folder, orig_filename)
    file.save(input_path)

    base_name = Path(orig_filename).stem
    target_ext = ".pdf" if mode == "word-to-pdf" else ".docx"
    output_filename = f"{base_name}{target_ext}"
    output_path = os.path.join(task_folder, output_filename)

    try:
        if mode == "word-to-pdf":
            convert_word_to_pdf(input_path, output_path)
        else:
            convert_pdf_to_word(input_path, output_path)

        if not os.path.isfile(output_path) or os.path.getsize(output_path) == 0:
            raise ConversionError("Output file was not generated.")

        file_size = os.path.getsize(output_path)
        size_str = f"{file_size / 1024:.1f} KB" if file_size < 1024 * 1024 else f"{file_size / (1024*1024):.2f} MB"

        return jsonify({
            "success": True,
            "task_id": task_id,
            "filename": output_filename,
            "file_size": size_str,
            "download_url": f"/api/download/{task_id}/{output_filename}"
        })

    except Exception as e:
        # Cleanup failed task folder
        shutil.rmtree(task_folder, ignore_errors=True)
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/download/<task_id>/<filename>", methods=["GET"])
def api_download(task_id, filename):
    """Securely downloads the converted document."""
    safe_task_id = secure_filename(task_id)
    safe_filename = secure_filename(filename)

    task_folder = os.path.join(UPLOAD_DIR, safe_task_id)
    target_file = os.path.join(task_folder, safe_filename)

    if not os.path.isfile(target_file):
        return "File not found or link has expired.", 404

    return send_from_directory(task_folder, safe_filename, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"=======================================================")
    print(f"  Starting Local Word <-> PDF Web Application")
    print(f"  Available at: http://127.0.0.1:{port}")
    print(f"=======================================================")
    app.run(host="0.0.0.0", port=port, debug=False)
