/**
 * DocMorph — Interactive Web Application Logic
 */

document.addEventListener("DOMContentLoaded", () => {
    // Elements
    const tabW2P = document.getElementById("tab-w2p");
    const tabP2W = document.getElementById("tab-p2w");
    const uploadZone = document.getElementById("upload-zone");
    const fileInput = document.getElementById("file-input");
    const uploadPrompt = document.getElementById("upload-prompt");
    const filePreview = document.getElementById("file-preview");
    const previewFilename = document.getElementById("preview-filename");
    const previewFilesize = document.getElementById("preview-filesize");
    const btnRemoveFile = document.getElementById("btn-remove-file");
    const btnConvert = document.getElementById("btn-convert");
    const btnText = document.getElementById("btn-text");
    const actionSection = document.getElementById("action-section");
    const progressSection = document.getElementById("progress-section");
    const progressStatusText = document.getElementById("progress-status-text");
    const progressPercentText = document.getElementById("progress-percent-text");
    const progressBarFill = document.getElementById("progress-bar-fill");
    const resultSection = document.getElementById("result-section");
    const resultMetaText = document.getElementById("result-meta-text");
    const downloadLink = document.getElementById("download-link");
    const btnReset = document.getElementById("btn-reset");
    const errorBox = document.getElementById("error-box");
    const errorMessage = document.getElementById("error-message");
    const formatHint = document.getElementById("format-hint");
    const badgeContainer = document.getElementById("badge-container");
    const engineStatusText = document.getElementById("engine-status-text");

    // State
    let currentMode = "word-to-pdf"; // 'word-to-pdf' or 'pdf-to-word'
    let selectedFile = null;

    // 1. Check System Engine Status
    fetch("/api/status")
        .then(res => res.json())
        .then(data => {
            let parts = [];
            if (data.word_com_available) {
                parts.push(data.word_com_info);
            } else if (data.libreoffice_available) {
                parts.push("LibreOffice Headless");
            } else {
                parts.push("Word Engine: Needs Word/LibreOffice");
            }
            parts.push("PDF Engine: pdf2docx");
            engineStatusText.textContent = parts.join(" • ");
        })
        .catch(() => {
            engineStatusText.textContent = "Conversion engine ready";
        });

    // 2. Mode Switching
    function setMode(mode, autoSwitched = false) {
        currentMode = mode;
        hideError();

        if (mode === "word-to-pdf") {
            tabW2P.classList.add("active");
            tabP2W.classList.remove("active");
            fileInput.accept = ".docx,.doc";
            formatHint.textContent = "Supports Microsoft Word documents (.docx, .doc)";
            badgeContainer.innerHTML = '<span class="format-badge">.DOCX</span><span class="format-badge">.DOC</span>';
            btnText.textContent = "Convert to PDF";
        } else {
            tabP2W.classList.add("active");
            tabW2P.classList.remove("active");
            fileInput.accept = ".pdf";
            formatHint.textContent = "Supports PDF documents (.pdf)";
            badgeContainer.innerHTML = '<span class="format-badge">.PDF</span>';
            btnText.textContent = "Convert to Word";
        }

        // Validate currently loaded file if any
        if (selectedFile && !autoSwitched) {
            validateFile(selectedFile);
        }
    }

    tabW2P.addEventListener("click", () => setMode("word-to-pdf"));
    tabP2W.addEventListener("click", () => setMode("pdf-to-word"));

    // 3. Drag and Drop Handling
    ["dragenter", "dragover"].forEach(event => {
        uploadZone.addEventListener(event, (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadZone.classList.add("dragover");
        });
    });

    ["dragleave", "drop"].forEach(event => {
        uploadZone.addEventListener(event, (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadZone.classList.remove("dragover");
        });
    });

    uploadZone.addEventListener("drop", (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileSelect(files[0]);
        }
    });

    uploadZone.addEventListener("click", (e) => {
        if (!e.target.closest("#btn-remove-file") && !filePreview.contains(e.target)) {
            fileInput.click();
        }
    });

    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    // 4. File Selection & Validation
    function handleFileSelect(file) {
        hideError();
        const ext = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();

        // Auto-switch mode based on file extension
        if ((ext === ".docx" || ext === ".doc") && currentMode !== "word-to-pdf") {
            setMode("word-to-pdf", true);
        } else if (ext === ".pdf" && currentMode !== "pdf-to-word") {
            setMode("pdf-to-word", true);
        }

        if (validateFile(file)) {
            selectedFile = file;
            previewFilename.textContent = file.name;
            previewFilesize.textContent = formatBytes(file.size);

            uploadPrompt.classList.add("hidden");
            filePreview.classList.remove("hidden");
            btnConvert.disabled = false;
        }
    }

    function validateFile(file) {
        const ext = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();

        if (currentMode === "word-to-pdf" && !['.docx', '.doc'].includes(ext)) {
            showError(`Invalid file format '${ext}'. Please upload a .docx or .doc file.`);
            resetFileInput();
            return false;
        }

        if (currentMode === "pdf-to-word" && ext !== '.pdf') {
            showError(`Invalid file format '${ext}'. Please upload a .pdf file.`);
            resetFileInput();
            return false;
        }

        return true;
    }

    btnRemoveFile.addEventListener("click", (e) => {
        e.stopPropagation();
        resetFileInput();
    });

    function resetFileInput() {
        selectedFile = null;
        fileInput.value = "";
        uploadPrompt.classList.remove("hidden");
        filePreview.classList.add("hidden");
        btnConvert.disabled = true;
    }

    function showError(msg) {
        errorMessage.textContent = msg;
        errorBox.classList.remove("hidden");
    }

    function hideError() {
        errorBox.classList.add("hidden");
    }

    function formatBytes(bytes) {
        if (bytes === 0) return "0 Bytes";
        const k = 1024;
        const sizes = ["Bytes", "KB", "MB", "GB"];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
    }

    // 5. Conversion Trigger
    btnConvert.addEventListener("click", () => {
        if (!selectedFile) return;

        hideError();
        actionSection.classList.add("hidden");
        uploadZone.style.pointerEvents = "none";
        uploadZone.style.opacity = "0.7";
        progressSection.classList.remove("hidden");
        resultSection.classList.add("hidden");

        updateProgress(15, "Uploading file to conversion engine...");

        const formData = new FormData();
        formData.append("file", selectedFile);
        formData.append("mode", currentMode);

        const xhr = new XMLHttpRequest();
        xhr.open("POST", "/api/convert", true);

        // Upload progress
        xhr.upload.onprogress = (e) => {
            if (e.lengthComputable) {
                const percent = Math.round((e.loaded / e.total) * 40);
                updateProgress(percent, "Uploading file...");
            }
        };

        xhr.onload = () => {
            if (xhr.status >= 200 && xhr.status < 300) {
                try {
                    const response = JSON.parse(xhr.responseText);
                    if (response.success) {
                        updateProgress(100, "Conversion complete!");
                        setTimeout(() => {
                            showResult(response);
                        }, 400);
                    } else {
                        handleFailure(response.error || "Conversion failed.");
                    }
                } catch (err) {
                    handleFailure("Malformed server response.");
                }
            } else {
                try {
                    const response = JSON.parse(xhr.responseText);
                    handleFailure(response.error || `Server error (${xhr.status})`);
                } catch {
                    handleFailure(`Server returned error code ${xhr.status}`);
                }
            }
        };

        xhr.onerror = () => {
            handleFailure("Network error communicating with conversion server.");
        };

        // Advance progress animation while processing
        let processingStep = 45;
        const processingInterval = setInterval(() => {
            if (processingStep < 90) {
                processingStep += 10;
                const statusMsg = currentMode === "word-to-pdf" 
                    ? "Rendering PDF layout & styles..." 
                    : "Extracting paragraphs & tables...";
                updateProgress(processingStep, statusMsg);
            } else {
                clearInterval(processingInterval);
            }
        }, 500);

        xhr.onloadend = () => {
            clearInterval(processingInterval);
        };

        xhr.send(formData);
    });

    function updateProgress(percent, text) {
        progressBarFill.style.width = `${percent}%`;
        progressPercentText.textContent = `${percent}%`;
        progressStatusText.textContent = text;
    }

    function showResult(data) {
        progressSection.classList.add("hidden");
        resultSection.classList.remove("hidden");
        uploadZone.classList.add("hidden");

        resultMetaText.textContent = `${data.filename} (${data.file_size})`;
        downloadLink.href = data.download_url;
        downloadLink.setAttribute("download", data.filename);
    }

    function handleFailure(err) {
        progressSection.classList.add("hidden");
        actionSection.classList.remove("hidden");
        uploadZone.style.pointerEvents = "auto";
        uploadZone.style.opacity = "1";
        showError(err);
    }

    // 6. Reset
    btnReset.addEventListener("click", () => {
        resetFileInput();
        resultSection.classList.add("hidden");
        uploadZone.classList.remove("hidden");
        uploadZone.style.pointerEvents = "auto";
        uploadZone.style.opacity = "1";
        actionSection.classList.remove("hidden");
    });
});
