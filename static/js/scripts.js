const uploadButton1 = document.getElementById("uploadButton1");
const generateButton1 = document.getElementById("generateButton1");
const fileInput1 = document.getElementById("fileInput1");
const fileNameDisplay1 = document.getElementById("fileName1");
const downloadButton1 = document.getElementById("downloadButton1");

const uploadButton2 = document.getElementById("uploadButton2");
const generateButton2 = document.getElementById("generateButton2");
const fileInput2 = document.getElementById("fileInput2");
const fileNameDisplay2 = document.getElementById("fileName2");
const downloadButton2 = document.getElementById("downloadButton2");

const uploadButton3 = document.getElementById("uploadButton3");
const fileInput3 = document.getElementById("fileInput3");
const fileNameDisplay3 = document.getElementById("fileName3");

const uploadSyllabusButton = document.getElementById("uploadSyllabusButton");
const syllabusFileInput = document.getElementById("syllabusFileInput");
const syllabusFileNameDisplay = document.getElementById("syllabusFileName");

const coInputs = document.querySelectorAll("#coInputs input");
const messageDisplay = document.getElementById("message");
const ctMessageDisplay = document.getElementById("ctMessage"); 

let selectedFiles1 = new Set();
let selectedFiles2 = new Set();
let selectedExcelFile = null;
let selectedSyllabusFile = null; 
let ctTemplate_File = null;
let tlp_File = null;

function updateGenerateButton1State() {
    if (selectedFiles1.size > 0 && selectedSyllabusFile !== null) {
        generateButton1.disabled = false;
    } else {
        generateButton1.disabled = true;
    }
}

function handleSyllabusFileSelection(fileInput, fileNameDisplay) {
    const file = fileInput.files[0];
    
    if (file) {
        const allowedExtensions = [".pdf", ".doc", ".docx", ".txt"];
        const isValid = allowedExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
        
        if (isValid) {
            selectedSyllabusFile = file;
            fileNameDisplay.innerHTML = `Selected File:<br>${file.name}`;
            updateGenerateButton1State(); 
        } else {
            alert("Only PDF, DOC, DOCX, or TXT files are allowed for syllabus.");
            selectedSyllabusFile = null;
            fileNameDisplay.textContent = "";
            updateGenerateButton1State();
        }
    } else {
        selectedSyllabusFile = null;
        fileNameDisplay.textContent = "";
        updateGenerateButton1State();
    }
}

function handleFileSelection(fileInput, fileNameDisplay, selectedFiles, generateButton) {
    const files = Array.from(fileInput.files);

    files.forEach(file => {
        const allowedExtensions = fileInput === fileInput2
            ? [".pdf", ".xls", ".xlsx"]
            : [".pdf"];
    
        const isValid = allowedExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
    
        if (isValid) {
            selectedFiles.add(file);
        } else {
            alert(`Only ${allowedExtensions.join(", ")} files are allowed.`);
        }
    });
        
    if (selectedFiles.size > 0) {
        fileNameDisplay.innerHTML = `Selected Files:<br><ul>${[...selectedFiles]
            .map(file => `<li>${file.name}</li>`).join("")}</ul>`;
        
        if (fileInput === fileInput1) {
            updateGenerateButton1State();
        } else {
            generateButton.disabled = false;
        }
    } else {
        fileNameDisplay.textContent = "";
        generateButton.disabled = true;
    }
}

function handleExcelFileSelection(fileInput, fileNameDisplay) {
    const file = fileInput.files[0];
    
    if (file) {
        const allowedExtensions = [".xls", ".xlsx"];
        const isValid = allowedExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
        
        if (isValid) {
            selectedExcelFile = file;
            fileNameDisplay.innerHTML = `Selected File:<br>${file.name}`;
        } else {
            alert("Only .xlsx and .xls files are allowed.");
            selectedExcelFile = null;
            fileNameDisplay.textContent = "";
        }
    } else {
        selectedExcelFile = null;
        fileNameDisplay.textContent = "";
    }
}

uploadButton1.addEventListener("click", () => fileInput1.click());
fileInput1.addEventListener("change", () => handleFileSelection(fileInput1, fileNameDisplay1, selectedFiles1, generateButton1));

uploadButton2.addEventListener("click", () => fileInput2.click());
fileInput2.addEventListener("change", () => handleFileSelection(fileInput2, fileNameDisplay2, selectedFiles2, generateButton2));

uploadButton3.addEventListener("click", () => fileInput3.click());
fileInput3.addEventListener("change", () => handleExcelFileSelection(fileInput3, fileNameDisplay3));

uploadSyllabusButton.addEventListener("click", () => syllabusFileInput.click());
syllabusFileInput.addEventListener("change", () => handleSyllabusFileSelection(syllabusFileInput, syllabusFileNameDisplay));

async function handleCTFileUpload(selectedFiles, url, messageElement, downloadButton) {
    if (selectedFiles.size === 0) {
        messageElement.textContent = "No files selected!";
        messageElement.style.display = "block";
        messageElement.style.color = "red";
        return;
    }

    if (!selectedSyllabusFile) {
        messageElement.textContent = "Please upload a syllabus file!";
        messageElement.style.display = "block";
        messageElement.style.color = "red";
        return;
    }

    messageElement.textContent = "Uploading files... Please wait.";
    messageElement.style.display = "block";
    messageElement.style.color = "blue";

    downloadButton.style.display = "none";
    
    const formData = new FormData();
    selectedFiles.forEach(file => formData.append("pdf_files", file));
    
    formData.append("syllabus_file", selectedSyllabusFile);

    try {
        const response = await fetch(url, {
            method: "POST",
            body: formData,
        });

        const data = await response.json();
        
        if (data.success) {
            messageElement.textContent = "Excel file created successfully!";
            messageElement.style.display = "block";
            messageElement.style.color = "green";
            downloadButton.style.display = "inline";
            
            selectedFiles.clear();
            selectedSyllabusFile = null;
            fileNameDisplay1.textContent = "";
            syllabusFileNameDisplay.textContent = "";
            fileInput1.value = "";
            syllabusFileInput.value = "";
            updateGenerateButton1State();
            
            ctTemplate_File = data.download_url;
        } else {
            throw new Error(data.message || "Upload failed");
        }
    } catch (error) {
        console.error("Upload error:", error);
        messageElement.textContent = "File upload failed: " + error.message;
        messageElement.style.display = "block";
        messageElement.style.color = "red";
    }
}

async function handleFileUpload(selectedFiles, url, messageElement, downloadButton, includeCO = false) {
    if (selectedFiles.size === 0) {
        messageElement.textContent = "No files selected!";
        messageElement.style.color = "red";
        return;
    }

    messageElement.textContent = "Uploading files... Please wait.";
    messageElement.style.color = "blue";

    downloadButton.style.display = "none";
    
    const formData = new FormData();
    selectedFiles.forEach(file => formData.append("pdf_files", file));

    if (includeCO) {
        coInputs.forEach(input => formData.append(input.id, input.value));
    }

    try {
        const response = await fetch(url, {
            method: "POST",
            body: formData,
        });

        const data = await response.json();
        
        if (data.success) {
            messageElement.textContent = "Excel file created successfully!";
            messageElement.style.color = "green";
            downloadButton.style.display = "inline";
            selectedFiles.clear(); 
        } else {
            throw new Error(data.message || "Upload failed");
        }
    } catch (error) {
        console.error("Upload error:", error);
        messageElement.textContent = "File upload failed: " + error.message;
        messageElement.style.color = "red";
    }
}

async function handleTLPUpload(selectedFiles, selectedExcelFile, url, messageElement, downloadButton) {
    if (selectedFiles.size === 0) {
        messageElement.textContent = "No TLP files selected!";
        messageElement.style.color = "red";
        return;
    }

    messageElement.textContent = "Processing TLP files... Please wait.";
    messageElement.style.color = "blue";

    downloadButton.style.display = "none";
    
    const formData = new FormData();
    selectedFiles.forEach(file => formData.append("pdf_files", file));

    coInputs.forEach(input => formData.append(input.id, input.value));

    if (selectedExcelFile) {
        formData.append("co_filled_excel", selectedExcelFile); 
        messageElement.textContent = "Processing TLP files with Excel sheet... Please wait.";
    }

    try {
        const response = await fetch(url, {
            method: "POST",
            body: formData,
        });

        const data = await response.json();
        
        if (data.success) {
            const sheetsText = selectedExcelFile ? "Excel file created successfully!" : "TLP Excel file created successfully!";
            messageElement.textContent = sheetsText;
            messageElement.style.color = "green";
            downloadButton.style.display = "inline";
            tlp_File = data.download_url;
            
            selectedFiles.clear();
            fileNameDisplay2.textContent = "";
            fileInput2.value = "";
            generateButton2.disabled = true;
            
            if (selectedExcelFile) {
                selectedExcelFile = null;
                fileNameDisplay3.textContent = "";
                fileInput3.value = "";
            }
        } else {
            throw new Error(data.message || "Upload failed");
        }
    } catch (error) {
        console.error("TLP upload error:", error);
        messageElement.textContent = "TLP processing failed: " + error.message;
        messageElement.style.color = "red";
    }
}

async function handleExcelFileUpload(selectedFile, url, messageElement, downloadButton) {
    if (!selectedFile) {
        messageElement.textContent = "No Excel file selected!";
        messageElement.style.color = "red";
        return;
    }

    messageElement.textContent = "Processing Excel file... Please wait.";
    messageElement.style.color = "blue";

    downloadButton.style.display = "none";
    
    const formData = new FormData();
    formData.append("co_filled_excel", selectedFile); 

    try {
        const response = await fetch(url, {
            method: "POST",
            body: formData,
        });

        const data = await response.json();
        
        if (data.success) {
            messageElement.textContent = "Excel file processed successfully!";
            messageElement.style.color = "green";
            downloadButton.style.display = "inline";

            selectedExcelFile = null;
            fileNameDisplay3.textContent = "";
        } else {
            throw new Error(data.message || "Processing failed");
        }
    } catch (error) {
        console.error("Excel processing error:", error);
        messageElement.textContent = "Excel processing failed: " + error.message;
        messageElement.style.color = "red";
    }
}

generateButton1.addEventListener("click", () => handleCTFileUpload(selectedFiles1, "/upload1", ctMessageDisplay, downloadButton1));
generateButton2.addEventListener("click", () => handleTLPUpload(selectedFiles2, selectedExcelFile, "/upload2", messageDisplay, downloadButton2));

async function downloadExcelFile(filePath, fileNameDisplay) {
    try {
        const response = await fetch(filePath);
        
        if (!response.ok) {
            throw new Error("File not found");
        }

        const blob = await response.blob();
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);

        const contentDisposition = response.headers.get("Content-Disposition");
        let filename = filePath.split("/").pop(); 

        if (contentDisposition) {
            const utf8Match = contentDisposition.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
            const asciiMatch = contentDisposition.match(/filename\s*=\s*["']?([^;"']+)/i);

            if (utf8Match) {
                filename = decodeURIComponent(utf8Match[1]);
            } else if (asciiMatch) {
                filename = asciiMatch[1];
            }
        }

        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        fileNameDisplay.innerHTML = '';
        fileNameDisplay.textContent = '';

    } catch (error) {
        console.error("Download failed:", error);
        messageDisplay.textContent = "Download failed!";
        messageDisplay.style.color = "red";
    }
}

downloadButton1.addEventListener("click", () => downloadExcelFile(ctTemplate_File, fileNameDisplay1));
downloadButton2.addEventListener("click", () => downloadExcelFile(tlp_File, fileNameDisplay2));