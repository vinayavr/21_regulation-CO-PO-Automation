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

const coInputs = document.querySelectorAll("#coInputs input");
const messageDisplay = document.getElementById("message");
const ctMessageDisplay = document.getElementById("ctMessage"); 

let selectedFiles1 = new Set();
let selectedFiles2 = new Set();
let selectedExcelFile = null;

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
        fileNameDisplay.innerHTML = `Selected Files: <ul>${[...selectedFiles]
            .map(file => `<li>${file.name}</li>`).join("")}</ul>`;
        generateButton.disabled = false;
    } else {
        fileNameDisplay.textContent = "No valid files selected.";
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
            fileNameDisplay.innerHTML = `Selected Excel File: <strong>${file.name}</strong>`;
        } else {
            alert("Only .xlsx and .xls files are allowed.");
            selectedExcelFile = null;
            fileNameDisplay.textContent = "No valid file selected.";
        }
    } else {
        selectedExcelFile = null;
        fileNameDisplay.textContent = "No file selected.";
    }
}

uploadButton1.addEventListener("click", () => fileInput1.click());
fileInput1.addEventListener("change", () => handleFileSelection(fileInput1, fileNameDisplay1, selectedFiles1, generateButton1));

uploadButton2.addEventListener("click", () => fileInput2.click());
fileInput2.addEventListener("change", () => handleFileSelection(fileInput2, fileNameDisplay2, selectedFiles2, generateButton2));

uploadButton3.addEventListener("click", () => fileInput3.click());
fileInput3.addEventListener("change", () => handleExcelFileSelection(fileInput3, fileNameDisplay3));

async function handleCTFileUpload(selectedFiles, url, messageElement, downloadButton) {
    if (selectedFiles.size === 0) {
        messageElement.textContent = "No files selected!";
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
            const sheetsText = selectedExcelFile ? "Excel file with 2 sheets created successfully!" : "TLP Excel file created successfully!";
            messageElement.textContent = sheetsText;
            messageElement.style.color = "green";
            downloadButton.style.display = "inline";

            selectedFiles.clear();
            if (selectedExcelFile) {
                selectedExcelFile = null;
                fileNameDisplay3.textContent = "";
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

async function downloadExcelFile(filePath) {
    try {
        const response = await fetch(filePath);
        
        if (!response.ok) {
            throw new Error("File not found");
        }

        const blob = await response.blob();
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);

        const contentDisposition = response.headers.get("Content-Disposition");
        const filename = contentDisposition ? contentDisposition.split("filename=")[1] : filePath.split("/").pop();
        link.download = filename.replace(/["']/g, "");

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } catch (error) {
        console.error("Download failed:", error);
        messageDisplay.textContent = "Download failed!";
        messageDisplay.style.color = "red";
    }
}

downloadButton1.addEventListener("click", () => downloadExcelFile("/download/CT_Template.xlsx"));
downloadButton2.addEventListener("click", () => downloadExcelFile("/download/co_allocation.xlsx"));