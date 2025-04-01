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

const coInputs = document.querySelectorAll("#coInputs input");

const messageDisplay = document.getElementById("message");

let selectedFiles1 = new Set();
let selectedFiles2 = new Set();

function handleFileSelection(fileInput, fileNameDisplay, selectedFiles, generateButton) {
    const files = Array.from(fileInput.files);

    files.forEach(file => {
        if (file.name.endsWith(".pdf")) {
            selectedFiles.add(file);
        } else {
            alert("Only PDF files are allowed.");
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

uploadButton1.addEventListener("click", () => fileInput1.click());
fileInput1.addEventListener("change", () => handleFileSelection(fileInput1, fileNameDisplay1, selectedFiles1, generateButton1));

uploadButton2.addEventListener("click", () => fileInput2.click());
fileInput2.addEventListener("change", () => handleFileSelection(fileInput2, fileNameDisplay2, selectedFiles2, generateButton2));

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

generateButton1.addEventListener("click", () => handleFileUpload(selectedFiles1, "/upload1", messageDisplay, downloadButton1));

generateButton2.addEventListener("click", () => handleFileUpload(selectedFiles2, "/upload2", messageDisplay, downloadButton2, true));

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

downloadButton1.addEventListener("click", () => downloadExcelFile("/download/result.xlsx"));
downloadButton2.addEventListener("click", () => downloadExcelFile("/download/co_allocation.xlsx"));
