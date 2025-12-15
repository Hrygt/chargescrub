// ChargeScrub JavaScript v1.0.0
// E/M Charge Duplicate & Overlap Review Tool

const API_ENDPOINT = '/api/upload';
let selectedFile = null;
let processedBlob = null;
let processedFilename = null;

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const clearFileBtn = document.getElementById('clearFileBtn');
const processBtn = document.getElementById('processBtn');
const resultsSection = document.getElementById('resultsSection');
const downloadBtn = document.getElementById('downloadBtn');
const downloadContainer = document.getElementById('downloadContainer');
const progressOverlay = document.getElementById('progress-overlay');
const progressFill = document.getElementById('progress-fill');
const progressTitle = document.getElementById('progress-title');
const progressMessage = document.getElementById('progress-message');
const apiStatus = document.getElementById('apiStatus');

// Toast notification
function showToast(message) {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}

// Format file size
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// Health check
async function checkHealth() {
    apiStatus.className = 'status-dot checking';
    try {
        const resp = await fetch('/api/health', { method: 'GET' });
        if (resp.ok) {
            apiStatus.className = 'status-dot';
        } else {
            apiStatus.className = 'status-dot offline';
        }
    } catch (e) {
        apiStatus.className = 'status-dot offline';
    }
}

// Check health on load and periodically
checkHealth();
setInterval(checkHealth, 30000);

// File selection handlers
function handleFileSelect(file) {
    if (!file) return;
    
    // Validate file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    const validExtensions = ['.xlsx', '.xls'];
    const hasValidExtension = validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
    
    if (!validTypes.includes(file.type) && !hasValidExtension) {
        showToast('Please select an Excel file (.xlsx or .xls)');
        return;
    }
    
    // Validate file size (max 10MB)
    if (file.size > 10 * 1024 * 1024) {
        showToast('File too large. Maximum size is 10MB.');
        return;
    }
    
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    
    dropZone.classList.add('has-file');
    fileInfo.style.display = 'flex';
    processBtn.disabled = false;
    
    // Reset results
    resultsSection.style.display = 'none';
    downloadContainer.style.display = 'none';
    processedBlob = null;
}

// Drop zone events
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelect(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

// Clear file
clearFileBtn.addEventListener('click', () => {
    selectedFile = null;
    processedBlob = null;
    fileInput.value = '';
    dropZone.classList.remove('has-file');
    fileInfo.style.display = 'none';
    processBtn.disabled = true;
    resultsSection.style.display = 'none';
    downloadContainer.style.display = 'none';
});

// Show/hide progress
function showProgress(show) {
    if (show) {
        progressOverlay.classList.add('active');
        progressFill.style.width = '0%';
    } else {
        progressOverlay.classList.remove('active');
    }
}

function updateProgress(percent, title, message) {
    progressFill.style.width = percent + '%';
    if (title) progressTitle.textContent = title;
    if (message) progressMessage.textContent = message;
}

// Process file
processBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    
    showProgress(true);
    updateProgress(10, 'Uploading file...', 'Sending to server');
    
    try {
        // Create FormData
        const formData = new FormData();
        formData.append('file', selectedFile);
        
        updateProgress(30, 'Processing charges...', 'Analyzing E/M code combinations');
        
        // Send to API
        const response = await fetch(API_ENDPOINT, {
            method: 'POST',
            body: formData
        });
        
        updateProgress(70, 'Processing charges...', 'Applying billing rules');
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.message || errorData.error || 'Processing failed');
        }
        
        updateProgress(90, 'Finalizing...', 'Preparing download');
        
        // Get stats from headers
        const stats = {
            total: response.headers.get('X-Stats-Total') || '-',
            charge: response.headers.get('X-Stats-Charge') || '-',
            deny: response.headers.get('X-Stats-Deny') || '-',
            review: response.headers.get('X-Stats-Review') || '-',
            patients: response.headers.get('X-Stats-Patients') || '-'
        };
        
        // Get filename from Content-Disposition
        const disposition = response.headers.get('Content-Disposition');
        if (disposition && disposition.includes('filename=')) {
            processedFilename = disposition.split('filename=')[1].replace(/"/g, '');
        } else {
            processedFilename = 'ChargeScrub_Output.xlsx';
        }
        
        // Get the blob
        processedBlob = await response.blob();
        
        updateProgress(100, 'Complete!', 'Analysis finished');
        
        // Update stats display
        document.getElementById('statTotal').textContent = stats.total;
        document.getElementById('statCharge').textContent = stats.charge;
        document.getElementById('statDeny').textContent = stats.deny;
        document.getElementById('statReview').textContent = stats.review;
        document.getElementById('statPatients').textContent = stats.patients;
        
        // Show results
        resultsSection.style.display = 'block';
        downloadContainer.style.display = 'block';
        
        setTimeout(() => {
            showProgress(false);
            // Scroll to download button
            downloadContainer.scrollIntoView({ behavior: 'smooth', block: 'center' });
            showToast('✓ Processing complete! Download your results below.');
        }, 500);
        
    } catch (error) {
        showProgress(false);
        showToast('Error: ' + error.message);
        console.error('Processing error:', error);
    }
});

// Download results
downloadBtn.addEventListener('click', () => {
    if (!processedBlob) {
        showToast('No processed file available');
        return;
    }
    
    const url = URL.createObjectURL(processedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = processedFilename || 'ChargeScrub_Output.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    showToast('Download started!');
});

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    // Ctrl+O to open file
    if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        fileInput.click();
    }
    // Ctrl+Enter to process
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        if (!processBtn.disabled) {
            processBtn.click();
        }
    }
});
