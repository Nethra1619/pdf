class PDFToPPTConverter {
    constructor() {
        this.currentFile = null;
        this.initializeElements();
        this.attachEventListeners();
        this.setupDragAndDrop();
    }

    initializeElements() {
        // Get all DOM elements
        this.uploadSection = document.getElementById('uploadSection');
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.browseBtn = document.getElementById('browseBtn');
        this.fileInfo = document.getElementById('fileInfo');
        this.fileName = document.getElementById('fileName');
        this.fileSize = document.getElementById('fileSize');
        this.removeFile = document.getElementById('removeFile');
        this.conversionOptions = document.getElementById('conversionOptions');
        this.convertSection = document.getElementById('convertSection');
        this.convertBtn = document.getElementById('convertBtn');
        this.progressSection = document.getElementById('progressSection');
        this.progressFill = document.getElementById('progressFill');
        this.progressPercentage = document.getElementById('progressPercentage');
        this.progressText = document.getElementById('progressText');
        this.successSection = document.getElementById('successSection');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.convertAnotherBtn = document.getElementById('convertAnotherBtn');
        this.slideLayout = document.getElementById('slideLayout');
        this.imageQuality = document.getElementById('imageQuality');
    }

    attachEventListeners() {
        // File input events
        this.browseBtn.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        this.removeFile.addEventListener('click', () => this.resetConverter());
        
        // Conversion events
        this.convertBtn.addEventListener('click', () => this.startConversion());
        this.downloadBtn.addEventListener('click', () => this.downloadFile());
        this.convertAnotherBtn.addEventListener('click', () => this.resetConverter());
    }

    setupDragAndDrop() {
        // Prevent default drag behaviors
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            this.uploadArea.addEventListener(eventName, this.preventDefaults, false);
            document.body.addEventListener(eventName, this.preventDefaults, false);
        });

        // Highlight drop area when item is dragged over it
        ['dragenter', 'dragover'].forEach(eventName => {
            this.uploadArea.addEventListener(eventName, () => this.highlight(), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            this.uploadArea.addEventListener(eventName, () => this.unhighlight(), false);
        });

        // Handle dropped files
        this.uploadArea.addEventListener('drop', (e) => this.handleDrop(e), false);
    }

    preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    highlight() {
        this.uploadArea.classList.add('dragover');
    }

    unhighlight() {
        this.uploadArea.classList.remove('dragover');
    }

    handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        this.handleFiles(files);
    }

    handleFileSelect(e) {
        const files = e.target.files;
        this.handleFiles(files);
    }

    handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            if (this.validateFile(file)) {
                this.currentFile = file;
                this.displayFileInfo(file);
                this.showConversionOptions();
            }
        }
    }

    validateFile(file) {
        // Check file type
        if (file.type !== 'application/pdf') {
            this.showNotification('Please select a PDF file.', 'error');
            return false;
        }

        // Check file size (50MB limit)
        const maxSize = 50 * 1024 * 1024; // 50MB in bytes
        if (file.size > maxSize) {
            this.showNotification('File size must be less than 50MB.', 'error');
            return false;
        }

        return true;
    }

    displayFileInfo(file) {
        this.fileName.textContent = file.name;
        this.fileSize.textContent = this.formatFileSize(file.size);
        
        // Hide upload section and show file info
        this.uploadSection.style.display = 'none';
        this.fileInfo.style.display = 'block';
    }

    showConversionOptions() {
        this.conversionOptions.style.display = 'block';
        this.convertSection.style.display = 'block';
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    async startConversion() {
        if (!this.currentFile) {
            this.showNotification('Please select a PDF file first.', 'error');
            return;
        }

        // Hide conversion section and show progress
        this.convertSection.style.display = 'none';
        this.conversionOptions.style.display = 'none';
        this.progressSection.style.display = 'block';

        try {
            await this.convertPDFToPPT();
        } catch (error) {
            console.error('Conversion error:', error);
            this.showNotification('Conversion failed. Please try again.', 'error');
            this.resetConverter();
        }
    }

    async convertPDFToPPT() {
        // Simulate conversion process with progress updates
        const steps = [
            { text: 'Reading PDF file...', progress: 10 },
            { text: 'Extracting pages...', progress: 25 },
            { text: 'Processing images...', progress: 50 },
            { text: 'Creating PowerPoint slides...', progress: 75 },
            { text: 'Finalizing presentation...', progress: 90 },
            { text: 'Preparing download...', progress: 100 }
        ];

        for (let i = 0; i < steps.length; i++) {
            const step = steps[i];
            this.updateProgress(step.progress, step.text);
            
            // Simulate processing time
            await this.delay(800 + Math.random() * 400);
        }

        // Create a mock PPT file for download
        await this.createPPTFile();
        
        // Show success section
        this.progressSection.style.display = 'none';
        this.successSection.style.display = 'block';
    }

    updateProgress(percentage, text) {
        this.progressFill.style.width = percentage + '%';
        this.progressPercentage.textContent = percentage + '%';
        this.progressText.textContent = text;
    }

    async createPPTFile() {
        // In a real implementation, this would use libraries like:
        // - PDF.js to extract PDF content
        // - PptxGenJS or similar to create PowerPoint files
        
        // For this demo, we'll create a simple text file as a placeholder
        const layout = this.slideLayout.value;
        const quality = this.imageQuality.value;
        
        const pptContent = this.generateMockPPTContent(layout, quality);
        
        // Create blob and prepare for download
        this.downloadBlob = new Blob([pptContent], { 
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' 
        });
        
        this.downloadFileName = this.currentFile.name.replace('.pdf', '.pptx');
    }

    generateMockPPTContent(layout, quality) {
        // This is a simplified mock - in reality, you'd use proper PPTX generation
        return `PowerPoint Presentation Generated from: ${this.currentFile.name}
        
Layout: ${layout}
Quality: ${quality}
Generated: ${new Date().toLocaleString()}

This is a demo conversion. In a real implementation, this would be a proper PPTX file
containing the slides extracted and converted from your PDF document.

Features that would be included:
- Each PDF page converted to a PowerPoint slide
- Images and text extracted and positioned appropriately
- Proper slide formatting based on selected layout
- High-quality image rendering based on quality settings
- Editable text elements where possible
- Preserved formatting and layout structure`;
    }

    downloadFile() {
        if (!this.downloadBlob) {
            this.showNotification('No file ready for download.', 'error');
            return;
        }

        // Create download link and trigger download
        const url = URL.createObjectURL(this.downloadBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = this.downloadFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        this.showNotification('Download started successfully!', 'success');
    }

    resetConverter() {
        // Reset all states
        this.currentFile = null;
        this.downloadBlob = null;
        this.downloadFileName = null;

        // Reset file input
        this.fileInput.value = '';

        // Show upload section, hide others
        this.uploadSection.style.display = 'block';
        this.fileInfo.style.display = 'none';
        this.conversionOptions.style.display = 'none';
        this.convertSection.style.display = 'none';
        this.progressSection.style.display = 'none';
        this.successSection.style.display = 'none';

        // Reset progress
        this.progressFill.style.width = '0%';
        this.progressPercentage.textContent = '0%';
        this.progressText.textContent = 'Initializing conversion...';
    }

    showNotification(message, type = 'info') {
        // Create notification element
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <i class="fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle'}"></i>
            <span>${message}</span>
        `;

        // Add notification styles
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: ${type === 'success' ? '#48bb78' : type === 'error' ? '#f56565' : '#4299e1'};
            color: white;
            padding: 15px 20px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 1000;
            display: flex;
            align-items: center;
            gap: 10px;
            font-weight: 500;
            animation: slideIn 0.3s ease-out;
        `;

        // Add animation keyframes
        if (!document.querySelector('#notification-styles')) {
            const style = document.createElement('style');
            style.id = 'notification-styles';
            style.textContent = `
                @keyframes slideIn {
                    from {
                        transform: translateX(100%);
                        opacity: 0;
                    }
                    to {
                        transform: translateX(0);
                        opacity: 1;
                    }
                }
            `;
            document.head.appendChild(style);
        }

        document.body.appendChild(notification);

        // Remove notification after 4 seconds
        setTimeout(() => {
            notification.style.animation = 'slideIn 0.3s ease-out reverse';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, 4000);
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// Initialize the converter when the page loads
document.addEventListener('DOMContentLoaded', () => {
    new PDFToPPTConverter();
});

// Add keyboard shortcuts
document.addEventListener('keydown', (e) => {
    // Ctrl+O or Cmd+O to open file dialog
    if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        document.getElementById('fileInput').click();
    }
    
    // Escape to reset converter
    if (e.key === 'Escape') {
        const converter = window.pdfToPptConverter;
        if (converter) {
            converter.resetConverter();
        }
    }
});
