import PptxGenJS from 'pptxgenjs';

class PDFToPPTConverter {
    constructor() {
        this.currentFile = null;
        this.pdfPages = [];
        this.initializeElements();
        this.attachEventListeners();
        this.setupDragAndDrop();
        
        // Configure PDF.js
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
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
        try {
            // Step 1: Read PDF file
            this.updateProgress(10, 'Reading PDF file...');
            const arrayBuffer = await this.currentFile.arrayBuffer();
            
            // Step 2: Load PDF document
            this.updateProgress(25, 'Loading PDF document...');
            const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
            
            // Step 3: Extract pages
            this.updateProgress(40, 'Extracting PDF pages...');
            this.pdfPages = [];
            
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const viewport = page.getViewport({ scale: 2.0 });
                
                // Create canvas to render page
                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;
                
                // Render page to canvas
                await page.render({
                    canvasContext: context,
                    viewport: viewport
                }).promise;
                
                // Convert canvas to image data
                const imageData = canvas.toDataURL('image/png');
                this.pdfPages.push({
                    pageNumber: pageNum,
                    imageData: imageData,
                    width: viewport.width,
                    height: viewport.height
                });
                
                // Update progress
                const progress = 40 + (pageNum / pdf.numPages) * 30;
                this.updateProgress(progress, `Processing page ${pageNum} of ${pdf.numPages}...`);
            }
            
            // Step 4: Create PowerPoint presentation
            this.updateProgress(75, 'Creating PowerPoint slides...');
            await this.createPPTFile();
            
            // Step 5: Finalize
            this.updateProgress(100, 'Finalizing presentation...');
            await this.delay(500);
            
            // Show success section
            this.progressSection.style.display = 'none';
            this.successSection.style.display = 'block';
            
        } catch (error) {
            console.error('PDF processing error:', error);
            throw new Error('Failed to process PDF file');
        }
    }

    async createPPTFile() {
        try {
            // Create new PowerPoint presentation
            const pptx = new PptxGenJS();
            
            // Set slide layout based on user selection
            const layout = this.slideLayout.value === 'widescreen' ? 'LAYOUT_16x9' : 'LAYOUT_4x3';
            pptx.layout = layout;
            
            // Add title slide
            const titleSlide = pptx.addSlide();
            titleSlide.addText('Converted from PDF', {
                x: 1,
                y: 1,
                w: 8,
                h: 1,
                fontSize: 32,
                bold: true,
                align: 'center'
            });
            
            titleSlide.addText(`Original file: ${this.currentFile.name}`, {
                x: 1,
                y: 2.5,
                w: 8,
                h: 0.5,
                fontSize: 16,
                align: 'center',
                color: '666666'
            });
            
            titleSlide.addText(`Converted on: ${new Date().toLocaleDateString()}`, {
                x: 1,
                y: 3,
                w: 8,
                h: 0.5,
                fontSize: 14,
                align: 'center',
                color: '888888'
            });
            
            // Add slides for each PDF page
            for (let i = 0; i < this.pdfPages.length; i++) {
                const pageData = this.pdfPages[i];
                const slide = pptx.addSlide();
                
                // Calculate image dimensions to fit slide
                const slideWidth = layout === 'LAYOUT_16x9' ? 10 : 10;
                const slideHeight = layout === 'LAYOUT_16x9' ? 5.625 : 7.5;
                
                const imageAspectRatio = pageData.width / pageData.height;
                const slideAspectRatio = slideWidth / slideHeight;
                
                let imgWidth, imgHeight, imgX, imgY;
                
                if (imageAspectRatio > slideAspectRatio) {
                    // Image is wider than slide
                    imgWidth = slideWidth * 0.9;
                    imgHeight = imgWidth / imageAspectRatio;
                    imgX = slideWidth * 0.05;
                    imgY = (slideHeight - imgHeight) / 2;
                } else {
                    // Image is taller than slide
                    imgHeight = slideHeight * 0.9;
                    imgWidth = imgHeight * imageAspectRatio;
                    imgX = (slideWidth - imgWidth) / 2;
                    imgY = slideHeight * 0.05;
                }
                
                // Add image to slide
                slide.addImage({
                    data: pageData.imageData,
                    x: imgX,
                    y: imgY,
                    w: imgWidth,
                    h: imgHeight
                });
                
                // Add page number
                slide.addText(`Page ${pageData.pageNumber}`, {
                    x: slideWidth - 1.5,
                    y: slideHeight - 0.5,
                    w: 1,
                    h: 0.3,
                    fontSize: 10,
                    align: 'center',
                    color: '666666'
                });
            }
            
            // Generate the PowerPoint file
            this.downloadBlob = await pptx.write('blob');
            this.downloadFileName = this.currentFile.name.replace('.pdf', '.pptx');
            
        } catch (error) {
            console.error('PowerPoint creation error:', error);
            throw new Error('Failed to create PowerPoint file');
        }
    }

    updateProgress(percentage, text) {
        this.progressFill.style.width = percentage + '%';
        this.progressPercentage.textContent = Math.round(percentage) + '%';
        this.progressText.textContent = text;
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
        this.pdfPages = [];
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