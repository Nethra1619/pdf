// Global Worker setup for PDF.js
if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

class PDFToPPTConverter {
  constructor() {
    this.currentFile = null;
    this.pdfPages = [];
    this.downloadBlob = null;
    this.downloadFileName = null;

    this.initializeElements();
    this.attachEventListeners();
    this.setupDragAndDrop();
  }

  initializeElements() {
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
    this.browseBtn.addEventListener('click', () => this.fileInput.click());
    this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
    this.removeFile.addEventListener('click', () => this.resetConverter());
    this.convertBtn.addEventListener('click', () => this.startConversion());
    this.downloadBtn.addEventListener('click', () => this.downloadFile());
    this.convertAnotherBtn.addEventListener('click', () => this.resetConverter());
  }

  setupDragAndDrop() {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
      this.uploadArea.addEventListener(eventName, this.preventDefaults, false);
      document.body.addEventListener(eventName, this.preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      this.uploadArea.addEventListener(eventName, () => this.uploadArea.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
      this.uploadArea.addEventListener(eventName, () => this.uploadArea.classList.remove('dragover'), false);
    });

    this.uploadArea.addEventListener('drop', (e) => this.handleDrop(e), false);
  }

  preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  handleDrop(e) {
    const files = e.dataTransfer.files;
    this.handleFiles(files);
  }

  handleFileSelect(e) {
    const files = e.target.files;
    this.handleFiles(files);
  }

  handleFiles(files) {
    const file = files[0];
    if (file && this.validateFile(file)) {
      this.currentFile = file;
      this.displayFileInfo(file);
      this.showConversionOptions();
    }
  }

  validateFile(file) {
    if (file.type !== 'application/pdf') {
      this.showNotification('Only PDF files are supported.', 'error');
      return false;
    }
    if (file.size > 50 * 1024 * 1024) {
      this.showNotification('File must be under 50MB.', 'error');
      return false;
    }
    return true;
  }

  displayFileInfo(file) {
    this.fileName.textContent = file.name;
    this.fileSize.textContent = (file.size / (1024 * 1024)).toFixed(2) + ' MB';
    this.uploadSection.style.display = 'none';
    this.fileInfo.style.display = 'block';
  }

  showConversionOptions() {
    this.conversionOptions.style.display = 'block';
    this.convertSection.style.display = 'block';
  }

  async startConversion() {
    if (!this.currentFile) return;

    this.conversionOptions.style.display = 'none';
    this.convertSection.style.display = 'none';
    this.progressSection.style.display = 'block';

    try {
      await this.convertPDFToPPT();
      this.progressSection.style.display = 'none';
      this.successSection.style.display = 'block';
    } catch (err) {
      this.showNotification('Conversion failed. Try again.', 'error');
      this.resetConverter();
    }
  }

  async convertPDFToPPT() {
    this.updateProgress(10, 'Reading file...');
    const arrayBuffer = await this.currentFile.arrayBuffer();

    this.updateProgress(30, 'Loading PDF...');
    const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      this.updateProgress(30 + (i / pdf.numPages) * 30, `Rendering page ${i}`);
      const page = await pdf.getPage(i);
      const scale = 2.0;
      const viewport = page.getViewport({ scale });

      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      canvas.width = viewport.width;
      canvas.height = viewport.height;

      await page.render({ canvasContext: context, viewport }).promise;
      const image = canvas.toDataURL('image/png');

      pages.push({
        image,
        width: viewport.width,
        height: viewport.height,
        number: i
      });
    }

    this.updateProgress(80, 'Creating PPT...');
    await this.createPPT(pages);

    this.updateProgress(100, 'Done!');
    await new Promise(resolve => setTimeout(resolve, 500));
  }

  async createPPT(pages) {
    const pptx = new PptxGenJS();
    pptx.layout = this.slideLayout.value === 'widescreen' ? 'LAYOUT_16x9' : 'LAYOUT_4x3';

    // Title slide
    const title = pptx.addSlide();
    title.addText('PDF to PPT Conversion', { x: 1, y: 1, fontSize: 30, w: 8, align: 'center' });

    // Slides for each page
    for (const page of pages) {
      const slide = pptx.addSlide();
      const layoutWidth = pptx.layout === 'LAYOUT_16x9' ? 10 : 10;
      const layoutHeight = pptx.layout === 'LAYOUT_16x9' ? 5.625 : 7.5;

      const imgRatio = page.width / page.height;
      const slideRatio = layoutWidth / layoutHeight;
      let imgW, imgH, imgX, imgY;

      if (imgRatio > slideRatio) {
        imgW = layoutWidth * 0.9;
        imgH = imgW / imgRatio;
        imgX = layoutWidth * 0.05;
        imgY = (layoutHeight - imgH) / 2;
      } else {
        imgH = layoutHeight * 0.9;
        imgW = imgH * imgRatio;
        imgX = (layoutWidth - imgW) / 2;
        imgY = layoutHeight * 0.05;
      }

      slide.addImage({ data: page.image, x: imgX, y: imgY, w: imgW, h: imgH });
      slide.addText(`Page ${page.number}`, { x: layoutWidth - 1.5, y: layoutHeight - 0.4, fontSize: 10, color: '666666' });
    }

    this.downloadBlob = await pptx.write('blob');
    this.downloadFileName = this.currentFile.name.replace('.pdf', '.pptx');
  }

  updateProgress(percent, text) {
    this.progressFill.style.width = percent + '%';
    this.progressPercentage.textContent = Math.round(percent) + '%';
    this.progressText.textContent = text;
  }

  downloadFile() {
    if (!this.downloadBlob) return;

    const url = URL.createObjectURL(this.downloadBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = this.downloadFileName;
    a.click();
    URL.revokeObjectURL(url);
  }

  resetConverter() {
    this.currentFile = null;
    this.pdfPages = [];
    this.downloadBlob = null;
    this.fileInput.value = '';

    this.uploadSection.style.display = 'block';
    this.fileInfo.style.display = 'none';
    this.conversionOptions.style.display = 'none';
    this.convertSection.style.display = 'none';
    this.progressSection.style.display = 'none';
    this.successSection.style.display = 'none';

    this.updateProgress(0, 'Initializing...');
  }

  showNotification(msg, type = 'info') {
    alert(msg); // You can replace this with a toast notification for better UX
  }
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  window.pdfToPptConverter = new PDFToPPTConverter();
});
