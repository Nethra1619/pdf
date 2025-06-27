import PptxGenJS from 'pptxgenjs';

class PDFToPPTConverter {
    constructor() {
        this.currentFile = null;
        this.pdfPages = [];
        this.pptx = new PptxGenJS();
        this.initializeElements();
        this.attachEventListeners();
        this.setupDragAndDrop();

        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
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
        this.convertBtn = document.getElementById('convertBtn');
    }

    attachEventListeners() {
        this.browseBtn.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFile(e.target.files[0]));
        this.removeFile.addEventListener('click', () => this.clearFile());
        this.convertBtn.addEventListener('click', () => this.convertToPPT());
    }

    setupDragAndDrop() {
        this.uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.uploadArea.classList.add('dragover');
        });

        this.uploadArea.addEventListener('dragleave', () => {
            this.uploadArea.classList.remove('dragover');
        });

        this.uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            this.uploadArea.classList.remove('dragover');
            this.handleFile(e.dataTransfer.files[0]);
        });
    }

    handleFile(file) {
        if (file && file.type === "application/pdf") {
            this.currentFile = file;
            this.fileName.textContent = file.name;
            this.fileSize.textContent = `${(file.size / 1024).toFixed(2)} KB`;
            this.fileInfo.classList.remove('hidden');
        } else {
            alert("Please upload a valid PDF file.");
        }
    }

    clearFile() {
        this.currentFile = null;
        this.fileInput.value = '';
        this.fileInfo.classList.add('hidden');
    }

    async convertToPPT() {
        if (!this.currentFile) {
            alert("Please upload a PDF first.");
            return;
        }

        const reader = new FileReader();
        reader.onload = async () => {
            const typedArray = new Uint8Array(reader.result);
            const pdf = await pdfjsLib.getDocument({ data: typedArray }).promise;

            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale: 2 });
                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                canvas.width = viewport.width;
                canvas.height = viewport.height;

                await page.render({ canvasContext: context, viewport }).promise;
                const imgData = canvas.toDataURL('image/jpeg');

                const slide = this.pptx.addSlide();
                slide.addImage({ data: imgData, x: 0, y: 0, w: '100%', h: '100%' });
            }

            this.pptx.writeFile({ fileName: "converted.pptx" });
        };
        reader.readAsArrayBuffer(this.currentFile);
    }
}

window.addEventListener('DOMContentLoaded', () => {
    new PDFToPPTConverter();
});
