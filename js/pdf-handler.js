/**
 * PDF Handler - Manages PDF loading and rendering using PDF.js
 */

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

export class PDFHandler {
    constructor() {
        this.pdfDoc = null;
        this.pdfBytes = null;
        this.pages = [];
        this.scale = 1.0;
        this.currentPage = 1;
        this.totalPages = 0;
    }

    /**
     * Load a PDF from an ArrayBuffer
     * @param {ArrayBuffer} arrayBuffer - The PDF file data
     * @returns {Promise<void>}
     */
    async loadPDF(arrayBuffer) {
        // Store original bytes for export
        this.pdfBytes = arrayBuffer.slice(0);

        // Load with PDF.js
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        this.pdfDoc = await loadingTask.promise;
        this.totalPages = this.pdfDoc.numPages;
        this.pages = [];

        return this.pdfDoc;
    }

    /**
     * Get a specific page from the PDF
     * @param {number} pageNum - Page number (1-indexed)
     * @returns {Promise<PDFPageProxy>}
     */
    async getPage(pageNum) {
        if (!this.pdfDoc) {
            throw new Error('No PDF loaded');
        }
        if (pageNum < 1 || pageNum > this.totalPages) {
            throw new Error(`Invalid page number: ${pageNum}`);
        }
        return await this.pdfDoc.getPage(pageNum);
    }

    /**
     * Render a page to a canvas element
     * @param {number} pageNum - Page number (1-indexed)
     * @param {HTMLCanvasElement} canvas - Target canvas element
     * @param {number} scale - Render scale
     * @returns {Promise<{width: number, height: number}>} - Rendered dimensions
     */
    async renderPage(pageNum, canvas, scale = this.scale) {
        const page = await this.getPage(pageNum);
        const viewport = page.getViewport({ scale });

        canvas.width = viewport.width;
        canvas.height = viewport.height;

        const context = canvas.getContext('2d');

        const renderContext = {
            canvasContext: context,
            viewport: viewport
        };

        await page.render(renderContext).promise;

        return {
            width: viewport.width,
            height: viewport.height,
            originalWidth: viewport.width / scale,
            originalHeight: viewport.height / scale
        };
    }

    /**
     * Get page dimensions at a given scale
     * @param {number} pageNum - Page number
     * @param {number} scale - Scale factor
     * @returns {Promise<{width: number, height: number}>}
     */
    async getPageDimensions(pageNum, scale = this.scale) {
        const page = await this.getPage(pageNum);
        const viewport = page.getViewport({ scale });
        return {
            width: viewport.width,
            height: viewport.height
        };
    }

    /**
     * Render all pages
     * @param {HTMLElement} container - Container element for pages
     * @param {Function} createOverlay - Function to create annotation overlay for each page
     * @param {number} scale - Render scale
     * @returns {Promise<Array>} - Array of page info objects
     */
    async renderAllPages(container, createOverlay, scale = this.scale) {
        this.scale = scale;
        container.innerHTML = '';
        this.pages = [];

        for (let pageNum = 1; pageNum <= this.totalPages; pageNum++) {
            // Create page wrapper
            const pageWrapper = document.createElement('div');
            pageWrapper.className = 'page-wrapper';
            pageWrapper.dataset.page = pageNum;

            // Create PDF canvas
            const pdfCanvas = document.createElement('canvas');
            pdfCanvas.className = 'pdf-canvas';
            pageWrapper.appendChild(pdfCanvas);

            // Render PDF page
            const dimensions = await this.renderPage(pageNum, pdfCanvas, scale);

            // Create annotation canvas container
            const annotationContainer = document.createElement('div');
            annotationContainer.className = 'annotation-container';
            annotationContainer.style.position = 'absolute';
            annotationContainer.style.top = '0';
            annotationContainer.style.left = '0';
            annotationContainer.style.width = dimensions.width + 'px';
            annotationContainer.style.height = dimensions.height + 'px';
            pageWrapper.appendChild(annotationContainer);

            // Create overlay canvas using the provided function
            const fabricCanvas = createOverlay(annotationContainer, dimensions.width, dimensions.height, pageNum);

            container.appendChild(pageWrapper);

            this.pages.push({
                pageNum,
                pdfCanvas,
                fabricCanvas,
                wrapper: pageWrapper,
                dimensions
            });
        }

        return this.pages;
    }

    /**
     * Re-render all pages at a new scale
     * @param {number} newScale - New scale factor
     * @param {Function} updateOverlay - Function to update overlay size
     */
    async rescalePages(newScale, updateOverlay) {
        this.scale = newScale;

        for (const page of this.pages) {
            const dimensions = await this.renderPage(page.pageNum, page.pdfCanvas, newScale);
            page.dimensions = dimensions;

            // Update annotation container size
            const container = page.wrapper.querySelector('.annotation-container');
            if (container) {
                container.style.width = dimensions.width + 'px';
                container.style.height = dimensions.height + 'px';
            }

            // Update overlay
            if (updateOverlay) {
                updateOverlay(page.fabricCanvas, dimensions.width, dimensions.height, newScale);
            }
        }
    }

    /**
     * Get the original PDF bytes
     * @returns {ArrayBuffer}
     */
    getOriginalBytes() {
        return this.pdfBytes;
    }

    /**
     * Check if a PDF is loaded
     * @returns {boolean}
     */
    isLoaded() {
        return this.pdfDoc !== null;
    }
}
