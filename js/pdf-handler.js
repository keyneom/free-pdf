/**
 * PDF Handler - Manages PDF loading and rendering using PDF.js
 */

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

export class PDFHandler {
    constructor() {
        // Multiple PDFs can be loaded (for append/merge)
        // docId -> { pdfDoc: PDFDocumentProxy, bytes: ArrayBuffer, name: string }
        this.docs = new Map();
        this.mainDocId = null;

        // Rendered view pages (in current order)
        this.pages = []; // [{ viewIndex, viewPageId, docId, sourcePageNum, rotation, ... }]
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
        // Reset and load as main document
        this.docs.clear();
        this.pages = [];

        const docId = 'main';
        const bytes = arrayBuffer.slice(0);
        const loadingTask = pdfjsLib.getDocument({ data: bytes });
        const pdfDoc = await loadingTask.promise;

        this.docs.set(docId, { pdfDoc, bytes, name: 'document.pdf' });
        this.mainDocId = docId;

        // totalPages is for the *main* document until view pages are set
        this.totalPages = pdfDoc.numPages;
        this.currentPage = 1;

        return pdfDoc;
    }

    /**
     * Add an additional PDF (for append/merge)
     * @param {ArrayBuffer} arrayBuffer
     * @param {string} name
     * @returns {Promise<{ docId: string; pageCount: number }>}
     */
    async addDocument(arrayBuffer, name = 'document.pdf') {
        const bytes = arrayBuffer.slice(0);
        const loadingTask = pdfjsLib.getDocument({ data: bytes });
        const pdfDoc = await loadingTask.promise;

        const docId = `doc_${Math.random().toString(36).slice(2, 10)}`;
        this.docs.set(docId, { pdfDoc, bytes, name });
        return { docId, pageCount: pdfDoc.numPages };
    }

    /**
     * Get a specific page from a given loaded document
     * @param {string} docId
     * @param {number} pageNum - Page number (1-indexed)
     * @returns {Promise<PDFPageProxy>}
     */
    async getPage(docId, pageNum) {
        const doc = this.docs.get(docId);
        if (!doc?.pdfDoc) {
            throw new Error('No PDF loaded');
        }
        if (pageNum < 1 || pageNum > doc.pdfDoc.numPages) {
            throw new Error(`Invalid page number: ${pageNum}`);
        }
        return await doc.pdfDoc.getPage(pageNum);
    }

    /**
     * Render a page to a canvas element
     * @param {string} docId
     * @param {number} pageNum - Page number (1-indexed)
     * @param {HTMLCanvasElement} canvas - Target canvas element
     * @param {number} scale - Render scale
     * @param {number} rotation - Rotation degrees (0/90/180/270)
     * @returns {Promise<{width: number, height: number}>} - Rendered dimensions
     */
    async renderPage(docId, pageNum, canvas, scale = this.scale, rotation = 0) {
        const page = await this.getPage(docId, pageNum);
        const viewport = page.getViewport({ scale, rotation });

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
     * @param {string} docId
     * @param {number} pageNum - Page number
     * @param {number} scale - Scale factor
     * @param {number} rotation - Rotation degrees (0/90/180/270)
     * @returns {Promise<{width: number, height: number}>}
     */
    async getPageDimensions(docId, pageNum, scale = this.scale, rotation = 0) {
        const page = await this.getPage(docId, pageNum);
        const viewport = page.getViewport({ scale, rotation });
        return {
            width: viewport.width,
            height: viewport.height
        };
    }

    /**
     * Render pages based on a view model (supports reorder/append/delete)
     * @param {HTMLElement} container - Container element for pages
     * @param {Array<{id: string; docId: string; sourcePageNum: number; rotation?: number}>} viewPages
     * @param {Function} createOverlay - Function to create annotation overlay for each page
     * @param {number} scale - Render scale
     * @returns {Promise<Array>} - Array of page info objects
     */
    async renderViewPages(container, viewPages, createOverlay, scale = this.scale) {
        this.scale = scale;
        container.innerHTML = '';
        this.pages = [];
        this.totalPages = viewPages.length;

        for (let i = 0; i < viewPages.length; i++) {
            const vp = viewPages[i];
            const viewIndex = i + 1;
            const rotation = vp.rotation || 0;

            // Create page wrapper (outer) and inner (for CSS rotation)
            const pageWrapper = document.createElement('div');
            pageWrapper.className = 'page-wrapper';
            pageWrapper.dataset.page = viewIndex;
            pageWrapper.dataset.pageId = vp.id;

            const inner = document.createElement('div');
            inner.className = 'page-wrapper-inner';

            // Create PDF canvas – always render at 0°; rotation applied via CSS
            const pdfCanvas = document.createElement('canvas');
            pdfCanvas.className = 'pdf-canvas';
            inner.appendChild(pdfCanvas);

            const dimensions = await this.renderPage(vp.docId, vp.sourcePageNum, pdfCanvas, scale, 0);

            inner.style.width = dimensions.width + 'px';
            inner.style.height = dimensions.height + 'px';

            const annotationContainer = document.createElement('div');
            annotationContainer.className = 'annotation-container';
            annotationContainer.style.position = 'absolute';
            annotationContainer.style.top = '0';
            annotationContainer.style.left = '0';
            annotationContainer.style.width = dimensions.width + 'px';
            annotationContainer.style.height = dimensions.height + 'px';
            inner.appendChild(annotationContainer);

            const fabricCanvas = createOverlay(annotationContainer, dimensions.width, dimensions.height, vp.id);
            pageWrapper.appendChild(inner);
            // Set wrapper size so main view shows (inner is position:absolute so doesn't take flow)
            const r = rotation % 360;
            const wrapperW = (r === 90 || r === 270) ? dimensions.height : dimensions.width;
            const wrapperH = (r === 90 || r === 270) ? dimensions.width : dimensions.height;
            pageWrapper.style.width = wrapperW + 'px';
            pageWrapper.style.height = wrapperH + 'px';
            if (inner) {
                inner.style.transform = r ? `translate(-50%, -50%) rotate(${r}deg)` : 'translate(-50%, -50%)';
            }
            container.appendChild(pageWrapper);

            this.pages.push({
                viewIndex,
                viewPageId: vp.id,
                docId: vp.docId,
                sourcePageNum: vp.sourcePageNum,
                rotation,
                pdfCanvas,
                fabricCanvas,
                wrapper: pageWrapper,
                inner,
                dimensions
            });
        }

        return this.pages;
    }

    /**
     * Render a single view page (for append/insert) without clearing existing pages.
     * @param {HTMLElement} container
     * @param {{id: string; docId: string; sourcePageNum: number; rotation?: number}} vp
     * @param {Function} createOverlay
     * @param {number} scale
     * @param {number} viewIndex - 1-indexed view position
     */
    async renderOneViewPage(container, vp, createOverlay, scale = this.scale, viewIndex = null) {
        const rotation = vp.rotation || 0;
        const index = viewIndex ?? (this.pages.length + 1);

        const pageWrapper = document.createElement('div');
        pageWrapper.className = 'page-wrapper';
        pageWrapper.dataset.page = index;
        pageWrapper.dataset.pageId = vp.id;

        const inner = document.createElement('div');
        inner.className = 'page-wrapper-inner';

        const pdfCanvas = document.createElement('canvas');
        pdfCanvas.className = 'pdf-canvas';
        inner.appendChild(pdfCanvas);

        const dimensions = await this.renderPage(vp.docId, vp.sourcePageNum, pdfCanvas, scale, 0);
        inner.style.width = dimensions.width + 'px';
        inner.style.height = dimensions.height + 'px';

        const annotationContainer = document.createElement('div');
        annotationContainer.className = 'annotation-container';
        annotationContainer.style.position = 'absolute';
        annotationContainer.style.top = '0';
        annotationContainer.style.left = '0';
        annotationContainer.style.width = dimensions.width + 'px';
        annotationContainer.style.height = dimensions.height + 'px';
        inner.appendChild(annotationContainer);

        const fabricCanvas = createOverlay(annotationContainer, dimensions.width, dimensions.height, vp.id);
        pageWrapper.appendChild(inner);
        pageWrapper.style.width = dimensions.width + 'px';
        pageWrapper.style.height = dimensions.height + 'px';
        container.appendChild(pageWrapper);

        const info = {
            viewIndex: index,
            viewPageId: vp.id,
            docId: vp.docId,
            sourcePageNum: vp.sourcePageNum,
            rotation,
            inner,
            pdfCanvas,
            fabricCanvas,
            wrapper: pageWrapper,
            dimensions
        };
        this.pages.push(info);
        return info;
    }

    /**
     * Re-render all pages at a new scale
     * @param {number} newScale - New scale factor
     * @param {Function} updateOverlay - Function to update overlay size
     */
    async rescalePages(newScale, updateOverlay) {
        this.scale = newScale;

        for (const page of this.pages) {
            const dimensions = await this.renderPage(page.docId, page.sourcePageNum, page.pdfCanvas, newScale, 0);
            page.dimensions = dimensions;

            const w = dimensions.width;
            const h = dimensions.height;
            if (page.inner) {
                page.inner.style.width = w + 'px';
                page.inner.style.height = h + 'px';
            }
            const container = page.wrapper?.querySelector('.annotation-container');
            if (container) {
                container.style.width = w + 'px';
                container.style.height = h + 'px';
            }

            if (updateOverlay) {
                updateOverlay(page.fabricCanvas, w, h, newScale);
            }
        }
    }

    /**
     * Get original bytes for a given document
     * @param {string} docId
     * @returns {ArrayBuffer | null}
     */
    getOriginalBytes(docId = this.mainDocId) {
        const doc = docId ? this.docs.get(docId) : null;
        return doc?.bytes || null;
    }

    /**
     * Get a map of all loaded document bytes
     * @returns {Map<string, ArrayBuffer>}
     */
    getAllOriginalBytes() {
        const out = new Map();
        for (const [docId, doc] of this.docs.entries()) {
            out.set(docId, doc.bytes);
        }
        return out;
    }

    /**
     * Check if a PDF is loaded
     * @returns {boolean}
     */
    isLoaded() {
        return this.docs.size > 0;
    }
}
