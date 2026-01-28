/**
 * PDF Editor - Main Application
 */

import { PDFHandler } from './pdf-handler.js';
import { CanvasManager } from './canvas-manager.js';
import { PDFExporter } from './export.js';
import { SignaturePad } from './signature-pad.js';

class PDFEditorApp {
    constructor() {
        this.pdfHandler = new PDFHandler();
        this.canvasManager = new CanvasManager();
        this.exporter = new PDFExporter();
        this.signaturePad = new SignaturePad();

        this.currentScale = 1.0;
        this.fileName = 'document.pdf';

        this.init();
    }

    /**
     * Initialize the application
     */
    init() {
        this.cacheElements();
        this.bindEvents();
        this.setupDragAndDrop();
        this.setupKeyboardShortcuts();
        this.initSignaturePad();
    }

    /**
     * Cache DOM elements
     */
    cacheElements() {
        // Screens
        this.welcomeScreen = document.getElementById('welcome-screen');
        this.pdfContainer = document.getElementById('pdf-container');
        this.loadingOverlay = document.getElementById('loading-overlay');
        this.loadingText = document.getElementById('loading-text');

        // File input
        this.fileInput = document.getElementById('file-input');

        // Buttons
        this.btnOpen = document.getElementById('btn-open');
        this.btnOpenWelcome = document.getElementById('btn-open-welcome');
        this.btnSave = document.getElementById('btn-save');
        this.btnUndo = document.getElementById('btn-undo');
        this.btnRedo = document.getElementById('btn-redo');
        this.btnDelete = document.getElementById('btn-delete');

        // Tool buttons
        this.toolButtons = document.querySelectorAll('[data-tool]');

        // Page navigation
        this.btnPrevPage = document.getElementById('btn-prev-page');
        this.btnNextPage = document.getElementById('btn-next-page');
        this.pageInput = document.getElementById('page-input');
        this.totalPagesSpan = document.getElementById('total-pages');

        // Zoom controls
        this.btnZoomIn = document.getElementById('btn-zoom-in');
        this.btnZoomOut = document.getElementById('btn-zoom-out');
        this.btnFitWidth = document.getElementById('btn-fit-width');
        this.zoomLevel = document.getElementById('zoom-level');

        // PDF area
        this.pdfScrollArea = document.getElementById('pdf-scroll-area');
        this.pdfPages = document.getElementById('pdf-pages');

        // Tool options
        this.toolOptions = document.getElementById('tool-options');

        // Signature modal
        this.signatureModal = document.getElementById('signature-modal');
    }

    /**
     * Bind event listeners
     */
    bindEvents() {
        // File open
        this.btnOpen.addEventListener('click', () => this.fileInput.click());
        this.btnOpenWelcome.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Save/Export
        this.btnSave.addEventListener('click', () => this.exportPDF());

        // Undo/Redo
        this.btnUndo.addEventListener('click', () => {
            this.canvasManager.undo();
            this.updateHistoryButtons();
        });
        this.btnRedo.addEventListener('click', () => {
            this.canvasManager.redo();
            this.updateHistoryButtons();
        });

        // Delete
        this.btnDelete.addEventListener('click', () => {
            this.canvasManager.deleteSelected();
        });

        // Tool selection
        this.toolButtons.forEach(btn => {
            btn.addEventListener('click', () => this.selectTool(btn));
        });

        // Page navigation
        this.btnPrevPage.addEventListener('click', () => this.goToPage(this.pdfHandler.currentPage - 1));
        this.btnNextPage.addEventListener('click', () => this.goToPage(this.pdfHandler.currentPage + 1));
        this.pageInput.addEventListener('change', (e) => this.goToPage(parseInt(e.target.value)));

        // Zoom controls
        this.btnZoomIn.addEventListener('click', () => this.zoom(this.currentScale + 0.25));
        this.btnZoomOut.addEventListener('click', () => this.zoom(this.currentScale - 0.25));
        this.btnFitWidth.addEventListener('click', () => this.fitWidth());

        // Signature modal
        this.setupSignatureModal();
    }

    /**
     * Set up drag and drop
     */
    setupDragAndDrop() {
        const dropZone = this.welcomeScreen;

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(event => {
            dropZone.addEventListener(event, (e) => {
                e.preventDefault();
                e.stopPropagation();
            });
        });

        dropZone.addEventListener('dragenter', () => {
            dropZone.classList.add('drag-over');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('drag-over');
        });

        dropZone.addEventListener('drop', (e) => {
            dropZone.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0 && files[0].type === 'application/pdf') {
                this.loadFile(files[0]);
            }
        });
    }

    /**
     * Set up keyboard shortcuts
     */
    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            // Ctrl/Cmd + Z = Undo
            if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
                e.preventDefault();
                this.canvasManager.undo();
                this.updateHistoryButtons();
            }

            // Ctrl/Cmd + Shift + Z or Ctrl/Cmd + Y = Redo
            if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
                e.preventDefault();
                this.canvasManager.redo();
                this.updateHistoryButtons();
            }

            // Delete = Delete selected
            if (e.key === 'Delete' || e.key === 'Backspace') {
                // Don't delete if we're in a text editing mode
                const activeEl = document.activeElement;
                if (activeEl.tagName !== 'INPUT' && activeEl.tagName !== 'TEXTAREA') {
                    const activeCanvas = this.canvasManager.activeCanvas;
                    if (activeCanvas && !activeCanvas.getActiveObject()?.isEditing) {
                        e.preventDefault();
                        this.canvasManager.deleteSelected();
                    }
                }
            }

            // Tool shortcuts
            if (!e.ctrlKey && !e.metaKey) {
                // Don't switch tools while typing in inputs/textareas/contentEditable
                const activeEl = document.activeElement;
                const isTypingInInput =
                    activeEl &&
                    (activeEl.tagName === 'INPUT' ||
                        activeEl.tagName === 'TEXTAREA' ||
                        activeEl.isContentEditable);

                // Don't switch tools while editing a canvas text object
                const activeCanvas = this.canvasManager.activeCanvas;
                const isEditingCanvasText =
                    !!(activeCanvas && activeCanvas.getActiveObject()?.isEditing);

                if (!isTypingInInput && !isEditingCanvasText) {
                    switch (e.key.toLowerCase()) {
                        case 'v':
                            this.selectToolByName('select');
                            break;
                        case 't':
                            this.selectToolByName('text');
                            break;
                        case 'w':
                            this.selectToolByName('whiteout');
                            break;
                        case 'd':
                            this.selectToolByName('draw');
                            break;
                        case 's':
                            this.selectToolByName('signature');
                            break;
                    }
                }
            }

            // Escape = Cancel current action
            if (e.key === 'Escape') {
                this.canvasManager.setTool('select');
                this.selectToolByName('select');
            }
        });
    }

    /**
     * Handle file selection
     */
    handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            this.loadFile(file);
        }
    }

    /**
     * Load a PDF file
     */
    async loadFile(file) {
        if (file.type !== 'application/pdf') {
            alert('Please select a PDF file.');
            return;
        }

        this.showLoading('Loading PDF...');
        this.fileName = file.name;

        try {
            const arrayBuffer = await file.arrayBuffer();
            await this.pdfHandler.loadPDF(arrayBuffer);

            // Clear any existing canvases
            this.canvasManager.clearAll();

            // Render all pages
            this.showLoading('Rendering pages...');
            await this.pdfHandler.renderAllPages(
                this.pdfPages,
                (container, width, height, pageNum) => {
                    return this.canvasManager.createCanvas(container, width, height, pageNum);
                },
                this.currentScale
            );

            // Update UI
            this.welcomeScreen.classList.add('hidden');
            this.pdfContainer.classList.remove('hidden');
            this.btnSave.disabled = false;
            this.updatePageNavigation();
            this.updateZoomDisplay();

            this.hideLoading();
        } catch (error) {
            console.error('Error loading PDF:', error);
            alert('Error loading PDF: ' + error.message);
            this.hideLoading();
        }
    }

    /**
     * Select a tool by button element
     */
    selectTool(btn) {
        const tool = btn.dataset.tool;

        // Handle signature tool specially
        if (tool === 'signature') {
            this.showSignatureModal();
            return;
        }

        // Update button states
        this.toolButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        // Set the tool
        this.canvasManager.setTool(tool);

        // Show tool options
        this.showToolOptions(tool);
    }

    /**
     * Select a tool by name
     */
    selectToolByName(toolName) {
        const btn = document.querySelector(`[data-tool="${toolName}"]`);
        if (btn) {
            this.selectTool(btn);
        }
    }

    /**
     * Show tool-specific options
     */
    showToolOptions(tool) {
        this.toolOptions.innerHTML = '';

        switch (tool) {
            case 'text':
                this.toolOptions.innerHTML = `
                    <div class="tool-option">
                        <label>Color:</label>
                        <input type="color" id="text-color" value="${this.canvasManager.settings.textColor}">
                    </div>
                    <div class="tool-option">
                        <label>Size:</label>
                        <input type="number" id="text-size" value="${this.canvasManager.settings.fontSize}" min="8" max="72">
                    </div>
                    <div class="tool-option">
                        <label>Font:</label>
                        <select id="text-font">
                            <option value="Arial">Arial</option>
                            <option value="Times New Roman">Times New Roman</option>
                            <option value="Courier New">Courier New</option>
                            <option value="Georgia">Georgia</option>
                            <option value="Verdana">Verdana</option>
                        </select>
                    </div>
                `;
                this.bindTextOptions();
                break;

            case 'draw':
                this.toolOptions.innerHTML = `
                    <div class="tool-option">
                        <label>Color:</label>
                        <input type="color" id="stroke-color" value="${this.canvasManager.settings.strokeColor}">
                    </div>
                    <div class="tool-option">
                        <label>Width:</label>
                        <input type="range" id="stroke-width" min="1" max="20" value="${this.canvasManager.settings.strokeWidth}">
                        <span id="stroke-width-value">${this.canvasManager.settings.strokeWidth}px</span>
                    </div>
                `;
                this.bindDrawOptions();
                break;

            case 'whiteout':
                this.toolOptions.innerHTML = `
                    <div class="tool-option">
                        <span>Click and drag to cover content</span>
                    </div>
                `;
                break;

            default:
                break;
        }
    }

    /**
     * Bind text tool options
     */
    bindTextOptions() {
        const colorInput = document.getElementById('text-color');
        const sizeInput = document.getElementById('text-size');
        const fontInput = document.getElementById('text-font');

        colorInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ textColor: e.target.value });
        });

        sizeInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ fontSize: parseInt(e.target.value) });
        });

        fontInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ fontFamily: e.target.value });
        });
    }

    /**
     * Bind draw tool options
     */
    bindDrawOptions() {
        const colorInput = document.getElementById('stroke-color');
        const widthInput = document.getElementById('stroke-width');
        const widthValue = document.getElementById('stroke-width-value');

        colorInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ strokeColor: e.target.value });
        });

        widthInput?.addEventListener('input', (e) => {
            const width = parseInt(e.target.value);
            this.canvasManager.updateSettings({ strokeWidth: width });
            if (widthValue) widthValue.textContent = `${width}px`;
        });
    }

    /**
     * Go to a specific page
     */
    goToPage(pageNum) {
        if (pageNum < 1 || pageNum > this.pdfHandler.totalPages) return;

        this.pdfHandler.currentPage = pageNum;
        this.pageInput.value = pageNum;

        // Scroll to page
        const pageWrapper = document.querySelector(`.page-wrapper[data-page="${pageNum}"]`);
        if (pageWrapper) {
            pageWrapper.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }

        this.updatePageNavigation();
    }

    /**
     * Update page navigation UI
     */
    updatePageNavigation() {
        const current = this.pdfHandler.currentPage;
        const total = this.pdfHandler.totalPages;

        this.pageInput.value = current;
        this.pageInput.max = total;
        this.totalPagesSpan.textContent = total;

        this.btnPrevPage.disabled = current <= 1;
        this.btnNextPage.disabled = current >= total;
    }

    /**
     * Zoom to a specific level
     */
    async zoom(scale) {
        scale = Math.max(0.25, Math.min(4, scale));
        this.currentScale = scale;

        this.showLoading('Adjusting zoom...');

        await this.pdfHandler.rescalePages(scale, (canvas, width, height, newScale) => {
            this.canvasManager.updateCanvasSize(canvas, width, height, newScale);
        });

        this.updateZoomDisplay();
        this.hideLoading();
    }

    /**
     * Fit PDF to width
     */
    async fitWidth() {
        const containerWidth = this.pdfScrollArea.clientWidth - 48; // Account for padding
        const page = await this.pdfHandler.getPage(1);
        const viewport = page.getViewport({ scale: 1 });
        const scale = containerWidth / viewport.width;

        await this.zoom(scale);
    }

    /**
     * Update zoom display
     */
    updateZoomDisplay() {
        this.zoomLevel.textContent = Math.round(this.currentScale * 100) + '%';
    }

    /**
     * Initialize signature pad
     */
    initSignaturePad() {
        const canvas = document.getElementById('signature-canvas');
        this.signaturePad.init(canvas);
    }

    /**
     * Set up signature modal
     */
    setupSignatureModal() {
        const modal = this.signatureModal;
        const closeBtn = document.getElementById('signature-modal-close');
        const cancelBtn = document.getElementById('sig-cancel');
        const applyBtn = document.getElementById('sig-apply');
        const clearBtn = document.getElementById('sig-clear');
        const tabs = modal.querySelectorAll('.sig-tab');
        const textInput = document.getElementById('sig-text-input');
        const fontOptions = modal.querySelectorAll('input[name="sig-font"]');
        const preview = document.getElementById('sig-preview');

        // Close modal
        closeBtn.addEventListener('click', () => this.hideSignatureModal());
        cancelBtn.addEventListener('click', () => this.hideSignatureModal());

        // Clear signature
        clearBtn.addEventListener('click', () => this.signaturePad.clear());

        // Tab switching
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');

                const tabName = tab.dataset.tab;
                document.getElementById('sig-draw-area').classList.toggle('active', tabName === 'draw');
                document.getElementById('sig-type-area').classList.toggle('active', tabName === 'type');

                this.signaturePad.setMode(tabName);
            });
        });

        // Text input
        textInput.addEventListener('input', (e) => {
            this.signaturePad.setTypedText(e.target.value);
            this.signaturePad.updatePreview(preview);
        });

        // Font options
        fontOptions.forEach(option => {
            option.addEventListener('change', (e) => {
                this.signaturePad.setFontStyle(e.target.value);
                this.signaturePad.updatePreview(preview);
            });
        });

        // Apply signature
        applyBtn.addEventListener('click', () => {
            const dataUrl = this.signaturePad.getDataUrl();
            if (dataUrl) {
                this.canvasManager.setSignature(dataUrl);
                this.canvasManager.setTool('signature');
                this.toolButtons.forEach(b => b.classList.remove('active'));
                document.querySelector('[data-tool="signature"]')?.classList.add('active');
                this.hideSignatureModal();
            }
        });

        // Close on backdrop click
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.hideSignatureModal();
            }
        });
    }

    /**
     * Show signature modal
     */
    showSignatureModal() {
        this.signatureModal.classList.remove('hidden');
        this.signaturePad.clear();
        document.getElementById('sig-text-input').value = '';
        document.getElementById('sig-preview').textContent = 'Preview';
    }

    /**
     * Hide signature modal
     */
    hideSignatureModal() {
        this.signatureModal.classList.add('hidden');
    }

    /**
     * Export the PDF with annotations
     */
    async exportPDF() {
        if (!this.pdfHandler.isLoaded()) return;

        this.showLoading('Exporting PDF...');

        try {
            const annotations = this.canvasManager.getAllAnnotations();
            const originalBytes = this.pdfHandler.getOriginalBytes();

            const modifiedPdfBytes = await this.exporter.exportPDF(
                originalBytes,
                annotations,
                this.currentScale
            );

            // Generate filename: use loaded name if present, otherwise a unique random name
            const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim();
            const hasRealFilename = baseName && baseName !== 'document';
            const exportName = hasRealFilename
                ? `${baseName}-edited.pdf`
                : `pdf-export-${Math.random().toString(36).slice(2, 10)}.pdf`;

            this.exporter.downloadPDF(modifiedPdfBytes, exportName);
            this.hideLoading();
        } catch (error) {
            console.error('Export error:', error);
            alert('Error exporting PDF: ' + error.message);
            this.hideLoading();
        }
    }

    /**
     * Update undo/redo button states
     */
    updateHistoryButtons() {
        // For now, just enable/disable based on whether there are any canvases
        const hasCanvases = this.canvasManager.canvases.size > 0;
        this.btnUndo.disabled = !hasCanvases;
        this.btnRedo.disabled = !hasCanvases;
    }

    /**
     * Show loading overlay
     */
    showLoading(text = 'Loading...') {
        this.loadingText.textContent = text;
        this.loadingOverlay.classList.remove('hidden');
    }

    /**
     * Hide loading overlay
     */
    hideLoading() {
        this.loadingOverlay.classList.add('hidden');
    }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.pdfEditor = new PDFEditorApp();
});
