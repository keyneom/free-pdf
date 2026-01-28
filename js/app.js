/**
 * PDF Editor - Main Application
 */

import { PDFHandler } from './pdf-handler.js';
import { CanvasManager } from './canvas-manager.js';
import { PDFExporter } from './export.js';
import { SignaturePad } from './signature-pad.js';
import { emailTemplates } from './email-templates.js';
import { BulkFillHandler } from './bulk-fill.js';

function escapeHtml(s) {
    if (s == null) return '';
    const t = String(s);
    return t
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

class PDFEditorApp {
    constructor() {
        this.pdfHandler = new PDFHandler();
        this.canvasManager = new CanvasManager();
        this.exporter = new PDFExporter();
        this.signaturePad = new SignaturePad();
        this.bulkFillHandler = new BulkFillHandler();

        this.currentScale = 1.0;
        this.fileName = 'document.pdf';
        this.documentHash = null;

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
        this.setupSendModal();
        this.setupTemplatesModal();
        this.setupTemplateEditModal();
        this.setupBulkFillModal();
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
        this.btnSend = document.getElementById('btn-send');
        this.btnTemplates = document.getElementById('btn-templates');
        this.btnBulkFill = document.getElementById('btn-bulk-fill');
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

        // Send / Templates modals
        this.sendModal = document.getElementById('send-modal');
        this.templatesModal = document.getElementById('templates-modal');
        this.templateEditModal = document.getElementById('template-edit-modal');
        this.bulkFillModal = document.getElementById('bulk-fill-modal');
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

        // Send / Templates / Bulk Fill
        this.btnSend?.addEventListener('click', () => this.showSendModal());
        this.btnTemplates?.addEventListener('click', () => this.showTemplatesModal());
        this.btnBulkFill?.addEventListener('click', async () => this.showBulkFillModal());

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
     * Compute SHA-256 hash of document bytes for audit association
     * @param {ArrayBuffer} buffer
     * @returns {Promise<string>} Hex-encoded hash
     */
    async computeDocumentHash(buffer) {
        const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
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
            this.documentHash = await this.computeDocumentHash(arrayBuffer);
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
            this.btnSend.disabled = false;
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
        const intentCheck = document.getElementById('sig-intent');
        const consentCheck = document.getElementById('sig-consent');
        const nameInput = document.getElementById('sig-name');
        const emailInput = document.getElementById('sig-email');
        const sigCanvas = document.getElementById('signature-canvas');

        const updateApplyState = () => this.updateSignatureApplyState();

        // Close modal
        closeBtn.addEventListener('click', () => this.hideSignatureModal());
        cancelBtn.addEventListener('click', () => this.hideSignatureModal());

        // Clear signature
        clearBtn.addEventListener('click', () => {
            this.signaturePad.clear();
            updateApplyState();
        });

        // Consent + identity
        [intentCheck, consentCheck].forEach((el) => el?.addEventListener('change', updateApplyState));
        nameInput?.addEventListener('input', updateApplyState);
        emailInput?.addEventListener('input', updateApplyState);

        // Tab switching
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');

                const tabName = tab.dataset.tab;
                document.getElementById('sig-draw-area').classList.toggle('active', tabName === 'draw');
                document.getElementById('sig-type-area').classList.toggle('active', tabName === 'type');

                this.signaturePad.setMode(tabName);
                updateApplyState();
            });
        });

        // Text input
        textInput.addEventListener('input', (e) => {
            this.signaturePad.setTypedText(e.target.value);
            this.signaturePad.updatePreview(preview);
            updateApplyState();
        });

        // Font options
        fontOptions.forEach(option => {
            option.addEventListener('change', (e) => {
                this.signaturePad.setFontStyle(e.target.value);
                this.signaturePad.updatePreview(preview);
            });
        });

        // Draw end (mouseup/touchend on pad) to refresh Apply state
        if (sigCanvas) {
            sigCanvas.addEventListener('mouseup', updateApplyState);
            sigCanvas.addEventListener('touchend', updateApplyState);
        }

        // Apply signature
        applyBtn.addEventListener('click', () => {
            if (!this.validateAndApplySignature()) return;
            this.hideSignatureModal();
        });

        // Close on backdrop click
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.hideSignatureModal();
            }
        });
    }

    /**
     * Update Insert Signature button enabled state based on form validity
     */
    updateSignatureApplyState() {
        const intentCheck = document.getElementById('sig-intent');
        const consentCheck = document.getElementById('sig-consent');
        const nameInput = document.getElementById('sig-name');
        const applyBtn = document.getElementById('sig-apply');
        if (!intentCheck || !consentCheck || !nameInput || !applyBtn) return;

        const intent = intentCheck.checked;
        const consent = consentCheck.checked;
        const name = (nameInput.value || '').trim();
        const hasSignature = this.signaturePad.mode === 'draw'
            ? !this.signaturePad.isEmpty()
            : !!this.signaturePad.typedText?.trim();

        applyBtn.disabled = !(intent && consent && name && hasSignature);
    }

    /**
     * Validate form, build meta, set signature, and switch to signature tool
     * @returns {boolean} true if applied
     */
    validateAndApplySignature() {
        const intentCheck = document.getElementById('sig-intent');
        const consentCheck = document.getElementById('sig-consent');
        const nameInput = document.getElementById('sig-name');
        const emailInput = document.getElementById('sig-email');
        const applyBtn = document.getElementById('sig-apply');
        if (!intentCheck || !consentCheck || !nameInput || !applyBtn) return false;

        const intent = intentCheck.checked;
        const consent = consentCheck.checked;
        const name = (nameInput.value || '').trim();
        const email = (emailInput?.value || '').trim();
        const dataUrl = this.signaturePad.getDataUrl();

        if (!intent || !consent || !name || !dataUrl) return false;

        const meta = {
            signerName: name,
            signerEmail: email || undefined,
            intentAccepted: intent,
            consentAccepted: consent,
            documentFilename: this.fileName || '',
            documentHash: this.documentHash || undefined
        };

        this.canvasManager.setSignature(dataUrl, meta);
        this.canvasManager.setTool('signature');
        this.toolButtons.forEach(b => b.classList.remove('active'));
        document.querySelector('[data-tool="signature"]')?.classList.add('active');
        return true;
    }

    /**
     * Show signature modal
     */
    showSignatureModal() {
        this.signatureModal.classList.remove('hidden');
        this.signaturePad.clear();
        this.signaturePad.setTypedText('');
        const textInput = document.getElementById('sig-text-input');
        if (textInput) textInput.value = '';
        const preview = document.getElementById('sig-preview');
        if (preview) preview.textContent = 'Preview';

        const intentCheck = document.getElementById('sig-intent');
        const consentCheck = document.getElementById('sig-consent');
        const nameInput = document.getElementById('sig-name');
        const emailInput = document.getElementById('sig-email');
        const applyBtn = document.getElementById('sig-apply');
        if (intentCheck) intentCheck.checked = false;
        if (consentCheck) consentCheck.checked = false;
        if (nameInput) nameInput.value = '';
        if (emailInput) emailInput.value = '';
        if (applyBtn) applyBtn.disabled = true;

        this.updateSignatureApplyState();
    }

    /**
     * Hide signature modal
     */
    hideSignatureModal() {
        this.signatureModal.classList.add('hidden');
    }

    /**
     * Export the PDF with annotations (shared logic)
     * @returns {Promise<{ bytes: Uint8Array; exportName: string } | null>}
     */
    async getExportedPDF() {
        if (!this.pdfHandler.isLoaded()) return null;

        const annotations = this.canvasManager.getAllAnnotations();
        const originalBytes = this.pdfHandler.getOriginalBytes();
        const modifiedPdfBytes = await this.exporter.exportPDF(
            originalBytes,
            annotations,
            this.currentScale
        );

        const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim();
        const hasRealFilename = baseName && baseName !== 'document';
        const exportName = hasRealFilename
            ? `${baseName}-edited.pdf`
            : `pdf-export-${Math.random().toString(36).slice(2, 10)}.pdf`;

        return { bytes: modifiedPdfBytes, exportName };
    }

    /**
     * Export the PDF with annotations
     */
    async exportPDF() {
        if (!this.pdfHandler.isLoaded()) return;

        this.showLoading('Exporting PDF...');
        try {
            const result = await this.getExportedPDF();
            if (result) {
                this.exporter.downloadPDF(result.bytes, result.exportName);
            }
        } catch (error) {
            console.error('Export error:', error);
            alert('Error exporting PDF: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    /**
     * Build context for email template placeholders
     * @param {string} exportName
     * @returns {{ filename: string; date: string; signatureSummary: string; signerNames: string; pageCount: number; documentHash: string }}
     */
    buildEmailContext(exportName) {
        const annotations = this.canvasManager.getAllAnnotations();
        const signatures = [];
        const signerSet = new Set();

        for (const page of annotations) {
            for (const ann of page.annotations) {
                if (ann.type === 'signature' && ann.object._signatureMeta) {
                    const m = ann.object._signatureMeta;
                    signatures.push({ name: m.signerName, ts: m.timestamp });
                    if (m.signerName) signerSet.add(m.signerName);
                }
            }
        }

        const signatureSummary =
            signatures.length === 0
                ? 'No signatures.'
                : signatures
                      .map((s) => `- Signed by ${s.name || '—'} on ${s.timestamp ? new Date(s.timestamp).toLocaleString() : '—'}`)
                      .join('\n');
        const signerNames = [...signerSet].join(', ') || '—';

        return {
            filename: exportName,
            date: new Date().toLocaleString(),
            signatureSummary,
            signerNames,
            pageCount: this.pdfHandler.totalPages || 0,
            documentHash: this.documentHash || ''
        };
    }

    /**
     * Send via email: download PDF, open mailto with template-filled subject/body
     */
    async sendViaEmail() {
        if (!this.pdfHandler.isLoaded()) return;

        const sel = document.getElementById('send-template-select');
        const subjectEl = document.getElementById('send-subject');
        const bodyEl = document.getElementById('send-body');
        if (!sel || !subjectEl || !bodyEl) return;

        const tpl = emailTemplates.getById(sel.value) || emailTemplates.getDefault();
        const subject = (subjectEl.value || '').trim();
        const body = (bodyEl.value || '').trim();
        if (!subject || !body) {
            alert('Please provide a subject and body.');
            return;
        }

        this.showLoading('Preparing email...');
        try {
            const result = await this.getExportedPDF();
            if (!result) {
                this.hideLoading();
                return;
            }
            this.exporter.downloadPDF(result.bytes, result.exportName);
            const ctx = this.buildEmailContext(result.exportName);
            const filled = emailTemplates.fill({ subject, body }, ctx);
            const mailto = `mailto:?subject=${encodeURIComponent(filled.subject)}&body=${encodeURIComponent(filled.body)}`;
            window.location.href = mailto;
            this.hideSendModal();
        } catch (error) {
            console.error('Send error:', error);
            alert('Error preparing email: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    /**
     * Populate Send modal with default template and fill subject/body
     */
    refreshSendModal() {
        const sel = document.getElementById('send-template-select');
        const subjectEl = document.getElementById('send-subject');
        const bodyEl = document.getElementById('send-body');
        if (!sel || !subjectEl || !bodyEl) return;

        const templates = emailTemplates.getTemplates();
        const defaultTpl = emailTemplates.getDefault();

        sel.innerHTML = templates.map((t) => `<option value="${t.id}" ${t.id === defaultTpl.id ? 'selected' : ''}>${escapeHtml(t.name)}</option>`).join('');

        const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim();
        const hasRealFilename = baseName && baseName !== 'document';
        const exportName = hasRealFilename ? `${baseName}-edited.pdf` : 'document.pdf';
        const ctx = this.buildEmailContext(exportName);
        const filled = emailTemplates.fill(defaultTpl, ctx);
        subjectEl.value = filled.subject;
        bodyEl.value = filled.body;
    }

    showSendModal() {
        if (!this.pdfHandler.isLoaded()) return;
        this.refreshSendModal();
        this.sendModal.classList.remove('hidden');
    }

    hideSendModal() {
        this.sendModal.classList.add('hidden');
    }

    setupSendModal() {
        const modal = this.sendModal;
        const closeBtn = document.getElementById('send-modal-close');
        const cancelBtn = document.getElementById('send-cancel');
        const doBtn = document.getElementById('send-do');
        const sel = document.getElementById('send-template-select');
        const subjectEl = document.getElementById('send-subject');
        const bodyEl = document.getElementById('send-body');
        const manageLink = document.getElementById('send-manage-templates');

        closeBtn?.addEventListener('click', () => this.hideSendModal());
        cancelBtn?.addEventListener('click', () => this.hideSendModal());
        doBtn?.addEventListener('click', () => this.sendViaEmail());

        sel?.addEventListener('change', () => {
            const t = emailTemplates.getById(sel.value);
            if (!t) return;
            const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim();
            const hasRealFilename = baseName && baseName !== 'document';
            const exportName = hasRealFilename ? `${baseName}-edited.pdf` : 'document.pdf';
            const ctx = this.buildEmailContext(exportName);
            const filled = emailTemplates.fill(t, ctx);
            subjectEl.value = filled.subject;
            bodyEl.value = filled.body;
        });

        manageLink?.addEventListener('click', (e) => {
            e.preventDefault();
            this.hideSendModal();
            this.showTemplatesModal();
        });

        modal?.addEventListener('click', (e) => {
            if (e.target === modal) this.hideSendModal();
        });
    }

    showTemplatesModal() {
        this.renderTemplatesList();
        this.templatesModal.classList.remove('hidden');
    }

    hideTemplatesModal() {
        this.templatesModal.classList.add('hidden');
    }

    renderTemplatesList() {
        const list = document.getElementById('templates-list');
        if (!list) return;

        const store = emailTemplates.getTemplates();
        const defaultId = emailTemplates.getDefault().id;

        list.innerHTML = store
            .map(
                (t) =>
                    `<li class="${t.id === defaultId ? 'default-item' : ''}" data-id="${escapeHtml(t.id)}">
  <div class="tpl-info">
    <span class="tpl-name">${escapeHtml(t.name)}</span>
    <div class="tpl-meta">${escapeHtml(t.subject || '(no subject)')}</div>
  </div>
  <div class="tpl-actions">
    <button type="button" class="tpl-btn edit-btn" data-id="${escapeHtml(t.id)}">Edit</button>${!t.builtin ? `<button type="button" class="tpl-btn delete-btn" data-id="${escapeHtml(t.id)}">Delete</button>` : ''}
    ${t.id !== defaultId ? `<button type="button" class="tpl-btn set-default" data-id="${escapeHtml(t.id)}">Set default</button>` : ''}
  </div>
</li>`
            )
            .join('');

        list.querySelectorAll('.edit-btn').forEach((btn) => {
            btn.addEventListener('click', () => this.openTemplateEdit(btn.dataset.id));
        });
        list.querySelectorAll('.delete-btn').forEach((btn) => {
            btn.addEventListener('click', () => this.deleteTemplate(btn.dataset.id));
        });
        list.querySelectorAll('.tpl-btn.set-default').forEach((btn) => {
            btn.addEventListener('click', () => this.setDefaultTemplate(btn.dataset.id));
        });
    }

    openTemplateEdit(id) {
        const title = document.getElementById('template-edit-title');
        const idEl = document.getElementById('tpl-edit-id');
        const nameEl = document.getElementById('tpl-edit-name');
        const subjectEl = document.getElementById('tpl-edit-subject');
        const bodyEl = document.getElementById('tpl-edit-body');

        if (id) {
            const t = emailTemplates.getById(id);
            if (!t) return;
            title.textContent = 'Edit template';
            idEl.value = t.id;
            nameEl.value = t.name;
            subjectEl.value = t.subject || '';
            bodyEl.value = t.body || '';
        } else {
            title.textContent = 'Add template';
            idEl.value = '';
            nameEl.value = '';
            subjectEl.value = '{{filename}}';
            bodyEl.value = `Please find attached {{filename}}.\n\nDocument summary:\n- Pages: {{pageCount}}\n{{signatureSummary}}\n{{documentHash}}\n\n{{attachmentNote}}`;
        }
        this.templateEditModal.classList.remove('hidden');
    }

    saveTemplateEdit() {
        const idEl = document.getElementById('tpl-edit-id');
        const nameEl = document.getElementById('tpl-edit-name');
        const subjectEl = document.getElementById('tpl-edit-subject');
        const bodyEl = document.getElementById('tpl-edit-body');
        const id = (idEl?.value || '').trim();
        const name = (nameEl?.value || '').trim();
        const subject = subjectEl?.value || '';
        const body = bodyEl?.value || '';

        if (!name) {
            alert('Please enter a template name.');
            return;
        }

        if (id) {
            emailTemplates.update(id, { name, subject, body });
        } else {
            emailTemplates.add({ name, subject, body });
        }
        this.hideTemplateEditModal();
        this.renderTemplatesList();
        this.refreshSendModal();
    }

    hideTemplateEditModal() {
        this.templateEditModal.classList.add('hidden');
    }

    deleteTemplate(id) {
        if (!id || !confirm('Delete this template?')) return;
        emailTemplates.remove(id);
        this.renderTemplatesList();
        this.refreshSendModal();
    }

    setDefaultTemplate(id) {
        if (!id) return;
        emailTemplates.setDefault(id);
        this.renderTemplatesList();
        this.refreshSendModal();
    }

    setupTemplatesModal() {
        const modal = this.templatesModal;
        const closeBtn = document.getElementById('templates-modal-close');
        const addBtn = document.getElementById('tpl-add');
        const exportBtn = document.getElementById('tpl-export');
        const importInput = document.getElementById('tpl-import-input');

        closeBtn?.addEventListener('click', () => this.hideTemplatesModal());
        addBtn?.addEventListener('click', () => this.openTemplateEdit(null));
        exportBtn?.addEventListener('click', () => {
            const json = emailTemplates.exportJson();
            const blob = new Blob([json], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `free-pdf-email-templates-${new Date().toISOString().slice(0, 10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        });
        importInput?.addEventListener('change', async (e) => {
            const f = e.target.files?.[0];
            if (!f) return;
            const replaceCb = document.getElementById('tpl-import-replace-cb');
            const replace = !!replaceCb?.checked;
            try {
                const text = await f.text();
                const { imported, errors } = emailTemplates.importJson(text, { replace });
                if (errors.length) alert(`Import completed with issues:\n${errors.join('\n')}`);
                else alert(`Imported ${imported} template(s).`);
                this.renderTemplatesList();
                this.refreshSendModal();
            } catch (err) {
                alert('Import failed: ' + (err.message || 'Invalid file'));
            }
            importInput.value = '';
        });

        modal?.addEventListener('click', (e) => {
            if (e.target === modal) this.hideTemplatesModal();
        });
    }

    setupTemplateEditModal() {
        const closeBtn = document.getElementById('template-edit-close');
        const cancelBtn = document.getElementById('tpl-edit-cancel');
        const saveBtn = document.getElementById('tpl-edit-save');
        const modal = this.templateEditModal;

        closeBtn?.addEventListener('click', () => this.hideTemplateEditModal());
        cancelBtn?.addEventListener('click', () => this.hideTemplateEditModal());
        saveBtn?.addEventListener('click', () => this.saveTemplateEdit());

        modal?.addEventListener('click', (e) => {
            if (e.target === modal) this.hideTemplateEditModal();
        });
    }

    /**
     * Setup bulk fill modal
     */
    setupBulkFillModal() {
        const modal = this.bulkFillModal;
        const closeBtn = document.getElementById('bulk-fill-modal-close');
        const cancelBtn = document.getElementById('bulk-fill-cancel');
        const processBtn = document.getElementById('bulk-fill-process');
        const templateInput = document.getElementById('bulk-fill-template');
        const csvInput = document.getElementById('bulk-fill-csv');
        const filenameInput = document.getElementById('bulk-fill-filename');
        const mappingDiv = document.getElementById('bulk-fill-mapping');
        const mappingList = document.getElementById('bulk-fill-mapping-list');
        const progressDiv = document.getElementById('bulk-fill-progress');
        const progressFill = document.getElementById('bulk-fill-progress-fill');
        const progressText = document.getElementById('bulk-fill-progress-text');
        const templateStatus = document.getElementById('bulk-fill-template-status');

        let templateBytes = null;
        let csvText = null;
        let pdfFieldNames = [];
        let csvHeaders = [];

        closeBtn?.addEventListener('click', () => this.hideBulkFillModal());
        cancelBtn?.addEventListener('click', () => this.hideBulkFillModal());

        const arrayBufferFromUint8 = (u8) => {
            if (!u8) return null;
            return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
        };

        const setTemplateFromBytes = async (bytes) => {
            templateBytes = bytes;
            pdfFieldNames = templateBytes ? await this.bulkFillHandler.extractFormFieldNames(templateBytes) : [];
            this.updateBulkFillMapping();
        };

        templateInput?.addEventListener('change', async (e) => {
            const file = e.target.files?.[0];
            if (!file) return;

            try {
                await setTemplateFromBytes(await file.arrayBuffer());
                
                if (pdfFieldNames.length === 0) {
                    alert('No form fields found in PDF. Make sure the PDF has form fields created with the Form Text Field or Form Checkbox tools.');
                } else {
                    this.updateBulkFillMapping();
                }
            } catch (error) {
                console.error('Error loading template:', error);
                alert('Error loading PDF template: ' + error.message);
            }
        });

        csvInput?.addEventListener('change', async (e) => {
            const file = e.target.files?.[0];
            if (!file) return;

            try {
                csvText = await file.text();
                const rows = this.bulkFillHandler.parseCSV(csvText);
                if (rows.length > 0) {
                    csvHeaders = Object.keys(rows[0]);
                    this.updateBulkFillMapping();
                } else {
                    alert('CSV file appears to be empty or invalid.');
                }
            } catch (error) {
                console.error('Error loading CSV:', error);
                alert('Error loading CSV file: ' + error.message);
            }
        });

        processBtn?.addEventListener('click', async () => {
            if (!csvText) {
                alert('Please upload a CSV file.');
                return;
            }

            // If the user didn't upload a template, default to the currently open document
            if (!templateBytes) {
                if (this.pdfHandler.isLoaded()) {
                    const exported = await this.getExportedPDF();
                    if (exported?.bytes) {
                        await setTemplateFromBytes(arrayBufferFromUint8(exported.bytes));
                    }
                }
            }

            if (!templateBytes) {
                alert('No PDF template available. Open a PDF in the editor (or upload a template in this modal).');
                return;
            }

            // Build field mapping
            const fieldMapping = {};
            const mappingInputs = mappingList.querySelectorAll('select');
            mappingInputs.forEach(select => {
                const csvColumn = select.dataset.csvColumn;
                const pdfField = select.value;
                if (pdfField && pdfField !== '') {
                    fieldMapping[csvColumn] = pdfField;
                }
            });

            const filenameTemplate = filenameInput?.value || 'document-{{row}}.pdf';

            // Show progress
            progressDiv.classList.remove('hidden');
            processBtn.disabled = true;
            cancelBtn.disabled = true;

            try {
                await this.bulkFillHandler.processBulkFill(
                    templateBytes,
                    csvText,
                    fieldMapping,
                    filenameTemplate,
                    (current, total) => {
                        const percent = Math.round((current / total) * 100);
                        progressFill.style.width = percent + '%';
                        progressText.textContent = `Processing ${current} of ${total}...`;
                    }
                );

                progressText.textContent = `Completed! ${csvHeaders.length > 0 ? this.bulkFillHandler.parseCSV(csvText).length : 0} PDFs downloaded.`;
                setTimeout(() => {
                    this.hideBulkFillModal();
                }, 2000);
            } catch (error) {
                console.error('Bulk fill error:', error);
                alert('Error processing bulk fill: ' + error.message);
                progressDiv.classList.add('hidden');
                processBtn.disabled = false;
                cancelBtn.disabled = false;
            }
        });

        modal?.addEventListener('click', (e) => {
            if (e.target === modal) this.hideBulkFillModal();
        });

        // Update mapping when fields are available
        this.updateBulkFillMapping = () => {
            if (pdfFieldNames.length === 0 || csvHeaders.length === 0) {
                mappingDiv.classList.add('hidden');
                processBtn.disabled = true;
                return;
            }

            mappingDiv.classList.remove('hidden');
            mappingList.innerHTML = '';

            csvHeaders.forEach(csvHeader => {
                const row = document.createElement('div');
                row.className = 'bulk-fill-mapping-row';
                row.innerHTML = `
                    <label>${escapeHtml(csvHeader)}</label>
                    <select class="bulk-fill-mapping-select" data-csv-column="${escapeHtml(csvHeader)}">
                        <option value="">-- Skip --</option>
                        ${pdfFieldNames.map(fieldName => 
                            `<option value="${escapeHtml(fieldName)}" ${csvHeader.toLowerCase() === fieldName.toLowerCase() ? 'selected' : ''}>${escapeHtml(fieldName)}</option>`
                        ).join('')}
                    </select>
                `;
                mappingList.appendChild(row);
            });

            processBtn.disabled = false;
        };

        // Expose minimal internals so showBulkFillModal() can set default template.
        this._bulkFillModalInternal = {
            reset: () => {
                templateBytes = null;
                csvText = null;
                pdfFieldNames = [];
                csvHeaders = [];
                if (templateStatus) templateStatus.textContent = '';
            },
            setTemplateFromOpenDoc: async () => {
                if (!this.pdfHandler.isLoaded()) {
                    if (templateStatus) templateStatus.textContent = 'No document open — upload a PDF template above.';
                    return;
                }
                if (templateStatus) templateStatus.textContent = 'Using the currently open document as the template.';
                const exported = await this.getExportedPDF();
                if (exported?.bytes) {
                    await setTemplateFromBytes(arrayBufferFromUint8(exported.bytes));
                }
            }
        };
    }

    async showBulkFillModal() {
        // Reset form
        document.getElementById('bulk-fill-template').value = '';
        document.getElementById('bulk-fill-csv').value = '';
        document.getElementById('bulk-fill-filename').value = 'document-{{row}}.pdf';
        document.getElementById('bulk-fill-mapping').classList.add('hidden');
        document.getElementById('bulk-fill-progress').classList.add('hidden');
        document.getElementById('bulk-fill-process').disabled = true;
        document.getElementById('bulk-fill-cancel').disabled = false;
        
        this.bulkFillModal.classList.remove('hidden');

        // Default template to the currently open/edited document.
        try {
            this._bulkFillModalInternal?.reset?.();
            await this._bulkFillModalInternal?.setTemplateFromOpenDoc?.();
        } catch (e) {
            console.warn('Failed to set bulk-fill template from open doc:', e);
        }
    }

    hideBulkFillModal() {
        this.bulkFillModal.classList.add('hidden');
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
