/**
 * PDF Editor - Main Application
 */

import { PDFHandler } from './pdf-handler.js';
import { CanvasManager } from './canvas-manager.js';
import { PDFExporter } from './export.js';
import { SignaturePad } from './signature-pad.js';
import { emailTemplates, setTemplatesBackend, getDefaultOnlyTemplatesStore } from './email-templates.js';
import { secureStorage } from './secure-storage.js';
import { BulkFillHandler } from './bulk-fill.js';
import { parseSigningMetadata, hasOurSigningMetadata, computeDocumentHash } from './signing-metadata.js';
import { loadFormFieldsFromPdf } from './load-form-fields.js';

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
        this.viewPages = []; // [{ id, docId, sourcePageNum, rotation }]
        /** Parsed from PDF Keywords when loading a doc that has our signing metadata */
        this.signingFlowMeta = null;
        /** Fields locked in an exported version; cannot be unlocked in this session (signatures: Set of labels, formFields: Set of field names) */
        this.lockedFromFile = { signatures: new Set(), formFields: new Set() };
        this._pendingSignatureImage = null;
        this._selectedSavedSig = null;
        
        /** Current mode: 'edit' (full editing) or 'fill' (signing/filling forms) */
        this.mode = 'edit';

        this.init();
    }

    /**
     * Initialize the application
     */
    init() {
        this.cacheElements();
        this.bindEvents();
        this.canvasManager.setOnHistoryChange(() => {
            this.updateHistoryButtons();
            this.updateSigningFlowBanner();
        });
        this.setupDragAndDrop();
        this.setupKeyboardShortcuts();
        this.initSignaturePad();
        this.setupVault();
        this.setupImageInsert();
        this.setupSendModal();
        this.setupTemplatesModal();
        this.setupTemplateEditModal();
        this.setupBulkFillModal();
        this.setupExpectedSignersModal();
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

        // Pages sidebar
        this.pagesSidebar = document.getElementById('pages-sidebar');
        this.pagesList = document.getElementById('pages-list');
        this.pagesAppendInput = document.getElementById('pages-append-input');
        this.pagesDeleteBtn = document.getElementById('pages-delete-btn');
        this.pagesExtractBtn = document.getElementById('pages-extract-btn');
        this.pagesSplitBtn = document.getElementById('pages-split-btn');
        this.pagesRotateBtn = document.getElementById('pages-rotate-btn');

        // Tool options
        this.toolOptions = document.getElementById('tool-options');

        // Dynamic hidden inputs
        this._imageInsertInput = null;

        // Signature modal
        this.signatureModal = document.getElementById('signature-modal');

        // Send / Templates modals
        this.sendModal = document.getElementById('send-modal');
        this.templatesModal = document.getElementById('templates-modal');
        this.templateEditModal = document.getElementById('template-edit-modal');
        this.bulkFillModal = document.getElementById('bulk-fill-modal');
    }

    setupImageInsert() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.hidden = true;
        input.id = 'image-insert-input';
        document.body.appendChild(input);
        this._imageInsertInput = input;

        input.addEventListener('change', async (e) => {
            const file = e.target.files?.[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                const dataUrl = reader.result;
                this.canvasManager.setPendingImage(dataUrl);
                this.canvasManager.setTool('image');
                this.toolButtons.forEach((b) => b.classList.remove('active'));
                document.querySelector('[data-tool="image"]')?.classList.add('active');
                this.showToolOptions('image');
            };
            reader.readAsDataURL(file);
            input.value = '';
        });
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
            const sigOpen = !this.signatureModal.classList.contains('hidden');
            if (sigOpen && this.signaturePad?.mode === 'draw') this.signaturePad.undo();
            else this.canvasManager.undo();
            this.updateHistoryButtons();
        });
        this.btnRedo.addEventListener('click', () => {
            const sigOpen = !this.signatureModal.classList.contains('hidden');
            if (sigOpen && this.signaturePad?.mode === 'draw') this.signaturePad.redo();
            else this.canvasManager.redo();
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

        // Mode toggle
        document.getElementById('mode-edit')?.addEventListener('click', () => this.switchMode('edit'));
        document.getElementById('mode-fill')?.addEventListener('click', () => this.switchMode('fill'));

        // Listen for signature field placement to hide hint
        window.addEventListener('signature-field-placed', () => {
            this.hideToolHint();
            this.applyToolCursor('select');
        });

        // Page navigation
        this.btnPrevPage.addEventListener('click', () => this.goToPage(this.pdfHandler.currentPage - 1));
        this.btnNextPage.addEventListener('click', () => this.goToPage(this.pdfHandler.currentPage + 1));
        this.pageInput.addEventListener('change', (e) => this.goToPage(parseInt(e.target.value)));

        // Zoom controls
        this.btnZoomIn.addEventListener('click', () => this.zoom(this.currentScale + 0.25));
        this.btnZoomOut.addEventListener('click', () => this.zoom(this.currentScale - 0.25));
        this.btnFitWidth.addEventListener('click', () => this.fitWidth());

        // Signature modals
        this.setupSignatureModal();
        this.setupSignatureFieldModal();

        // Pages sidebar
        this.setupPagesSidebar();
        
        // Fields sidebar
        this.setupFieldsSidebar();
        
        // Field Properties sidebar (Edit mode)
        this.setupFieldPropertiesSidebar();
        
        // Listen for field selection to show/hide Field Properties sidebar
        window.addEventListener('field-properties-show', (e) => {
            if (this.mode === 'edit' && e.detail?.object && e.detail?.canvas) {
                this.showFieldPropertiesSidebar(e.detail.object, e.detail.canvas, e.detail.annotationType);
            }
        });
        window.addEventListener('field-properties-hide', () => {
            this.hideFieldPropertiesSidebar();
        });
        
        // Listen for field updates to refresh sidebar
        window.addEventListener('field-updated', () => {
            if (this.mode === 'fill') {
                this.refreshFieldsSidebar();
            }
        });
        
        // Listen for signature field fills to refresh sidebar
        window.addEventListener('signature-field-filled', () => {
            if (this.mode === 'fill') {
                this.refreshFieldsSidebar();
            }
        });
        
        // Listen for form field selection
        window.addEventListener('form-field-selected', (e) => {
            if (this.mode === 'fill') {
                if (e.detail.annotationType === 'signature-field') {
                    const field = {
                        type: 'signature-field',
                        signatureFieldLabel: e.detail.object?._signatureFieldLabel || '',
                        object: e.detail.object,
                        pageId: e.detail.pageId
                    };
                    this.showSignatureModal({ targetField: field });
                } else {
                    this.focusSidebarInputForField(e.detail.object);
                }
            }
        });
    }

    setupPagesSidebar() {
        this.pagesAppendInput?.addEventListener('change', async (e) => {
            const file = e.target.files?.[0];
            if (!file) return;
            try {
                await this.appendPdfFile(file);
            } catch (err) {
                console.error('Append PDF failed:', err);
                alert('Failed to append PDF: ' + (err.message || err));
            } finally {
                this.pagesAppendInput.value = '';
            }
        });

        this.pagesDeleteBtn?.addEventListener('click', () => this.deleteSelectedPages());
        this.pagesExtractBtn?.addEventListener('click', () => this.extractSelectedPages());
        this.pagesSplitBtn?.addEventListener('click', () => this.splitPdfPrompt());
        this.pagesRotateBtn?.addEventListener('click', () => this.rotateSelectedPages());
    }

    /**
     * Set up fields sidebar
     */
    setupFieldsSidebar() {
        const closeBtn = document.getElementById('fields-sidebar-close');
        closeBtn?.addEventListener('click', () => {
            this.hideFieldsDetail();
            document.getElementById('fields-sidebar')?.classList.add('hidden');
        });
        
        const backBtn = document.getElementById('fields-back-btn');
        backBtn?.addEventListener('click', () => {
            this.hideFieldsDetail();
        });
    }

    /**
     * Set up Field Properties sidebar (Edit mode)
     */
    setupFieldPropertiesSidebar() {
        const closeBtn = document.getElementById('field-properties-close');
        closeBtn?.addEventListener('click', () => {
            this.hideFieldPropertiesSidebar();
        });
    }

    /**
     * Show the Field Properties sidebar for a form field
     */
    showFieldPropertiesSidebar(fieldObject, canvas, annotationType) {
        this._editingField = { object: fieldObject, canvas, annotationType };
        
        const sidebar = document.getElementById('field-properties-sidebar');
        const bodyEl = document.getElementById('field-properties-body');
        if (!sidebar || !bodyEl) return;
        
        // Build sidebar content based on field type
        const fieldName = annotationType === 'signature-field' 
            ? (fieldObject._signatureFieldLabel || '')
            : (fieldObject._fieldName || '');
        const labelText = annotationType === 'signature-field' ? 'Field Label' : 'Field Name';
        const placeholder = annotationType === 'signature-field' ? 'e.g. Tenant 1, Landlord' : 'e.g. tenant_name';
        const hintText = annotationType === 'signature-field' ? 'Used to match signers' : 'Used for CSV bulk filling';
        
        const isTextBased = ['textfield', 'date', 'dropdown'].includes(annotationType);
        const currentFontSize = isTextBased ? Math.round((fieldObject._valueFontSize || 12) * this.canvasManager.currentScale) : 12;
        const fontFamily = fieldObject._fontFamily || 'Arial';
        
        let html = `
            <div class="field-property-group">
                <label for="field-prop-name">${labelText}</label>
                <input type="text" id="field-prop-name" value="${this.escapeHtml(fieldName)}" placeholder="${placeholder}">
                <small id="field-prop-name-hint">${hintText}</small>
            </div>
        `;
        
        if (isTextBased) {
            html += `
                <div class="field-property-group">
                    <label for="field-prop-font-size">Font Size</label>
                    <input type="number" id="field-prop-font-size" value="${currentFontSize}" min="8" max="36">
                </div>
                <div class="field-property-group">
                    <label for="field-prop-font-family">Font</label>
                    <select id="field-prop-font-family">
                        <option value="Arial" ${fontFamily === 'Arial' ? 'selected' : ''}>Arial</option>
                        <option value="Times New Roman" ${fontFamily === 'Times New Roman' ? 'selected' : ''}>Times New Roman</option>
                        <option value="Courier New" ${fontFamily === 'Courier New' ? 'selected' : ''}>Courier New</option>
                        <option value="Georgia" ${fontFamily === 'Georgia' ? 'selected' : ''}>Georgia</option>
                    </select>
                </div>
            `;
        }
        
        if (annotationType === 'dropdown') {
            html += `
                <div class="field-property-group">
                    <label>Options</label>
                    <p class="field-properties-hint">Add options that users can select.</p>
                    <div class="options-list" id="dropdown-options-list"></div>
                    <div class="add-option-row" style="margin-top: 8px;">
                        <input type="text" id="new-option-input" placeholder="New option..." class="new-option-input">
                        <button type="button" id="add-option-btn" class="btn btn-sm btn-primary" title="Add option">
                            <svg viewBox="0 0 24 24" width="16" height="16"><path fill="currentColor" d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
                        </button>
                    </div>
                    <small class="field-properties-footer-hint" style="margin-top: 4px;">Drag to reorder. Click × to remove.</small>
                </div>
            `;
        }
        
        bodyEl.innerHTML = html;
        
        // Bind field name
        const nameInput = document.getElementById('field-prop-name');
        const nameHint = document.getElementById('field-prop-name-hint');
        if (nameInput) {
            nameInput.addEventListener('input', (e) => {
                this.updateFieldProperty('name', e.target.value.trim());
            });
            nameInput.addEventListener('blur', (e) => {
                const value = e.target.value.trim();
                const isDup = value && this.canvasManager.isFieldNameDuplicate(value, fieldObject);
                if (nameHint) {
                    nameHint.textContent = isDup ? '⚠ This name is already used' : hintText;
                    nameHint.style.color = isDup ? '#ef4444' : '';
                }
                nameInput.style.borderColor = isDup ? '#ef4444' : '';
            });
        }
        
        // Bind font controls
        if (isTextBased) {
            const fontSizeInput = document.getElementById('field-prop-font-size');
            if (fontSizeInput) {
                fontSizeInput.addEventListener('change', (e) => {
                    const val = parseInt(e.target.value);
                    if (val >= 8 && val <= 36) this.updateFieldProperty('fontSize', val);
                });
            }
            const fontFamilyInput = document.getElementById('field-prop-font-family');
            if (fontFamilyInput) {
                fontFamilyInput.addEventListener('change', (e) => {
                    this.updateFieldProperty('fontFamily', e.target.value);
                });
            }
        }
        
        // Bind dropdown options
        if (annotationType === 'dropdown') {
            const addBtn = document.getElementById('add-option-btn');
            const newInput = document.getElementById('new-option-input');
            addBtn?.addEventListener('click', () => this.addDropdownOption());
            newInput?.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') { e.preventDefault(); this.addDropdownOption(); }
            });
            this.refreshDropdownOptionsList();
        }
        
        sidebar.classList.remove('hidden');
    }

    /**
     * Hide the Field Properties sidebar
     */
    hideFieldPropertiesSidebar() {
        const sidebar = document.getElementById('field-properties-sidebar');
        sidebar?.classList.add('hidden');
        this._editingField = null;
    }

    /**
     * Update a field property from the sidebar
     */
    updateFieldProperty(prop, value) {
        if (!this._editingField) return;
        const { object: obj, canvas, annotationType } = this._editingField;
        
        if (prop === 'name') {
            if (annotationType === 'signature-field') {
                obj._signatureFieldLabel = value;
                const textObj = obj.getObjects?.()?.[1];
                if (textObj?.type === 'text') {
                    textObj.set('text', `${value || 'Signature'}\n(Double-click to sign)`);
                }
            } else {
                obj._fieldName = value;
                if (annotationType === 'radio') obj._radioGroup = value;
                const objects = obj.getObjects?.() || [];
                const minObjects = annotationType === 'dropdown' ? 4 : 3;
                if (objects.length >= minObjects) {
                    const labelObj = objects[1];
                    if (labelObj?.type === 'text') {
                        let defaultLabel = 'Text Field';
                        if (annotationType === 'date') defaultLabel = 'Date Field';
                        if (annotationType === 'dropdown') defaultLabel = 'Dropdown';
                        labelObj.set('text', value || defaultLabel);
                    }
                }
            }
        } else if (prop === 'fontSize') {
            obj._valueFontSize = value / this.canvasManager.currentScale;
            const objects = obj.getObjects?.() || [];
            const valueText = objects[2];
            if (valueText && (valueText.type === 'text' || valueText.type === 'i-text')) {
                valueText.set('fontSize', value / this.canvasManager.currentScale);
            }
        } else if (prop === 'fontFamily') {
            obj._fontFamily = value;
            const objects = obj.getObjects?.() || [];
            const valueText = objects[2];
            if (valueText && (valueText.type === 'text' || valueText.type === 'i-text')) {
                valueText.set('fontFamily', value);
            }
        }
        canvas.renderAll();
    }

    /**
     * Refresh the dropdown options list in the sidebar
     */
    refreshDropdownOptionsList() {
        const listEl = document.getElementById('dropdown-options-list');
        if (!listEl || !this._editingField || this._editingField.annotationType !== 'dropdown') return;
        
        const options = this._editingField.object._options || [];
        
        listEl.innerHTML = '';
        
        if (options.length === 0) {
            listEl.innerHTML = '<div style="color: var(--text-secondary); font-size: 0.8rem; padding: 8px 0;">No options yet. Add one below.</div>';
            return;
        }
        
        options.forEach((opt, index) => {
            const item = document.createElement('div');
            item.className = 'option-item';
            item.draggable = true;
            item.dataset.index = index;
            
            item.innerHTML = `
                <span class="option-item-drag" title="Drag to reorder">
                    <svg viewBox="0 0 24 24" width="16" height="16"><path fill="currentColor" d="M11 18c0 1.1-.9 2-2 2s-2-.9-2-2 .9-2 2-2 2 .9 2 2zm-2-8c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm0-6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm6 4c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm0 2c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm0 6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"/></svg>
                </span>
                <input type="text" class="option-item-text" value="${this.escapeHtml(opt)}" data-index="${index}">
                <button type="button" class="option-item-remove" title="Remove option" data-index="${index}">
                    <svg viewBox="0 0 24 24" width="16" height="16"><path fill="currentColor" d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
                </button>
            `;
            
            // Edit option text
            const textInput = item.querySelector('.option-item-text');
            textInput?.addEventListener('input', (e) => {
                this.updateDropdownOption(index, e.target.value);
            });
            
            // Remove option
            const removeBtn = item.querySelector('.option-item-remove');
            removeBtn?.addEventListener('click', () => {
                this.removeDropdownOption(index);
            });
            
            // Drag and drop
            item.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('text/plain', index);
                item.classList.add('dragging');
            });
            
            item.addEventListener('dragend', () => {
                item.classList.remove('dragging');
            });
            
            item.addEventListener('dragover', (e) => {
                e.preventDefault();
            });
            
            item.addEventListener('drop', (e) => {
                e.preventDefault();
                const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
                const toIndex = index;
                if (fromIndex !== toIndex) {
                    this.reorderDropdownOptions(fromIndex, toIndex);
                }
            });
            
            listEl.appendChild(item);
        });
    }

    /**
     * Add a new dropdown option
     */
    addDropdownOption() {
        const input = document.getElementById('new-option-input');
        if (!input || !this._editingField) return;
        
        const value = input.value.trim();
        if (!value) return;
        
        const options = this._editingField.object._options || [];
        options.push(value);
        this._editingField.object._options = options;
        
        input.value = '';
        this.refreshDropdownOptionsList();
        this._editingField.canvas.renderAll();
        
        // Focus input for next entry
        input.focus();
    }

    /**
     * Update a dropdown option
     */
    updateDropdownOption(index, newValue) {
        if (!this._editingField) return;
        
        const options = this._editingField.object._options || [];
        if (index >= 0 && index < options.length) {
            options[index] = newValue;
            this._editingField.object._options = options;
            this._editingField.canvas.renderAll();
        }
    }

    /**
     * Remove a dropdown option
     */
    removeDropdownOption(index) {
        if (!this._editingField) return;
        
        const options = this._editingField.object._options || [];
        if (index >= 0 && index < options.length) {
            options.splice(index, 1);
            this._editingField.object._options = options;
            this.refreshDropdownOptionsList();
            this._editingField.canvas.renderAll();
        }
    }

    /**
     * Reorder dropdown options
     */
    reorderDropdownOptions(fromIndex, toIndex) {
        if (!this._editingField) return;
        
        const options = this._editingField.object._options || [];
        const [moved] = options.splice(fromIndex, 1);
        options.splice(toIndex, 0, moved);
        this._editingField.object._options = options;
        this.refreshDropdownOptionsList();
        this._editingField.canvas.renderAll();
    }

    /**
     * Escape HTML for safe rendering
     */
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text || '';
        return div.innerHTML;
    }

    /**
     * Refresh the fields sidebar with current form fields
     */
    refreshFieldsSidebar() {
        const fieldsList = document.getElementById('fields-list');
        if (!fieldsList) return;

        // Get all form fields from all pages
        const allFields = [];
        const annotations = this.canvasManager.getAllAnnotations?.() || [];
        
        annotations.forEach((pageData, pageIndex) => {
            const pageAnnotations = pageData.annotations || [];
            pageAnnotations.forEach(ann => {
                if (['textfield', 'checkbox', 'radio', 'dropdown', 'date', 'signature-field'].includes(ann.type)) {
                    const fieldName = ann.data?._fieldName || '';
                    const locked = !!ann.object?._fieldLocked;
                    const fieldId = fieldName.trim();
                    const permanentlyLocked = !!fieldId && this.lockedFromFile?.formFields?.has(fieldId);
                    allFields.push({
                        type: ann.type,
                        fieldName,
                        fieldValue: ann.data?._fieldValue || '',
                        checked: ann.data?._checked || false,
                        selectedOption: ann.data?._selectedOption || '',
                        signatureFieldLabel: ann.data?._signatureFieldLabel || '',
                        pageNum: pageIndex + 1,
                        pageId: pageData.pageId,
                        object: ann.object,
                        locked,
                        permanentlyLocked
                    });
                } else if (ann.type === 'signature' && ann.data?._signatureMeta?.replacedFieldLabel) {
                    // Filled signature field: show in sidebar so the field doesn't "disappear" after signing
                    const meta = ann.data._signatureMeta;
                    const sigLabel = (meta.replacedFieldLabel || 'Signature').trim();
                    const permanentlyLocked = !!sigLabel && this.lockedFromFile?.signatures?.has(sigLabel);
                    allFields.push({
                        type: 'signature-filled',
                        fieldName: '',
                        fieldValue: '',
                        checked: false,
                        selectedOption: '',
                        signatureFieldLabel: meta.replacedFieldLabel || 'Signature',
                        signerName: meta.signerName || '',
                        signedAt: meta.timestamp || '',
                        pageNum: pageIndex + 1,
                        pageId: pageData.pageId,
                        object: ann.object,
                        locked: true,
                        permanentlyLocked
                    });
                }
            });
        });

        // Sort by page, then reading order (top-left to bottom-right)
        allFields.sort((a, b) => {
            if (a.pageNum !== b.pageNum) return a.pageNum - b.pageNum;
            const rectA = a.object?.getBoundingRect?.() ?? { top: 0, left: 0 };
            const rectB = b.object?.getBoundingRect?.() ?? { top: 0, left: 0 };
            if (Math.abs(rectA.top - rectB.top) > 8) return rectA.top - rectB.top;
            return rectA.left - rectB.left;
        });

        // Render fields
        fieldsList.innerHTML = '';
        
        if (allFields.length === 0) {
            fieldsList.innerHTML = '<div style="padding: 20px; text-align: center; color: var(--text-secondary);">No form fields found in this document.</div>';
            return;
        }

        allFields.forEach((field, index) => {
            const fieldItem = document.createElement('div');
            fieldItem.className = 'field-item';
            fieldItem.dataset.fieldIndex = index;
            fieldItem.dataset.pageId = field.pageId;
            
            const label = field.fieldName || field.signatureFieldLabel || `${field.type} ${index + 1}`;
            const typeLabel = field.type === 'signature-filled' ? 'signature (signed)' : field.type;
            
            // Create header
            const header = document.createElement('div');
            header.className = 'field-item-header';
            header.innerHTML = `
                <span class="field-item-label">${escapeHtml(label)}</span>
                <span class="field-item-actions">
                    <button type="button" class="field-configure-btn" title="Configure field" data-field-index="${index}">⚙</button>
                    <span class="field-item-type">${escapeHtml(typeLabel)}</span>
                </span>
            `;
            fieldItem.appendChild(header);
            
            // Configure button - store field reference for the click handler
            const configBtn = header.querySelector('.field-configure-btn');
            if (configBtn) {
                configBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    this.showFieldsDetail(field);
                });
            }
            
            // Create editable input based on field type
            const inputContainer = document.createElement('div');
            inputContainer.className = 'field-item-input';
            
            if (field.type === 'textfield' || field.type === 'date') {
                const input = document.createElement('input');
                input.type = field.type === 'date' ? 'date' : 'text';
                input.className = 'field-sidebar-input';
                input.placeholder = field.type === 'date' ? 'YYYY-MM-DD' : 'Enter value...';
                input.value = field.fieldValue || '';
                input.dataset.fieldIndex = index;
                input.disabled = !!field.locked;
                input.addEventListener('input', (e) => {
                    if (!field.locked) this.updateFieldValue(field, e.target.value);
                });
                input.addEventListener('focus', () => {
                    this.navigateToField(field);
                });
                inputContainer.appendChild(input);
                if (this.mode === 'fill') this._appendLockUnlock(inputContainer, field);
            } else if (field.type === 'checkbox') {
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.className = 'field-sidebar-checkbox';
                checkbox.checked = field.checked || false;
                checkbox.dataset.fieldIndex = index;
                checkbox.disabled = !!field.locked;
                checkbox.addEventListener('change', (e) => {
                    if (!field.locked) this.updateFieldChecked(field, e.target.checked);
                });
                checkbox.addEventListener('focus', () => {
                    this.navigateToField(field);
                });
                const checkLabel = document.createElement('label');
                checkLabel.textContent = field.checked ? 'Checked' : 'Unchecked';
                checkLabel.style.marginLeft = '8px';
                checkLabel.style.color = 'var(--text-secondary)';
                inputContainer.appendChild(checkbox);
                inputContainer.appendChild(checkLabel);
                if (this.mode === 'fill') this._appendLockUnlock(inputContainer, field);
            } else if (field.type === 'radio') {
                const radio = document.createElement('input');
                radio.type = 'checkbox'; // Use checkbox to allow toggle
                radio.className = 'field-sidebar-checkbox';
                radio.checked = field.checked || false;
                radio.dataset.fieldIndex = index;
                radio.disabled = !!field.locked;
                radio.addEventListener('change', (e) => {
                    if (!field.locked) this.updateFieldChecked(field, e.target.checked);
                });
                radio.addEventListener('focus', () => {
                    this.navigateToField(field);
                });
                const radioLabel = document.createElement('label');
                radioLabel.textContent = field.checked ? 'Selected' : 'Not selected';
                radioLabel.style.marginLeft = '8px';
                radioLabel.style.color = 'var(--text-secondary)';
                inputContainer.appendChild(radio);
                inputContainer.appendChild(radioLabel);
                if (this.mode === 'fill') this._appendLockUnlock(inputContainer, field);
            } else if (field.type === 'dropdown') {
                const select = document.createElement('select');
                select.className = 'field-sidebar-select';
                select.dataset.fieldIndex = index;
                
                // Add empty option
                const emptyOpt = document.createElement('option');
                emptyOpt.value = '';
                emptyOpt.textContent = 'Select...';
                select.appendChild(emptyOpt);
                
                // Add options from field
                const options = field.object?._options || [];
                options.forEach(opt => {
                    const option = document.createElement('option');
                    option.value = opt;
                    option.textContent = opt;
                    if (opt === field.selectedOption) option.selected = true;
                    select.appendChild(option);
                });
                
                select.disabled = !!field.locked;
                select.addEventListener('change', (e) => {
                    if (!field.locked) this.updateFieldDropdown(field, e.target.value);
                });
                select.addEventListener('focus', () => {
                    this.navigateToField(field);
                });
                inputContainer.appendChild(select);
                if (this.mode === 'fill') this._appendLockUnlock(inputContainer, field);
            } else if (field.type === 'signature-field') {
                const signBtn = document.createElement('button');
                signBtn.className = 'field-sidebar-sign-btn';
                signBtn.textContent = 'Click to Sign';
                signBtn.dataset.fieldIndex = index;
                
                signBtn.addEventListener('click', () => {
                    this.navigateToField(field);
                    this.highlightFieldOnCanvas(field);
                    this.showSignatureModal({ targetField: field });
                });
                
                inputContainer.appendChild(signBtn);
            } else if (field.type === 'signature-filled') {
                const signedInfo = document.createElement('div');
                signedInfo.className = 'field-sidebar-signed-info';
                const signer = field.signerName ? `Signed by ${escapeHtml(field.signerName)}` : 'Signed';
                const dateStr = field.signedAt ? (() => {
                    try {
                        const d = new Date(field.signedAt);
                        return isNaN(d.getTime()) ? '' : d.toLocaleString();
                    } catch (_) { return ''; }
                })() : '';
                signedInfo.innerHTML = signer + (dateStr ? `<br><small class="field-sidebar-signed-date">${escapeHtml(dateStr)}</small>` : '');
                inputContainer.appendChild(signedInfo);
            }
            
            fieldItem.appendChild(inputContainer);
            
            // Page info
            const pageInfo = document.createElement('div');
            pageInfo.className = 'field-item-page';
            pageInfo.textContent = `Page ${field.pageNum}`;
            fieldItem.appendChild(pageInfo);
            
            fieldsList.appendChild(fieldItem);
        });
    }

    /**
     * Update a text/date field value from sidebar
     */
    updateFieldValue(field, value) {
        if (!field.object) return;
        
        field.object._fieldValue = value;
        
        // Update visual text on canvas
        const objects = field.object.getObjects?.() || [];
        if (objects.length >= 3) {
            const valueText = objects[2];
            const isDate = field.type === 'date';
            const placeholder = isDate ? 'YYYY-MM-DD' : 'Click to fill';
            valueText.set({
                text: value || placeholder,
                fill: value ? '#000000' : '#9ca3af',
                fontStyle: value ? 'normal' : 'italic'
            });
        }
        
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (canvas) canvas.renderAll();
    }

    /**
     * Update a checkbox/radio field from sidebar
     */
    updateFieldChecked(field, checked) {
        if (!field.object) return;
        
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (!canvas) return;
        
        if (field.type === 'checkbox') {
            this.canvasManager.toggleCheckbox(field.object, canvas);
        } else if (field.type === 'radio') {
            this.canvasManager.toggleRadio(field.object, canvas);
        }
        
        // Refresh to update sidebar
        setTimeout(() => this.refreshFieldsSidebar(), 50);
    }

    /**
     * Update a dropdown field from sidebar
     */
    updateFieldDropdown(field, value) {
        if (!field.object) return;
        
        field.object._selectedOption = value;
        
        // Update visual text on canvas
        const objects = field.object.getObjects?.() || [];
        const hasNewStructure = objects.length >= 4;
        const valueText = hasNewStructure ? objects[2] : objects[1];
        
        if (valueText) {
            valueText.set({
                text: value || 'Select...',
                fill: value ? '#000000' : '#9ca3af',
                fontStyle: value ? 'normal' : 'italic'
            });
        }
        
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (canvas) canvas.renderAll();
    }

    /**
     * Append Lock or Unlock button for a form field in the sidebar (fill mode only).
     * @param {HTMLElement} container - Parent to append the button to
     * @param {Object} field - Field descriptor with locked, permanentlyLocked
     */
    _appendLockUnlock(container, field) {
        const wrap = document.createElement('div');
        wrap.className = 'field-lock-wrap';
        if (field.locked) {
            if (!field.permanentlyLocked) {
                const unlockBtn = document.createElement('button');
                unlockBtn.type = 'button';
                unlockBtn.className = 'field-sidebar-unlock-btn';
                unlockBtn.textContent = 'Unlock';
                unlockBtn.title = 'Unlock this field (only before exporting)';
                unlockBtn.addEventListener('click', () => this.unlockField(field));
                wrap.appendChild(unlockBtn);
            } else {
                const lockedLabel = document.createElement('span');
                lockedLabel.className = 'field-sidebar-locked-label';
                lockedLabel.textContent = 'Locked (from exported version)';
                lockedLabel.title = 'This field was locked in an exported version and cannot be unlocked';
                wrap.appendChild(lockedLabel);
            }
        } else {
            const lockBtn = document.createElement('button');
            lockBtn.type = 'button';
            lockBtn.className = 'field-sidebar-lock-btn';
            lockBtn.textContent = 'Lock';
            lockBtn.title = 'Lock this field so it cannot be changed after export';
            lockBtn.addEventListener('click', () => this.lockField(field));
            wrap.appendChild(lockBtn);
        }
        container.appendChild(wrap);
    }

    /**
     * Lock a form field (cannot be edited until unlocked; after export, lock is permanent for that version).
     */
    lockField(field) {
        if (!field?.object || field.type === 'signature-filled') return;
        field.object._fieldLocked = true;
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (canvas) {
            this.canvasManager._applyObjectInteractivity(field.object, true, this.canvasManager.activeTool);
            canvas.renderAll();
        }
        this.refreshFieldsSidebar();
    }

    /**
     * Unlock a form field. Only allowed if the field was not locked in an exported version (permanentlyLocked).
     */
    unlockField(field) {
        if (!field?.object || field.permanentlyLocked) return;
        field.object._fieldLocked = false;
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (canvas) {
            this.canvasManager._applyObjectInteractivity(field.object, true, this.canvasManager.activeTool);
            canvas.renderAll();
        }
        this.refreshFieldsSidebar();
    }

    /**
     * Highlight a field on the canvas and show font controls
     */
    highlightFieldOnCanvas(field) {
        if (!field.object) return;
        
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (canvas) {
            canvas.setActiveObject(field.object);
            canvas.renderAll();
            // Trigger selection change to update Field Properties sidebar
            this.canvasManager.onSelectionChanged(canvas);
        }
    }

    /**
     * Focus the sidebar input for a selected field object
     */
    focusSidebarInputForField(fieldObject) {
        const fieldsList = document.getElementById('fields-list');
        if (!fieldsList) return;
        
        // Find the field item that corresponds to this object (include signature-filled read-only row)
        const inputs = fieldsList.querySelectorAll('.field-sidebar-input, .field-sidebar-checkbox, .field-sidebar-select, .field-sidebar-sign-btn, .field-sidebar-signed-info');
        
        for (const input of inputs) {
            const fieldItem = input.closest('.field-item');
            if (!fieldItem) continue;
            
            const index = parseInt(fieldItem.dataset.fieldIndex);
            
            // Get field data to compare
            const annotations = this.canvasManager.getAllAnnotations?.() || [];
            let fieldIndex = 0;
            
            for (const pageData of annotations) {
                for (const ann of pageData.annotations || []) {
                    const isFormField = ['textfield', 'checkbox', 'radio', 'dropdown', 'date', 'signature-field'].includes(ann.type);
                    const isFilledSignature = ann.type === 'signature' && ann.data?._signatureMeta?.replacedFieldLabel;
                    if (isFormField || isFilledSignature) {
                        if (ann.object === fieldObject) {
                            // Found the matching field - focus its input (or scroll to row for signature-filled)
                            if (input.focus) input.focus();
                            input.closest('.field-item')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                            return;
                        }
                        fieldIndex++;
                    }
                }
            }
        }
    }

    /**
     * Show field configuration detail view in the sidebar
     */
    showFieldsDetail(field) {
        this._sidebarConfigField = field;
        const listWrap = document.getElementById('fields-list-wrap');
        const detailWrap = document.getElementById('fields-detail-wrap');
        const titleEl = document.getElementById('fields-sidebar-title');
        const hintEl = document.getElementById('fields-sidebar-hint');
        const detailBody = document.getElementById('fields-detail-body');
        
        if (!listWrap || !detailWrap) return;
        
        listWrap.classList.add('hidden');
        detailWrap.classList.remove('hidden');
        if (titleEl) titleEl.textContent = 'Configure Field';
        if (hintEl) hintEl.textContent = '';
        
        // Filled signature: read-only detail (field label, signer, date)
        if (field.type === 'signature-filled') {
            const signedDate = field.signedAt ? (() => {
                try {
                    const d = new Date(field.signedAt);
                    return isNaN(d.getTime()) ? field.signedAt : d.toLocaleString();
                } catch (_) { return field.signedAt || ''; }
            })() : '';
            if (detailBody) {
                detailBody.innerHTML = `
                    <div class="field-property-group">
                        <label>Field Label</label>
                        <p class="field-detail-readonly">${this.escapeHtml(field.signatureFieldLabel || 'Signature')}</p>
                    </div>
                    <div class="field-property-group">
                        <label>Signed by</label>
                        <p class="field-detail-readonly">${this.escapeHtml(field.signerName || '—')}</p>
                    </div>
                    ${signedDate ? `<div class="field-property-group"><label>Date</label><p class="field-detail-readonly">${this.escapeHtml(signedDate)}</p></div>` : ''}
                `;
            }
            return;
        }
        
        const label = field.fieldName || field.signatureFieldLabel || field.type;
        const isTextBased = ['textfield', 'date', 'dropdown'].includes(field.type);
        const fieldName = field.type === 'signature-field' ? (field.signatureFieldLabel || '') : (field.fieldName || '');
        const currentFontSize = isTextBased ? Math.round((field.object?._valueFontSize || 12) * this.canvasManager.currentScale) : 12;
        const fontFamily = field.object?._fontFamily || 'Arial';
        
        const isFillMode = this.mode === 'fill';
        const nameReadonly = isFillMode ? ' readonly' : '';
        let html = `
            <div class="field-property-group">
                <label>${field.type === 'signature-field' ? 'Field Label' : 'Field Name'}</label>
                <input type="text" id="fields-detail-name" value="${this.escapeHtml(fieldName)}"${nameReadonly} class="${isFillMode ? 'readonly' : ''}">
                ${isFillMode ? '<small class="field-properties-hint">Field name cannot be changed in Fill mode.</small>' : ''}
            </div>
        `;
        
        if (isTextBased) {
            html += `
                <div class="field-property-group">
                    <label>Font Size</label>
                    <input type="number" id="fields-detail-font-size" value="${currentFontSize}" min="8" max="36">
                </div>
                <div class="field-property-group">
                    <label>Font</label>
                    <select id="fields-detail-font-family">
                        <option value="Arial" ${fontFamily === 'Arial' ? 'selected' : ''}>Arial</option>
                        <option value="Times New Roman" ${fontFamily === 'Times New Roman' ? 'selected' : ''}>Times New Roman</option>
                        <option value="Courier New" ${fontFamily === 'Courier New' ? 'selected' : ''}>Courier New</option>
                        <option value="Georgia" ${fontFamily === 'Georgia' ? 'selected' : ''}>Georgia</option>
                    </select>
                </div>
            `;
        }
        
        if (field.type === 'dropdown' && !isFillMode) {
            const options = field.object?._options || [];
            html += `
                <div class="field-property-group">
                    <label>Options</label>
                    <p class="field-properties-hint">Add options that users can select.</p>
                    <div class="options-list" id="fields-detail-options-list"></div>
                    <div class="add-option-row" style="margin-top: 8px;">
                        <input type="text" id="fields-detail-new-option" placeholder="New option..." class="new-option-input">
                        <button type="button" id="fields-detail-add-option" class="btn btn-sm btn-primary" title="Add option">+</button>
                    </div>
                </div>
            `;
        } else if (field.type === 'dropdown' && isFillMode) {
            const options = field.object?._options || [];
            html += `
                <div class="field-property-group">
                    <label>Options</label>
                    <p class="field-properties-hint">Options cannot be changed in Fill mode.</p>
                    <div class="options-list" id="fields-detail-options-list"></div>
                </div>
            `;
        }
        
        document.getElementById('fields-detail-body').innerHTML = html;
        
        // Bind name (only when not read-only)
        const nameInput = document.getElementById('fields-detail-name');
        if (nameInput && !isFillMode) {
            nameInput.addEventListener('input', () => {
                this._applySidebarConfigToField('name', nameInput.value.trim());
            });
        }
        
        // Bind font controls
        if (isTextBased) {
            const fontSizeInput = document.getElementById('fields-detail-font-size');
            if (fontSizeInput) {
                fontSizeInput.addEventListener('change', () => {
                    const val = parseInt(fontSizeInput.value);
                    if (val >= 8 && val <= 36) this._applySidebarConfigToField('fontSize', val);
                });
            }
            const fontFamilyInput = document.getElementById('fields-detail-font-family');
            if (fontFamilyInput) {
                fontFamilyInput.addEventListener('change', () => {
                    this._applySidebarConfigToField('fontFamily', fontFamilyInput.value);
                });
            }
        }
        
        // Bind dropdown options (only in Edit mode)
        if (field.type === 'dropdown') {
            this._refreshFieldsDetailOptions(field);
            if (!isFillMode) {
            document.getElementById('fields-detail-add-option')?.addEventListener('click', () => {
                const input = document.getElementById('fields-detail-new-option');
                const val = input?.value?.trim();
                if (val && field.object) {
                    field.object._options = field.object._options || [];
                    field.object._options.push(val);
                    input.value = '';
                    this._refreshFieldsDetailOptions(field);
                    this.canvasManager.canvases.get(field.pageId)?.renderAll();
                }
            });
            document.getElementById('fields-detail-new-option')?.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    document.getElementById('fields-detail-add-option')?.click();
                }
            });
            }
        }
    }

    _applySidebarConfigToField(prop, value) {
        const field = this._sidebarConfigField;
        if (!field?.object) return;
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (!canvas) return;
        
        if (prop === 'name') {
            if (field.type === 'signature-field') {
                field.object._signatureFieldLabel = value;
                const textObj = field.object.getObjects?.()?.[1];
                if (textObj?.type === 'text') {
                    textObj.set('text', `${value || 'Signature'}\n(Double-click to sign)`);
                }
            } else {
                field.object._fieldName = value;
                if (field.type === 'radio') field.object._radioGroup = value;
                const objects = field.object.getObjects?.() || [];
                const minObjects = field.type === 'dropdown' ? 4 : 3;
                if (objects.length >= minObjects) {
                    const labelObj = objects[1];
                    if (labelObj?.type === 'text') {
                        labelObj.set('text', value || (field.type === 'date' ? 'Date Field' : field.type === 'dropdown' ? 'Dropdown' : 'Text Field'));
                    }
                }
            }
        } else if (prop === 'fontSize') {
            field.object._valueFontSize = value / this.canvasManager.currentScale;
            const objects = field.object.getObjects?.() || [];
            const valueText = objects[2];
            if (valueText && (valueText.type === 'text' || valueText.type === 'i-text')) {
                valueText.set('fontSize', value / this.canvasManager.currentScale);
            }
        } else if (prop === 'fontFamily') {
            field.object._fontFamily = value;
            const objects = field.object.getObjects?.() || [];
            const valueText = objects[2];
            if (valueText && (valueText.type === 'text' || valueText.type === 'i-text')) {
                valueText.set('fontFamily', value);
            }
        }
        canvas.renderAll();
        this.refreshFieldsSidebar();
    }

    _refreshFieldsDetailOptions(field) {
        const listEl = document.getElementById('fields-detail-options-list');
        if (!listEl || field.type !== 'dropdown') return;
        const options = field.object?._options || [];
        const isFillMode = this.mode === 'fill';
        listEl.innerHTML = options.map((opt, i) => `
            <div class="option-item">
                ${!isFillMode ? '<span class="option-item-drag">⋮⋮</span>' : ''}
                <span>${this.escapeHtml(opt)}</span>
                ${!isFillMode ? `<button type="button" class="option-item-remove" data-index="${i}">×</button>` : ''}
            </div>
        `).join('');
        if (!isFillMode) {
            listEl.querySelectorAll('.option-item-remove').forEach(btn => {
                btn.addEventListener('click', () => {
                    const i = parseInt(btn.dataset.index);
                    field.object._options.splice(i, 1);
                    this._refreshFieldsDetailOptions(field);
                    this.canvasManager.canvases.get(field.pageId)?.renderAll();
                });
            });
        }
    }

    /**
     * Hide field config detail and return to list view
     */
    hideFieldsDetail() {
        this._sidebarConfigField = null;
        const listWrap = document.getElementById('fields-list-wrap');
        const detailWrap = document.getElementById('fields-detail-wrap');
        const titleEl = document.getElementById('fields-sidebar-title');
        const hintEl = document.getElementById('fields-sidebar-hint');
        
        if (listWrap) listWrap.classList.remove('hidden');
        if (detailWrap) detailWrap.classList.add('hidden');
        if (titleEl) titleEl.textContent = 'Form Fields';
        if (hintEl) hintEl.textContent = 'Use the inputs above to fill fields. Tab to move between fields. Click a field on the document to select it.';
        
        this.refreshFieldsSidebar();
    }

    /**
     * Navigate to a field in the document
     */
    navigateToField(field) {
        // Go to the field's page
        this.goToPage(field.pageNum);
        
        // Highlight the field (optional - could add visual feedback)
        // For now, just ensure the page is visible
    }

    /**
     * Edit a field from the sidebar
     */
    editFieldFromSidebar(field) {
        // Navigate to the field first
        this.navigateToField(field);
        
        const canvas = this.canvasManager.canvases.get(field.pageId);
        if (!canvas) return;
        
        // Filled signature: just navigate and select the signature on canvas (no editing)
        if (field.type === 'signature-filled' && field.object) {
            canvas.setActiveObject(field.object);
            canvas.renderAll();
            this.canvasManager.onSelectionChanged(canvas);
            return;
        }
        
        // Find the actual field object on the canvas and trigger its edit handler
        const objects = canvas.getObjects();
        const fieldObj = objects.find(obj => {
            return obj._annotationType === field.type &&
                   (obj._fieldName === field.fieldName || obj._signatureFieldLabel === field.signatureFieldLabel);
        });
        
        if (fieldObj) {
            // Trigger the appropriate interaction based on field type
            if (field.type === 'textfield' || field.type === 'date') {
                this.canvasManager.editTextField(fieldObj, canvas);
            } else if (field.type === 'checkbox') {
                this.canvasManager.toggleCheckbox(fieldObj, canvas);
            } else if (field.type === 'radio') {
                this.canvasManager.toggleRadio(fieldObj, canvas);
            } else if (field.type === 'dropdown') {
                this.canvasManager.selectDropdownValue(fieldObj, canvas);
            } else if (field.type === 'signature-field') {
                // Open signature modal for this specific field
                this.showSignatureModal({ targetField: field });
            }
            
            // Refresh sidebar after edit
            setTimeout(() => this.refreshFieldsSidebar(), 100);
        }
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
            const sigOpen = !this.signatureModal.classList.contains('hidden');

            // Ctrl/Cmd + Z = Undo
            if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
                e.preventDefault();
                if (sigOpen && this.signaturePad?.mode === 'draw') this.signaturePad.undo();
                else this.canvasManager.undo();
                this.updateHistoryButtons();
            }

            // Ctrl/Cmd + Shift + Z or Ctrl/Cmd + Y = Redo
            if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
                e.preventDefault();
                if (sigOpen && this.signaturePad?.mode === 'draw') this.signaturePad.redo();
                else this.canvasManager.redo();
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

            // Build initial view page model from the main document
            this.viewPages = [];
            const mainDocId = this.pdfHandler.mainDocId || 'main';
            for (let p = 1; p <= this.pdfHandler.totalPages; p++) {
                this.viewPages.push({
                    id: `${mainDocId}:${p}`,
                    docId: mainDocId,
                    sourcePageNum: p,
                    rotation: 0
                });
            }
            this.pdfHandler.currentPage = 1;
            this.pdfHandler.totalPages = this.viewPages.length;

            // Render all pages
            this.showLoading('Rendering pages...');
            await this.pdfHandler.renderViewPages(
                this.pdfPages,
                this.viewPages,
                (container, width, height, pageNum) => {
                    return this.canvasManager.createCanvas(container, width, height, pageNum);
                },
                this.currentScale
            );

            this.renderPagesSidebar();
            this.applyPageRotationUI();

            // Load form fields from PDF into canvas so they appear in the sidebar and can be filled/signed
            try {
                const bytes = this.pdfHandler.getOriginalBytes(mainDocId);
                if (bytes) {
                    const descriptors = await loadFormFieldsFromPdf(bytes);
                    for (const desc of descriptors) {
                        const vp = this.viewPages[desc.pageIndex];
                        if (!vp) continue;
                        const canvas = this.canvasManager.canvases.get(vp.id);
                        if (!canvas) continue;
                        this.canvasManager.addFormFieldFromPdfDescriptor(canvas, desc, this.currentScale);
                    }
                    if (descriptors.length > 0) {
                        this.canvasManager.canvases.forEach((c) => c.renderAll());
                        this.canvasManager.canvases.forEach((c, pageId) => this.canvasManager.saveState(pageId));
                        if (this.mode === 'fill') this.refreshFieldsSidebar();
                    }
                }
            } catch (e) {
                console.warn('Could not load form fields from PDF:', e);
            }

            // Detect our signing metadata (Keywords or Producer) so we can show multi-signer flow
            this.signingFlowMeta = null;
            this.lockedFromFile = { signatures: new Set(), formFields: new Set() };
            try {
                const metadata = await this.pdfHandler.getMetadata(mainDocId);
                if (hasOurSigningMetadata(metadata)) {
                    const info = metadata?.info || {};
                    const parsed = parseSigningMetadata(info.Keywords);
                    if (parsed) {
                        this.signingFlowMeta = parsed;
                        (parsed.lockedSignatureFields || []).forEach((l) => this.lockedFromFile.signatures.add(l));
                        (parsed.lockedFormFields || []).forEach((l) => this.lockedFromFile.formFields.add(l));
                    } else {
                        // Producer says our app but no Keywords yet (e.g. first save)
                        this.signingFlowMeta = { signers: [], expectedSigners: [] };
                    }
                }
            } catch (e) {
                console.warn('Could not read PDF metadata for signing flow:', e);
            }

            // Auto-detect mode: if signing metadata present, default to Fill mode
            if (this.signingFlowMeta) {
                this.switchMode('fill');
            } else {
                this.switchMode('edit');
            }

            // Update UI
            this.welcomeScreen.classList.add('hidden');
            this.pdfContainer.classList.remove('hidden');
            this.btnSave.disabled = false;
            this.btnSend.disabled = false;
            this.updatePageNavigation();
            this.updateZoomDisplay();
            this.updateHistoryButtons();
            this.updateSigningFlowBanner();

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

        // Handle signature field tool (for placing empty signature boxes when editing)
        // Place directly like other form fields; edit name in Field Properties sidebar
        if (tool === 'signature-field') {
            this.canvasManager.setTool(tool);
            this.toolButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            this.applyToolCursor(tool);
            this.showToolHint('Click on the document where you want to place the signature field');
            return;
        }

        // Handle sign tool (for actually signing the document)
        if (tool === 'sign' || tool === 'signature') {
            this.showSignatureModal();
            return;
        }

        // Image tool picks a file first
        if (tool === 'image') {
            this._imageInsertInput?.click();
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
     * Switch between Edit and Fill modes
     */
    switchMode(newMode) {
        if (this.mode === newMode) return;

        // Warn if switching from Fill to Edit with active signing flow; clear signing flow on switch
        if (newMode === 'edit' && this.mode === 'fill') {
            if (this.signingFlowMeta) {
                const ok = confirm(
                    'Switching to Edit Mode will clear the signing flow. You can make structural changes and set up a new signing flow in Fill mode later.\n\n' +
                    'Continue to Edit Mode?'
                );
                if (!ok) return;
            }
            // Clear signing flow so they can start a new one
            this.signingFlowMeta = null;
            this.lockedFromFile = { signatures: new Set(), formFields: new Set() };
        }

        this.mode = newMode;
        
        // Update mode toggle buttons
        document.getElementById('mode-edit')?.classList.toggle('active', newMode === 'edit');
        document.getElementById('mode-fill')?.classList.toggle('active', newMode === 'fill');
        
        // Update canvas manager fill mode
        this.canvasManager.setFillMode(newMode === 'fill');
        
        // Update toolbar tools visibility
        this.updateToolsForMode();
        
        // Show/hide sidebars and hint
        const fieldsSidebar = document.getElementById('fields-sidebar');
        if (newMode === 'fill') {
            this.hideFieldPropertiesSidebar();
            fieldsSidebar?.classList.remove('hidden');
            this.refreshFieldsSidebar();
            // Show helpful hint for Fill mode
            this.showToolHint('Use the sidebar to fill in fields. Click a field on the document to select it.');
        } else {
            this.hideToolHint();
            fieldsSidebar?.classList.add('hidden');
        }
        
        // Switch to select tool
        this.selectToolByName('select');
    }

    /**
     * Show/hide tools based on current mode
     */
    updateToolsForMode() {
        const editOnlyTools = [
            'text', 'whiteout', 'draw', 'eraser',
            'signature-field', 'textfield', 'checkbox', 'radio', 'dropdown', 'date',
            'highlight', 'underline', 'strike', 'rect', 'ellipse', 'arrow', 'note', 'stamp', 'image'
        ];
        
        const fillModeTools = ['select', 'sign'];

        if (this.mode === 'fill') {
            // In Fill mode: hide edit-only tools
            editOnlyTools.forEach(tool => {
                const btn = document.querySelector(`[data-tool="${tool}"]`);
                if (btn) btn.style.display = 'none';
            });
            // Show fill mode tools
            fillModeTools.forEach(tool => {
                const btn = document.querySelector(`[data-tool="${tool}"]`);
                if (btn) btn.style.display = '';
            });
        } else {
            // In Edit mode: show all tools
            [...editOnlyTools, ...fillModeTools].forEach(tool => {
                const btn = document.querySelector(`[data-tool="${tool}"]`);
                if (btn) btn.style.display = '';
            });
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
                    <div class="tool-option">
                        <label>Style:</label>
                        <button type="button" class="pages-btn" id="text-bold-btn" aria-pressed="${this.canvasManager.settings.fontWeight === 'bold'}">B</button>
                        <button type="button" class="pages-btn" id="text-italic-btn" aria-pressed="${this.canvasManager.settings.fontStyle === 'italic'}">I</button>
                        <select id="text-align">
                            <option value="left" ${this.canvasManager.settings.textAlign === 'left' ? 'selected' : ''}>Left</option>
                            <option value="center" ${this.canvasManager.settings.textAlign === 'center' ? 'selected' : ''}>Center</option>
                            <option value="right" ${this.canvasManager.settings.textAlign === 'right' ? 'selected' : ''}>Right</option>
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

            case 'highlight':
                this.toolOptions.innerHTML = `
                    <div class="tool-option">
                        <label>Color:</label>
                        <input type="color" id="highlight-color" value="${this.canvasManager.settings.highlightColor}">
                    </div>
                    <div class="tool-option">
                        <label>Opacity:</label>
                        <input type="range" id="highlight-opacity" min="0.1" max="1" step="0.05" value="${this.canvasManager.settings.highlightOpacity}">
                        <span id="highlight-opacity-value">${Math.round(this.canvasManager.settings.highlightOpacity * 100)}%</span>
                    </div>
                `;
                this.bindHighlightOptions();
                break;

            case 'rect':
            case 'ellipse':
            case 'arrow':
            case 'underline':
            case 'strike':
                this.toolOptions.innerHTML = `
                    <div class="tool-option">
                        <label>Stroke:</label>
                        <input type="color" id="shape-stroke-color" value="${this.canvasManager.settings.strokeColor}">
                    </div>
                    <div class="tool-option">
                        <label>Width:</label>
                        <input type="range" id="shape-stroke-width" min="1" max="20" value="${this.canvasManager.settings.strokeWidth}">
                        <span id="shape-stroke-width-value">${this.canvasManager.settings.strokeWidth}px</span>
                    </div>
                    <div class="tool-option">
                        <label>Fill:</label>
                        <select id="shape-fill">
                            <option value="transparent" ${this.canvasManager.settings.shapeFill === 'transparent' ? 'selected' : ''}>None</option>
                            <option value="#ffffff">White</option>
                            <option value="#000000">Black</option>
                            <option value="#fff59d">Yellow</option>
                            <option value="#dcfce7">Green</option>
                            <option value="#dbeafe">Blue</option>
                        </select>
                    </div>
                `;
                this.bindShapeOptions();
                break;

            case 'stamp':
                this.toolOptions.innerHTML = `
                    <div class="tool-option">
                        <label>Stamp:</label>
                        <select id="stamp-text">
                            ${['APPROVED','DRAFT','REJECTED','CONFIDENTIAL'].map((t) => `<option value="${t}" ${t === (this.canvasManager.settings.stampText || 'APPROVED') ? 'selected' : ''}>${t}</option>`).join('')}
                        </select>
                    </div>
                    <div class="tool-option">
                        <small>Click on the page to place.</small>
                    </div>
                `;
                this.bindStampOptions();
                break;

            case 'note':
                this.toolOptions.innerHTML = `<div class="tool-option"><small>Click to add a note (double-click to edit).</small></div>`;
                break;

            case 'image':
                this.toolOptions.innerHTML = `<div class="tool-option"><small>Click to place the selected image.</small></div>`;
                break;

            case 'eraser':
                this.toolOptions.innerHTML = `<div class="tool-option"><small>Click an annotation to remove it.</small></div>`;
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

    bindHighlightOptions() {
        const colorInput = document.getElementById('highlight-color');
        const opInput = document.getElementById('highlight-opacity');
        const opValue = document.getElementById('highlight-opacity-value');
        colorInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ highlightColor: e.target.value });
        });
        opInput?.addEventListener('input', (e) => {
            const v = parseFloat(e.target.value);
            this.canvasManager.updateSettings({ highlightOpacity: v });
            if (opValue) opValue.textContent = `${Math.round(v * 100)}%`;
        });
    }

    bindShapeOptions() {
        const colorInput = document.getElementById('shape-stroke-color');
        const widthInput = document.getElementById('shape-stroke-width');
        const widthValue = document.getElementById('shape-stroke-width-value');
        const fillSel = document.getElementById('shape-fill');
        colorInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ strokeColor: e.target.value });
        });
        widthInput?.addEventListener('input', (e) => {
            const w = parseInt(e.target.value);
            this.canvasManager.updateSettings({ strokeWidth: w });
            if (widthValue) widthValue.textContent = `${w}px`;
        });
        fillSel?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ shapeFill: e.target.value });
        });
    }

    bindStampOptions() {
        const sel = document.getElementById('stamp-text');
        sel?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ stampText: e.target.value });
        });
    }

    /**
     * Bind text tool options
     */
    bindTextOptions() {
        const colorInput = document.getElementById('text-color');
        const sizeInput = document.getElementById('text-size');
        const fontInput = document.getElementById('text-font');
        const boldBtn = document.getElementById('text-bold-btn');
        const italicBtn = document.getElementById('text-italic-btn');
        const alignSel = document.getElementById('text-align');

        colorInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ textColor: e.target.value });
        });

        sizeInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ fontSize: parseInt(e.target.value) });
        });

        fontInput?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ fontFamily: e.target.value });
        });

        boldBtn?.addEventListener('click', () => {
            const next = this.canvasManager.settings.fontWeight === 'bold' ? 'normal' : 'bold';
            this.canvasManager.updateSettings({ fontWeight: next });
            boldBtn.setAttribute('aria-pressed', String(next === 'bold'));
        });

        italicBtn?.addEventListener('click', () => {
            const next = this.canvasManager.settings.fontStyle === 'italic' ? 'normal' : 'italic';
            this.canvasManager.updateSettings({ fontStyle: next });
            italicBtn.setAttribute('aria-pressed', String(next === 'italic'));
        });

        alignSel?.addEventListener('change', (e) => {
            this.canvasManager.updateSettings({ textAlign: e.target.value });
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
        if (pageNum < 1 || pageNum > (this.viewPages?.length || this.pdfHandler.totalPages)) return;

        this.pdfHandler.currentPage = pageNum;
        this.pageInput.value = pageNum;

        // Set active canvas to this page so undo/redo apply to the page we're viewing
        const vp = this.viewPages?.[pageNum - 1];
        if (vp) this.canvasManager.setActivePage(vp.id);

        // Scroll to page
        const pageWrapper = document.querySelector(`.page-wrapper[data-page="${pageNum}"]`);
        if (pageWrapper) {
            pageWrapper.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }

        this.updatePageNavigation();
        this.updateHistoryButtons();
    }

    /**
     * Update page navigation UI
     */
    updatePageNavigation() {
        const current = this.pdfHandler.currentPage;
        const total = this.viewPages?.length || this.pdfHandler.totalPages;

        this.pageInput.value = current;
        this.pageInput.max = total;
        this.totalPagesSpan.textContent = total;

        this.btnPrevPage.disabled = current <= 1;
        this.btnNextPage.disabled = current >= total;
    }

    renderPagesSidebar() {
        if (!this.pagesList) return;
        this.pagesList.innerHTML = '';

        // Build items in view order; thumbnails come from already-rendered pdf canvases
        const pagesInfo = this.pdfHandler.pages || [];
        const infoById = new Map(pagesInfo.map((p) => [p.viewPageId, p]));

        const makeThumb = (vp, idx) => {
            const pageId = vp.id;
            const pageInfo = infoById.get(pageId);

            const item = document.createElement('div');
            item.className = 'page-thumb';
            item.draggable = true;
            item.dataset.pageId = pageId;

            item.innerHTML = `
                <input type="checkbox" class="page-thumb-cb" aria-label="Select page ${idx + 1}">
                <div class="page-thumb-body">
                    <canvas class="page-thumb-canvas" width="160" height="220"></canvas>
                    <div class="page-thumb-meta">
                        <span>Page ${idx + 1}</span>
                        <span>${(vp.docId === (this.pdfHandler.mainDocId || 'main')) ? 'Doc' : 'Appended'}${vp.rotation ? ` · ${vp.rotation}°` : ''}</span>
                    </div>
                </div>
            `;

            const cb = item.querySelector('.page-thumb-cb');
            const thumbCanvas = item.querySelector('.page-thumb-canvas');
            cb.addEventListener('change', () => this.onPageSelectionChanged());

            const handleDrop = (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
                const fromId = ev.dataTransfer.getData('text/plain') || ev.dataTransfer.getData('application/x-page-id');
                const toId = pageId;
                if (!fromId || fromId === toId) return;
                this.reorderPagesById(fromId, toId);
            };
            const handleDragOver = (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
                if (ev.dataTransfer.types.includes('text/plain')) ev.dataTransfer.dropEffect = 'move';
            };

            // Click body to go to page
            item.addEventListener('click', (ev) => {
                if (ev.target === cb) return;
                this.goToPage(idx + 1);
            });

            // Drag/drop reorder
            item.addEventListener('dragstart', (ev) => {
                item.classList.add('dragging');
                ev.dataTransfer.effectAllowed = 'move';
                ev.dataTransfer.setData('text/plain', pageId);
                ev.dataTransfer.setData('application/x-page-id', pageId);
            });
            item.addEventListener('dragend', () => item.classList.remove('dragging'));
            item.addEventListener('dragenter', (ev) => {
                ev.preventDefault();
                if (ev.dataTransfer.types.includes('text/plain')) ev.dataTransfer.dropEffect = 'move';
            });
            item.addEventListener('dragover', handleDragOver);
            item.addEventListener('drop', handleDrop);
            cb.addEventListener('dragover', handleDragOver);
            cb.addEventListener('drop', handleDrop);

            // Draw thumbnail from rendered page canvas (if available)
            if (pageInfo?.pdfCanvas && thumbCanvas) {
                const ctx = thumbCanvas.getContext('2d');
                const src = pageInfo.pdfCanvas;
                const rot = vp.rotation || 0;
                const sw = src.width;
                const sh = src.height;
                const tw = thumbCanvas.width;
                const th = thumbCanvas.height;
                const scale = Math.min(tw / sw, th / sh);
                const w = Math.floor(sw * scale);
                const h = Math.floor(sh * scale);
                const cx = tw / 2;
                const cy = th / 2;
                const dx = (tw - w) / 2;
                const dy = (th - h) / 2;
                ctx.clearRect(0, 0, tw, th);
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, tw, th);
                ctx.save();
                ctx.translate(cx, cy);
                if (rot) ctx.rotate((rot * Math.PI) / 180);
                ctx.translate(-cx, -cy);
                ctx.drawImage(src, dx, dy, w, h);
                ctx.restore();
            }

            return item;
        };

        this.viewPages.forEach((vp, idx) => {
            this.pagesList.appendChild(makeThumb(vp, idx));
        });

        if (this.pagesRotateBtn) this.pagesRotateBtn.disabled = this.viewPages.length === 0;
        this.onPageSelectionChanged();
    }

    getSelectedPageIds() {
        const ids = [];
        this.pagesList?.querySelectorAll('.page-thumb').forEach((el) => {
            const cb = el.querySelector('.page-thumb-cb');
            if (cb?.checked) ids.push(el.dataset.pageId);
        });
        return ids;
    }

    onPageSelectionChanged() {
        const selected = this.getSelectedPageIds();
        if (this.pagesDeleteBtn) this.pagesDeleteBtn.disabled = selected.length === 0;
        if (this.pagesExtractBtn) this.pagesExtractBtn.disabled = selected.length === 0;
        this.pagesList?.querySelectorAll('.page-thumb').forEach((el) => {
            const cb = el.querySelector('.page-thumb-cb');
            el.classList.toggle('selected', !!cb?.checked);
        });
    }

    reorderPagesById(fromId, toId) {
        const fromIdx = this.viewPages.findIndex((p) => p.id === fromId);
        const toIdx = this.viewPages.findIndex((p) => p.id === toId);
        if (fromIdx < 0 || toIdx < 0 || fromIdx === toIdx) return;

        const [moved] = this.viewPages.splice(fromIdx, 1);
        this.viewPages.splice(toIdx, 0, moved);

        // Reorder DOM page wrappers to match new view order (use pdfHandler refs so we move the right nodes)
        const wrapperById = new Map();
        for (const p of this.pdfHandler.pages || []) {
            if (p.wrapper && p.viewPageId) wrapperById.set(p.viewPageId, p.wrapper);
        }
        this.viewPages.forEach((vp, i) => {
            const w = wrapperById.get(vp.id);
            if (!w) return;
            w.dataset.page = String(i + 1);
            this.pdfPages.appendChild(w);
        });

        // Sync pdfHandler.pages ordering & indices
        const infoById = new Map((this.pdfHandler.pages || []).map((p) => [p.viewPageId, p]));
        this.pdfHandler.pages = this.viewPages
            .map((vp, i) => {
                const info = infoById.get(vp.id);
                if (!info) return null;
                info.viewIndex = i + 1;
                return info;
            })
            .filter(Boolean);

        this.pdfHandler.totalPages = this.viewPages.length;
        this.updatePageNavigation();
        this.applyPageRotationUI();
        this.renderPagesSidebar();
    }

    async appendPdfFile(file) {
        if (!file || file.type !== 'application/pdf') return;
        const bytes = await file.arrayBuffer();
        const { docId, pageCount } = await this.pdfHandler.addDocument(bytes, file.name);

        // Add view pages for appended doc
        const newPages = [];
        for (let p = 1; p <= pageCount; p++) {
            newPages.push({ id: `${docId}:${p}`, docId, sourcePageNum: p, rotation: 0 });
        }

        this.showLoading('Appending pages...');

        try {
            for (const vp of newPages) {
                this.viewPages.push(vp);
                await this.pdfHandler.renderOneViewPage(
                    this.pdfPages,
                    vp,
                    (container, width, height, pageId) => this.canvasManager.createCanvas(container, width, height, pageId),
                    this.currentScale,
                    this.viewPages.length
                );
            }

            this.pdfHandler.totalPages = this.viewPages.length;
            this.updatePageNavigation();
            this.renderPagesSidebar();
        } finally {
            this.hideLoading();
        }
    }

    deleteSelectedPages() {
        const ids = this.getSelectedPageIds();
        if (ids.length === 0) return;
        if (!confirm(`Delete ${ids.length} page(s)? This cannot be undone.`)) return;

        const toDelete = new Set(ids);
        this.viewPages = this.viewPages.filter((p) => !toDelete.has(p.id));
        this.pdfHandler.pages = (this.pdfHandler.pages || []).filter((p) => !toDelete.has(p.viewPageId));

        // Remove wrappers and canvases
        ids.forEach((id) => {
            const w = this.pdfPages.querySelector(`.page-wrapper[data-page-id="${CSS.escape(id)}"]`);
            if (w) w.remove();
            this.canvasManager.removePage(id);
        });

        // Re-number wrappers
        this.viewPages.forEach((vp, i) => {
            const w = this.pdfPages.querySelector(`.page-wrapper[data-page-id="${CSS.escape(vp.id)}"]`);
            if (w) w.dataset.page = String(i + 1);
        });

        // Clamp current page
        const total = this.viewPages.length;
        this.pdfHandler.totalPages = total;
        this.pdfHandler.currentPage = Math.min(this.pdfHandler.currentPage, Math.max(1, total));

        this.updatePageNavigation();
        this.renderPagesSidebar();
        this.updateHistoryButtons();
    }

    rotateSelectedPages() {
        let ids = this.getSelectedPageIds();
        if (ids.length === 0 && this.viewPages.length > 0) {
            const cur = this.pdfHandler.currentPage;
            const vp = this.viewPages[cur - 1];
            if (vp) ids = [vp.id];
        }
        if (ids.length === 0) return;
        const idSet = new Set(ids);
        for (const vp of this.viewPages) {
            if (!idSet.has(vp.id)) continue;
            vp.rotation = ((vp.rotation || 0) + 90) % 360;
        }
        for (const p of this.pdfHandler.pages || []) {
            const v = this.viewPages.find((vp) => vp.id === p.viewPageId);
            if (v) p.rotation = v.rotation;
        }
        this.applyPageRotationUI();
        this.renderPagesSidebar();
    }

    /**
     * Update main-view wrappers and inner divs to reflect viewPages[].rotation (CSS transform + size).
     */
    applyPageRotationUI() {
        const pages = this.pdfHandler.pages || [];
        for (const p of pages) {
            const vp = this.viewPages.find((v) => v.id === p.viewPageId);
            const r = (vp?.rotation ?? p.rotation ?? 0) % 360;
            const wrapper = p.wrapper;
            const inner = p.inner;
            const d = p.dimensions;
            if (!wrapper || !inner || !d) continue;
            const w = d.width;
            const h = d.height;
            if (r === 90 || r === 270) {
                wrapper.style.width = h + 'px';
                wrapper.style.height = w + 'px';
            } else {
                wrapper.style.width = w + 'px';
                wrapper.style.height = h + 'px';
            }
            inner.style.transform = r ? `translate(-50%, -50%) rotate(${r}deg)` : 'translate(-50%, -50%)';
        }
    }

    async extractSelectedPages() {
        const ids = this.getSelectedPageIds();
        if (ids.length === 0) return;
        const selectedSet = new Set(ids);
        const subset = this.viewPages.filter((p) => selectedSet.has(p.id));

        const annotationsArr = this.canvasManager.getAllAnnotations();
        const annotationsByPageId = new Map();
        for (const p of annotationsArr) {
            if (selectedSet.has(p.pageId)) annotationsByPageId.set(p.pageId, p.annotations);
        }

        this.showLoading('Extracting pages...');
        try {
            const bytes = await this.exporter.exportPDF({
                docBytesById: this.pdfHandler.getAllOriginalBytes(),
                viewPages: subset,
                annotationsByPageId,
                scale: this.currentScale
            });
            const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim() || 'document';
            this.exporter.downloadPDF(bytes, `${baseName}-extracted.pdf`);
        } finally {
            this.hideLoading();
        }
    }

    parsePageRanges(input, maxPage) {
        const out = [];
        const parts = String(input || '')
            .split(',')
            .map((s) => s.trim())
            .filter(Boolean);
        for (const part of parts) {
            const m = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/);
            if (!m) continue;
            const a = parseInt(m[1], 10);
            const b = m[2] ? parseInt(m[2], 10) : a;
            const start = Math.max(1, Math.min(a, b));
            const end = Math.min(maxPage, Math.max(a, b));
            if (start <= end) out.push([start, end]);
        }
        return out;
    }

    async splitPdfPrompt() {
        const total = this.viewPages.length;
        if (total === 0) return;
        const input = prompt(
            `Split into multiple PDFs by ranges.\n\nExamples:\n- 1-3,4-6\n- 1-2,5\n\nTotal pages: ${total}\nEnter ranges:`,
            '1-1'
        );
        if (!input) return;
        const ranges = this.parsePageRanges(input, total);
        if (ranges.length === 0) {
            alert('No valid ranges.');
            return;
        }

        this.showLoading('Splitting PDF...');
        try {
            const allAnn = this.canvasManager.getAllAnnotations();
            const annByIdAll = new Map(allAnn.map((p) => [p.pageId, p.annotations]));
            const docBytesById = this.pdfHandler.getAllOriginalBytes();
            const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim() || 'document';

            let partNum = 1;
            for (const [start, end] of ranges) {
                const subset = this.viewPages.slice(start - 1, end);
                const annSubset = new Map();
                for (const vp of subset) {
                    annSubset.set(vp.id, annByIdAll.get(vp.id) || []);
                }
                const bytes = await this.exporter.exportPDF({
                    docBytesById,
                    viewPages: subset,
                    annotationsByPageId: annSubset,
                    scale: this.currentScale
                });
                this.exporter.downloadPDF(bytes, `${baseName}-part-${partNum}-${start}-${end}.pdf`);
                partNum += 1;
            }
        } finally {
            this.hideLoading();
        }
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

        this.applyPageRotationUI();
        this.updateZoomDisplay();
        this.hideLoading();
    }

    /**
     * Fit PDF to width
     */
    async fitWidth() {
        const containerWidth = this.pdfScrollArea.clientWidth - 48; // Account for padding
        const first = this.viewPages?.[0];
        if (!first) return;
        const page = await this.pdfHandler.getPage(first.docId, first.sourcePageNum);
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
            this.updateHistoryButtons();
        });

        // Consent + identity
        [intentCheck, consentCheck].forEach((el) => el?.addEventListener('change', updateApplyState));
        nameInput?.addEventListener('input', updateApplyState);
        emailInput?.addEventListener('input', updateApplyState);

        // Tab switching
        const drawArea = document.getElementById('sig-draw-area');
        const typeArea = document.getElementById('sig-type-area');
        const imageArea = document.getElementById('sig-image-area');
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');

                const tabName = tab.dataset.tab;
                drawArea?.classList.toggle('active', tabName === 'draw');
                typeArea?.classList.toggle('active', tabName === 'type');
                imageArea?.classList.toggle('active', tabName === 'image');

                this._selectedSavedSig = null;
                if (tabName !== 'image') this._pendingSignatureImage = null;
                this._clearImagePreview();
                if (tabName === 'draw' || tabName === 'type') this.signaturePad.setMode(tabName);
                this._refreshSignatureModalExtras();
                updateApplyState();
                this.updateHistoryButtons();
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
            const onSigChanged = () => {
                updateApplyState();
                this.updateHistoryButtons();
            };
            sigCanvas.addEventListener('mouseup', onSigChanged);
            sigCanvas.addEventListener('touchend', onSigChanged);
        }

        // Image tab: file input, preview, clear
        const imageInput = document.getElementById('sig-image-input');
        const imagePreviewWrap = document.getElementById('sig-image-preview-wrap');
        const imagePreview = document.getElementById('sig-image-preview');
        const imageClear = document.getElementById('sig-image-clear');
        imageInput?.addEventListener('change', (e) => {
            const f = e.target.files?.[0];
            if (!f) return;
            const r = new FileReader();
            r.onload = () => {
                this._pendingSignatureImage = r.result;
                if (imagePreview) imagePreview.src = this._pendingSignatureImage;
                imagePreviewWrap?.classList.remove('hidden');
                updateApplyState();
                this._refreshSignatureModalExtras();
            };
            r.readAsDataURL(f);
            imageInput.value = '';
        });
        imageClear?.addEventListener('click', () => {
            this._pendingSignatureImage = null;
            this._clearImagePreview();
            updateApplyState();
            this._refreshSignatureModalExtras();
        });

        // Save to My Signatures
        const saveName = document.getElementById('sig-save-name');
        const saveBtn = document.getElementById('sig-save-btn');
        saveBtn?.addEventListener('click', async () => {
            const name = (saveName?.value || '').trim();
            if (!name) { alert('Enter a name for this signature.'); return; }
            const dataUrl = this._signatureDataUrlForSave();
            if (!dataUrl) return;
            const type = this._activeSignatureTab() === 'image' ? 'image' : (this.signaturePad.mode === 'type' ? 'type' : 'draw');
            try {
                await secureStorage.addSignature({ name, dataUrl, type });
                if (saveName) saveName.value = '';
                this._refreshSignatureModalExtras();
            } catch (e) {
                alert('Failed to save: ' + (e instanceof Error ? e.message : String(e)));
            }
        });

        // Apply signature
        applyBtn.addEventListener('click', () => {
            if (!this.validateAndApplySignature()) return;
            this.hideSignatureModal();
            this.updateSigningFlowBanner();
        });

        // Close on backdrop click
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.hideSignatureModal();
            }
        });
    }

    _clearImagePreview() {
        const wrap = document.getElementById('sig-image-preview-wrap');
        const img = document.getElementById('sig-image-preview');
        const input = document.getElementById('sig-image-input');
        if (img) img.src = '';
        wrap?.classList.add('hidden');
        if (input) input.value = '';
    }

    _activeSignatureTab() {
        const t = this.signatureModal?.querySelector('.sig-tab.active');
        return t?.dataset?.tab || 'draw';
    }

    _signatureDataUrlForSave() {
        if (this._selectedSavedSig) return this._selectedSavedSig.dataUrl;
        const tab = this._activeSignatureTab();
        if (tab === 'image') return this._pendingSignatureImage || null;
        if (tab === 'draw' && this.signaturePad.isEmpty()) return null;
        if (tab === 'type' && !this.signaturePad.typedText?.trim()) return null;
        return this.signaturePad.getDataUrl() || null;
    }

    _refreshSignatureModalExtras() {
        const unlocked = secureStorage.hasVault() && secureStorage.isUnlocked();
        const mySig = document.getElementById('sig-my-signatures');
        const saveCurrent = document.getElementById('sig-save-current');
        const hasSig = !!this._signatureDataUrlForSave();

        if (mySig) mySig.classList.toggle('hidden', !unlocked);
        if (saveCurrent) saveCurrent.classList.toggle('hidden', !unlocked || !hasSig);

        if (unlocked) {
            const list = document.getElementById('sig-saved-list');
            if (!list) return;
            const sigs = secureStorage.getSignatures();
            list.innerHTML = sigs.length === 0
                ? '<span class="sig-saved-empty">No saved signatures.</span>'
                : sigs.map((s) => `<button type="button" class="sig-saved-item" data-id="${escapeHtml(s.id)}" title="${escapeHtml(s.name)}"><img src="${(s.dataUrl || '').replace(/"/g, '&quot;')}" alt=""><span>${escapeHtml(s.name)}</span></button>`).join('');
            list.querySelectorAll('.sig-saved-item').forEach((btn) => {
                btn.addEventListener('click', () => {
                    const id = btn.dataset.id;
                    const sig = secureStorage.getSignatures().find((x) => x.id === id);
                    if (!sig) return;
                    this._selectedSavedSig = sig;
                    this.signatureModal?.querySelectorAll('.sig-tab').forEach((t) => t.classList.remove('active'));
                    document.getElementById('sig-draw-area')?.classList.remove('active');
                    document.getElementById('sig-type-area')?.classList.remove('active');
                    document.getElementById('sig-image-area')?.classList.remove('active');
                    this.updateSignatureApplyState();
                });
            });
        }
    }

    /**
     * Set up signature field modal (for placing empty signature boxes when editing)
     */
    setupSignatureFieldModal() {
        const modal = document.getElementById('signature-field-modal');
        const closeBtn = document.getElementById('signature-field-modal-close');
        const cancelBtn = document.getElementById('sig-field-cancel');
        const addBtn = document.getElementById('sig-field-add');
        const labelInput = document.getElementById('sig-field-label');

        if (!modal || !closeBtn || !cancelBtn || !addBtn || !labelInput) return;

        const updateAddState = () => {
            const label = (labelInput.value || '').trim();
            addBtn.disabled = !label;
        };

        // Close modal
        closeBtn.addEventListener('click', () => this.hideSignatureFieldModal());
        cancelBtn.addEventListener('click', () => this.hideSignatureFieldModal());

        // Update Add button state when label changes
        labelInput.addEventListener('input', updateAddState);

        // Add field button
        addBtn.addEventListener('click', () => {
            const label = (labelInput.value || '').trim();
            if (!label) return;

            // Set the signature field label and switch to signature-field tool
            this.canvasManager.setSignatureFieldLabel(label);
            this.canvasManager.setTool('signature-field');
            this.toolButtons.forEach(b => b.classList.remove('active'));
            document.querySelector('[data-tool="signature-field"]')?.classList.add('active');
            this.hideSignatureFieldModal();
            
            // Show hint banner and apply cursor
            this.showToolHint('Click on the document where you want to place the signature field');
            this.applyToolCursor('signature-field');
        });

        // Close on outside click
        modal.addEventListener('click', (e) => {
            if (e.target === modal) this.hideSignatureFieldModal();
        });
    }

    /**
     * Show signature field modal
     */
    showSignatureFieldModal() {
        const modal = document.getElementById('signature-field-modal');
        const labelInput = document.getElementById('sig-field-label');
        const addBtn = document.getElementById('sig-field-add');
        if (!modal || !labelInput || !addBtn) return;

        labelInput.value = '';
        addBtn.disabled = true;
        modal.classList.remove('hidden');
        labelInput.focus();
    }

    /**
     * Hide signature field modal
     */
    hideSignatureFieldModal() {
        const modal = document.getElementById('signature-field-modal');
        if (modal) modal.classList.add('hidden');
    }

    /**
     * Show tool hint banner with message
     */
    showToolHint(message) {
        const banner = document.getElementById('tool-hint-banner');
        const text = document.getElementById('tool-hint-banner-text');
        if (!banner || !text) return;
        
        text.textContent = message;
        banner.classList.remove('hidden');
    }

    /**
     * Hide tool hint banner
     */
    hideToolHint() {
        const banner = document.getElementById('tool-hint-banner');
        if (banner) banner.classList.add('hidden');
    }

    /**
     * Apply tool-specific cursor to canvas containers
     */
    applyToolCursor(tool) {
        const containers = document.querySelectorAll('.canvas-container');
        containers.forEach(container => {
            // Remove all tool cursor classes
            container.classList.remove('text-mode', 'signature-field-mode');
            
            // Add specific cursor class
            if (tool === 'text') {
                container.classList.add('text-mode');
            } else if (tool === 'signature-field') {
                container.classList.add('signature-field-mode');
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
        const tab = this._activeSignatureTab();
        let hasSignature = !!this._selectedSavedSig;
        if (!hasSignature) {
            if (tab === 'image') hasSignature = !!this._pendingSignatureImage;
            else if (tab === 'draw') hasSignature = !this.signaturePad.isEmpty();
            else hasSignature = !!this.signaturePad.typedText?.trim();
        }

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
        const printedNameInput = document.getElementById('sig-printed-name');
        const emailInput = document.getElementById('sig-email');
        const applyBtn = document.getElementById('sig-apply');
        if (!intentCheck || !consentCheck || !nameInput || !applyBtn) return false;

        const intent = intentCheck.checked;
        const consent = consentCheck.checked;
        const fieldLabel = (nameInput.value || '').trim();  // Which signature field (slot)
        const printedName = (printedNameInput?.value || '').trim();  // Signer's legal name for audit
        const email = (emailInput?.value || '').trim();
        const dataUrl = this._selectedSavedSig
            ? this._selectedSavedSig.dataUrl
            : (this._activeSignatureTab() === 'image' ? this._pendingSignatureImage : this.signaturePad.getDataUrl());

        if (!intent || !consent || !fieldLabel || !dataUrl) return false;

        // signerName = printed name for audit trail; fallback to field label if not provided
        const meta = {
            signerName: printedName || fieldLabel,
            signerEmail: email || undefined,
            intentAccepted: intent,
            consentAccepted: consent,
            documentFilename: this.fileName || '',
            documentHash: this.documentHash || undefined
        };

        this.canvasManager.setSignature(dataUrl, meta);
        
        // In Fill mode: fill signature field(s)
        if (this.mode === 'fill') {
            if (this._targetSignatureField?.object) {
                // We have a specific field to fill
                const canvas = this.canvasManager.canvases.get(this._targetSignatureField.pageId);
                if (canvas) {
                    this.canvasManager.replaceSignatureField(this._targetSignatureField.object, canvas);
                    this._targetSignatureField = null;
                    this.refreshFieldsSidebar();
                    return true;
                }
            }
            const filled = this.fillMatchingSignatureFields(fieldLabel, email);
            if (filled > 0) {
                this.refreshFieldsSidebar();
                return true;
            } else {
                alert('No matching signature field found for this signer. Please ensure the document has a signature field with your name or contact the sender.');
                return false;
            }
        }
        
        // In Edit mode: set tool to place signature manually
        this.canvasManager.setTool('signature');
        this.toolButtons.forEach(b => b.classList.remove('active'));
        document.querySelector('[data-tool="signature"]')?.classList.add('active');
        return true;
    }

    /**
     * Find and fill signature fields that match the field label and optional email
     * @param {string} fieldLabelToMatch - Signature field label (slot identifier) to match
     * @param {string} signerEmail - Email of the signer (optional, for expected-signer matching)
     * @returns {number} Number of signature fields filled
     */
    fillMatchingSignatureFields(fieldLabelToMatch, signerEmail) {
        let filledCount = 0;
        
        // Get all signature fields from all canvases
        this.canvasManager.canvases.forEach((canvas, pageId) => {
            const objects = canvas.getObjects();
            objects.forEach(obj => {
                if (obj._annotationType === 'signature-field') {
                    const fieldLabel = obj._signatureFieldLabel || '';
                    
                    // Match by field label (exact or partial)
                    const matchesFieldLabel = fieldLabel.toLowerCase().includes(fieldLabelToMatch.toLowerCase()) ||
                                             fieldLabelToMatch.toLowerCase().includes(fieldLabel.toLowerCase());
                    const matchesEmail = signerEmail && fieldLabel.toLowerCase().includes(signerEmail.toLowerCase());
                    
                    // Also check if this field matches expected signers (by field label / email)
                    const matchesExpected = this.signatureFieldMatchesExpectedSigner(fieldLabel, fieldLabelToMatch, signerEmail);
                    
                    if (matchesFieldLabel || matchesEmail || matchesExpected) {
                        this.canvasManager.replaceSignatureField(obj, canvas);
                        filledCount++;
                    }
                }
            });
        });
        
        return filledCount;
    }

    /**
     * Check if a signature field matches an expected signer from metadata
     */
    signatureFieldMatchesExpectedSigner(fieldLabel, signerName, signerEmail) {
        if (!this.signingFlowMeta || !this.signingFlowMeta.expectedSigners) return false;
        
        const expectedSigners = this.signingFlowMeta.expectedSigners || [];
        
        // Find a matching expected signer
        const matchingSigner = expectedSigners.find(signer => {
            const nameMatches = signer.name && (
                signer.name.toLowerCase() === signerName.toLowerCase() ||
                fieldLabel.toLowerCase().includes(signer.name.toLowerCase())
            );
            const emailMatches = signerEmail && signer.email && 
                                signer.email.toLowerCase() === signerEmail.toLowerCase();
            
            return nameMatches || emailMatches;
        });
        
        return !!matchingSigner;
    }

    /**
     * Show signature modal
     * @param {Object} [options]
     * @param {Object} [options.targetField] - When signing a specific field, pass it. Name will be field ID (read-only), email pre-filled if known.
     */
    showSignatureModal(options = {}) {
        const { targetField } = options;
        this._targetSignatureField = targetField || null;
        this._pendingSignatureImage = null;
        this._selectedSavedSig = null;
        this._clearImagePreview();

        this.signatureModal.classList.remove('hidden');
        this.signaturePad.clear();
        this.signaturePad.setTypedText('');
        // Resize canvas to match display after modal is laid out (fixes coordinate offset)
        requestAnimationFrame(() => {
            requestAnimationFrame(() => {
                this.signaturePad.resize();
            });
        });
        const textInput = document.getElementById('sig-text-input');
        if (textInput) textInput.value = '';
        const preview = document.getElementById('sig-preview');
        if (preview) preview.textContent = 'Preview';

        this.signatureModal?.querySelectorAll('.sig-tab').forEach((t) => t.classList.remove('active'));
        const drawTab = this.signatureModal?.querySelector('.sig-tab[data-tab="draw"]');
        if (drawTab) drawTab.classList.add('active');
        document.getElementById('sig-draw-area')?.classList.add('active');
        document.getElementById('sig-type-area')?.classList.remove('active');
        document.getElementById('sig-image-area')?.classList.remove('active');
        this.signaturePad.setMode('draw');

        const intentCheck = document.getElementById('sig-intent');
        const consentCheck = document.getElementById('sig-consent');
        const nameInput = document.getElementById('sig-name');
        const nameLabel = document.querySelector('label[for="sig-name"]');
        const nameHint = nameInput?.closest('.sig-field')?.querySelector('.sig-field-hint');
        const printedNameInput = document.getElementById('sig-printed-name');
        const emailInput = document.getElementById('sig-email');
        const applyBtn = document.getElementById('sig-apply');
        if (intentCheck) intentCheck.checked = false;
        if (consentCheck) consentCheck.checked = false;
        if (printedNameInput) printedNameInput.value = '';
        if (applyBtn) applyBtn.disabled = true;

        if (targetField) {
            // Signing a specific field: field ID read-only, printed name optional
            const fieldLabel = targetField.signatureFieldLabel || targetField.object?._signatureFieldLabel || '';
            if (nameInput) {
                nameInput.value = fieldLabel;
                nameInput.readOnly = true;
                nameInput.classList.add('readonly');
            }
            if (nameLabel) nameLabel.textContent = 'Signature field';
            if (nameHint) nameHint.textContent = 'This field will receive your signature.';
            if (printedNameInput) {
                printedNameInput.placeholder = 'e.g. John Doe';
                printedNameInput.disabled = false;
            }
            const signerInfo = this.detectCurrentSignerForField(targetField);
            if (emailInput) {
                emailInput.value = signerInfo.email || '';
                emailInput.readOnly = !!signerInfo.email;
                emailInput.classList.toggle('readonly', !!signerInfo.email);
            }
        } else {
            // General sign: user enters field label for matching, printed name optional
            if (nameInput) {
                nameInput.value = '';
                nameInput.readOnly = false;
                nameInput.classList.remove('readonly');
            }
            if (nameLabel) nameLabel.innerHTML = 'Signature field <span class="required">*</span>';
            if (nameHint) nameHint.textContent = 'Identifies which signature slot. Must match the field label on the document.';
            if (printedNameInput) {
                printedNameInput.placeholder = 'e.g. John Doe';
                printedNameInput.disabled = false;
            }
            if (emailInput) {
                emailInput.readOnly = false;
                emailInput.classList.remove('readonly');
                emailInput.value = '';
            }
            if (this.mode === 'fill') {
                const signerInfo = this.detectCurrentSigner();
                if (signerInfo.fieldLabel && nameInput) nameInput.value = signerInfo.fieldLabel;
                if (signerInfo.email && emailInput) emailInput.value = signerInfo.email;
            }
        }

        this._refreshSignatureModalExtras();
        this.updateSignatureApplyState();
        this.updateHistoryButtons();
    }

    /**
     * Get signer email for a specific field (from URL params or signing flow metadata)
     */
    detectCurrentSignerForField(field) {
        const urlParams = new URLSearchParams(window.location.search);
        const urlEmail = urlParams.get('email');
        if (urlEmail) return { email: urlEmail };

        const fieldLabel = (field.signatureFieldLabel || field.object?._signatureFieldLabel || '').toLowerCase();
        if (!fieldLabel || !this.signingFlowMeta?.expectedSigners) return { email: '' };
        const signer = this.signingFlowMeta.expectedSigners.find(s => {
            const sName = (s.name || '').toLowerCase();
            const sEmail = (s.email || '').toLowerCase();
            if (!sName && !sEmail) return false;
            if (sName && (fieldLabel.includes(sName) || sName.includes(fieldLabel))) return true;
            if (sEmail && fieldLabel.includes(sEmail)) return true;
            return false;
        });
        return { email: signer?.email || '' };
    }

    /**
     * Detect current signer from URL parameters or signing metadata
     * Returns fieldLabel (which signature slot) and email for pre-filling the modal
     */
    detectCurrentSigner() {
        const urlParams = new URLSearchParams(window.location.search);
        const urlField = urlParams.get('field') || urlParams.get('signer') || urlParams.get('name');
        const urlEmail = urlParams.get('email');
        
        if (urlField || urlEmail) {
            return {
                fieldLabel: urlField || '',
                email: urlEmail || ''
            };
        }
        
        if (this.signingFlowMeta && this.signingFlowMeta.expectedSigners) {
            const withEmail = this.signingFlowMeta.expectedSigners.filter(s => s.email);
            if (withEmail.length === 1) {
                return {
                    fieldLabel: withEmail[0].name || '',  // expectedSigner.name = field label
                    email: withEmail[0].email || ''
                };
            }
        }
        
        return { fieldLabel: '', email: '' };
    }

    /**
     * Hide signature modal
     */
    hideSignatureModal() {
        this.signatureModal.classList.add('hidden');
        this.updateHistoryButtons();
    }

    /**
     * Export the PDF with annotations (shared logic)
     * @returns {Promise<{ bytes: Uint8Array; exportName: string } | null>}
     */
    async getExportedPDF() {
        if (!this.pdfHandler.isLoaded()) return null;

        const annotationsArr = this.canvasManager.getAllAnnotations();
        const annotationsByPageId = new Map();
        for (const p of annotationsArr) {
            annotationsByPageId.set(p.pageId, p.annotations);
        }

        // Collect locked signature field labels and locked form field names (permanently locked on export)
        const lockedSignatureFields = [];
        const lockedFormFields = [];
        for (const p of annotationsArr) {
            for (const ann of p.annotations || []) {
                if (ann.type === 'signature' && ann.object?._signatureLocked && ann.data?._signatureMeta?.replacedFieldLabel) {
                    const label = (ann.data._signatureMeta.replacedFieldLabel || '').trim();
                    if (label && !lockedSignatureFields.includes(label)) lockedSignatureFields.push(label);
                }
                if (['textfield', 'checkbox', 'radio', 'dropdown', 'date'].includes(ann.type) && ann.object?._fieldLocked && ann.data?._fieldName) {
                    const name = (ann.data._fieldName || '').trim();
                    if (name && !lockedFormFields.includes(name)) lockedFormFields.push(name);
                }
            }
        }
        const mergedLockedSignatures = [...new Set([...(this.signingFlowMeta?.lockedSignatureFields || []), ...lockedSignatureFields])];
        const mergedLockedFormFields = [...new Set([...(this.signingFlowMeta?.lockedFormFields || []), ...lockedFormFields])];

        // Document stage: signed if any locked fields; otherwise preserve previous or draft
        const documentStage = (mergedLockedSignatures.length > 0 || mergedLockedFormFields.length > 0) ? 'signed' : (this.signingFlowMeta?.documentStage || 'draft');

        // Build canonical digest for hash chain (proves when changes were made)
        const canonical = {
            viewPages: (this.viewPages || []).map((vp) => ({ id: vp.id, docId: vp.docId, sourcePageNum: vp.sourcePageNum, rotation: vp.rotation })),
            pages: (this.viewPages || []).map((vp) => {
                const anns = annotationsByPageId.get(vp.id) || [];
                const digests = anns.map((a) => {
                    const type = a.type || '';
                    const label = a.data?._signatureMeta?.replacedFieldLabel ?? a.data?._fieldName ?? a.data?._signatureFieldLabel ?? '';
                    return { type, label: (label || '').trim() };
                }).sort((a, b) => (a.type + a.label).localeCompare(b.type + b.label));
                return digests;
            })
        };
        const documentHash = await computeDocumentHash(JSON.stringify(canonical));
        const previousHash = this.signingFlowMeta?.hashChain?.hash || undefined;
        const hashChain = documentHash ? { hash: documentHash, timestamp: new Date().toISOString(), previousHash } : undefined;
        this.documentHash = documentHash || this.documentHash;

        const exportSigningFlowMeta = {
            ...this.signingFlowMeta,
            lockedSignatureFields: mergedLockedSignatures.length > 0 ? mergedLockedSignatures : (this.signingFlowMeta?.lockedSignatureFields),
            lockedFormFields: mergedLockedFormFields.length > 0 ? mergedLockedFormFields : (this.signingFlowMeta?.lockedFormFields),
            documentStage,
            hashChain
        };

        const docBytesById = this.pdfHandler.getAllOriginalBytes();
        const modifiedPdfBytes = await this.exporter.exportPDF({
            docBytesById,
            viewPages: this.viewPages,
            annotationsByPageId,
            scale: this.currentScale,
            mainDocId: this.pdfHandler.mainDocId,
            signingFlowMeta: exportSigningFlowMeta
        });

        // Once exported, those locked fields become permanent for this session (cannot be unlocked)
        mergedLockedSignatures.forEach((l) => this.lockedFromFile.signatures.add(l));
        mergedLockedFormFields.forEach((l) => this.lockedFromFile.formFields.add(l));

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

        const editLink =
            typeof window !== 'undefined' && window.location
                ? (window.location.origin + window.location.pathname).replace(/\/$/, '')
                : '';
        return {
            filename: exportName,
            date: new Date().toLocaleString(),
            signatureSummary,
            signerNames,
            pageCount: this.pdfHandler.totalPages || 0,
            documentHash: this.documentHash || '',
            editLink
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

        const isDocTemplate = sel.value === '__doc__';
        const tpl = isDocTemplate ? { subject: subjectEl.value || '', body: bodyEl.value || '' } : (emailTemplates.getById(sel.value) || emailTemplates.getDefault());
        const subject = (subjectEl.value || '').trim();
        const body = (bodyEl.value || '').trim();
        if (!subject || !body) {
            alert('Please provide a subject and body.');
            return;
        }

        this.showLoading('Preparing email...');
        try {
            // Attach full template to document so it follows the doc across machines (saved in PDF Keywords)
            if (!this.signingFlowMeta) this.signingFlowMeta = { signers: [], expectedSigners: [] };
            if (!isDocTemplate) this.signingFlowMeta.emailTemplate = { subject: tpl.subject, body: tpl.body };
            const result = await this.getExportedPDF();
            if (!result) {
                this.hideLoading();
                return;
            }
            this.exporter.downloadPDF(result.bytes, result.exportName);
            const ctx = this.buildEmailContext(result.exportName);
            const filled = emailTemplates.fill({ subject, body }, ctx);
            // When document has completion flow, pre-fill To with original sender or all signers
            const toEmails = this.getCompletionToEmails();
            const toPart = toEmails.length > 0 ? encodeURIComponent(toEmails.join(',')) + '?' : '?';
            const ccEl = document.getElementById('send-cc');
            const bccEl = document.getElementById('send-bcc');
            const ccVal = (ccEl?.value || '').trim();
            const bccVal = (bccEl?.value || '').trim();
            let mailto = `mailto:${toPart}subject=${encodeURIComponent(filled.subject)}&body=${encodeURIComponent(filled.body)}`;
            if (ccVal) mailto += `&cc=${encodeURIComponent(ccVal)}`;
            if (bccVal) mailto += `&bcc=${encodeURIComponent(bccVal)}`;
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
     * Emails to use for "Send completed document to" (original sender or all signers). Used in mailto To.
     * @returns {string[]}
     */
    getCompletionToEmails() {
        const meta = this.signingFlowMeta;
        if (!meta) return [];
        if (typeof meta.originalSenderEmail === 'string' && meta.originalSenderEmail.trim()) return [meta.originalSenderEmail.trim()];
        if (Array.isArray(meta.completionToEmails) && meta.completionToEmails.length > 0) return meta.completionToEmails;
        return [];
    }

    getCompletionCcEmails() {
        const meta = this.signingFlowMeta;
        if (!meta || !Array.isArray(meta.completionCcEmails)) return [];
        return meta.completionCcEmails;
    }

    getCompletionBccEmails() {
        const meta = this.signingFlowMeta;
        if (!meta || !Array.isArray(meta.completionBccEmails)) return [];
        return meta.completionBccEmails;
    }

    /**
     * Populate Send modal: use full template embedded in doc when present (works across machines), else default template.
     */
    refreshSendModal() {
        const sel = document.getElementById('send-template-select');
        const subjectEl = document.getElementById('send-subject');
        const bodyEl = document.getElementById('send-body');
        const sendToHint = document.getElementById('send-to-hint');
        if (!sel || !subjectEl || !bodyEl) return;

        const templates = emailTemplates.getTemplates();
        const defaultTpl = emailTemplates.getDefault();
        const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim();
        const hasRealFilename = baseName && baseName !== 'document';
        const exportName = hasRealFilename ? `${baseName}-edited.pdf` : 'document.pdf';
        const ctx = this.buildEmailContext(exportName);

        // Prefer full template embedded in document (same on all machines)
        const docTemplate = this.signingFlowMeta?.emailTemplate;
        if (docTemplate && typeof docTemplate.subject === 'string' && typeof docTemplate.body === 'string') {
            const filled = emailTemplates.fill({ subject: docTemplate.subject, body: docTemplate.body }, ctx);
            subjectEl.value = filled.subject;
            bodyEl.value = filled.body;
            sel.innerHTML = templates.map((t) => `<option value="${t.id}">${escapeHtml(t.name)}</option>`).join('');
            const opt = document.createElement('option');
            opt.value = '__doc__';
            opt.selected = true;
            opt.textContent = '(Document template)';
            sel.insertBefore(opt, sel.firstChild);
        } else {
            const tplToUse = defaultTpl;
            sel.innerHTML = templates.map((t) => `<option value="${t.id}" ${t.id === tplToUse.id ? 'selected' : ''}>${escapeHtml(t.name)}</option>`).join('');
            const filled = emailTemplates.fill(tplToUse, ctx);
            subjectEl.value = filled.subject;
            bodyEl.value = filled.body;
        }

        // Show "Send to" hint and pre-fill CC/BCC when completion emails are set
        const toEmails = this.getCompletionToEmails();
        if (sendToHint) {
            if (toEmails.length > 0) {
                sendToHint.textContent = `This email will be addressed to: ${toEmails.join(', ')}`;
                sendToHint.classList.remove('hidden');
            } else {
                sendToHint.classList.add('hidden');
                sendToHint.textContent = '';
            }
        }
        const sendCc = document.getElementById('send-cc');
        const sendBcc = document.getElementById('send-bcc');
        if (sendCc) sendCc.value = (this.getCompletionCcEmails() || []).join(', ');
        if (sendBcc) sendBcc.value = (this.getCompletionBccEmails() || []).join(', ');
    }

    showSendModal() {
        if (!this.pdfHandler.isLoaded()) return;
        this.refreshSendModal();
        this.sendModal.classList.remove('hidden');
    }

    hideSendModal() {
        this.sendModal.classList.add('hidden');
    }

    /**
     * Vault: password-protected storage for templates & signatures.
     * No vault -> legacy localStorage. Vault exists -> locked (default-only) or unlocked (vault backend).
     */
    setupVault() {
        secureStorage.migrateFromLegacyIfNeeded();

        const lockedBackend = {
            loadStore: getDefaultOnlyTemplatesStore,
            saveStore: () => Promise.resolve()
        };
        const vaultBackend = () => ({
            loadStore: () => secureStorage.getTemplatesStore(),
            saveStore: (s) => secureStorage.saveTemplatesStore(s)
        });

        if (!secureStorage.hasVault()) {
            setTemplatesBackend(null);
        } else {
            setTemplatesBackend(secureStorage.isUnlocked() ? vaultBackend() : lockedBackend);
        }
        this.updateVaultUI();

        const vaultModal = document.getElementById('vault-modal');
        const vaultClose = document.getElementById('vault-modal-close');
        const createPanel = document.getElementById('vault-create-panel');
        const unlockPanel = document.getElementById('vault-unlock-panel');
        const createName = document.getElementById('vault-create-name');
        const createPw = document.getElementById('vault-create-password');
        const createConfirm = document.getElementById('vault-create-confirm');
        const createError = document.getElementById('vault-create-error');
        const createBtn = document.getElementById('vault-create-btn');
        const unlockSelect = document.getElementById('vault-unlock-select');
        const unlockSelectWrap = document.getElementById('vault-unlock-select-wrap');
        const unlockSingleWrap = document.getElementById('vault-unlock-single-wrap');
        const unlockSingleName = document.getElementById('vault-unlock-single-name');
        const unlockHint = document.getElementById('vault-unlock-hint');
        const unlockPw = document.getElementById('vault-unlock-password');
        const unlockError = document.getElementById('vault-unlock-error');
        const unlockBtn = document.getElementById('vault-unlock-btn');
        const vaultDeleteBtn = document.getElementById('vault-delete-btn');
        const createAnotherLink = document.getElementById('vault-create-another-link');
        const unlockedPanel = document.getElementById('vault-unlocked-panel');
        const unlockedNameEl = document.getElementById('vault-unlocked-name');
        const vaultLockBtn = document.getElementById('vault-lock-btn');
        const vaultSwitchBtn = document.getElementById('vault-switch-btn');
        const vaultRenameBtn = document.getElementById('vault-rename-btn');
        const vaultExportBtn = document.getElementById('vault-export-btn');
        const vaultTransferQrBtn = document.getElementById('vault-transfer-qr-btn');
        const vaultTransferCopyBtn = document.getElementById('vault-transfer-copy-btn');
        const vaultTransferSendPanel = document.getElementById('vault-transfer-send-panel');
        const vaultTransferQrWrap = document.getElementById('vault-transfer-qr-wrap');
        const vaultTransferProgress = document.getElementById('vault-transfer-progress');
        const vaultTransferSendCancel = document.getElementById('vault-transfer-send-cancel');
        const vaultReceiveBtn = document.getElementById('vault-receive-btn');
        const vaultReceivePasteBtn = document.getElementById('vault-receive-paste-btn');
        const vaultPastePanel = document.getElementById('vault-paste-panel');
        const vaultPasteInput = document.getElementById('vault-paste-input');
        const vaultPasteError = document.getElementById('vault-paste-error');
        const vaultPasteImportBtn = document.getElementById('vault-paste-import-btn');
        const vaultPasteCancel = document.getElementById('vault-paste-cancel');
        const vaultReceivePanel = document.getElementById('vault-receive-panel');
        const vaultReceiveVideo = document.getElementById('vault-receive-video');
        const vaultReceiveCanvas = document.getElementById('vault-receive-canvas');
        const vaultReceiveProgress = document.getElementById('vault-receive-progress');
        const vaultReceiveCancel = document.getElementById('vault-receive-cancel');
        const vaultReceiveDone = document.getElementById('vault-receive-done');
        const vaultReceivePassword = document.getElementById('vault-receive-password');
        const vaultReceiveError = document.getElementById('vault-receive-error');
        const vaultReceiveImportBtn = document.getElementById('vault-receive-import-btn');
        const vaultDeleteCurrentBtn = document.getElementById('vault-delete-current-btn');
        const vaultRenameForm = document.getElementById('vault-rename-form');
        const vaultRenamePassword = document.getElementById('vault-rename-password');
        const vaultRenameNew = document.getElementById('vault-rename-new');
        const vaultRenameError = document.getElementById('vault-rename-error');
        const vaultRenameSave = document.getElementById('vault-rename-save');
        const vaultExportPlainSectionBtn = document.getElementById('vault-export-plain-section-btn');
        const vaultImportPlainInput = document.getElementById('vault-import-plain-input');
        const vaultImportInput = document.getElementById('vault-import-input');
        const vaultImportForm = document.getElementById('vault-import-form');
        const vaultImportPassword = document.getElementById('vault-import-password');
        const vaultImportError = document.getElementById('vault-import-error');
        const vaultImportNewBtn = document.getElementById('vault-import-new-btn');
        const vaultImportReplaceBtn = document.getElementById('vault-import-replace-btn');
        const btnVault = document.getElementById('btn-vault');
        const btnVaultLabel = document.getElementById('btn-vault-label');
        const templatesUnlockLink = document.getElementById('templates-unlock-link');

        let pendingImportData = null;
        let vaultTransferIntervalId = null;
        let vaultReceiveStream = null;
        let vaultReceiveAnimationId = null;
        const VAULT_QR_PREFIX = 'fpv:';
        const VAULT_QR_CHUNK_SIZE = 1100;

        const hideCreate = () => {
            createPanel?.classList.add('hidden');
            if (createName) createName.value = '';
            if (createPw) createPw.value = '';
            if (createConfirm) createConfirm.value = '';
            createError?.classList.add('hidden');
        };
        const hideUnlock = () => {
            unlockPanel?.classList.add('hidden');
            if (unlockPw) unlockPw.value = '';
            unlockError?.classList.add('hidden');
        };
        const hideUnlocked = () => {
            unlockedPanel?.classList.add('hidden');
            vaultRenameForm?.classList.add('hidden');
            if (vaultRenamePassword) vaultRenamePassword.value = '';
            if (vaultRenameNew) vaultRenameNew.value = '';
            vaultRenameError?.classList.add('hidden');
        };
        const hideImportForm = () => {
            vaultImportForm?.classList.add('hidden');
            pendingImportData = null;
            if (vaultImportPassword) vaultImportPassword.value = '';
            vaultImportError?.classList.add('hidden');
            if (vaultImportInput) vaultImportInput.value = '';
        };
        const hideTransferPanels = () => {
            if (vaultTransferIntervalId != null) {
                clearInterval(vaultTransferIntervalId);
                vaultTransferIntervalId = null;
            }
            vaultTransferSendPanel?.classList.add('hidden');
            if (vaultTransferQrWrap) vaultTransferQrWrap.innerHTML = '';
            vaultReceivePanel?.classList.add('hidden');
            vaultReceiveDone?.classList.add('hidden');
            if (vaultReceiveStream) {
                vaultReceiveStream.getTracks().forEach((t) => t.stop());
                vaultReceiveStream = null;
            }
            if (vaultReceiveVideo) vaultReceiveVideo.srcObject = null;
            if (vaultReceiveAnimationId != null) {
                cancelAnimationFrame(vaultReceiveAnimationId);
                vaultReceiveAnimationId = null;
            }
            if (vaultReceivePassword) vaultReceivePassword.value = '';
            vaultReceiveError?.classList.add('hidden');
            vaultPastePanel?.classList.add('hidden');
            if (vaultPasteInput) vaultPasteInput.value = '';
            vaultPasteError?.classList.add('hidden');
        };
        const populateUnlockSelect = () => {
            const reg = secureStorage.getRegistry();
            if (!unlockSelect) return;
            unlockSelect.innerHTML = reg.map((r) => `<option value="${escapeHtml(r.id)}">${escapeHtml(r.name)}</option>`).join('');
            const single = reg.length === 1;
            if (unlockSelectWrap) unlockSelectWrap.classList.toggle('hidden', single);
            if (unlockSingleWrap) unlockSingleWrap.classList.toggle('hidden', !single);
            if (unlockSingleName && single) unlockSingleName.textContent = reg[0].name;
            if (unlockHint) unlockHint.textContent = single ? 'Enter the vault password.' : 'Select a vault and enter its password.';
        };
        const getSelectedUnlockVaultId = () => unlockSelect?.value || (secureStorage.getRegistry()[0]?.id ?? null);
        const getSelectedUnlockVaultName = () => {
            const id = getSelectedUnlockVaultId();
            const r = secureStorage.getRegistry().find((x) => x.id === id);
            return r?.name ?? '';
        };
        const showUnlockedPanel = () => {
            if (unlockedNameEl) unlockedNameEl.textContent = secureStorage.getActiveVaultName();
            unlockedPanel?.classList.remove('hidden');
        };
        const showVaultModal = (panel) => {
            hideCreate();
            hideUnlock();
            hideUnlocked();
            hideImportForm();
            hideTransferPanels();
            if (panel === 'create') {
                createPanel?.classList.remove('hidden');
            } else if (panel === 'unlock') {
                populateUnlockSelect();
                unlockPanel?.classList.remove('hidden');
            } else if (panel === 'unlocked') {
                showUnlockedPanel();
            }
            vaultModal?.classList.remove('hidden');
        };
        const hideVaultModal = () => {
            vaultModal?.classList.add('hidden');
            hideCreate();
            hideUnlock();
            hideUnlocked();
            hideImportForm();
            hideTransferPanels();
        };
        const refreshVaultModalState = () => {
            this.updateVaultUI();
            this.renderTemplatesList();
            this.refreshSendModal();
        };

        vaultClose?.addEventListener('click', hideVaultModal);
        vaultModal?.addEventListener('click', (e) => { if (e.target === vaultModal) hideVaultModal(); });

        btnVault?.addEventListener('click', () => {
            if (!secureStorage.hasVault()) showVaultModal('create');
            else if (secureStorage.isUnlocked()) showVaultModal('unlocked');
            else showVaultModal('unlock');
        });

        createBtn?.addEventListener('click', async () => {
            const name = (createName?.value || '').trim();
            const pw = (createPw?.value || '').trim();
            const conf = (createConfirm?.value || '').trim();
            createError?.classList.add('hidden');
            if (!pw) { createError.textContent = 'Enter a password.'; createError?.classList.remove('hidden'); return; }
            if (pw !== conf) { createError.textContent = 'Passwords do not match.'; createError?.classList.remove('hidden'); return; }
            try {
                await secureStorage.createVault(name || 'Unnamed', pw);
                setTemplatesBackend(vaultBackend());
                hideVaultModal();
                this.updateVaultUI();
                this.renderTemplatesList();
                this.refreshSendModal();
            } catch (e) {
                createError.textContent = e instanceof Error ? e.message : 'Create failed.';
                createError?.classList.remove('hidden');
            }
        });

        unlockBtn?.addEventListener('click', async () => {
            const id = getSelectedUnlockVaultId();
            const pw = (unlockPw?.value || '').trim();
            unlockError?.classList.add('hidden');
            if (!id) { unlockError.textContent = 'Select a vault.'; unlockError?.classList.remove('hidden'); return; }
            if (!pw) { unlockError.textContent = 'Enter your password.'; unlockError?.classList.remove('hidden'); return; }
            try {
                await secureStorage.unlock(id, pw);
                setTemplatesBackend(vaultBackend());
                hideUnlock();
                showUnlockedPanel();
                refreshVaultModalState();
            } catch (e) {
                unlockError.textContent = e instanceof Error ? e.message : 'Unlock failed.';
                unlockError?.classList.remove('hidden');
            }
        });

        vaultDeleteBtn?.addEventListener('click', async () => {
            const id = getSelectedUnlockVaultId();
            const name = getSelectedUnlockVaultName();
            const pw = (unlockPw?.value || '').trim();
            unlockError?.classList.add('hidden');
            if (!id) { unlockError.textContent = 'Select a vault.'; unlockError?.classList.remove('hidden'); return; }
            if (!pw) { unlockError.textContent = 'Enter password to confirm deletion.'; unlockError?.classList.remove('hidden'); return; }
            if (!confirm(`Delete vault "${name}"? This cannot be undone.`)) return;
            try {
                await secureStorage.deleteVault(id, pw);
                if (!secureStorage.hasVault()) {
                    setTemplatesBackend(null);
                    hideVaultModal();
                } else {
                    populateUnlockSelect();
                    if (unlockPw) unlockPw.value = '';
                }
                refreshVaultModalState();
            } catch (e) {
                unlockError.textContent = e instanceof Error ? e.message : 'Delete failed.';
                unlockError?.classList.remove('hidden');
            }
        });

        vaultLockBtn?.addEventListener('click', () => {
            secureStorage.lock();
            setTemplatesBackend(lockedBackend);
            hideUnlocked();
            if (secureStorage.hasVault()) {
                populateUnlockSelect();
                unlockPanel?.classList.remove('hidden');
            }
            hideVaultModal();
            refreshVaultModalState();
        });

        vaultSwitchBtn?.addEventListener('click', () => {
            secureStorage.lock();
            setTemplatesBackend(lockedBackend);
            hideUnlocked();
            populateUnlockSelect();
            unlockPanel?.classList.remove('hidden');
            if (unlockPw) unlockPw.value = '';
            unlockError?.classList.add('hidden');
            refreshVaultModalState();
        });

        vaultRenameBtn?.addEventListener('click', () => {
            vaultRenameForm?.classList.toggle('hidden', vaultRenameForm?.classList.contains('hidden'));
            if (vaultRenameNew) vaultRenameNew.value = secureStorage.getActiveVaultName();
        });

        vaultRenameSave?.addEventListener('click', async () => {
            const pw = (vaultRenamePassword?.value || '').trim();
            const newName = (vaultRenameNew?.value || '').trim();
            vaultRenameError?.classList.add('hidden');
            if (!pw) { vaultRenameError.textContent = 'Enter current password.'; vaultRenameError?.classList.remove('hidden'); return; }
            if (!newName) { vaultRenameError.textContent = 'Enter a new name.'; vaultRenameError?.classList.remove('hidden'); return; }
            try {
                await secureStorage.renameVault(secureStorage.getActiveVaultId(), pw, newName);
                vaultRenameForm?.classList.add('hidden');
                if (vaultRenamePassword) vaultRenamePassword.value = '';
                if (unlockedNameEl) unlockedNameEl.textContent = secureStorage.getActiveVaultName();
                refreshVaultModalState();
            } catch (e) {
                vaultRenameError.textContent = e instanceof Error ? e.message : 'Rename failed.';
                vaultRenameError?.classList.remove('hidden');
            }
        });

        vaultExportBtn?.addEventListener('click', () => {
            try {
                const data = secureStorage.exportVault();
                const name = (data.name || 'vault').replace(/[^a-zA-Z0-9-_]/g, '-');
                const date = new Date().toISOString().slice(0, 10);
                const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `free-pdf-vault-${name}-${date}.json`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            } catch (e) {
                alert('Export failed: ' + (e instanceof Error ? e.message : String(e)));
            }
        });

        vaultExportPlainSectionBtn?.addEventListener('click', () => {
            try {
                const templatesJson = emailTemplates.exportJson();
                const data = JSON.parse(templatesJson);
                data.signatures = secureStorage.isUnlocked() ? secureStorage.getSignatures() : [];
                const date = new Date().toISOString().slice(0, 10);
                const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `free-pdf-backup-${date}.json`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            } catch (e) {
                alert('Export failed: ' + (e instanceof Error ? e.message : String(e)));
            }
        });

        vaultTransferQrBtn?.addEventListener('click', () => {
            if (typeof QRCode === 'undefined') {
                alert('QR code library not loaded.');
                return;
            }
            try {
                const data = secureStorage.exportVault();
                const jsonStr = JSON.stringify(data);
                const b64 = btoa(unescape(encodeURIComponent(jsonStr)));
                const chunks = [];
                for (let i = 0; i < b64.length; i += VAULT_QR_CHUNK_SIZE) {
                    chunks.push(b64.slice(i, i + VAULT_QR_CHUNK_SIZE));
                }
                const total = chunks.length;
                let index = 0;
                const showChunk = () => {
                    if (!vaultTransferQrWrap) return;
                    vaultTransferQrWrap.innerHTML = '';
                    const text = VAULT_QR_PREFIX + index + ':' + total + ':' + chunks[index];
                    new QRCode(vaultTransferQrWrap, { text, width: 256, height: 256 });
                    if (vaultTransferProgress) vaultTransferProgress.textContent = `Part ${index + 1} of ${total}`;
                };
                showChunk();
                vaultTransferSendPanel?.classList.remove('hidden');
                vaultTransferIntervalId = setInterval(() => {
                    index = (index + 1) % total;
                    showChunk();
                }, 2000);
            } catch (e) {
                alert('Transfer failed: ' + (e instanceof Error ? e.message : String(e)));
            }
        });

        vaultTransferCopyBtn?.addEventListener('click', async () => {
            try {
                const data = secureStorage.exportVault();
                const jsonStr = JSON.stringify(data);
                const b64 = btoa(unescape(encodeURIComponent(jsonStr)));
                await navigator.clipboard.writeText(b64);
                alert('Vault data copied to clipboard. Paste it on the other device in “Paste transfer data”.');
            } catch (e) {
                alert('Copy failed: ' + (e instanceof Error ? e.message : String(e)));
            }
        });

        vaultTransferSendCancel?.addEventListener('click', () => {
            if (vaultTransferIntervalId != null) clearInterval(vaultTransferIntervalId);
            vaultTransferIntervalId = null;
            vaultTransferSendPanel?.classList.add('hidden');
            if (vaultTransferQrWrap) vaultTransferQrWrap.innerHTML = '';
        });

        vaultReceiveBtn?.addEventListener('click', async () => {
            if (typeof jsQR === 'undefined') {
                alert('QR scanner not loaded.');
                return;
            }
            try {
                vaultReceivePanel?.classList.remove('hidden');
                vaultReceiveDone?.classList.add('hidden');
                vaultReceiveProgress?.classList.remove('hidden');
                if (vaultReceiveProgress) vaultReceiveProgress.textContent = 'Starting camera…';
                const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
                vaultReceiveStream = stream;
                if (vaultReceiveVideo) {
                    vaultReceiveVideo.srcObject = stream;
                    vaultReceiveVideo.setAttribute('playsinline', true);
                    await vaultReceiveVideo.play();
                }
                const receivedParts = {};
                let totalParts = -1;
                const checkDone = () => {
                    if (totalParts < 0) return false;
                    for (let i = 0; i < totalParts; i++) if (!receivedParts[i]) return false;
                    return true;
                };
                const tryDecode = () => {
                    if (!vaultReceiveVideo || vaultReceiveVideo.readyState < 2 || !vaultReceiveCanvas) return;
                    const w = vaultReceiveVideo.videoWidth;
                    const h = vaultReceiveVideo.videoHeight;
                    if (w === 0 || h === 0) return;
                    vaultReceiveCanvas.width = w;
                    vaultReceiveCanvas.height = h;
                    const ctx = vaultReceiveCanvas.getContext('2d');
                    ctx.drawImage(vaultReceiveVideo, 0, 0);
                    const imageData = ctx.getImageData(0, 0, w, h);
                    const result = jsQR(imageData.data, w, h);
                    if (result && result.data.startsWith(VAULT_QR_PREFIX)) {
                        const rest = result.data.slice(VAULT_QR_PREFIX.length);
                        const firstColon = rest.indexOf(':');
                        const secondColon = rest.indexOf(':', firstColon + 1);
                        if (firstColon >= 0 && secondColon >= 0) {
                            const partIndex = parseInt(rest.slice(0, firstColon), 10);
                            const total = parseInt(rest.slice(firstColon + 1, secondColon), 10);
                            const chunk = rest.slice(secondColon + 1);
                            if (!isNaN(partIndex) && !isNaN(total)) {
                                totalParts = total;
                                receivedParts[partIndex] = chunk;
                                const have = Object.keys(receivedParts).length;
                                if (vaultReceiveProgress) vaultReceiveProgress.textContent = `Received ${have} of ${total} parts`;
                                if (checkDone()) {
                                    if (vaultReceiveStream) {
                                        vaultReceiveStream.getTracks().forEach((t) => t.stop());
                                        vaultReceiveStream = null;
                                    }
                                    if (vaultReceiveVideo) vaultReceiveVideo.srcObject = null;
                                    if (vaultReceiveAnimationId != null) cancelAnimationFrame(vaultReceiveAnimationId);
                                    vaultReceiveAnimationId = null;
                                    const combined = Array.from({ length: total }, (_, i) => receivedParts[i]).join('');
                                    try {
                                        const jsonStr = decodeURIComponent(escape(atob(combined)));
                                        const data = JSON.parse(jsonStr);
                                        if (!data.name || !data.salt || !data.payload) throw new Error('Invalid data');
                                        pendingImportData = data;
                                        vaultReceivePanel?.classList.add('hidden');
                                        vaultReceiveProgress?.classList.add('hidden');
                                        vaultReceiveDone?.classList.remove('hidden');
                                        if (vaultReceivePassword) vaultReceivePassword.value = '';
                                        vaultReceiveError?.classList.add('hidden');
                                    } catch (err) {
                                        if (vaultReceiveProgress) vaultReceiveProgress.textContent = 'Invalid data. Try again.';
                                    }
                                    return;
                                }
                            }
                        }
                    }
                };
                const tick = () => {
                    tryDecode();
                    vaultReceiveAnimationId = requestAnimationFrame(tick);
                };
                tick();
            } catch (e) {
                vaultReceivePanel?.classList.add('hidden');
                alert('Camera failed: ' + (e instanceof Error ? e.message : String(e)));
            }
        });

        vaultReceiveCancel?.addEventListener('click', () => {
            if (vaultReceiveStream) {
                vaultReceiveStream.getTracks().forEach((t) => t.stop());
                vaultReceiveStream = null;
            }
            if (vaultReceiveVideo) vaultReceiveVideo.srcObject = null;
            if (vaultReceiveAnimationId != null) cancelAnimationFrame(vaultReceiveAnimationId);
            vaultReceiveAnimationId = null;
            vaultReceivePanel?.classList.add('hidden');
            vaultReceiveDone?.classList.add('hidden');
        });

        vaultReceivePasteBtn?.addEventListener('click', () => {
            vaultPastePanel?.classList.remove('hidden');
            if (vaultPasteInput) vaultPasteInput.value = '';
            vaultPasteError?.classList.add('hidden');
        });

        vaultPasteCancel?.addEventListener('click', () => {
            vaultPastePanel?.classList.add('hidden');
            if (vaultPasteInput) vaultPasteInput.value = '';
        });

        vaultPasteImportBtn?.addEventListener('click', () => {
            vaultPasteError?.classList.add('hidden');
            const raw = (vaultPasteInput?.value || '').trim();
            if (!raw) {
                vaultPasteError?.classList.remove('hidden');
                if (vaultPasteError) vaultPasteError.textContent = 'Paste the transfer data first.';
                return;
            }
            try {
                const jsonStr = decodeURIComponent(escape(atob(raw)));
                const data = JSON.parse(jsonStr);
                if (!data.name || !data.salt || !data.payload) throw new Error('Invalid vault data.');
                pendingImportData = data;
                vaultPastePanel?.classList.add('hidden');
                if (vaultPasteInput) vaultPasteInput.value = '';
                vaultReceiveDone?.classList.remove('hidden');
                if (vaultReceivePassword) vaultReceivePassword.value = '';
                vaultReceiveError?.classList.add('hidden');
            } catch (err) {
                vaultPasteError?.classList.remove('hidden');
                vaultPasteError.textContent = err instanceof Error ? err.message : 'Invalid data.';
            }
        });

        vaultReceiveImportBtn?.addEventListener('click', async () => {
            const pw = (vaultReceivePassword?.value || '').trim();
            vaultReceiveError?.classList.add('hidden');
            if (!pw) { vaultReceiveError.textContent = 'Enter password for vault.'; vaultReceiveError?.classList.remove('hidden'); return; }
            if (!pendingImportData) return;
            try {
                await secureStorage.importVaultAsNew(pendingImportData, pw);
                setTemplatesBackend(vaultBackend());
                vaultReceiveDone?.classList.add('hidden');
                pendingImportData = null;
                if (vaultReceivePassword) vaultReceivePassword.value = '';
                hideUnlock();
                showUnlockedPanel();
                refreshVaultModalState();
            } catch (e) {
                vaultReceiveError.textContent = e instanceof Error ? e.message : 'Import failed.';
                vaultReceiveError?.classList.remove('hidden');
            }
        });

        vaultDeleteCurrentBtn?.addEventListener('click', async () => {
            const name = secureStorage.getActiveVaultName();
            const pw = prompt(`Enter password for "${name}" to delete this vault. This cannot be undone.`);
            if (pw == null) return;
            try {
                await secureStorage.deleteVault(secureStorage.getActiveVaultId(), pw.trim());
                setTemplatesBackend(secureStorage.hasVault() ? lockedBackend : null);
                hideVaultModal();
                refreshVaultModalState();
            } catch (e) {
                alert('Delete failed: ' + (e instanceof Error ? e.message : String(e)));
            }
        });

        vaultImportInput?.addEventListener('change', async (e) => {
            const f = e.target.files?.[0];
            if (!f) return;
            try {
                const text = await f.text();
                const data = JSON.parse(text);
                if (!data.name || !data.salt || !data.payload) throw new Error('Invalid vault file.');
                pendingImportData = data;
                vaultImportForm?.classList.remove('hidden');
                if (vaultImportPassword) vaultImportPassword.value = '';
                vaultImportError?.classList.add('hidden');
                vaultImportReplaceBtn?.classList.toggle('hidden', !secureStorage.isUnlocked());
            } catch (err) {
                alert('Invalid file: ' + (err instanceof Error ? err.message : String(err)));
            }
            e.target.value = '';
        });

        vaultImportPlainInput?.addEventListener('change', async (e) => {
            const f = e.target.files?.[0];
            if (!f) return;
            try {
                const text = await f.text();
                const data = JSON.parse(text);
                if (!data.templates || !Array.isArray(data.templates)) throw new Error('Invalid backup file: missing templates.');
                if (secureStorage.hasVault() && !secureStorage.isUnlocked()) {
                    alert('Unlock a vault to import. Plain backup includes templates and signatures.');
                    e.target.value = '';
                    return;
                }
                const payload = { version: data.version || 1, defaultId: data.defaultId || 'default', templates: data.templates };
                const { imported, errors } = await emailTemplates.importJson(JSON.stringify(payload), { replace: true });
                if (secureStorage.isUnlocked() && Array.isArray(data.signatures) && data.signatures.length > 0) {
                    await secureStorage.setSignatures(data.signatures);
                }
                if (errors.length > 0) alert('Imported with notes: ' + errors.join(' '));
                else alert(`Imported ${imported} template(s)` + (secureStorage.isUnlocked() && data.signatures?.length ? ` and ${data.signatures.length} signature(s).` : '.'));
                refreshVaultModalState();
                this.renderTemplatesList();
            } catch (err) {
                alert('Import failed: ' + (err instanceof Error ? err.message : String(err)));
            }
            e.target.value = '';
        });

        vaultImportNewBtn?.addEventListener('click', async () => {
            const pw = (vaultImportPassword?.value || '').trim();
            vaultImportError?.classList.add('hidden');
            if (!pw) { vaultImportError.textContent = 'Enter password for vault file.'; vaultImportError?.classList.remove('hidden'); return; }
            if (!pendingImportData) return;
            try {
                await secureStorage.importVaultAsNew(pendingImportData, pw);
                setTemplatesBackend(vaultBackend());
                hideImportForm();
                hideUnlock();
                showUnlockedPanel();
                refreshVaultModalState();
            } catch (e) {
                vaultImportError.textContent = e instanceof Error ? e.message : 'Import failed.';
                vaultImportError?.classList.remove('hidden');
            }
        });

        vaultImportReplaceBtn?.addEventListener('click', async () => {
            const pw = (vaultImportPassword?.value || '').trim();
            vaultImportError?.classList.add('hidden');
            if (!pw) { vaultImportError.textContent = 'Enter password for vault file.'; vaultImportError?.classList.remove('hidden'); return; }
            if (!pendingImportData) return;
            try {
                await secureStorage.replaceVaultWithImport(pendingImportData, pw);
                hideImportForm();
                refreshVaultModalState();
            } catch (e) {
                vaultImportError.textContent = e instanceof Error ? e.message : 'Replace failed.';
                vaultImportError?.classList.remove('hidden');
            }
        });

        createAnotherLink?.addEventListener('click', (e) => {
            e.preventDefault();
            hideUnlock();
            createPanel?.classList.remove('hidden');
            if (createName) createName.value = '';
        });

        templatesUnlockLink?.addEventListener('click', (e) => {
            e.preventDefault();
            this.hideTemplatesModal();
            showVaultModal('unlock');
        });
    }

    updateVaultUI() {
        const btn = document.getElementById('btn-vault');
        const label = document.getElementById('btn-vault-label');
        if (!btn || !label) return;
        if (!secureStorage.hasVault()) label.textContent = 'Create vault';
        else if (secureStorage.isUnlocked()) label.textContent = 'Lock ' + (secureStorage.getActiveVaultName() || 'vault');
        else label.textContent = 'Unlock vault';
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
            const baseName = (this.fileName || '').replace(/\.pdf$/i, '').trim();
            const hasRealFilename = baseName && baseName !== 'document';
            const exportName = hasRealFilename ? `${baseName}-edited.pdf` : 'document.pdf';
            const ctx = this.buildEmailContext(exportName);
            if (sel.value === '__doc__' && this.signingFlowMeta?.emailTemplate) {
                const filled = emailTemplates.fill(this.signingFlowMeta.emailTemplate, ctx);
                subjectEl.value = filled.subject;
                bodyEl.value = filled.body;
                return;
            }
            const t = emailTemplates.getById(sel.value);
            if (!t) return;
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
        const locked = secureStorage.hasVault() && !secureStorage.isUnlocked();
        const banner = document.getElementById('templates-vault-locked');
        const addBtn = document.getElementById('tpl-add');
        const importInput = document.getElementById('tpl-import-input');
        const importLabel = document.querySelector('label[for="tpl-import-input"]');
        if (banner) banner.classList.toggle('hidden', !locked);
        if (addBtn) addBtn.disabled = locked;
        if (importInput) importInput.disabled = locked;
        if (importLabel) importLabel.classList.toggle('disabled', locked);
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
        const locked = secureStorage.hasVault() && !secureStorage.isUnlocked();

        list.innerHTML = store
            .map(
                (t) =>
                    `<li class="${t.id === defaultId ? 'default-item' : ''}" data-id="${escapeHtml(t.id)}">
  <div class="tpl-info">
    <span class="tpl-name">${escapeHtml(t.name)}</span>
    <div class="tpl-meta">${escapeHtml(t.subject || '(no subject)')}</div>
  </div>
  <div class="tpl-actions">
    <button type="button" class="tpl-btn edit-btn" data-id="${escapeHtml(t.id)}" ${locked ? 'disabled' : ''}>Edit</button>${!t.builtin ? `<button type="button" class="tpl-btn delete-btn" data-id="${escapeHtml(t.id)}" ${locked ? 'disabled' : ''}>Delete</button>` : ''}
    ${t.id !== defaultId ? `<button type="button" class="tpl-btn set-default" data-id="${escapeHtml(t.id)}" ${locked ? 'disabled' : ''}>Set default</button>` : ''}
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

    async saveTemplateEdit() {
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

        try {
            if (id) {
                await emailTemplates.update(id, { name, subject, body });
            } else {
                await emailTemplates.add({ name, subject, body });
            }
            this.hideTemplateEditModal();
            this.renderTemplatesList();
            this.refreshSendModal();
        } catch (e) {
            alert('Failed to save template: ' + (e instanceof Error ? e.message : String(e)));
        }
    }

    hideTemplateEditModal() {
        this.templateEditModal.classList.add('hidden');
    }

    async deleteTemplate(id) {
        if (!id || !confirm('Delete this template?')) return;
        try {
            await emailTemplates.remove(id);
            this.renderTemplatesList();
            this.refreshSendModal();
        } catch (e) {
            alert('Failed to delete template: ' + (e instanceof Error ? e.message : String(e)));
        }
    }

    async setDefaultTemplate(id) {
        if (!id) return;
        try {
            await emailTemplates.setDefault(id);
            this.renderTemplatesList();
            this.refreshSendModal();
        } catch (e) {
            alert('Failed to set default: ' + (e instanceof Error ? e.message : String(e)));
        }
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
                const { imported, errors } = await emailTemplates.importJson(text, { replace });
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
        const emailSection = document.getElementById('bulk-fill-email-section');
        const emailColumnSelect = document.getElementById('bulk-fill-email-column');
        const emailTemplateSelect = document.getElementById('bulk-fill-email-template');
        const bulkFillRowsDiv = document.getElementById('bulk-fill-rows');
        const bulkFillRowsList = document.getElementById('bulk-fill-rows-list');
        const sendAllBtn = document.getElementById('bulk-fill-send-all');
        const sendAllStatus = document.getElementById('bulk-fill-send-all-status');
        const progressDiv = document.getElementById('bulk-fill-progress');
        const progressFill = document.getElementById('bulk-fill-progress-fill');
        const progressText = document.getElementById('bulk-fill-progress-text');
        const templateStatus = document.getElementById('bulk-fill-template-status');

        let templateBytes = null;
        let csvText = null;
        let pdfFieldNames = [];
        let csvHeaders = [];
        const sentRowIndices = new Set();
        /** Per-row manually entered email when column(s) are missing (rowIndex -> email string) */
        const manualEmailByRow = {};

        const getFieldMappingFromDOM = () => {
            const fieldMapping = {};
            if (!mappingList) return fieldMapping;
            mappingList.querySelectorAll('select').forEach((select) => {
                const csvColumn = select.dataset.csvColumn;
                const pdfField = select.value;
                if (pdfField && pdfField !== '') fieldMapping[csvColumn] = pdfField;
            });
            return fieldMapping;
        };

        const getSelectedEmailColumns = () => {
            if (!emailColumnSelect || !emailColumnSelect.multiple) return [];
            return Array.from(emailColumnSelect.selectedOptions).map((o) => o.value).filter(Boolean);
        };

        /** All emails from selected columns + manual, for display/validation. */
        const getEmailsForRow = (rowData, index) => {
            const cols = getSelectedEmailColumns();
            const fromCols = cols.flatMap((c) => (rowData[c] || '').trim()).filter(Boolean);
            const manual = (manualEmailByRow[index] || '').trim();
            const combined = manual ? [...fromCols, manual] : fromCols;
            return [...new Set(combined)];
        };

        /** Email(s) to use when sending: only the first signer (sequential flow). Uses first selected column, then manual if needed. */
        const getSendToEmailsForRow = (rowData, index) => {
            const cols = getSelectedEmailColumns();
            const manual = (manualEmailByRow[index] || '').trim();
            if (manual) return [manual];
            if (cols.length === 0) return [];
            const firstCol = cols[0];
            const first = (rowData[firstCol] || '').trim();
            return first ? [first] : [];
        };

        /** Selected email columns that have no value for this row (missing signer slot). */
        const getMissingEmailColumnsForRow = (rowData) => {
            const cols = getSelectedEmailColumns();
            if (cols.length <= 1) return [];
            return cols.filter((c) => !(rowData[c] || '').trim());
        };

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

        // Update mapping when fields are available; also show email section and rows list
        this.updateBulkFillMapping = () => {
            if (pdfFieldNames.length === 0 || csvHeaders.length === 0) {
                mappingDiv.classList.add('hidden');
                if (emailSection) emailSection.classList.add('hidden');
                if (bulkFillRowsDiv) bulkFillRowsDiv.classList.add('hidden');
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

            // Email section: multi-select column(s) + template dropdowns
            if (emailSection) {
                emailSection.classList.remove('hidden');
                if (emailColumnSelect) {
                    const currentSelected = getSelectedEmailColumns();
                    emailColumnSelect.innerHTML = csvHeaders.map((h) => `<option value="${escapeHtml(h)}" ${currentSelected.includes(h) ? 'selected' : ''}>${escapeHtml(h)}</option>`).join('');
                }
                if (emailTemplateSelect) {
                    const templates = emailTemplates.getTemplates();
                    const defaultTpl = emailTemplates.getDefault();
                    const currentTpl = emailTemplateSelect.value;
                    emailTemplateSelect.innerHTML = templates
                        .map((t) => `<option value="${t.id}" ${currentTpl === t.id || (!currentTpl && t.id === defaultTpl.id) ? 'selected' : ''}>${escapeHtml(t.name)}</option>`)
                        .join('');
                }
            }

            // Rows list: one row per CSV row with optional manual email, Download and Send email
            if (bulkFillRowsDiv && bulkFillRowsList && csvText) {
                // Preserve manual email inputs before rebuilding
                bulkFillRowsList.querySelectorAll('.bulk-fill-row-email-input').forEach((input) => {
                    const idx = parseInt(input.dataset.rowIndex, 10);
                    if (!Number.isNaN(idx)) manualEmailByRow[idx] = (input.value || '').trim();
                });
                const rows = this.bulkFillHandler.parseCSV(csvText);
                const filenameTemplate = filenameInput?.value || 'document-{{row}}.pdf';
                bulkFillRowsDiv.classList.remove('hidden');
                bulkFillRowsList.innerHTML = '';

                rows.forEach((rowData, index) => {
                    const firstCol = csvHeaders[0];
                    const label = firstCol ? (rowData[firstCol] || '').trim() || `Row ${index + 1}` : `Row ${index + 1}`;
                    const emails = getEmailsForRow(rowData, index);
                    const hasEmail = emails.length > 0;
                    const emailDisplay = emails.length > 0 ? emails.join(', ') : 'No email';
                    const isSent = sentRowIndices.has(index);
                    const manualVal = manualEmailByRow[index] || '';
                    const rowEl = document.createElement('div');
                    rowEl.className = 'bulk-fill-row-item' + (isSent ? ' bulk-fill-row-sent' : '');
                    rowEl.dataset.rowIndex = String(index);
                    rowEl.innerHTML = `
                        <span class="bulk-fill-row-label" title="${escapeHtml(label)}">${escapeHtml(String(label).slice(0, 50))}${String(label).length > 50 ? '…' : ''}${hasEmail ? ' — ' + escapeHtml(emailDisplay) : ''}</span>
                        <input type="email" class="bulk-fill-row-email-input" data-row-index="${index}" placeholder="Add email if missing" value="${escapeHtml(manualVal)}" title="Add or override email for this row">
                        <div class="bulk-fill-row-actions">
                            ${isSent ? '<span class="bulk-fill-row-sent-badge" aria-label="Email opened for this row">✓ Sent</span>' : ''}
                            <button type="button" class="btn btn-secondary bulk-fill-row-download" data-row-index="${index}" title="Download filled PDF">Download</button>
                            <button type="button" class="btn btn-primary bulk-fill-row-send" data-row-index="${index}" title="Download PDF and open email" ${!hasEmail ? 'disabled' : ''}>Send email</button>
                        </div>
                    `;
                    bulkFillRowsList.appendChild(rowEl);
                });

                // Enable Send all when at least one row has an email
                const rowsWithEmail = rows.filter((r, i) => getEmailsForRow(r, i).length > 0);
                if (sendAllBtn) {
                    sendAllBtn.disabled = rowsWithEmail.length === 0;
                }
                if (sendAllStatus) {
                    sendAllStatus.classList.add('hidden');
                    sendAllStatus.textContent = '';
                }
            }
        };

        const markRowSentInDOM = (index) => {
            const rowEl = bulkFillRowsList?.querySelector(`.bulk-fill-row-item[data-row-index="${index}"]`);
            if (!rowEl || rowEl.querySelector('.bulk-fill-row-sent-badge')) return;
            const actions = rowEl.querySelector('.bulk-fill-row-actions');
            if (actions) {
                const badge = document.createElement('span');
                badge.className = 'bulk-fill-row-sent-badge';
                badge.setAttribute('aria-label', 'Email opened for this row');
                badge.textContent = '✓ Sent';
                actions.insertBefore(badge, actions.firstChild);
            }
            rowEl.classList.add('bulk-fill-row-sent');
        };

        // When user types in per-row email input, update manualEmailByRow and enable/disable Send for that row
        bulkFillRowsList?.addEventListener('input', (e) => {
            const input = e.target.closest('.bulk-fill-row-email-input');
            if (!input || !csvText) return;
            const index = parseInt(input.dataset.rowIndex, 10);
            if (Number.isNaN(index)) return;
            manualEmailByRow[index] = (input.value || '').trim();
            const rows = this.bulkFillHandler.parseCSV(csvText);
            const rowData = rows[index];
            if (rowData == null) return;
            const emails = getEmailsForRow(rowData, index);
            const rowEl = input.closest('.bulk-fill-row-item');
            const sendBtn = rowEl?.querySelector('.bulk-fill-row-send');
            if (sendBtn) sendBtn.disabled = emails.length === 0;
            const anyWithEmail = rows.some((r, i) => getEmailsForRow(r, i).length > 0);
            if (sendAllBtn) sendAllBtn.disabled = !anyWithEmail;
        });

        // Delegated handlers for per-row Download and Send email
        bulkFillRowsList?.addEventListener('click', async (e) => {
            const btn = e.target.closest('.bulk-fill-row-download, .bulk-fill-row-send');
            if (!btn || !templateBytes || !csvText) return;
            const index = parseInt(btn.dataset.rowIndex, 10);
            if (Number.isNaN(index)) return;

            const rows = this.bulkFillHandler.parseCSV(csvText);
            if (index < 0 || index >= rows.length) return;
            const rowData = rows[index];
            const fieldMapping = getFieldMappingFromDOM();
            const filenameTemplate = filenameInput?.value || 'document-{{row}}.pdf';

            const doRow = async () => {
                const filledBytes = await this.bulkFillHandler.fillPDF(templateBytes, rowData, fieldMapping);
                const filename = this.bulkFillHandler.generateFilename(filenameTemplate, rowData, index);
                this.bulkFillHandler.downloadPDF(filledBytes, filename);
                return { filledBytes, filename };
            };

            if (btn.classList.contains('bulk-fill-row-download')) {
                try {
                    await doRow();
                } catch (err) {
                    console.error('Bulk fill row download error:', err);
                    alert('Error generating PDF: ' + err.message);
                }
                return;
            }

            if (btn.classList.contains('bulk-fill-row-send')) {
                const emails = getSendToEmailsForRow(rowData, index);
                if (emails.length === 0) {
                    alert('Please enter an email for this row before sending. Add an email in the row or select email column(s) above.');
                    return;
                }
                const missingCols = getMissingEmailColumnsForRow(rowData);
                if (missingCols.length > 0) {
                    const ok = confirm(
                        `This row has no email for: ${missingCols.join(', ')}. Send anyway with ${emails.join(', ')}?`
                    );
                    if (!ok) return;
                }
                const tplId = emailTemplateSelect?.value || '';
                const tpl = emailTemplates.getById(tplId) || emailTemplates.getDefault();
                try {
                    const { filename } = await doRow();
                    const pageCount = await this.bulkFillHandler.getPageCount(templateBytes);
                    const editLink =
                        typeof window !== 'undefined' && window.location
                            ? (window.location.origin + window.location.pathname).replace(/\/$/, '')
                            : '';
                    const ctx = {
                        filename,
                        date: new Date().toLocaleString(),
                        signatureSummary: '—',
                        signerNames: '—',
                        pageCount,
                        documentHash: '',
                        editLink
                    };
                    const filled = emailTemplates.fill(tpl, ctx);
                    const toPart = encodeURIComponent(emails.join(','));
                    const mailto = `mailto:${toPart}?subject=${encodeURIComponent(filled.subject)}&body=${encodeURIComponent(filled.body)}`;
                    window.location.href = mailto;
                    sentRowIndices.add(index);
                    markRowSentInDOM(index);
                } catch (err) {
                    console.error('Bulk fill send email error:', err);
                    alert('Error preparing email: ' + err.message);
                }
            }
        });

        // Send all: open mailto for each row that has at least one email (from columns or manual)
        const SEND_ALL_DELAY_MS = 3000;
        sendAllBtn?.addEventListener('click', async () => {
            if (!templateBytes || !csvText) return;
            const rows = this.bulkFillHandler.parseCSV(csvText);
            const indicesToSend = rows
                .map((r, i) => (getEmailsForRow(r, i).length > 0 ? i : -1))
                .filter((i) => i >= 0);
            const missingCount = rows.length - indicesToSend.length;
            if (indicesToSend.length === 0) {
                alert('No rows have an email. Select email column(s) and/or add an email for each row.');
                return;
            }
            if (missingCount > 0) {
                const ok = confirm(`${missingCount} row(s) have no email and will be skipped. Continue with ${indicesToSend.length} row(s)?`);
                if (!ok) return;
            }
            const rowsWithMissingSigner = indicesToSend
                .map((i) => ({ index: i, rowData: rows[i], missing: getMissingEmailColumnsForRow(rows[i]) }))
                .filter((x) => x.missing.length > 0);
            if (rowsWithMissingSigner.length > 0) {
                const summary = rowsWithMissingSigner
                    .slice(0, 5)
                    .map((x) => `Row ${x.index + 1} (missing: ${x.missing.join(', ')})`)
                    .join('; ');
                const more = rowsWithMissingSigner.length > 5 ? ` … and ${rowsWithMissingSigner.length - 5} more` : '';
                const ok = confirm(
                    `${rowsWithMissingSigner.length} row(s) have missing signer email(s): ${summary}${more}. Send anyway with the email(s) they have?`
                );
                if (!ok) return;
            }

            const fieldMapping = getFieldMappingFromDOM();
            const filenameTemplate = filenameInput?.value || 'document-{{row}}.pdf';
            const tplId = emailTemplateSelect?.value || '';
            const tpl = emailTemplates.getById(tplId) || emailTemplates.getDefault();
            const pageCount = await this.bulkFillHandler.getPageCount(templateBytes);
            const editLink =
                typeof window !== 'undefined' && window.location
                    ? (window.location.origin + window.location.pathname).replace(/\/$/, '')
                    : '';

            sendAllBtn.disabled = true;
            if (sendAllStatus) {
                sendAllStatus.classList.remove('hidden');
            }

            for (let i = 0; i < indicesToSend.length; i++) {
                const index = indicesToSend[i];
                const rowData = rows[index];
                const emails = getSendToEmailsForRow(rowData, index);
                if (emails.length === 0) continue;

                if (sendAllStatus) {
                    sendAllStatus.textContent = `Opening email ${i + 1} of ${indicesToSend.length}…`;
                }
                try {
                    const filledBytes = await this.bulkFillHandler.fillPDF(templateBytes, rowData, fieldMapping);
                    const filename = this.bulkFillHandler.generateFilename(filenameTemplate, rowData, index);
                    this.bulkFillHandler.downloadPDF(filledBytes, filename);
                    const ctx = {
                        filename,
                        date: new Date().toLocaleString(),
                        signatureSummary: '—',
                        signerNames: '—',
                        pageCount,
                        documentHash: '',
                        editLink
                    };
                    const filled = emailTemplates.fill(tpl, ctx);
                    const toPart = encodeURIComponent(emails.join(','));
                    const mailto = `mailto:${toPart}?subject=${encodeURIComponent(filled.subject)}&body=${encodeURIComponent(filled.body)}`;
                    window.location.href = mailto;
                    sentRowIndices.add(index);
                    markRowSentInDOM(index);
                } catch (err) {
                    console.error('Send all error for row ' + (index + 1), err);
                    if (sendAllStatus) sendAllStatus.textContent = `Error on row ${index + 1}: ${err.message}`;
                    break;
                }
                if (i < indicesToSend.length - 1) {
                    await new Promise((r) => setTimeout(r, SEND_ALL_DELAY_MS));
                }
            }

            if (sendAllStatus) {
                sendAllStatus.textContent = indicesToSend.length > 0 ? `Opened ${indicesToSend.length} email(s). Attach each downloaded PDF and send.` : '';
            }
            sendAllBtn.disabled = false;
        });

        // When email column changes, re-enable/disable Send email buttons and refresh row labels
        emailColumnSelect?.addEventListener('change', () => this.updateBulkFillMapping());

        // Expose minimal internals so showBulkFillModal() can set default template.
        this._bulkFillModalInternal = {
            reset: () => {
                templateBytes = null;
                csvText = null;
                pdfFieldNames = [];
                csvHeaders = [];
                sentRowIndices.clear();
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

    setupExpectedSignersModal() {
        const modal = document.getElementById('expected-signers-modal');
        const closeBtn = document.getElementById('expected-signers-modal-close');
        const cancelBtn = document.getElementById('expected-signers-cancel');
        const saveBtn = document.getElementById('expected-signers-save');
        const listEl = document.getElementById('expected-signers-list');
        const addBtn = document.getElementById('expected-signers-add');
        const completionToInput = document.getElementById('expected-signers-completion-to');
        const completionCcInput = document.getElementById('expected-signers-completion-cc');
        const completionBccInput = document.getElementById('expected-signers-completion-bcc');
        const setExpectedBtn = document.getElementById('signing-flow-set-expected');

        const countSignaturePlacements = () => {
            let n = 0;
            for (const page of this.canvasManager.getAllAnnotations?.() || []) {
                for (const ann of page.annotations || []) {
                    if (ann.type === 'signature') n++;
                }
            }
            return n;
        };

        const renderList = (entries) => {
            if (!listEl) return;
            listEl.innerHTML = '';
            (entries || []).forEach((entry, i) => {
                const row = document.createElement('div');
                row.className = 'expected-signers-row';
                row.setAttribute('role', 'listitem');
                row.innerHTML = `
                    <span class="expected-signers-ordinal">${i + 1}.</span>
                    <input type="text" class="expected-signers-name send-input" placeholder="Signature field (e.g. Tenant, Landlord)" value="${escapeHtml(entry.name || '')}" data-index="${i}" title="Must match the field label on the document">
                    <input type="email" class="expected-signers-email send-input" placeholder="Email" value="${escapeHtml(entry.email || '')}" data-index="${i}">
                    <button type="button" class="btn btn-secondary expected-signers-remove" data-index="${i}" title="Remove signer">Remove</button>
                `;
                listEl.appendChild(row);
            });
        };

        const getEntriesFromList = () => {
            if (!listEl) return [];
            const entries = [];
            listEl.querySelectorAll('.expected-signers-row').forEach((row, i) => {
                const nameInput = row.querySelector('.expected-signers-name');
                const emailInput = row.querySelector('.expected-signers-email');
                const name = (nameInput?.value || '').trim();
                const email = (emailInput?.value || '').trim();
                entries.push({ name, email: email || undefined, order: i + 1 });
            });
            return entries;
        };

        const show = () => {
            if (!this.signingFlowMeta) this.signingFlowMeta = { signers: [], expectedSigners: [] };
            let entries = this.signingFlowMeta.expectedSigners || [];
            if (entries.length === 0) {
                const n = countSignaturePlacements();
                if (n > 0) {
                    entries = Array.from({ length: n }, (_, i) => ({ name: `Signer ${i + 1}`, email: undefined, order: i + 1 }));
                }
            }
            renderList(entries);
            const meta = this.signingFlowMeta;
            if (completionToInput) {
                if (meta.originalSenderEmail) completionToInput.value = meta.originalSenderEmail;
                else if (Array.isArray(meta.completionToEmails) && meta.completionToEmails.length > 0) completionToInput.value = meta.completionToEmails.join(', ');
                else completionToInput.value = '';
            }
            if (completionCcInput) completionCcInput.value = (meta?.completionCcEmails || []).join(', ');
            if (completionBccInput) completionBccInput.value = (meta?.completionBccEmails || []).join(', ');
            modal?.classList.remove('hidden');
        };

        const hide = () => modal?.classList.add('hidden');

        addBtn?.addEventListener('click', () => {
            const entries = getEntriesFromList();
            entries.push({ name: '', email: undefined, order: entries.length + 1 });
            renderList(entries);
        });

        listEl?.addEventListener('click', (e) => {
            const removeBtn = e.target.closest('.expected-signers-remove');
            if (!removeBtn) return;
            const row = removeBtn.closest('.expected-signers-row');
            if (!row) return;
            const entries = getEntriesFromList();
            const idx = Array.from(listEl.querySelectorAll('.expected-signers-row')).indexOf(row);
            if (idx < 0 || idx >= entries.length) return;
            entries.splice(idx, 1);
            entries.forEach((ent, i) => { ent.order = i + 1; });
            renderList(entries);
        });

        setExpectedBtn?.addEventListener('click', () => show());
        closeBtn?.addEventListener('click', () => hide());
        cancelBtn?.addEventListener('click', () => hide());
        saveBtn?.addEventListener('click', () => {
            const raw = getEntriesFromList();
            const expectedSigners = raw.filter((e) => (e.name || '').trim()).map((e, i) => ({ ...e, name: (e.name || '').trim(), order: i + 1 }));
            const completionToText = (completionToInput?.value ?? '').trim();
            const completionToEmails = completionToText
                ? completionToText.split(',').map((e) => e.trim()).filter(Boolean)
                : [];
            const completionCcText = (completionCcInput?.value ?? '').trim();
            const completionCcEmails = completionCcText ? completionCcText.split(',').map((e) => e.trim()).filter(Boolean) : [];
            const completionBccText = (completionBccInput?.value ?? '').trim();
            const completionBccEmails = completionBccText ? completionBccText.split(',').map((e) => e.trim()).filter(Boolean) : [];
            if (!this.signingFlowMeta) this.signingFlowMeta = { signers: [], expectedSigners: [] };
            this.signingFlowMeta.expectedSigners = expectedSigners;
            if (completionToEmails.length === 1) {
                this.signingFlowMeta.originalSenderEmail = completionToEmails[0];
                this.signingFlowMeta.completionToEmails = completionToEmails;
            } else if (completionToEmails.length > 1) {
                this.signingFlowMeta.originalSenderEmail = undefined;
                this.signingFlowMeta.completionToEmails = completionToEmails;
            } else {
                this.signingFlowMeta.originalSenderEmail = undefined;
                this.signingFlowMeta.completionToEmails = undefined;
            }
            this.signingFlowMeta.completionCcEmails = completionCcEmails.length > 0 ? completionCcEmails : undefined;
            this.signingFlowMeta.completionBccEmails = completionBccEmails.length > 0 ? completionBccEmails : undefined;
            this.updateSigningFlowBanner();
            hide();
        });
        modal?.addEventListener('click', (e) => {
            if (e.target === modal) hide();
        });
    }

    async showBulkFillModal() {
        // Reset form
        document.getElementById('bulk-fill-template').value = '';
        document.getElementById('bulk-fill-csv').value = '';
        document.getElementById('bulk-fill-filename').value = 'document-{{row}}.pdf';
        document.getElementById('bulk-fill-mapping').classList.add('hidden');
        document.getElementById('bulk-fill-email-section')?.classList.add('hidden');
        document.getElementById('bulk-fill-rows')?.classList.add('hidden');
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
        const sigOpen = !this.signatureModal.classList.contains('hidden');
        if (sigOpen && this.signaturePad?.mode === 'draw') {
            this.btnUndo.disabled = !this.signaturePad.canUndo();
            this.btnRedo.disabled = !this.signaturePad.canRedo();
            return;
        }
        this.btnUndo.disabled = !this.canvasManager.canUndo();
        this.btnRedo.disabled = !this.canvasManager.canRedo();
    }

    /**
     * Show or hide the signing-flow banner based on signingFlowMeta and current canvas signatures.
     * Ensures a smooth multi-signer flow: open → sign → send to next → repeat until complete.
     */
    updateSigningFlowBanner() {
        const banner = document.getElementById('signing-flow-banner');
        const textEl = document.getElementById('signing-flow-banner-text');
        if (!banner || !textEl) return;

        if (!this.signingFlowMeta) {
            banner.classList.add('hidden');
            return;
        }

        const fromMeta = this.signingFlowMeta.signers || [];
        const fromCanvas = [];
        for (const page of this.canvasManager.getAllAnnotations?.() || []) {
            for (const ann of page.annotations || []) {
                if (ann.type === 'signature' && ann.object?._signatureMeta?.signerName) {
                    fromCanvas.push({
                        name: ann.object._signatureMeta.signerName,
                        timestamp: ann.object._signatureMeta.timestamp || ''
                    });
                }
            }
        }
        const seen = new Set();
        const allSigners = [];
        for (const s of fromMeta) {
            if (s.name && !seen.has(s.name)) {
                seen.add(s.name);
                allSigners.push(s.name);
            }
        }
        for (const s of fromCanvas) {
            if (s.name && !seen.has(s.name)) {
                seen.add(s.name);
                allSigners.push(s.name);
            }
        }

        const expected = this.signingFlowMeta.expectedSigners || [];
        // Only count expected signers who have an email (flow completes when each populated signer has signed)
        const expectedWithEmail = expected.filter((e) => (e.email || '').trim());
        const totalExpected = expectedWithEmail.length;
        const countSoFar = allSigners.length;
        const completionTo = this.signingFlowMeta?.originalSenderEmail
            ? [this.signingFlowMeta.originalSenderEmail]
            : (this.signingFlowMeta?.completionToEmails || []);
        const hasCompletionTo = completionTo.length > 0;
        const allSigned = totalExpected > 0 && countSoFar >= totalExpected;
        const lastSigner = totalExpected > 0 && countSoFar === totalExpected - 1;
        const lockedFields = this.signingFlowMeta.lockedSignatureFields || [];
        const hasLockedFields = lockedFields.length > 0 || this.signingFlowMeta.documentStage === 'signed';

        let msg = 'This document is part of a signing flow. ';
        if (hasLockedFields) {
            msg += 'Signed fields are locked and cannot be modified by another participant. ';
        }
        if (allSigners.length > 0) {
            msg += `Signatures so far: ${allSigners.join(', ')}. `;
        }
        if (totalExpected > 0) {
            msg += `${countSoFar} of ${totalExpected} signers. `;
        }
        if (allSigned) {
            if (hasCompletionTo) {
                msg += `All signers have signed. Send the completed document to: ${completionTo.join(', ')}.`;
            } else {
                msg += 'All signers have signed. Use Send to email the completed document.';
            }
        } else if (lastSigner && hasCompletionTo) {
            msg += `You're the last signer. Sign below and send the completed document to: ${completionTo.join(', ')}.`;
        } else if (lastSigner) {
            msg += "You're the last signer. Sign below and send the completed document back to the original sender.";
        } else {
            msg += 'Sign below and send to the next person to get a complete document.';
        }

        textEl.textContent = msg;
        banner.classList.remove('hidden');
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
