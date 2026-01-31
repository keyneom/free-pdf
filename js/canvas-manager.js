/**
 * Canvas Manager - Manages Fabric.js canvas overlays for PDF annotation
 */

export class CanvasManager {
    constructor() {
        this.canvases = new Map(); // pageId -> fabric.Canvas
        this.activeCanvas = null;
        this.activePageId = null;
        this.activeTool = 'select';
        this.fillMode = false; // Whether in fill mode (form filling/signing)
        this.history = new Map(); // pageId -> {undoStack, redoStack}
        this.currentScale = 1.0;
        this._restoringPages = new Set(); // pageId currently being restored from history
        /** Optional callback when undo/redo stacks change (e.g. to update toolbar buttons) */
        this._onHistoryChange = null;

        // Tool settings
        this.settings = {
            textColor: '#000000',
            fontSize: 16,
            fontFamily: 'Arial',
            fontWeight: 'normal',
            fontStyle: 'normal',
            textAlign: 'left',
            strokeColor: '#000000',
            strokeWidth: 2,
            whiteoutColor: '#ffffff',
            highlightColor: '#fff59d',
            highlightOpacity: 0.55,
            shapeFill: 'transparent',
            shapeOpacity: 1.0,
            stampText: 'APPROVED'
        };

        // Signature data
        this.signatureImage = null;
        this.signatureMeta = null;

        // Signature field data
        this.signatureFieldLabel = null;

        // Image insert
        this.pendingImageDataUrl = null;

        // Temp drawing state for drag-based tools
        this._temp = null; // { kind, startX, startY, obj, extra?, moveHandler? }
    }

    /**
     * Create a Fabric.js canvas overlay for a page
     * @param {HTMLElement} container - Container element
     * @param {number} width - Canvas width
     * @param {number} height - Canvas height
     * @param {string} pageId - Stable view page id
     * @returns {fabric.Canvas}
     */
    createCanvas(container, width, height, pageId) {
        const safeId = String(pageId).replace(/[^a-zA-Z0-9_-]/g, '_');
        // Create canvas element
        const canvasEl = document.createElement('canvas');
        canvasEl.id = `annotation-canvas-${safeId}`;
        canvasEl.width = width;
        canvasEl.height = height;
        container.appendChild(canvasEl);

        // Create Fabric.js canvas
        const fabricCanvas = new fabric.Canvas(canvasEl, {
            width: width,
            height: height,
            selection: true,
            preserveObjectStacking: true
        });

        // Store reference
        this.canvases.set(pageId, fabricCanvas);
        this.history.set(pageId, { undoStack: [], redoStack: [] });

        // Set up event listeners
        this.setupCanvasEvents(fabricCanvas, pageId);

        // Seed initial "blank" state so the first action can be undone
        this.saveState(pageId);

        // Default active canvas/page (first created)
        if (this.activeCanvas === null) {
            this.activeCanvas = fabricCanvas;
            this.activePageId = pageId;
        }

        return fabricCanvas;
    }

    canUndo() {
        const pageId = this.activePageId;
        const history = pageId != null ? this.history.get(pageId) : null;
        return !!(history && history.undoStack.length > 1);
    }

    canRedo() {
        const pageId = this.activePageId;
        const history = pageId != null ? this.history.get(pageId) : null;
        return !!(history && history.redoStack.length > 0);
    }

    /**
     * Set the active page (e.g. when user navigates). Undo/redo will apply to this page.
     * @param {string} pageId - View page id
     */
    setActivePage(pageId) {
        const canvas = this.canvases.get(pageId);
        if (canvas) {
            this.activeCanvas = canvas;
            this.activePageId = pageId;
        }
    }

    /**
     * Register callback to run when history changes (undo/redo stacks).
     * @param {() => void} callback
     */
    setOnHistoryChange(callback) {
        this._onHistoryChange = typeof callback === 'function' ? callback : null;
    }

    /**
     * Set up canvas event listeners
     * @param {fabric.Canvas} canvas - Fabric.js canvas
     * @param {string} pageId - View page id
     */
    setupCanvasEvents(canvas, pageId) {
        // Track modifications for undo/redo
        canvas.on('object:added', (e) => {
            if (this._restoringPages.has(pageId)) return;
            if (e?.target && !e.target._fromHistory) {
                this.saveState(pageId);
            }
        });

        canvas.on('object:modified', () => {
            if (this._restoringPages.has(pageId)) return;
            this.saveState(pageId);
        });

        canvas.on('object:removed', (e) => {
            if (this._restoringPages.has(pageId)) return;
            if (e?.target && !e.target._fromHistory) {
                this.saveState(pageId);
            }
        });

        // Click handler for tools
        canvas.on('mouse:down', (e) => {
            this.activeCanvas = canvas;
            this.activePageId = pageId;
            this.handleMouseDown(e, canvas, pageId);
        });

        canvas.on('mouse:up', (e) => {
            this.handleMouseUp(e, canvas, pageId);
        });

        // Selection events
        canvas.on('selection:created', () => {
            this.onSelectionChanged(canvas);
        });

        canvas.on('selection:updated', () => {
            this.onSelectionChanged(canvas);
        });

        canvas.on('selection:cleared', () => {
            this.onSelectionChanged(canvas);
        });
    }

    /**
     * Handle mouse down event based on active tool
     */
    handleMouseDown(e, canvas, pageId) {
        const pointer = canvas.getPointer(e.e);

        // In Fill mode, handle single-click on form fields
        if (this.fillMode && e.target) {
            const annotationType = e.target._annotationType;
            
            // Select the field to focus sidebar input
            if (['textfield', 'date', 'dropdown', 'checkbox', 'radio', 'signature-field'].includes(annotationType)) {
                canvas.setActiveObject(e.target);
                this.onSelectionChanged(canvas); // Force update toolbar
            }
            
            if (annotationType === 'textfield' || annotationType === 'date') {
                // Show inline HTML input overlay for direct editing
                this._showInlineTextEditor(e.target, canvas);
                return;
            }
            if (annotationType === 'dropdown') {
                // Show inline HTML select overlay
                this._showInlineDropdownEditor(e.target, canvas);
                return;
            }
            if (annotationType === 'checkbox' || annotationType === 'radio') {
                // Toggle directly on canvas
                if (annotationType === 'checkbox') {
                    this.toggleCheckbox(e.target, canvas);
                } else {
                    this.toggleRadio(e.target, canvas);
                }
                this._highlightField(e.target, canvas);
                window.dispatchEvent(new CustomEvent('field-updated'));
                return;
            }
            if (annotationType === 'signature-field') {
                // Dispatch to open signature modal
                window.dispatchEvent(new CustomEvent('form-field-selected', {
                    detail: { object: e.target, annotationType, pageId }
                }));
                this._highlightField(e.target, canvas);
                return;
            }
        }

        switch (this.activeTool) {
            case 'text':
                this.addTextBox(canvas, pointer.x, pointer.y);
                break;
            case 'whiteout':
                this.startWhiteout(canvas, pointer.x, pointer.y);
                break;
            case 'highlight':
                this.startHighlight(canvas, pointer.x, pointer.y);
                break;
            case 'underline':
                this.startLine(canvas, pointer.x, pointer.y, 'underline');
                break;
            case 'strike':
                this.startLine(canvas, pointer.x, pointer.y, 'strike');
                break;
            case 'rect':
                this.startRect(canvas, pointer.x, pointer.y);
                break;
            case 'ellipse':
                this.startEllipse(canvas, pointer.x, pointer.y);
                break;
            case 'arrow':
                this.startArrow(canvas, pointer.x, pointer.y);
                break;
            case 'note':
                this.addNote(canvas, pointer.x, pointer.y);
                break;
            case 'stamp':
                this.addStamp(canvas, pointer.x, pointer.y);
                break;
            case 'image':
                this.insertPendingImage(canvas, pointer.x, pointer.y);
                break;
            case 'textfield':
                this.addFormTextField(canvas, pointer.x, pointer.y);
                break;
            case 'checkbox':
                this.addFormCheckbox(canvas, pointer.x, pointer.y);
                break;
            case 'radio':
                this.addFormRadio(canvas, pointer.x, pointer.y);
                break;
            case 'dropdown':
                this.addFormDropdown(canvas, pointer.x, pointer.y);
                break;
            case 'date':
                this.addFormDateField(canvas, pointer.x, pointer.y);
                break;
            case 'signature':
                if (this.signatureImage) {
                    this.insertSignature(canvas, pointer.x, pointer.y);
                }
                break;
            case 'signature-field':
                this.insertSignatureField(canvas, pointer.x, pointer.y);
                break;
            case 'eraser':
                this.eraseAtPoint(canvas, pointer.x, pointer.y);
                break;
        }
    }

    /**
     * Handle mouse up event
     */
    handleMouseUp(e, canvas, pageId) {
        if (this.activeTool === 'whiteout' && this.tempWhiteout) {
            this.finishWhiteout(canvas);
        }
        if (this._temp) {
            this.finishTempTool(canvas);
        }
    }

    /**
     * Set the active tool
     * @param {string} tool - Tool name
     */
    /**
     * Apply fill-mode or edit-mode interactivity to an object
     * @param {fabric.Object} obj - The object to update
     * @param {boolean} fillMode - Whether we're in fill mode
     * @param {string} activeTool - Current active tool
     */
    _applyObjectInteractivity(obj, fillMode, activeTool) {
        const isFormField = obj._annotationType === 'textfield' ||
                           obj._annotationType === 'checkbox' ||
                           obj._annotationType === 'radio' ||
                           obj._annotationType === 'dropdown' ||
                           obj._annotationType === 'date';
        const isLockedFormField = isFormField && obj._fieldLocked;
        const isSignatureField = obj._annotationType === 'signature-field';
        const isLockedSignature = obj._annotationType === 'signature' && obj._signatureLocked;

        if (fillMode) {
            if (isLockedFormField || isLockedSignature) {
                // Locked form fields and signed fields: read-only, not editable
                obj.set({
                    selectable: false,
                    evented: true,
                    hasControls: false,
                    hasBorders: false,
                    lockMovementX: true,
                    lockMovementY: true,
                    lockScalingX: true,
                    lockScalingY: true,
                    lockRotation: true
                });
            } else if (isFormField || isSignatureField) {
                obj.set({
                    selectable: false,
                    evented: true,
                    hasControls: false,
                    hasBorders: false,
                    lockMovementX: true,
                    lockMovementY: true,
                    lockScalingX: true,
                    lockScalingY: true,
                    lockRotation: true
                });
            } else {
                obj.set({ selectable: false, evented: false });
            }
        } else if (activeTool === 'select') {
            if (isLockedSignature || isLockedFormField) {
                // Locked signatures/form fields: selectable (e.g. show in sidebar) but not moved/resized/deleted
                obj.set({
                    selectable: true,
                    evented: true,
                    hasControls: false,
                    hasBorders: true,
                    lockMovementX: true,
                    lockMovementY: true,
                    lockScalingX: true,
                    lockScalingY: true,
                    lockRotation: true
                });
            } else {
                obj.set({
                    selectable: true,
                    evented: true,
                    hasControls: true,
                    hasBorders: true,
                    lockMovementX: false,
                    lockMovementY: false,
                    lockScalingX: false,
                    lockScalingY: false,
                    lockRotation: false
                });
            }
        } else {
            obj.set({ selectable: false, evented: false });
        }
    }

    /**
     * Set fill mode (form filling/signing mode)
     * @param {boolean} enabled - Whether fill mode is enabled
     */
    setFillMode(enabled) {
        this.fillMode = enabled;

        this.canvases.forEach((canvas) => {
            canvas.forEachObject((obj) => {
                this._applyObjectInteractivity(obj, enabled, this.activeTool);
            });
            canvas.discardActiveObject();
            canvas.renderAll();
        });
    }

    setTool(tool) {
        this.activeTool = tool;

        // Update all canvases based on tool
        this.canvases.forEach((canvas) => {
            if (tool === 'draw') {
                canvas.isDrawingMode = true;
                canvas.freeDrawingBrush.color = this.settings.strokeColor;
                canvas.freeDrawingBrush.width = this.settings.strokeWidth;
            } else {
                canvas.isDrawingMode = false;
            }

            if (tool === 'select') {
                canvas.selection = !this.fillMode; // Disable box selection in fill mode
                canvas.forEachObject((obj) => {
                    this._applyObjectInteractivity(obj, this.fillMode, 'select');
                });
            } else if (tool !== 'draw') {
                canvas.selection = false;
                canvas.discardActiveObject();
                canvas.forEachObject((obj) => {
                    this._applyObjectInteractivity(obj, this.fillMode, tool);
                });
                canvas.renderAll();
            }
        });
    }

    setPendingImage(dataUrl) {
        this.pendingImageDataUrl = dataUrl || null;
    }

    /**
     * Add a text box at the specified position
     */
    addTextBox(canvas, x, y) {
        const textbox = new fabric.IText('Click to edit', {
            left: x,
            top: y,
            fontSize: this.settings.fontSize / this.currentScale,
            fontFamily: this.settings.fontFamily,
            fontWeight: this.settings.fontWeight,
            fontStyle: this.settings.fontStyle,
            textAlign: this.settings.textAlign,
            textBaseline: 'alphabetic',
            fill: this.settings.textColor,
            editable: true,
            selectable: true,
            _annotationType: 'text'
        });

        canvas.add(textbox);
        canvas.setActiveObject(textbox);
        textbox.enterEditing();
        textbox.selectAll();
        canvas.renderAll();

        // Switch back to select tool after adding
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="text"]')?.classList.remove('active');
    }

    /**
     * Start drawing a whiteout rectangle
     */
    startWhiteout(canvas, x, y) {
        this.whiteoutStart = { x, y };
        this.tempWhiteout = new fabric.Rect({
            left: x,
            top: y,
            width: 0,
            height: 0,
            fill: this.settings.whiteoutColor,
            selectable: false,
            evented: false,
            _annotationType: 'whiteout'
        });
        canvas.add(this.tempWhiteout);

        // Track mouse move to resize
        const moveHandler = (e) => {
            const pointer = canvas.getPointer(e.e);
            const width = pointer.x - this.whiteoutStart.x;
            const height = pointer.y - this.whiteoutStart.y;

            this.tempWhiteout.set({
                width: Math.abs(width),
                height: Math.abs(height),
                left: width < 0 ? pointer.x : this.whiteoutStart.x,
                top: height < 0 ? pointer.y : this.whiteoutStart.y
            });
            canvas.renderAll();
        };

        canvas.on('mouse:move', moveHandler);
        this._whiteoutMoveHandler = moveHandler;
    }

    /**
     * Finish whiteout rectangle
     */
    finishWhiteout(canvas) {
        if (this._whiteoutMoveHandler) {
            canvas.off('mouse:move', this._whiteoutMoveHandler);
            this._whiteoutMoveHandler = null;
        }

        if (this.tempWhiteout) {
            // If too small, remove it
            if (this.tempWhiteout.width < 5 || this.tempWhiteout.height < 5) {
                canvas.remove(this.tempWhiteout);
            } else {
                // Make it selectable
                this.tempWhiteout.set({
                    selectable: true,
                    evented: true
                });
            }
            this.tempWhiteout = null;
        }

        canvas.renderAll();
    }

    /**
     * Start drawing a highlight rectangle
     */
    startHighlight(canvas, x, y) {
        this._temp = { kind: 'highlight', startX: x, startY: y, obj: null, moveHandler: null };
        const rect = new fabric.Rect({
            left: x,
            top: y,
            width: 0,
            height: 0,
            fill: this.settings.highlightColor,
            opacity: this.settings.highlightOpacity,
            selectable: false,
            evented: false,
            _annotationType: 'highlight'
        });
        this._temp.obj = rect;
        canvas.add(rect);

        const moveHandler = (ev) => {
            const p = canvas.getPointer(ev.e);
            this.updateDragRect(rect, x, y, p.x, p.y);
            canvas.renderAll();
        };
        canvas.on('mouse:move', moveHandler);
        this._temp.moveHandler = moveHandler;
    }

    startRect(canvas, x, y) {
        this._temp = { kind: 'rect', startX: x, startY: y, obj: null, moveHandler: null };
        const rect = new fabric.Rect({
            left: x,
            top: y,
            width: 0,
            height: 0,
            fill: this.settings.shapeFill,
            opacity: this.settings.shapeOpacity,
            stroke: this.settings.strokeColor,
            strokeWidth: this.settings.strokeWidth / this.currentScale,
            selectable: false,
            evented: false,
            _annotationType: 'rect'
        });
        this._temp.obj = rect;
        canvas.add(rect);
        const moveHandler = (ev) => {
            const p = canvas.getPointer(ev.e);
            this.updateDragRect(rect, x, y, p.x, p.y);
            canvas.renderAll();
        };
        canvas.on('mouse:move', moveHandler);
        this._temp.moveHandler = moveHandler;
    }

    startEllipse(canvas, x, y) {
        this._temp = { kind: 'ellipse', startX: x, startY: y, obj: null, moveHandler: null };
        const ellipse = new fabric.Ellipse({
            left: x,
            top: y,
            rx: 0,
            ry: 0,
            fill: this.settings.shapeFill,
            opacity: this.settings.shapeOpacity,
            stroke: this.settings.strokeColor,
            strokeWidth: this.settings.strokeWidth / this.currentScale,
            selectable: false,
            evented: false,
            originX: 'center',
            originY: 'center',
            _annotationType: 'ellipse'
        });
        this._temp.obj = ellipse;
        canvas.add(ellipse);

        const moveHandler = (ev) => {
            const p = canvas.getPointer(ev.e);
            const left = Math.min(x, p.x);
            const top = Math.min(y, p.y);
            const w = Math.abs(p.x - x);
            const h = Math.abs(p.y - y);
            ellipse.set({
                left: left + w / 2,
                top: top + h / 2,
                rx: w / 2,
                ry: h / 2
            });
            canvas.renderAll();
        };
        canvas.on('mouse:move', moveHandler);
        this._temp.moveHandler = moveHandler;
    }

    startLine(canvas, x, y, kind) {
        this._temp = { kind, startX: x, startY: y, obj: null, moveHandler: null };
        const line = new fabric.Line([x, y, x, y], {
            stroke: this.settings.strokeColor,
            strokeWidth: this.settings.strokeWidth / this.currentScale,
            selectable: false,
            evented: false,
            _annotationType: kind
        });
        this._temp.obj = line;
        canvas.add(line);
        const moveHandler = (ev) => {
            const p = canvas.getPointer(ev.e);
            line.set({ x2: p.x, y2: kind === 'underline' || kind === 'strike' ? y : p.y });
            canvas.renderAll();
        };
        canvas.on('mouse:move', moveHandler);
        this._temp.moveHandler = moveHandler;
    }

    startArrow(canvas, x, y) {
        this._temp = { kind: 'arrow', startX: x, startY: y, obj: null, extra: null, moveHandler: null };
        const line = new fabric.Line([x, y, x, y], {
            stroke: this.settings.strokeColor,
            strokeWidth: this.settings.strokeWidth / this.currentScale,
            selectable: false,
            evented: false
        });
        const head = new fabric.Triangle({
            width: 10 / this.currentScale,
            height: 12 / this.currentScale,
            fill: this.settings.strokeColor,
            left: x,
            top: y,
            originX: 'center',
            originY: 'center',
            angle: 0,
            selectable: false,
            evented: false
        });
        const group = new fabric.Group([line, head], {
            selectable: false,
            evented: false,
            _annotationType: 'arrow'
        });
        this._temp.obj = group;
        this._temp.extra = { line, head };
        canvas.add(group);

        const moveHandler = (ev) => {
            const p = canvas.getPointer(ev.e);
            line.set({ x1: x, y1: y, x2: p.x, y2: p.y });
            const angle = (Math.atan2(p.y - y, p.x - x) * 180) / Math.PI + 90;
            head.set({ left: p.x, top: p.y, angle });
            group.addWithUpdate();
            canvas.renderAll();
        };
        canvas.on('mouse:move', moveHandler);
        this._temp.moveHandler = moveHandler;
    }

    updateDragRect(rect, x1, y1, x2, y2) {
        const left = Math.min(x1, x2);
        const top = Math.min(y1, y2);
        const width = Math.abs(x2 - x1);
        const height = Math.abs(y2 - y1);
        rect.set({ left, top, width, height });
    }

    finishTempTool(canvas) {
        const t = this._temp;
        if (!t) return;
        if (t.moveHandler) canvas.off('mouse:move', t.moveHandler);
        const obj = t.obj;
        if (obj) {
            const bounds = obj.getBoundingRect();
            if (bounds.width < 5 || bounds.height < 5) {
                canvas.remove(obj);
            } else {
                obj.set({ selectable: true, evented: true });
                if (t.kind === 'highlight') obj.sendToBack();
            }
        }
        this._temp = null;
        canvas.renderAll();
    }

    addNote(canvas, x, y) {
        const text = prompt('Note text:', '');
        if (text == null) return;
        const w = 140 / this.currentScale;
        const h = 70 / this.currentScale;
        const bg = new fabric.Rect({
            width: w,
            height: h,
            fill: '#fff9c4',
            stroke: '#f59e0b',
            strokeWidth: 1 / this.currentScale,
            rx: 6 / this.currentScale,
            ry: 6 / this.currentScale
        });
        const label = new fabric.Textbox(text || 'Note', {
            width: w - 12 / this.currentScale,
            fontSize: 12 / this.currentScale,
            fill: '#111827',
            textBaseline: 'alphabetic',
            left: 6 / this.currentScale,
            top: 6 / this.currentScale,
            editable: false
        });
        const group = new fabric.Group([bg, label], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'note',
            _noteText: text || ''
        });
        group.on('mousedblclick', () => {
            const next = prompt('Edit note:', group._noteText || '');
            if (next == null) return;
            group._noteText = next;
            label.text = next || 'Note';
            canvas.renderAll();
        });
        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="note"]')?.classList.remove('active');
    }

    addStamp(canvas, x, y) {
        const text = String(this.settings.stampText || 'APPROVED').toUpperCase();
        const w = 170 / this.currentScale;
        const h = 44 / this.currentScale;
        const border = new fabric.Rect({
            width: w,
            height: h,
            fill: 'transparent',
            stroke: '#dc2626',
            strokeWidth: 2 / this.currentScale,
            rx: 6 / this.currentScale,
            ry: 6 / this.currentScale
        });
        const label = new fabric.Text(text, {
            fontSize: 20 / this.currentScale,
            fill: '#dc2626',
            fontWeight: 'bold',
            textBaseline: 'alphabetic',
            originX: 'center',
            originY: 'center',
            left: w / 2,
            top: h / 2
        });
        const group = new fabric.Group([border, label], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'stamp',
            _stampText: text
        });
        group.on('mousedblclick', () => {
            const next = prompt('Stamp text:', group._stampText || text);
            if (next == null) return;
            group._stampText = String(next).toUpperCase();
            label.text = group._stampText;
            canvas.renderAll();
        });
        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="stamp"]')?.classList.remove('active');
    }

    insertPendingImage(canvas, x, y) {
        const src = this.pendingImageDataUrl;
        if (!src) {
            alert('Choose an image first.');
            return;
        }
        fabric.Image.fromURL(src, (img) => {
            const maxWidth = 240 / this.currentScale;
            const scale = maxWidth / img.width;
            img.set({
                left: x,
                top: y,
                scaleX: scale,
                scaleY: scale,
                selectable: true,
                _annotationType: 'image'
            });
            canvas.add(img);
            canvas.setActiveObject(img);
            canvas.renderAll();
        });
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="image"]')?.classList.remove('active');
    }

    eraseAtPoint(canvas, x, y) {
        const p = new fabric.Point(x, y);
        const objs = canvas.getObjects();
        let hit = null;
        for (let i = objs.length - 1; i >= 0; i--) {
            const obj = objs[i];
            if (obj.containsPoint && obj.containsPoint(p)) {
                hit = obj;
                break;
            }
        }
        if (hit) {
            canvas.remove(hit);
            canvas.renderAll();
        }
    }

    /**
     * Add a form text field
     */
    /**
     * Highlight a field briefly to show it was selected (Fill mode)
     */
    _highlightField(field, canvas) {
        // Store original stroke
        const originalStroke = field.stroke;
        const originalStrokeWidth = field.strokeWidth;
        
        // Apply highlight
        field.set({
            stroke: '#2563eb',
            strokeWidth: 2
        });
        canvas.renderAll();
        
        // Remove highlight after a short delay
        setTimeout(() => {
            field.set({
                stroke: originalStroke || '#d1d5db',
                strokeWidth: originalStrokeWidth || 1
            });
            canvas.renderAll();
        }, 300);
    }

    /**
     * Show an inline HTML input over a text/date field for direct editing (Fill mode)
     */
    _showInlineTextEditor(group, canvas) {
        // Remove any existing inline editor
        this._removeInlineEditor();
        
        const rect = group.getBoundingRect(true);
        const container = (canvas.lowerCanvas?.parentElement?.parentElement) || (canvas.wrapperEl?.parentElement);
        if (!container) return;
        
        const isDateField = group._annotationType === 'date';
        const currentValue = group._fieldValue || '';
        const objects = group.getObjects();
        const hasNewStructure = objects.length >= 3;
        const valueText = hasNewStructure ? objects[2] : objects[1];
        const fontSize = valueText ? Math.max(10, Math.round(valueText.fontSize * this.currentScale)) : 12;
        
        const input = document.createElement('input');
        input.type = isDateField ? 'date' : 'text';
        input.value = currentValue;
        input.className = 'inline-field-editor';
        input.style.cssText = `
            position: absolute;
            left: ${rect.left}px;
            top: ${rect.top}px;
            width: ${Math.max(60, rect.width - 4)}px;
            height: ${Math.max(24, rect.height - 4)}px;
            font-size: ${fontSize}px;
            padding: 2px 6px;
            border: 2px solid #2563eb;
            border-radius: 4px;
            outline: none;
            box-sizing: border-box;
            z-index: 1000;
        `;
        
        const commit = (val) => {
            this._removeInlineEditor();
            group._fieldValue = val || '';
            const placeholder = isDateField ? 'YYYY-MM-DD' : 'Enter value';
            if (valueText) {
                valueText.set({
                    text: val || placeholder,
                    fill: val ? '#000000' : '#9ca3af',
                    fontStyle: val ? 'normal' : 'italic'
                });
            }
            canvas.renderAll();
            window.dispatchEvent(new CustomEvent('field-updated'));
        };
        
        input.addEventListener('blur', () => commit(input.value));
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                input.blur();
            }
            if (e.key === 'Escape') {
                e.preventDefault();
                this._removeInlineEditor();
            }
        });
        
        container.appendChild(input);
        input.focus();
        input.select();
        this._inlineEditorEl = input;
    }

    /**
     * Show an inline HTML select over a dropdown field (Fill mode)
     */
    _showInlineDropdownEditor(group, canvas) {
        this._removeInlineEditor();
        
        const options = group._options || [];
        if (options.length === 0) {
            return;
        }
        
        const rect = group.getBoundingRect(true);
        const container = (canvas.lowerCanvas?.parentElement?.parentElement) || (canvas.wrapperEl?.parentElement);
        if (!container) return;
        
        const objects = group.getObjects();
        const hasNewStructure = objects.length >= 4;
        const valueText = hasNewStructure ? objects[2] : objects[1];
        
        const select = document.createElement('select');
        select.className = 'inline-field-editor';
        select.style.cssText = `
            position: absolute;
            left: ${rect.left}px;
            top: ${rect.top}px;
            width: ${Math.max(80, rect.width - 4)}px;
            height: ${Math.max(24, rect.height - 4)}px;
            font-size: 12px;
            padding: 2px 6px;
            border: 2px solid #2563eb;
            border-radius: 4px;
            outline: none;
            box-sizing: border-box;
            z-index: 1000;
            cursor: pointer;
        `;
        
        const emptyOpt = document.createElement('option');
        emptyOpt.value = '';
        emptyOpt.textContent = 'Select...';
        select.appendChild(emptyOpt);
        options.forEach(opt => {
            const o = document.createElement('option');
            o.value = opt;
            o.textContent = opt;
            if (opt === group._selectedOption) o.selected = true;
            select.appendChild(o);
        });
        
        const commit = (val) => {
            this._removeInlineEditor();
            group._selectedOption = val || '';
            if (valueText) {
                valueText.set({
                    text: val || 'Select...',
                    fill: val ? '#000000' : '#9ca3af',
                    fontStyle: val ? 'normal' : 'italic'
                });
            }
            canvas.renderAll();
            window.dispatchEvent(new CustomEvent('field-updated'));
        };
        
        select.addEventListener('change', () => commit(select.value));
        select.addEventListener('blur', () => commit(select.value));
        select.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                e.preventDefault();
                this._removeInlineEditor();
            }
        });
        
        container.appendChild(select);
        select.focus();
        this._inlineEditorEl = select;
    }

    /**
     * Remove the inline editor overlay if present
     */
    _removeInlineEditor() {
        if (this._inlineEditorEl && this._inlineEditorEl.parentNode) {
            this._inlineEditorEl.parentNode.removeChild(this._inlineEditorEl);
        }
        this._inlineEditorEl = null;
    }

    /**
     * Helper to fix text scaling in form field groups.
     * When a group is scaled, text gets distorted. This handler
     * resets text scale and repositions text to stay within bounds.
     */
    _setupFormFieldScaling(group, canvas, padding) {
        group.on('scaling', () => {
            // During scaling, we need to counteract the scale on text objects
            const objects = group.getObjects();
            const scaleX = group.scaleX;
            const scaleY = group.scaleY;
            
            // Reset text scale to maintain original size
            for (let i = 1; i < objects.length; i++) {
                const obj = objects[i];
                if (obj.type === 'text' || obj.type === 'i-text') {
                    // Counteract group scale to keep text unscaled
                    obj.set({
                        scaleX: 1 / scaleX,
                        scaleY: 1 / scaleY
                    });
                }
            }
        });
        
        group.on('scaled', () => {
            const objects = group.getObjects();
            const background = objects[0];
            const scaleX = group.scaleX;
            const scaleY = group.scaleY;
            
            // Get the new actual dimensions
            const newWidth = background.width * scaleX;
            const newHeight = background.height * scaleY;
            
            // Reset group scale to 1 and update background size directly
            group.set({ scaleX: 1, scaleY: 1 });
            background.set({ width: newWidth, height: newHeight });
            
            // Reset text scale and reposition
            for (let i = 1; i < objects.length; i++) {
                const obj = objects[i];
                if (obj.type === 'text' || obj.type === 'i-text') {
                    // Reset text scale to 1 (normal)
                    obj.set({ scaleX: 1, scaleY: 1 });
                    
                    // Reposition within new bounds
                    if (i === 1) {
                        // Label - top left
                        obj.set({
                            left: -newWidth / 2 + padding,
                            top: -newHeight / 2 + padding
                        });
                    } else if (i === 2) {
                        // Value - below label
                        const labelHeight = objects[1].fontSize || 10;
                        obj.set({
                            left: -newWidth / 2 + padding,
                            top: -newHeight / 2 + padding + labelHeight + 2
                        });
                    }
                }
            }
            
            group.setCoords();
            canvas.renderAll();
        });
    }

    addFormTextField(canvas, x, y) {
        const defaultName = this.getNextDefaultFieldId('textfield');
        const fieldWidth = 150 / this.currentScale; // Smaller default width
        const fieldHeight = 32 / this.currentScale; // Smaller height
        const labelFontSize = 10 / this.currentScale; // Larger label for readability
        const valueFontSize = 12 / this.currentScale;
        const padding = 4 / this.currentScale;

        // Create field background (centered at origin)
        const background = new fabric.Rect({
            width: fieldWidth,
            height: fieldHeight,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 1 / this.currentScale,
            rx: 3 / this.currentScale,
            ry: 3 / this.currentScale,
            originX: 'center',
            originY: 'center',
            left: 0,
            top: 0
        });

        // Create label text (small, at top-left inside field)
        const label = new fabric.Text(defaultName, {
            fontSize: labelFontSize,
            fontFamily: 'Arial',
            fill: '#6b7280',
            textBaseline: 'alphabetic',
            originX: 'left',
            originY: 'top',
            left: -fieldWidth / 2 + padding,
            top: -fieldHeight / 2 + padding
        });

        // Create value text (larger, below label) - show placeholder hint
        const valueText = new fabric.Text('Double-click to fill', {
            fontSize: valueFontSize,
            fontFamily: 'Arial',
            fill: '#9ca3af', // Light gray for placeholder
            fontStyle: 'italic',
            textBaseline: 'alphabetic',
            originX: 'left',
            originY: 'top',
            left: -fieldWidth / 2 + padding,
            top: -fieldHeight / 2 + padding + labelFontSize + 2 / this.currentScale
        });

        // Group them
        const group = new fabric.Group([background, label, valueText], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'textfield',
            _fieldValue: '',
            _fieldName: defaultName,
            _fieldLocked: false,
            _labelFontSize: labelFontSize,
            _valueFontSize: valueFontSize,
            _padding: padding
        });

        // Prevent text distortion on resize
        this._setupFormFieldScaling(group, canvas, padding);

        group.on('mousedblclick', () => {
            if (group._fieldLocked) return;
            this._showInlineTextEditor(group, canvas);
        });

        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();

        // Switch to select
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="textfield"]')?.classList.remove('active');
    }

    /**
     * Add a form checkbox
     */
    addFormCheckbox(canvas, x, y) {
        const defaultName = this.getNextDefaultFieldId('checkbox');
        const size = 20 / this.currentScale;

        // Create checkbox background
        const box = new fabric.Rect({
            width: size,
            height: size,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 2 / this.currentScale,
            rx: 3 / this.currentScale,
            ry: 3 / this.currentScale
        });

        const group = new fabric.Group([box], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'checkbox',
            _checked: false,
            _fieldName: defaultName,
            _fieldLocked: false
        });

        group.on('mousedblclick', () => {
            if (group._fieldLocked) return;
            this.toggleCheckbox(group, canvas);
        });

        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();

        // Switch to select
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="checkbox"]')?.classList.remove('active');
    }

    /**
     * Add a form radio button
     */
    addFormRadio(canvas, x, y) {
        const defaultName = this.getNextDefaultFieldId('radio');
        const size = 20 / this.currentScale;
        const outer = new fabric.Circle({
            radius: size / 2,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 2 / this.currentScale,
            originX: 'center',
            originY: 'center',
            left: 0,
            top: 0
        });

        const group = new fabric.Group([outer], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'radio',
            _checked: false,
            _fieldName: defaultName,
            _fieldLocked: false,
            _radioGroup: defaultName,
            _radioValue: `option_${Math.random().toString(36).slice(2, 10)}`
        });

        group.on('mousedblclick', () => {
            if (group._fieldLocked) return;
            this.toggleRadio(group, canvas);
        });

        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();

        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="radio"]')?.classList.remove('active');
    }

    toggleRadio(group, canvas) {
        const fieldName = group._fieldName || group._radioGroup || '';
        const makeDot = () =>
            new fabric.Circle({
                radius: 5 / this.currentScale,
                fill: '#2563eb',
                originX: 'center',
                originY: 'center',
                left: 0,
                top: 0
            });

        // If no name, just toggle locally
        if (!fieldName) {
            group._checked = !group._checked;
            const objs = group.getObjects();
            if (group._checked && objs.length === 1) group.addWithUpdate(makeDot());
            if (!group._checked && objs.length > 1) group.remove(objs[1]);
            canvas.renderAll();
            return;
        }

        // Uncheck other radios in same group
        canvas.getObjects().forEach((obj) => {
            if (obj === group) return;
            if (obj._annotationType === 'radio' && (obj._fieldName || obj._radioGroup) === fieldName) {
                obj._checked = false;
                const objs = obj.getObjects?.() || [];
                if (objs.length > 1) obj.remove(objs[1]);
            }
        });

        group._checked = true;
        const objs = group.getObjects();
        if (objs.length === 1) group.addWithUpdate(makeDot());
        canvas.renderAll();
    }

    /**
     * Add a form dropdown field
     */
    addFormDropdown(canvas, x, y) {
        const defaultName = this.getNextDefaultFieldId('dropdown');
        const fieldWidth = 180 / this.currentScale; // Smaller default width
        const fieldHeight = 32 / this.currentScale; // Smaller height
        const labelFontSize = 10 / this.currentScale; // Larger label for readability
        const valueFontSize = 12 / this.currentScale;
        const padding = 4 / this.currentScale;

        const background = new fabric.Rect({
            width: fieldWidth,
            height: fieldHeight,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 1 / this.currentScale,
            rx: 3 / this.currentScale,
            ry: 3 / this.currentScale,
            originX: 'center',
            originY: 'center',
            left: 0,
            top: 0
        });

        // Small label at top-left
        const label = new fabric.Text(defaultName, {
            fontSize: labelFontSize,
            fontFamily: 'Arial',
            fill: '#6b7280',
            textBaseline: 'alphabetic',
            originX: 'left',
            originY: 'top',
            left: -fieldWidth / 2 + padding,
            top: -fieldHeight / 2 + padding
        });

        // Value text (larger, below label) - show placeholder hint
        const valueText = new fabric.Text('Double-click to select', {
            fontSize: valueFontSize,
            fontFamily: 'Arial',
            fill: '#9ca3af',
            fontStyle: 'italic',
            textBaseline: 'alphabetic',
            originX: 'left',
            originY: 'top',
            left: -fieldWidth / 2 + padding,
            top: -fieldHeight / 2 + padding + labelFontSize + 2 / this.currentScale
        });

        const chevron = new fabric.Triangle({
            width: 8 / this.currentScale,
            height: 6 / this.currentScale,
            fill: '#6b7280',
            originX: 'center',
            originY: 'center',
            left: fieldWidth / 2 - 12 / this.currentScale,
            top: 0,
            angle: 180
        });

        const group = new fabric.Group([background, label, valueText, chevron], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'dropdown',
            _fieldName: defaultName,
            _fieldLocked: false,
            _options: ['Option 1', 'Option 2'],
            _selectedOption: '',
            _labelFontSize: labelFontSize,
            _valueFontSize: valueFontSize,
            _padding: padding
        });
        
        this._setupFormFieldScaling(group, canvas, padding);
        
        group.on('mousedblclick', () => {
            if (group._fieldLocked) return;
            if (this.fillMode) this._showInlineDropdownEditor(group, canvas);
        });
        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();

        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="dropdown"]')?.classList.remove('active');
    }

    /**
     * Add a form date field (visual text field with date placeholder)
     */
    addFormDateField(canvas, x, y) {
        const defaultName = this.getNextDefaultFieldId('date');
        const fieldWidth = 150 / this.currentScale; // Smaller default width
        const fieldHeight = 32 / this.currentScale; // Smaller height
        const labelFontSize = 10 / this.currentScale; // Larger label for readability
        const valueFontSize = 12 / this.currentScale;
        const padding = 4 / this.currentScale;

        const background = new fabric.Rect({
            width: fieldWidth,
            height: fieldHeight,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 1 / this.currentScale,
            rx: 3 / this.currentScale,
            ry: 3 / this.currentScale,
            originX: 'center',
            originY: 'center',
            left: 0,
            top: 0
        });

        // Create label text (small, at top-left inside field)
        const label = new fabric.Text(defaultName, {
            fontSize: labelFontSize,
            fontFamily: 'Arial',
            fill: '#6b7280',
            textBaseline: 'alphabetic',
            originX: 'left',
            originY: 'top',
            left: -fieldWidth / 2 + padding,
            top: -fieldHeight / 2 + padding
        });

        // Create value text with placeholder (larger, below label)
        const valueText = new fabric.Text('YYYY-MM-DD', {
            fontSize: valueFontSize,
            fontFamily: 'Arial',
            fill: '#9ca3af', // Lighter color for placeholder
            textBaseline: 'alphabetic',
            originX: 'left',
            originY: 'top',
            left: -fieldWidth / 2 + padding,
            top: -fieldHeight / 2 + padding + labelFontSize + 2 / this.currentScale
        });

        const group = new fabric.Group([background, label, valueText], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'date',
            _fieldValue: '',
            _fieldName: defaultName,
            _fieldLocked: false,
            _labelFontSize: labelFontSize,
            _valueFontSize: valueFontSize,
            _padding: padding
        });
        
        this._setupFormFieldScaling(group, canvas, padding);
        
        group.on('mousedblclick', () => {
            if (group._fieldLocked) return;
            this._showInlineTextEditor(group, canvas);
        });
        
        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();

        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="date"]')?.classList.remove('active');
    }

    /**
     * Toggle a checkbox state
     */
    toggleCheckbox(group, canvas) {
        const isChecked = !group._checked;
        group._checked = isChecked;

        // Remove old checkmark if exists
        const objects = group.getObjects();
        if (objects.length > 1) {
            group.remove(objects[1]);
        }

        // Add checkmark if checked
        if (isChecked) {
            const size = objects[0].width;
            const checkmark = new fabric.Path('M 3 10 L 8 15 L 17 4', {
                stroke: '#2563eb',
                strokeWidth: 2 / this.currentScale,
                fill: 'transparent',
                scaleX: size / 20,
                scaleY: size / 20,
                left: -size / 2,
                top: -size / 2
            });
            group.addWithUpdate(checkmark);
        }

        canvas.renderAll();
    }

    /**
     * Edit a text field (textfield or date field) - shows inline editor instead of prompt
     */
    editTextField(group, canvas) {
        this._showInlineTextEditor(group, canvas);
    }

    /**
     * Update the label text of a form field
     */
    updateFieldLabel(group, canvas, newLabel) {
        const objects = group.getObjects();
        const hasNewStructure = objects.length >= 3;
        
        if (hasNewStructure) {
            const labelText = objects[1]; // Label is the second object
            if (labelText && labelText.type === 'text') {
                labelText.set('text', newLabel || 'Text Field');
                canvas.renderAll();
            }
        }
    }

    /**
     * Select a value from a dropdown field - shows inline select overlay instead of prompt
     */
    selectDropdownValue(group, canvas) {
        const options = group._options || [];
        if (options.length === 0) {
            return;
        }
        this._showInlineDropdownEditor(group, canvas);
    }

    /**
     * Set a signature image and metadata to be inserted
     * @param {string} dataUrl - Signature image data URL
     * @param {Object} meta - Audit metadata (signerName, signerEmail?, intentAccepted, consentAccepted, documentFilename, documentHash?)
     */
    setSignature(dataUrl, meta) {
        this.signatureImage = dataUrl;
        this.signatureMeta = meta || null;
    }

    /**
     * Insert the signature at the specified position
     */
    insertSignature(canvas, x, y) {
        if (!this.signatureImage) return;

        const meta = this.signatureMeta ? { ...this.signatureMeta } : {};
        meta.timestamp = new Date().toISOString();

        fabric.Image.fromURL(this.signatureImage, (img) => {
            // Scale signature to reasonable size
            const maxWidth = 200 / this.currentScale;
            const scale = maxWidth / img.width;

            img.set({
                left: x,
                top: y,
                scaleX: scale,
                scaleY: scale,
                selectable: true,
                _annotationType: 'signature',
                _signatureMeta: meta
            });

            canvas.add(img);
            canvas.setActiveObject(img);
            canvas.renderAll();
        });

        // Switch to select tool
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="signature"]')?.classList.remove('active');
    }

    /**
     * Replace a signature field with an actual signature
     * @param {fabric.Group} signatureField - The signature field to replace
     * @param {fabric.Canvas} canvas - The canvas containing the field
     */
    replaceSignatureField(signatureField, canvas) {
        if (!this.signatureImage || !signatureField) return;

        const meta = this.signatureMeta ? { ...this.signatureMeta } : {};
        meta.timestamp = new Date().toISOString();
        meta.replacedFieldLabel = signatureField._signatureFieldLabel;

        // Get the position and size of the signature field
        const bounds = signatureField.getBoundingRect();
        const left = bounds.left;
        const top = bounds.top;
        const width = bounds.width;
        const height = bounds.height;

        fabric.Image.fromURL(this.signatureImage, (img) => {
            // Scale signature to fit within the field bounds
            const scaleX = width / img.width;
            const scaleY = height / img.height;
            const scale = Math.min(scaleX, scaleY) * 0.9; // 90% to leave some padding

            img.set({
                left: left,
                top: top,
                scaleX: scale,
                scaleY: scale,
                selectable: this.fillMode ? false : true,
                _annotationType: 'signature',
                _signatureMeta: meta,
                _signatureLocked: true // Immutable once signed; cannot be moved, resized, or deleted by another participant
            });

            // Remove the signature field
            canvas.remove(signatureField);
            
            // Add the signature
            canvas.add(img);
            canvas.renderAll();
            
            // Trigger field update event
            if (typeof window !== 'undefined') {
                window.dispatchEvent(new CustomEvent('signature-field-filled', { 
                    detail: { fieldLabel: signatureField._signatureFieldLabel, signerName: meta.signerName }
                }));
            }
        });
    }

    /**
     * Set the label for the next signature field to be placed
     * @param {string} label - Label for the signature field (e.g. "Tenant 1", "Parent 1")
     */
    setSignatureFieldLabel(label) {
        this.signatureFieldLabel = label || 'Signature';
    }

    /**
     * Insert an empty signature field (box with label) at the specified position
     */
    insertSignatureField(canvas, x, y) {
        const label = this.getNextDefaultFieldId('signature-field');
        
        // Create a group with a rectangle and text label
        const width = 200 / this.currentScale;
        const height = 60 / this.currentScale;
        const fontSize = 14 / this.currentScale;

        const rect = new fabric.Rect({
            width: width,
            height: height,
            fill: 'rgba(255, 255, 200, 0.3)',
            stroke: '#999',
            strokeWidth: 2 / this.currentScale,
            strokeDashArray: [5 / this.currentScale, 5 / this.currentScale],
            rx: 4 / this.currentScale,
            ry: 4 / this.currentScale
        });

        const text = new fabric.Text(`${label}\n(Double-click to sign)`, {
            fontSize: fontSize,
            fill: '#666',
            textAlign: 'center',
            textBaseline: 'alphabetic',
            originX: 'center',
            originY: 'center',
            left: width / 2,
            top: height / 2
        });

        const group = new fabric.Group([rect, text], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'signature-field',
            _signatureFieldLabel: label
        });

        canvas.add(group);
        canvas.setActiveObject(group);
        canvas.renderAll();

        // Notify that field was placed (for hiding hint banner)
        if (typeof window !== 'undefined') {
            window.dispatchEvent(new CustomEvent('signature-field-placed'));
        }

        // Trigger selection change so Field Properties sidebar shows
        this.onSelectionChanged(canvas);

        // Switch to select tool
        this.setTool('select');
        document.querySelector('[data-tool="select"]')?.classList.add('active');
        document.querySelector('[data-tool="signature-field"]')?.classList.remove('active');
    }

    /**
     * Update settings
     */
    updateSettings(newSettings) {
        Object.assign(this.settings, newSettings);

        // Update drawing brush if in draw mode
        if (this.activeTool === 'draw') {
            this.canvases.forEach((canvas) => {
                if (canvas.freeDrawingBrush) {
                    canvas.freeDrawingBrush.color = this.settings.strokeColor;
                    canvas.freeDrawingBrush.width = this.settings.strokeWidth;
                }
            });
        }

        // If a text object is selected, apply relevant text settings immediately
        this.canvases.forEach((canvas) => {
            const activeObject = canvas.getActiveObject();
            if (activeObject && activeObject._annotationType === 'text') {
                if ('textColor' in newSettings) {
                    activeObject.set('fill', this.settings.textColor);
                }
                if ('fontSize' in newSettings) {
                    // Account for current zoom so visual size matches control value
                    activeObject.set('fontSize', this.settings.fontSize / this.currentScale);
                }
                if ('fontFamily' in newSettings) {
                    activeObject.set('fontFamily', this.settings.fontFamily);
                }
                if ('fontWeight' in newSettings) {
                    activeObject.set('fontWeight', this.settings.fontWeight);
                }
                if ('fontStyle' in newSettings) {
                    activeObject.set('fontStyle', this.settings.fontStyle);
                }
                if ('textAlign' in newSettings) {
                    activeObject.set('textAlign', this.settings.textAlign);
                }
                activeObject.setCoords();
                canvas.renderAll();
            }
        });
    }

    /**
     * Delete selected objects
     */
    deleteSelected() {
        this.canvases.forEach((canvas, pageId) => {
            const activeObjects = canvas.getActiveObjects();
            if (activeObjects.length > 0) {
                activeObjects.forEach((obj) => {
                    // Signed/locked fields are immutable; do not allow deletion
                    if (obj._signatureLocked || obj._fieldLocked) return;
                    canvas.remove(obj);
                });
                canvas.discardActiveObject();
                canvas.renderAll();
                this.saveState(pageId);
            }
        });
    }

    /**
     * Save current state for undo
     */
    saveState(pageId) {
        const canvas = this.canvases.get(pageId);
        const history = this.history.get(pageId);
        if (!canvas || !history) return;

        const state = JSON.stringify(
            canvas.toJSON([
                '_annotationType',
                '_checked',
                '_fieldValue',
                '_fieldName',
                '_signatureMeta',
                '_signatureFieldLabel',
                '_signatureLocked',
                '_fieldLocked',
                '_noteText',
                '_stampText',
                '_options',
                '_selectedOption',
                '_radioGroup',
                '_radioValue'
            ])
        );
        history.undoStack.push(state);
        history.redoStack = []; // Clear redo stack on new action

        // Limit history size
        if (history.undoStack.length > 50) {
            history.undoStack.shift();
        }
        this._onHistoryChange?.();
    }

    /**
     * Undo last action
     */
    undo() {
        const pageId = this.activePageId;
        const canvas = pageId != null ? this.canvases.get(pageId) : null;
        const history = pageId != null ? this.history.get(pageId) : null;
        if (!canvas || !history || history.undoStack.length <= 1) return;

        const currentState = history.undoStack.pop();
        history.redoStack.push(currentState);

        const prevState = history.undoStack[history.undoStack.length - 1];
        if (!prevState) return;

        this._restoringPages.add(pageId);
        canvas.loadFromJSON(prevState, () => {
            canvas.forEachObject((obj) => {
                this._fixTextBaseline(obj);
                obj._fromHistory = true;
                this._applyObjectInteractivity(obj, this.fillMode, this.activeTool);
            });
            canvas.discardActiveObject();
            canvas.renderAll();
            this._restoringPages.delete(pageId);
            this._onHistoryChange?.();
        });
    }

    /**
     * Redo last undone action
     */
    redo() {
        const pageId = this.activePageId;
        const canvas = pageId != null ? this.canvases.get(pageId) : null;
        const history = pageId != null ? this.history.get(pageId) : null;
        if (!canvas || !history || history.redoStack.length === 0) return;

        const nextState = history.redoStack.pop();
        history.undoStack.push(nextState);

        this._restoringPages.add(pageId);
        canvas.loadFromJSON(nextState, () => {
            canvas.forEachObject((obj) => {
                this._fixTextBaseline(obj);
                obj._fromHistory = true;
                this._applyObjectInteractivity(obj, this.fillMode, this.activeTool);
            });
            canvas.discardActiveObject();
            canvas.renderAll();
            this._restoringPages.delete(pageId);
            this._onHistoryChange?.();
        });
    }

    /**
     * Fix invalid textBaseline (e.g. 'alphabetical' → 'alphabetic') on Fabric text objects.
     * Fabric.js/Canvas expects 'alphabetic', not 'alphabetical'.
     */
    _fixTextBaseline(obj) {
        const fix = (o) => {
            if (o && typeof o.textBaseline === 'string' && o.textBaseline !== 'alphabetic') {
                if (o.textBaseline === 'alphabetical' || !['top', 'hanging', 'middle', 'alphabetic', 'ideographic', 'bottom'].includes(o.textBaseline)) {
                    o.textBaseline = 'alphabetic';
                }
            }
            if (o && o.getObjects) {
                o.getObjects().forEach(fix);
            }
        };
        fix(obj);
    }

    /**
     * Update canvas dimensions on zoom
     */
    updateCanvasSize(fabricCanvas, width, height, scale) {
        this.currentScale = scale;

        const prevScale = fabricCanvas.width / width * scale;
        const scaleChange = scale / prevScale;

        fabricCanvas.setWidth(width);
        fabricCanvas.setHeight(height);

        // Scale all objects
        fabricCanvas.forEachObject((obj) => {
            obj.scaleX *= scaleChange;
            obj.scaleY *= scaleChange;
            obj.left *= scaleChange;
            obj.top *= scaleChange;
            obj.setCoords();
        });

        // Update brush width
        if (fabricCanvas.freeDrawingBrush) {
            fabricCanvas.freeDrawingBrush.width = this.settings.strokeWidth;
        }

        fabricCanvas.renderAll();
    }

    /**
     * Get all annotations for export
     */
    getAllAnnotations() {
        const annotations = [];

        this.canvases.forEach((canvas, pageId) => {
            const pageAnnotations = [];
            canvas.forEachObject((obj) => {
                pageAnnotations.push({
                    type: obj._annotationType || obj.type,
                    data: obj.toJSON(['_annotationType', '_checked', '_fieldValue', '_fieldName', '_signatureMeta', '_signatureFieldLabel', '_signatureLocked', '_fieldLocked']),
                    object: obj
                });
            });
            annotations.push({
                pageId,
                annotations: pageAnnotations,
                canvas
            });
        });

        return annotations;
    }

    /**
     * Handle selection changes
     */
    onSelectionChanged(canvas) {
        const activeObject = canvas.getActiveObject();
        const deleteBtn = document.getElementById('btn-delete');
        if (deleteBtn) {
            deleteBtn.disabled = !activeObject;
        }

        // Dispatch event for form field selection (for sidebar focus in Fill mode)
        if (activeObject && this.fillMode) {
            const annotationType = activeObject._annotationType;
            if (['textfield', 'checkbox', 'radio', 'dropdown', 'date', 'signature-field'].includes(annotationType)) {
                window.dispatchEvent(new CustomEvent('form-field-selected', {
                    detail: { object: activeObject, annotationType }
                }));
            }
        }

        // Dispatch for Field Properties sidebar (Edit mode only)
        if (this.fillMode) {
            window.dispatchEvent(new CustomEvent('field-properties-hide'));
        } else if (activeObject) {
            const annotationType = activeObject._annotationType;
            if (['textfield', 'checkbox', 'radio', 'dropdown', 'date', 'signature-field'].includes(annotationType)) {
                window.dispatchEvent(new CustomEvent('field-properties-show', {
                    detail: { object: activeObject, canvas, annotationType }
                }));
            } else {
                window.dispatchEvent(new CustomEvent('field-properties-hide'));
            }
        } else {
            window.dispatchEvent(new CustomEvent('field-properties-hide'));
        }
    }

    /**
     * Get the next available default field ID for a given type.
     * Returns e.g. "text_field_1", "text_field_2", "signature_field_1", etc.
     */
    getNextDefaultFieldId(type) {
        const prefixMap = {
            'textfield': 'text_field_',
            'date': 'date_field_',
            'dropdown': 'dropdown_',
            'checkbox': 'checkbox_',
            'radio': 'radio_',
            'signature-field': 'signature_field_'
        };
        const prefix = prefixMap[type] || 'field_';
        const existing = new Set();

        for (const [, canvas] of this.canvases) {
            for (const obj of canvas.getObjects()) {
                if (obj._annotationType === 'signature-field') {
                    const name = (obj._signatureFieldLabel || '').trim();
                    if (name && name.toLowerCase().startsWith(prefix.toLowerCase())) {
                        existing.add(name.toLowerCase());
                    }
                } else if (obj._fieldName) {
                    const name = (obj._fieldName || '').trim();
                    if (name && name.toLowerCase().startsWith(prefix.toLowerCase())) {
                        existing.add(name.toLowerCase());
                    }
                }
            }
        }

        let n = 1;
        while (existing.has(`${prefix}${n}`.toLowerCase())) {
            n++;
        }
        return `${prefix}${n}`;
    }

    /**
     * Add a form field from a PDF descriptor (when loading a PDF that has form fields).
     * @param {fabric.Canvas} canvas
     * @param {Object} descriptor - { type, name, rect: { left, top, width, height }, value?, checked?, options?, signatureLabel? }
     * @param {number} scale - Scale factor (PDF points to canvas pixels)
     */
    addFormFieldFromPdfDescriptor(canvas, descriptor, scale) {
        const { type, name, rect, value = '', checked = false, options = [], signatureLabel } = descriptor;
        const left = rect.left * scale;
        const top = rect.top * scale;
        const width = Math.max(20, rect.width * scale);
        const height = Math.max(16, rect.height * scale);
        const pad = 4 / this.currentScale;

        if (type === 'signature-field') {
            const label = signatureLabel || 'Signature';
            const rectObj = new fabric.Rect({
                width,
                height,
                fill: 'rgba(255, 255, 200, 0.3)',
                stroke: '#999',
                strokeWidth: 2 / this.currentScale,
                strokeDashArray: [5 / this.currentScale, 5 / this.currentScale],
                rx: 4 / this.currentScale,
                ry: 4 / this.currentScale
            });
            const fontSize = Math.min(14, height / 4) / this.currentScale;
            const text = new fabric.Text(`${label}\n(Double-click to sign)`, {
                fontSize,
                fill: '#666',
                textAlign: 'center',
                textBaseline: 'alphabetic',
                originX: 'center',
                originY: 'center',
                left: width / 2,
                top: height / 2
            });
            const group = new fabric.Group([rectObj, text], {
                left,
                top,
                selectable: true,
                _annotationType: 'signature-field',
                _signatureFieldLabel: label
            });
            canvas.add(group);
        } else if (type === 'textfield' || type === 'date') {
            const labelFontSize = Math.min(10, height / 3) / this.currentScale;
            const valueFontSize = Math.min(12, height / 2.5) / this.currentScale;
            const background = new fabric.Rect({
                width,
                height,
                fill: '#ffffff',
                stroke: '#2563eb',
                strokeWidth: 1 / this.currentScale,
                rx: 3 / this.currentScale,
                ry: 3 / this.currentScale,
                originX: 'left',
                originY: 'top',
                left: 0,
                top: 0
            });
            const label = new fabric.Text(name, {
                fontSize: labelFontSize,
                fontFamily: 'Arial',
                fill: '#6b7280',
                textBaseline: 'alphabetic',
                originX: 'left',
                originY: 'top',
                left: pad,
                top: pad
            });
            const displayVal = value || (type === 'date' ? 'YYYY-MM-DD' : 'Double-click to fill');
            const valueText = new fabric.Text(displayVal, {
                fontSize: valueFontSize,
                fontFamily: 'Arial',
                fill: value ? '#000000' : '#9ca3af',
                fontStyle: value ? 'normal' : 'italic',
                textBaseline: 'alphabetic',
                originX: 'left',
                originY: 'top',
                left: pad,
                top: pad + labelFontSize + 2 / this.currentScale
            });
            const group = new fabric.Group([background, label, valueText], {
                left,
                top,
                selectable: true,
                _annotationType: type,
                _fieldValue: value,
                _fieldName: name,
                _fieldLocked: false,
                _labelFontSize: labelFontSize,
                _valueFontSize: valueFontSize,
                _padding: pad
            });
            this._setupFormFieldScaling(group, canvas, pad);
            group.on('mousedblclick', () => {
                if (group._fieldLocked) return;
                this._showInlineTextEditor(group, canvas);
            });
            canvas.add(group);
        } else if (type === 'checkbox') {
            const size = Math.min(width, height, 24 / this.currentScale);
            const box = new fabric.Rect({
                width: size,
                height: size,
                fill: '#ffffff',
                stroke: '#2563eb',
                strokeWidth: 2 / this.currentScale,
                rx: 3 / this.currentScale,
                ry: 3 / this.currentScale,
                originX: 'center',
                originY: 'center',
                left: width / 2,
                top: height / 2
            });
            const group = new fabric.Group([box], {
                left: left + width / 2,
                top: top + height / 2,
                originX: 'center',
                originY: 'center',
                selectable: true,
                _annotationType: 'checkbox',
                _checked: checked,
                _fieldName: name,
                _fieldLocked: false
            });
            if (checked) {
                const dot = new fabric.Circle({
                    radius: 5 / this.currentScale,
                    fill: '#2563eb',
                    originX: 'center',
                    originY: 'center',
                    left: 0,
                    top: 0
                });
                group.addWithUpdate(dot);
            }
            group.on('mousedblclick', () => {
                if (group._fieldLocked) return;
                this.toggleCheckbox(group, canvas);
            });
            canvas.add(group);
        } else if (type === 'dropdown') {
            const labelFontSize = Math.min(10, height / 3) / this.currentScale;
            const valueFontSize = Math.min(12, height / 2.5) / this.currentScale;
            const background = new fabric.Rect({
                width,
                height,
                fill: '#ffffff',
                stroke: '#2563eb',
                strokeWidth: 1 / this.currentScale,
                rx: 3 / this.currentScale,
                ry: 3 / this.currentScale,
                originX: 'left',
                originY: 'top',
                left: 0,
                top: 0
            });
            const label = new fabric.Text(name, {
                fontSize: labelFontSize,
                fontFamily: 'Arial',
                fill: '#6b7280',
                textBaseline: 'alphabetic',
                originX: 'left',
                originY: 'top',
                left: pad,
                top: pad
            });
            const displayVal = value || 'Double-click to select';
            const valueText = new fabric.Text(displayVal, {
                fontSize: valueFontSize,
                fontFamily: 'Arial',
                fill: value ? '#000000' : '#9ca3af',
                fontStyle: value ? 'normal' : 'italic',
                textBaseline: 'alphabetic',
                originX: 'left',
                originY: 'top',
                left: pad,
                top: pad + labelFontSize + 2 / this.currentScale
            });
            const chevron = new fabric.Triangle({
                width: 8 / this.currentScale,
                height: 6 / this.currentScale,
                fill: '#6b7280',
                originX: 'center',
                originY: 'center',
                left: width - 12 / this.currentScale,
                top: height / 2,
                angle: 180
            });
            const group = new fabric.Group([background, label, valueText, chevron], {
                left,
                top,
                selectable: true,
                _annotationType: 'dropdown',
                _fieldName: name,
                _fieldLocked: false,
                _options: options.length ? options : ['Option 1', 'Option 2'],
                _selectedOption: value,
                _labelFontSize: labelFontSize,
                _valueFontSize: valueFontSize,
                _padding: pad
            });
            this._setupFormFieldScaling(group, canvas, pad);
            group.on('mousedblclick', () => {
                if (group._fieldLocked) return;
                if (this.fillMode) this._showInlineDropdownEditor(group, canvas);
            });
            canvas.add(group);
        }
        canvas.renderAll();
    }

    /**
     * Check if a field name is already used by another field
     * @param {string} name - The field name to check
     * @param {fabric.Object} excludeObject - The object to exclude from the check (the one being edited)
     * @returns {boolean} True if the name is already used by another field
     */
    isFieldNameDuplicate(name, excludeObject) {
        if (!name) return false;
        
        const lowerName = name.toLowerCase();
        
        for (const [pageId, canvas] of this.canvases) {
            for (const obj of canvas.getObjects()) {
                if (obj === excludeObject) continue;
                
                // Check _fieldName for form fields
                if (obj._fieldName && obj._fieldName.toLowerCase() === lowerName) {
                    return true;
                }
                
                // Check _signatureFieldLabel for signature fields
                if (obj._signatureFieldLabel && obj._signatureFieldLabel.toLowerCase() === lowerName) {
                    return true;
                }
            }
        }
        
        return false;
    }

    /**
     * Clear all canvases
     */
    clearAll() {
        this.canvases.forEach((canvas) => {
            canvas.clear();
            canvas.dispose();
        });
        this.canvases.clear();
        this.history.clear();
        this.activeCanvas = null;
        this.activePageId = null;
    }

    /**
     * Remove a single page canvas and its history (used for page delete)
     * @param {string} pageId
     */
    removePage(pageId) {
        const canvas = this.canvases.get(pageId);
        if (canvas) {
            try {
                canvas.clear();
                canvas.dispose();
            } catch {
                // ignore
            }
        }
        this.canvases.delete(pageId);
        this.history.delete(pageId);
        this._restoringPages.delete(pageId);
        if (this.activePageId === pageId) {
            this.activeCanvas = null;
            this.activePageId = null;
        }
    }
}
