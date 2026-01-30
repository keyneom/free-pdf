/**
 * Canvas Manager - Manages Fabric.js canvas overlays for PDF annotation
 */

export class CanvasManager {
    constructor() {
        this.canvases = new Map(); // pageId -> fabric.Canvas
        this.activeCanvas = null;
        this.activePageId = null;
        this.activeTool = 'select';
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
                canvas.selection = true;
                canvas.forEachObject((obj) => {
                    obj.selectable = true;
                    obj.evented = true;
                });
            } else if (tool !== 'draw') {
                canvas.selection = false;
                canvas.discardActiveObject();
                canvas.forEachObject((obj) => {
                    obj.selectable = false;
                    obj.evented = false;
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
    addFormTextField(canvas, x, y) {
        const fieldWidth = 200 / this.currentScale;
        const fieldHeight = 30 / this.currentScale;

        // Create field background
        const background = new fabric.Rect({
            width: fieldWidth,
            height: fieldHeight,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 1 / this.currentScale,
            rx: 3 / this.currentScale,
            ry: 3 / this.currentScale
        });

        // Create placeholder text
        const placeholder = new fabric.IText('Text field', {
            fontSize: 12 / this.currentScale,
            fontFamily: 'Arial',
            fill: '#6b7280',
            editable: true
        });

        // Group them
        const group = new fabric.Group([background, placeholder], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'textfield',
            _fieldValue: '',
            _fieldName: '' // Field name for bulk filling
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
            _fieldName: '' // Field name for bulk filling
        });

        // Toggle on double click
        group.on('mousedblclick', () => {
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
            _fieldName: '',
            _radioGroup: '', // alias to _fieldName for serialization compatibility
            _radioValue: `option_${Math.random().toString(36).slice(2, 10)}`
        });

        group.on('mousedblclick', () => {
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
        const fieldWidth = 220 / this.currentScale;
        const fieldHeight = 32 / this.currentScale;
        const background = new fabric.Rect({
            width: fieldWidth,
            height: fieldHeight,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 1 / this.currentScale,
            rx: 3 / this.currentScale,
            ry: 3 / this.currentScale
        });
        const label = new fabric.Textbox('Dropdown', {
            width: fieldWidth - 34 / this.currentScale,
            fontSize: 12 / this.currentScale,
            fontFamily: 'Arial',
            fill: '#6b7280',
            left: 8 / this.currentScale,
            top: 8 / this.currentScale,
            editable: false
        });
        const chevron = new fabric.Triangle({
            width: 10 / this.currentScale,
            height: 8 / this.currentScale,
            fill: '#6b7280',
            left: fieldWidth - 16 / this.currentScale,
            top: fieldHeight / 2,
            originX: 'center',
            originY: 'center',
            angle: 180
        });
        const group = new fabric.Group([background, label, chevron], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'dropdown',
            _fieldName: '',
            _options: ['Option 1', 'Option 2'],
            _selectedOption: ''
        });
        group.on('mousedblclick', () => {
            const raw = prompt('Dropdown options (comma separated):', (group._options || []).join(', '));
            if (raw == null) return;
            group._options = raw
                .split(',')
                .map((s) => s.trim())
                .filter(Boolean);
            label.text = group._selectedOption || 'Dropdown';
            canvas.renderAll();
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
        const fieldWidth = 200 / this.currentScale;
        const fieldHeight = 30 / this.currentScale;
        const background = new fabric.Rect({
            width: fieldWidth,
            height: fieldHeight,
            fill: '#ffffff',
            stroke: '#2563eb',
            strokeWidth: 1 / this.currentScale,
            rx: 3 / this.currentScale,
            ry: 3 / this.currentScale
        });
        const placeholder = new fabric.IText('YYYY-MM-DD', {
            fontSize: 12 / this.currentScale,
            fontFamily: 'Arial',
            fill: '#6b7280',
            editable: true
        });
        const group = new fabric.Group([background, placeholder], {
            left: x,
            top: y,
            selectable: true,
            _annotationType: 'date',
            _fieldValue: '',
            _fieldName: ''
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
                obj._fromHistory = true;
            });
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
                obj._fromHistory = true;
            });
            canvas.renderAll();
            this._restoringPages.delete(pageId);
            this._onHistoryChange?.();
        });
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
                    data: obj.toJSON(['_annotationType', '_checked', '_fieldValue', '_fieldName', '_signatureMeta']),
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

        // Show field name editor for form fields
        const toolOptions = document.getElementById('tool-options');
        if (toolOptions && activeObject) {
            const annotationType = activeObject._annotationType;
            if (annotationType === 'textfield' || annotationType === 'checkbox' || annotationType === 'radio' || annotationType === 'dropdown' || annotationType === 'date') {
                const fieldName = activeObject._fieldName || '';
                const extra =
                    annotationType === 'dropdown'
                        ? `
                    <div class="tool-option" style="margin-top: 8px;">
                        <label>Options:</label>
                        <input type="text" id="field-options-input" value="${(activeObject._options || []).join(', ')}" placeholder="Option 1, Option 2" style="min-width: 240px;">
                    </div>
                `
                        : '';
                toolOptions.innerHTML = `
                    <div class="tool-option">
                        <label>Field Name:</label>
                        <input type="text" id="field-name-input" value="${fieldName}" placeholder="e.g. tenant_name" style="min-width: 200px;">
                        <small style="display: block; color: #6b7280; margin-top: 4px;">Used for CSV bulk filling</small>
                    </div>
                    ${extra}
                `;
                const nameInput = document.getElementById('field-name-input');
                if (nameInput) {
                    nameInput.addEventListener('input', (e) => {
                        activeObject._fieldName = e.target.value.trim();
                        if (annotationType === 'radio') activeObject._radioGroup = activeObject._fieldName;
                        canvas.renderAll();
                    });
                }
                const optsInput = document.getElementById('field-options-input');
                if (optsInput) {
                    optsInput.addEventListener('input', (e) => {
                        activeObject._options = e.target.value
                            .split(',')
                            .map((s) => s.trim())
                            .filter(Boolean);
                        canvas.renderAll();
                    });
                }
            }
        }
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
