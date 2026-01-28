/**
 * Canvas Manager - Manages Fabric.js canvas overlays for PDF annotation
 */

export class CanvasManager {
    constructor() {
        this.canvases = new Map(); // pageNum -> fabric.Canvas
        this.activeCanvas = null;
        this.activeTool = 'select';
        this.history = new Map(); // pageNum -> {undoStack, redoStack}
        this.currentScale = 1.0;

        // Tool settings
        this.settings = {
            textColor: '#000000',
            fontSize: 16,
            fontFamily: 'Arial',
            strokeColor: '#000000',
            strokeWidth: 2,
            whiteoutColor: '#ffffff'
        };

        // Signature data
        this.signatureImage = null;
        this.signatureMeta = null;
    }

    /**
     * Create a Fabric.js canvas overlay for a page
     * @param {HTMLElement} container - Container element
     * @param {number} width - Canvas width
     * @param {number} height - Canvas height
     * @param {number} pageNum - Page number
     * @returns {fabric.Canvas}
     */
    createCanvas(container, width, height, pageNum) {
        // Create canvas element
        const canvasEl = document.createElement('canvas');
        canvasEl.id = `annotation-canvas-${pageNum}`;
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
        this.canvases.set(pageNum, fabricCanvas);
        this.history.set(pageNum, { undoStack: [], redoStack: [] });

        // Set up event listeners
        this.setupCanvasEvents(fabricCanvas, pageNum);

        return fabricCanvas;
    }

    /**
     * Set up canvas event listeners
     * @param {fabric.Canvas} canvas - Fabric.js canvas
     * @param {number} pageNum - Page number
     */
    setupCanvasEvents(canvas, pageNum) {
        // Track modifications for undo/redo
        canvas.on('object:added', (e) => {
            if (!e.target._fromHistory) {
                this.saveState(pageNum);
            }
        });

        canvas.on('object:modified', () => {
            this.saveState(pageNum);
        });

        canvas.on('object:removed', (e) => {
            if (!e.target._fromHistory) {
                this.saveState(pageNum);
            }
        });

        // Click handler for tools
        canvas.on('mouse:down', (e) => {
            this.activeCanvas = canvas;
            this.handleMouseDown(e, canvas, pageNum);
        });

        canvas.on('mouse:up', (e) => {
            this.handleMouseUp(e, canvas, pageNum);
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
    handleMouseDown(e, canvas, pageNum) {
        const pointer = canvas.getPointer(e.e);

        switch (this.activeTool) {
            case 'text':
                this.addTextBox(canvas, pointer.x, pointer.y);
                break;
            case 'whiteout':
                this.startWhiteout(canvas, pointer.x, pointer.y);
                break;
            case 'textfield':
                this.addFormTextField(canvas, pointer.x, pointer.y);
                break;
            case 'checkbox':
                this.addFormCheckbox(canvas, pointer.x, pointer.y);
                break;
            case 'signature':
                if (this.signatureImage) {
                    this.insertSignature(canvas, pointer.x, pointer.y);
                }
                break;
        }
    }

    /**
     * Handle mouse up event
     */
    handleMouseUp(e, canvas, pageNum) {
        if (this.activeTool === 'whiteout' && this.tempWhiteout) {
            this.finishWhiteout(canvas);
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

    /**
     * Add a text box at the specified position
     */
    addTextBox(canvas, x, y) {
        const textbox = new fabric.IText('Click to edit', {
            left: x,
            top: y,
            fontSize: this.settings.fontSize / this.currentScale,
            fontFamily: this.settings.fontFamily,
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
            _fieldValue: ''
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
            _checked: false
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
                activeObject.setCoords();
                canvas.renderAll();
            }
        });
    }

    /**
     * Delete selected objects
     */
    deleteSelected() {
        this.canvases.forEach((canvas, pageNum) => {
            const activeObjects = canvas.getActiveObjects();
            if (activeObjects.length > 0) {
                activeObjects.forEach((obj) => {
                    canvas.remove(obj);
                });
                canvas.discardActiveObject();
                canvas.renderAll();
                this.saveState(pageNum);
            }
        });
    }

    /**
     * Save current state for undo
     */
    saveState(pageNum) {
        const canvas = this.canvases.get(pageNum);
        const history = this.history.get(pageNum);
        if (!canvas || !history) return;

        const state = JSON.stringify(canvas.toJSON(['_annotationType', '_checked', '_fieldValue', '_signatureMeta']));
        history.undoStack.push(state);
        history.redoStack = []; // Clear redo stack on new action

        // Limit history size
        if (history.undoStack.length > 50) {
            history.undoStack.shift();
        }
    }

    /**
     * Undo last action
     */
    undo() {
        this.canvases.forEach((canvas, pageNum) => {
            const history = this.history.get(pageNum);
            if (!history || history.undoStack.length <= 1) return;

            const currentState = history.undoStack.pop();
            history.redoStack.push(currentState);

            const prevState = history.undoStack[history.undoStack.length - 1];
            if (prevState) {
                canvas.loadFromJSON(prevState, () => {
                    canvas.forEachObject((obj) => {
                        obj._fromHistory = true;
                    });
                    canvas.renderAll();
                });
            }
        });
    }

    /**
     * Redo last undone action
     */
    redo() {
        this.canvases.forEach((canvas, pageNum) => {
            const history = this.history.get(pageNum);
            if (!history || history.redoStack.length === 0) return;

            const nextState = history.redoStack.pop();
            history.undoStack.push(nextState);

            canvas.loadFromJSON(nextState, () => {
                canvas.forEachObject((obj) => {
                    obj._fromHistory = true;
                });
                canvas.renderAll();
            });
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

        this.canvases.forEach((canvas, pageNum) => {
            const pageAnnotations = [];
            canvas.forEachObject((obj) => {
                pageAnnotations.push({
                    type: obj._annotationType || obj.type,
                    data: obj.toJSON(['_annotationType', '_checked', '_fieldValue', '_signatureMeta']),
                    object: obj
                });
            });
            annotations.push({
                pageNum,
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
    }
}
