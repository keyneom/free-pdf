/**
 * Signature Pad - Handles signature creation (draw and type modes)
 */

export class SignaturePad {
    constructor() {
        this.canvas = null;
        this.ctx = null;
        this.isDrawing = false;
        this.lastX = 0;
        this.lastY = 0;
        this.mode = 'draw'; // 'draw' or 'type'
        this.typedText = '';
        this.fontStyle = 'cursive';
    }

    /**
     * Initialize the signature pad with a canvas element
     * @param {HTMLCanvasElement} canvas - Canvas element for drawing
     */
    init(canvas) {
        this.canvas = canvas;
        this.ctx = canvas.getContext('2d');

        // Set up drawing styles
        this.ctx.strokeStyle = '#000000';
        this.ctx.lineWidth = 2;
        this.ctx.lineCap = 'round';
        this.ctx.lineJoin = 'round';

        // Set up event listeners for drawing
        this.setupDrawingEvents();
    }

    /**
     * Set up mouse and touch events for drawing
     */
    setupDrawingEvents() {
        // Mouse events
        this.canvas.addEventListener('mousedown', (e) => this.startDrawing(e));
        this.canvas.addEventListener('mousemove', (e) => this.draw(e));
        this.canvas.addEventListener('mouseup', () => this.stopDrawing());
        this.canvas.addEventListener('mouseout', () => this.stopDrawing());

        // Touch events
        this.canvas.addEventListener('touchstart', (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            this.startDrawing(touch);
        });

        this.canvas.addEventListener('touchmove', (e) => {
            e.preventDefault();
            const touch = e.touches[0];
            this.draw(touch);
        });

        this.canvas.addEventListener('touchend', () => this.stopDrawing());
    }

    /**
     * Start drawing
     */
    startDrawing(e) {
        this.isDrawing = true;
        const rect = this.canvas.getBoundingClientRect();
        this.lastX = e.clientX - rect.left;
        this.lastY = e.clientY - rect.top;
    }

    /**
     * Draw on the canvas
     */
    draw(e) {
        if (!this.isDrawing) return;

        const rect = this.canvas.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;

        this.ctx.beginPath();
        this.ctx.moveTo(this.lastX, this.lastY);
        this.ctx.lineTo(x, y);
        this.ctx.stroke();

        this.lastX = x;
        this.lastY = y;
    }

    /**
     * Stop drawing
     */
    stopDrawing() {
        this.isDrawing = false;
    }

    /**
     * Clear the signature canvas
     */
    clear() {
        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    }

    /**
     * Set drawing color
     * @param {string} color - Color string
     */
    setColor(color) {
        this.ctx.strokeStyle = color;
    }

    /**
     * Set line width
     * @param {number} width - Line width in pixels
     */
    setLineWidth(width) {
        this.ctx.lineWidth = width;
    }

    /**
     * Check if the canvas has any drawing
     * @returns {boolean}
     */
    isEmpty() {
        const imageData = this.ctx.getImageData(0, 0, this.canvas.width, this.canvas.height);
        const data = imageData.data;

        // Check if any pixel is not transparent
        for (let i = 3; i < data.length; i += 4) {
            if (data[i] > 0) return false;
        }
        return true;
    }

    /**
     * Get the signature as a data URL (PNG)
     * @returns {string} - Data URL of the signature
     */
    getDataUrl() {
        if (this.mode === 'draw') {
            return this.getTrimmedDataUrl();
        } else {
            return this.getTypedSignatureDataUrl();
        }
    }

    /**
     * Get trimmed drawing as data URL (removes whitespace)
     * @returns {string}
     */
    getTrimmedDataUrl() {
        const imageData = this.ctx.getImageData(0, 0, this.canvas.width, this.canvas.height);
        const data = imageData.data;

        let minX = this.canvas.width;
        let minY = this.canvas.height;
        let maxX = 0;
        let maxY = 0;

        // Find bounding box of non-transparent pixels
        for (let y = 0; y < this.canvas.height; y++) {
            for (let x = 0; x < this.canvas.width; x++) {
                const alpha = data[(y * this.canvas.width + x) * 4 + 3];
                if (alpha > 0) {
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }

        // Add padding
        const padding = 10;
        minX = Math.max(0, minX - padding);
        minY = Math.max(0, minY - padding);
        maxX = Math.min(this.canvas.width, maxX + padding);
        maxY = Math.min(this.canvas.height, maxY + padding);

        const width = maxX - minX;
        const height = maxY - minY;

        if (width <= 0 || height <= 0) {
            return this.canvas.toDataURL();
        }

        // Create trimmed canvas
        const trimmedCanvas = document.createElement('canvas');
        trimmedCanvas.width = width;
        trimmedCanvas.height = height;
        const trimmedCtx = trimmedCanvas.getContext('2d');

        trimmedCtx.drawImage(
            this.canvas,
            minX, minY, width, height,
            0, 0, width, height
        );

        return trimmedCanvas.toDataURL('image/png');
    }

    /**
     * Generate typed signature as data URL
     * @returns {string}
     */
    getTypedSignatureDataUrl() {
        if (!this.typedText) return null;

        // Create canvas for typed signature
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');

        // Set font based on style
        let fontFamily;
        switch (this.fontStyle) {
            case 'cursive':
                fontFamily = "'Brush Script MT', cursive";
                break;
            case 'handwriting':
                fontFamily = "'Comic Sans MS', cursive";
                break;
            case 'formal':
                fontFamily = "'Times New Roman', serif";
                break;
            default:
                fontFamily = 'cursive';
        }

        const fontSize = 48;
        ctx.font = `${fontSize}px ${fontFamily}`;

        // Measure text
        const metrics = ctx.measureText(this.typedText);
        const textWidth = metrics.width;
        const textHeight = fontSize * 1.2;

        // Set canvas size with padding
        const padding = 20;
        canvas.width = textWidth + padding * 2;
        canvas.height = textHeight + padding * 2;

        // Re-set font after resize
        ctx.font = `${fontSize}px ${fontFamily}`;
        ctx.fillStyle = '#000000';
        ctx.textBaseline = 'middle';

        // Draw text
        ctx.fillText(this.typedText, padding, canvas.height / 2);

        return canvas.toDataURL('image/png');
    }

    /**
     * Set the typed text
     * @param {string} text - Signature text
     */
    setTypedText(text) {
        this.typedText = text;
    }

    /**
     * Set the font style for typed signature
     * @param {string} style - Font style ('cursive', 'handwriting', 'formal')
     */
    setFontStyle(style) {
        this.fontStyle = style;
    }

    /**
     * Set the mode
     * @param {string} mode - 'draw' or 'type'
     */
    setMode(mode) {
        this.mode = mode;
    }

    /**
     * Preview typed signature
     * @param {HTMLElement} previewElement - Element to show preview
     */
    updatePreview(previewElement) {
        previewElement.textContent = this.typedText || 'Preview';
        previewElement.className = 'sig-preview ' + this.fontStyle;
    }
}
