/**
 * PDF Export - Handles exporting annotated PDFs using pdf-lib
 */

const { PDFDocument, rgb, StandardFonts, degrees } = PDFLib;

export class PDFExporter {
    constructor() {
        this.fonts = {};
    }

    /**
     * Export the PDF with all annotations
     * @param {ArrayBuffer} originalPdfBytes - Original PDF bytes
     * @param {Array} allAnnotations - Annotations from all pages
     * @param {number} scale - Current canvas scale
     * @returns {Promise<Uint8Array>} - Modified PDF bytes
     */
    async exportPDF(originalPdfBytes, allAnnotations, scale) {
        // Load the original PDF
        const pdfDoc = await PDFDocument.load(originalPdfBytes);

        // Register fontkit for custom fonts
        pdfDoc.registerFontkit(fontkit);

        // Embed standard fonts
        this.fonts.helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
        this.fonts.helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        this.fonts.timesRoman = await pdfDoc.embedFont(StandardFonts.TimesRoman);
        this.fonts.courier = await pdfDoc.embedFont(StandardFonts.Courier);

        const pages = pdfDoc.getPages();
        const scaleFactor = 1 / scale;
        const auditEntries = [];

        // Process each page's annotations
        for (const pageData of allAnnotations) {
            const pageIndex = pageData.pageNum - 1;
            if (pageIndex >= pages.length) continue;

            const page = pages[pageIndex];
            const { width: pageWidth, height: pageHeight } = page.getSize();

            for (const annotation of pageData.annotations) {
                await this.drawAnnotation(pdfDoc, page, annotation, scaleFactor, pageHeight, auditEntries, pageData.pageNum);
            }
        }

        if (auditEntries.length > 0) {
            const auditJson = JSON.stringify(auditEntries, null, 2);
            const auditBytes = new TextEncoder().encode(auditJson);
            await pdfDoc.attach(auditBytes, 'signatures-audit.json', {
                mimeType: 'application/json',
                description: 'Signature audit trail'
            });
        }

        pdfDoc.setModificationDate(new Date());
        pdfDoc.setProducer('Free PDF Editor');
        pdfDoc.setCreator('Free PDF Editor');

        // Save and return the modified PDF
        const modifiedPdfBytes = await pdfDoc.save();
        return modifiedPdfBytes;
    }

    /**
     * Draw a single annotation on a PDF page
     */
    async drawAnnotation(pdfDoc, page, annotation, scaleFactor, pageHeight, auditEntries = null, pageNum = null) {
        const obj = annotation.object;
        const type = annotation.type;

        // Get object bounds (accounting for Fabric.js coordinate system)
        const bounds = obj.getBoundingRect();

        // Convert from canvas coordinates to PDF coordinates
        // PDF origin is bottom-left, canvas origin is top-left
        const pdfX = bounds.left * scaleFactor;
        const pdfY = pageHeight - (bounds.top + bounds.height) * scaleFactor;
        const pdfWidth = bounds.width * scaleFactor;
        const pdfHeight = bounds.height * scaleFactor;

        switch (type) {
            case 'text':
            case 'i-text':
                await this.drawText(page, obj, scaleFactor, pageHeight);
                break;

            case 'whiteout':
            case 'rect':
                this.drawRect(page, obj, scaleFactor, pageHeight);
                break;

            case 'path':
            case 'draw':
                await this.drawPath(pdfDoc, page, obj, scaleFactor, pageHeight);
                break;

            case 'signature':
            case 'image':
                await this.drawImage(pdfDoc, page, obj, scaleFactor, pageHeight);
                // Handle signature audit trail
                if (type === 'signature' && obj._signatureMeta && auditEntries !== null && pageNum !== null) {
                    const meta = { ...obj._signatureMeta };
                    meta.pageNum = pageNum;
                    meta.bounds = {
                        left: bounds.left * scaleFactor,
                        top: pageHeight - (bounds.top + bounds.height) * scaleFactor,
                        width: bounds.width * scaleFactor,
                        height: bounds.height * scaleFactor
                    };
                    auditEntries.push(meta);
                }
                break;

            case 'textfield':
                await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'text');
                break;

            case 'checkbox':
                await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'checkbox');
                break;

            case 'group':
                // Handle grouped objects
                if (obj._annotationType === 'textfield') {
                    await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'text');
                } else if (obj._annotationType === 'checkbox') {
                    await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'checkbox');
                }
                break;
        }
    }

    /**
     * Draw text annotation
     */
    async drawText(page, obj, scaleFactor, pageHeight) {
        const text = obj.text || '';
        if (!text.trim()) return;

        // Map font family to embedded font
        let font = this.fonts.helvetica;
        const fontFamily = (obj.fontFamily || 'Arial').toLowerCase();
        if (fontFamily.includes('times')) {
            font = this.fonts.timesRoman;
        } else if (fontFamily.includes('courier') || fontFamily.includes('mono')) {
            font = this.fonts.courier;
        }

        // Calculate position
        const left = (obj.left || 0) * scaleFactor;
        const fontSize = (obj.fontSize || 16) * (obj.scaleY || 1) * scaleFactor;

        // PDF text is drawn from baseline, so adjust Y position
        const top = pageHeight - (obj.top || 0) * scaleFactor - fontSize;

        // Parse color
        const color = this.parseColor(obj.fill || '#000000');

        // Draw each line of text
        const lines = text.split('\n');
        let currentY = top;

        for (const line of lines) {
            if (line.trim()) {
                page.drawText(line, {
                    x: left,
                    y: currentY,
                    size: fontSize,
                    font: font,
                    color: rgb(color.r, color.g, color.b)
                });
            }
            currentY -= fontSize * 1.2; // Line height
        }
    }

    /**
     * Draw rectangle (whiteout)
     */
    drawRect(page, obj, scaleFactor, pageHeight) {
        const left = (obj.left || 0) * scaleFactor;
        const width = (obj.width || 0) * (obj.scaleX || 1) * scaleFactor;
        const height = (obj.height || 0) * (obj.scaleY || 1) * scaleFactor;
        const top = pageHeight - (obj.top || 0) * scaleFactor - height;

        const fillColor = this.parseColor(obj.fill || '#ffffff');

        page.drawRectangle({
            x: left,
            y: top,
            width: width,
            height: height,
            color: rgb(fillColor.r, fillColor.g, fillColor.b)
        });

        // Draw border if present
        if (obj.stroke && obj.strokeWidth) {
            const strokeColor = this.parseColor(obj.stroke);
            page.drawRectangle({
                x: left,
                y: top,
                width: width,
                height: height,
                borderColor: rgb(strokeColor.r, strokeColor.g, strokeColor.b),
                borderWidth: obj.strokeWidth * scaleFactor
            });
        }
    }

    /**
     * Draw freehand path
     */
    async drawPath(pdfDoc, page, obj, scaleFactor, pageHeight) {
        try {
            const dataUrl = obj.toDataURL?.({ format: 'png', multiplier: 2 });
            if (!dataUrl || !dataUrl.startsWith('data:image')) return;

            const bounds = obj.getBoundingRect();
            const left = bounds.left * scaleFactor;
            const width = bounds.width * scaleFactor;
            const height = bounds.height * scaleFactor;
            const top = pageHeight - bounds.top * scaleFactor - height;

            const imageBytes = this.dataUrlToBytes(dataUrl);
            const pdfImage = await pdfDoc.embedPng(imageBytes);

            page.drawImage(pdfImage, {
                x: left,
                y: top,
                width: width,
                height: height
            });
        } catch (e) {
            console.warn('Export path as image failed:', e);
        }
    }

    /**
     * Convert Fabric.js path to SVG path string
     */
    fabricPathToSvg(obj, scaleFactor, pageHeight) {
        const pathData = obj.path;
        if (!pathData) return '';

        let svgPath = '';
        const offsetX = (obj.left || 0);
        const offsetY = (obj.top || 0);

        for (const cmd of pathData) {
            switch (cmd[0]) {
                case 'M':
                    const mx = (cmd[1] + offsetX) * scaleFactor;
                    const my = pageHeight - (cmd[2] + offsetY) * scaleFactor;
                    svgPath += `M ${mx} ${my} `;
                    break;
                case 'L':
                    const lx = (cmd[1] + offsetX) * scaleFactor;
                    const ly = pageHeight - (cmd[2] + offsetY) * scaleFactor;
                    svgPath += `L ${lx} ${ly} `;
                    break;
                case 'Q':
                    const qcx = (cmd[1] + offsetX) * scaleFactor;
                    const qcy = pageHeight - (cmd[2] + offsetY) * scaleFactor;
                    const qx = (cmd[3] + offsetX) * scaleFactor;
                    const qy = pageHeight - (cmd[4] + offsetY) * scaleFactor;
                    svgPath += `Q ${qcx} ${qcy} ${qx} ${qy} `;
                    break;
                case 'C':
                    const c1x = (cmd[1] + offsetX) * scaleFactor;
                    const c1y = pageHeight - (cmd[2] + offsetY) * scaleFactor;
                    const c2x = (cmd[3] + offsetX) * scaleFactor;
                    const c2y = pageHeight - (cmd[4] + offsetY) * scaleFactor;
                    const cx = (cmd[5] + offsetX) * scaleFactor;
                    const cy = pageHeight - (cmd[6] + offsetY) * scaleFactor;
                    svgPath += `C ${c1x} ${c1y} ${c2x} ${c2y} ${cx} ${cy} `;
                    break;
            }
        }

        return svgPath;
    }

    /**
     * Draw image (signature)
     */
    async drawImage(pdfDoc, page, obj, scaleFactor, pageHeight) {
        try {
            // Get image source
            const src = obj.getSrc ? obj.getSrc() : obj._element?.src;
            if (!src) return;

            // Embed image in PDF
            let pdfImage;
            if (src.startsWith('data:image/png')) {
                const imageBytes = this.dataUrlToBytes(src);
                pdfImage = await pdfDoc.embedPng(imageBytes);
            } else if (src.startsWith('data:image/jpeg') || src.startsWith('data:image/jpg')) {
                const imageBytes = this.dataUrlToBytes(src);
                pdfImage = await pdfDoc.embedJpg(imageBytes);
            } else {
                // Try to fetch and embed
                const response = await fetch(src);
                const imageBytes = await response.arrayBuffer();
                const contentType = response.headers.get('content-type');
                if (contentType?.includes('png')) {
                    pdfImage = await pdfDoc.embedPng(imageBytes);
                } else {
                    pdfImage = await pdfDoc.embedJpg(imageBytes);
                }
            }

            // Calculate position and size
            const width = (obj.width || pdfImage.width) * (obj.scaleX || 1) * scaleFactor;
            const height = (obj.height || pdfImage.height) * (obj.scaleY || 1) * scaleFactor;
            const left = (obj.left || 0) * scaleFactor;
            const top = pageHeight - (obj.top || 0) * scaleFactor - height;

            page.drawImage(pdfImage, {
                x: left,
                y: top,
                width: width,
                height: height
            });
        } catch (error) {
            console.error('Error embedding image:', error);
        }
    }

    /**
     * Create actual PDF form field (for bulk filling)
     */
    async createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, fieldType) {
        const bounds = obj.getBoundingRect();
        const left = bounds.left * scaleFactor;
        const width = bounds.width * scaleFactor;
        const height = bounds.height * scaleFactor;
        const top = pageHeight - bounds.top * scaleFactor - height;
        const fieldName = obj._fieldName || '';

        // Only create form field if it has a name
        if (!fieldName) {
            // Fall back to visual representation if no name
            this.drawFormField(page, obj, scaleFactor, pageHeight, fieldType);
            return;
        }

        try {
            if (fieldType === 'text') {
                const textField = pdfDoc.form.createTextField(fieldName);
                textField.addToPage(page, {
                    x: left,
                    y: top,
                    width: width,
                    height: height,
                    borderColor: rgb(0.15, 0.39, 0.92),
                    borderWidth: 1,
                    backgroundColor: rgb(1, 1, 1)
                });
                // Set value if present
                if (obj._fieldValue) {
                    textField.setText(obj._fieldValue);
                }
            } else if (fieldType === 'checkbox') {
                const checkbox = pdfDoc.form.createCheckBox(fieldName);
                checkbox.addToPage(page, {
                    x: left,
                    y: top,
                    width: width,
                    height: height,
                    borderColor: rgb(0.15, 0.39, 0.92),
                    borderWidth: 1,
                    backgroundColor: rgb(1, 1, 1)
                });
                if (obj._checked) {
                    checkbox.check();
                }
            }
        } catch (error) {
            console.warn('Failed to create PDF form field, falling back to visual:', error);
            // Fall back to visual representation
            this.drawFormField(page, obj, scaleFactor, pageHeight, fieldType);
        }
    }

    /**
     * Draw form field visual representation
     */
    drawFormField(page, obj, scaleFactor, pageHeight, fieldType) {
        const bounds = obj.getBoundingRect();
        const left = bounds.left * scaleFactor;
        const width = bounds.width * scaleFactor;
        const height = bounds.height * scaleFactor;
        const top = pageHeight - bounds.top * scaleFactor - height;

        // Draw field background
        page.drawRectangle({
            x: left,
            y: top,
            width: width,
            height: height,
            color: rgb(1, 1, 1),
            borderColor: rgb(0.15, 0.39, 0.92),
            borderWidth: 1
        });

        if (fieldType === 'checkbox' && obj._checked) {
            // Draw checkmark
            const padding = width * 0.2;
            page.drawLine({
                start: { x: left + padding, y: top + height * 0.5 },
                end: { x: left + width * 0.4, y: top + padding },
                thickness: 2,
                color: rgb(0.15, 0.39, 0.92)
            });
            page.drawLine({
                start: { x: left + width * 0.4, y: top + padding },
                end: { x: left + width - padding, y: top + height - padding },
                thickness: 2,
                color: rgb(0.15, 0.39, 0.92)
            });
        }
    }

    /**
     * Parse color string to RGB values (0-1 range)
     */
    parseColor(colorStr) {
        if (!colorStr) return { r: 0, g: 0, b: 0 };

        // Handle hex colors
        if (colorStr.startsWith('#')) {
            const hex = colorStr.slice(1);
            const r = parseInt(hex.substr(0, 2), 16) / 255;
            const g = parseInt(hex.substr(2, 2), 16) / 255;
            const b = parseInt(hex.substr(4, 2), 16) / 255;
            return { r, g, b };
        }

        // Handle rgb/rgba
        const rgbMatch = colorStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
        if (rgbMatch) {
            return {
                r: parseInt(rgbMatch[1]) / 255,
                g: parseInt(rgbMatch[2]) / 255,
                b: parseInt(rgbMatch[3]) / 255
            };
        }

        // Handle named colors (basic)
        const namedColors = {
            'black': { r: 0, g: 0, b: 0 },
            'white': { r: 1, g: 1, b: 1 },
            'red': { r: 1, g: 0, b: 0 },
            'green': { r: 0, g: 0.5, b: 0 },
            'blue': { r: 0, g: 0, b: 1 }
        };

        return namedColors[colorStr.toLowerCase()] || { r: 0, g: 0, b: 0 };
    }

    /**
     * Convert data URL to byte array
     */
    dataUrlToBytes(dataUrl) {
        const base64 = dataUrl.split(',')[1];
        const binaryString = atob(base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes;
    }

    /**
     * Trigger download of the PDF
     */
    downloadPDF(pdfBytes, filename = 'edited-document.pdf') {
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }
}
