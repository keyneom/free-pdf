/**
 * PDF Export - Handles exporting annotated PDFs using pdf-lib
 */

const { PDFDocument, rgb, StandardFonts, degrees } = PDFLib;
import { parseSigningMetadata, buildSigningKeywords } from './signing-metadata.js';

/**
 * pdf-lib setKeywords() requires an array of strings. buildSigningKeywords returns a string.
 * Ensure we always pass a proper array.
 */
function toKeywordsArray(payload) {
    if (!payload) return null;
    if (Array.isArray(payload)) {
        const arr = payload.filter((x) => typeof x === 'string' && x.length > 0).map(String);
        return arr.length > 0 ? arr : null;
    }
    const s = String(payload).trim();
    return s ? [s] : null;
}

export class PDFExporter {
    constructor() {
        this.fonts = {};
    }

    /**
     * Export a PDF with annotations.
     *
     * New mode (supports reorder/append via viewPages):
     * @param {{ docBytesById: Map<string, ArrayBuffer>; viewPages: Array<{id: string; docId: string; sourcePageNum: number; rotation?: number}>; annotationsByPageId: Map<string, any[]>; scale: number }} input
     * @returns {Promise<Uint8Array>} - Modified PDF bytes
     */
    async exportPDF(input, allAnnotationsLegacy, scaleLegacy) {
        // Backward compatibility (older call signature)
        if (input instanceof ArrayBuffer) {
            const pdfDoc = await PDFDocument.load(input);
            pdfDoc.registerFontkit(fontkit);

            // Embed standard fonts
            this.fonts.helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
            this.fonts.helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
            this.fonts.timesRoman = await pdfDoc.embedFont(StandardFonts.TimesRoman);
            this.fonts.courier = await pdfDoc.embedFont(StandardFonts.Courier);

            const pages = pdfDoc.getPages();
            const scaleFactor = 1 / (scaleLegacy || 1);
            const auditEntries = [];

            for (const pageData of allAnnotationsLegacy || []) {
                const pageIndex = pageData.pageNum - 1;
                if (pageIndex >= pages.length) continue;
                const page = pages[pageIndex];
                const { height: pageHeight } = page.getSize();
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
                const signers = auditEntries.map((e) => ({ name: e.signerName || '', timestamp: e.timestamp || '' }));
                const keywordsPayload = buildSigningKeywords({ signers });
                const kwArray = toKeywordsArray(keywordsPayload);
                if (kwArray) pdfDoc.setKeywords(kwArray);
            }

            pdfDoc.setModificationDate(new Date());
            pdfDoc.setProducer('Free PDF Editor');
            pdfDoc.setCreator('Free PDF Editor');
            return await pdfDoc.save();
        }

        const { docBytesById, viewPages, annotationsByPageId, scale, mainDocId, signingFlowMeta } = input;

        // Load all source PDFs with pdf-lib
        const srcDocs = new Map();
        for (const [docId, bytes] of docBytesById.entries()) {
            srcDocs.set(docId, await PDFDocument.load(bytes));
        }
        const mainDocIdResolved = mainDocId || viewPages?.[0]?.docId;

        // Create output PDF
        const outDoc = await PDFDocument.create();

        // Register fontkit for custom fonts
        outDoc.registerFontkit(fontkit);

        // Embed standard fonts
        this.fonts.helvetica = await outDoc.embedFont(StandardFonts.Helvetica);
        this.fonts.helveticaBold = await outDoc.embedFont(StandardFonts.HelveticaBold);
        this.fonts.timesRoman = await outDoc.embedFont(StandardFonts.TimesRoman);
        this.fonts.courier = await outDoc.embedFont(StandardFonts.Courier);

        const scaleFactor = 1 / scale;
        const auditEntries = [];
        // Per-export run cache for PDF form fields that may be referenced multiple times
        this._formFieldCache = { dropdown: new Map(), radio: new Map() };

        // Build output pages in view order
        for (const vp of viewPages) {
            const src = srcDocs.get(vp.docId);
            if (!src) continue;
            const [copied] = await outDoc.copyPages(src, [vp.sourcePageNum - 1]);
            if (vp.rotation) {
                copied.setRotation(degrees(vp.rotation));
            }
            outDoc.addPage(copied);
        }

        // Draw annotations in view order
        const outPages = outDoc.getPages();
        for (let i = 0; i < viewPages.length; i++) {
            const vp = viewPages[i];
            const page = outPages[i];
            if (!page) continue;

            const { height: pageHeight } = page.getSize();
            const pageAnnotations = annotationsByPageId.get(vp.id) || [];
            for (const annotation of pageAnnotations) {
                await this.drawAnnotation(outDoc, page, annotation, scaleFactor, pageHeight, auditEntries, i + 1);
            }
        }

        if (auditEntries.length > 0) {
            const auditJson = JSON.stringify(auditEntries, null, 2);
            const auditBytes = new TextEncoder().encode(auditJson);
            await outDoc.attach(auditBytes, 'signatures-audit.json', {
                mimeType: 'application/json',
                description: 'Signature audit trail'
            });
        }

        // Read existing signing metadata from main source (so expectedSigners persist) and write merged payload to Keywords
        let previousMeta = null;
        if (mainDocIdResolved) {
            const mainSrc = srcDocs.get(mainDocIdResolved);
            if (mainSrc && typeof mainSrc.getKeywords === 'function') {
                const kw = mainSrc.getKeywords();
                if (typeof kw === 'string') previousMeta = parseSigningMetadata(kw);
            }
        }
        const expectedSigners = signingFlowMeta?.expectedSigners ?? previousMeta?.expectedSigners ?? [];
        const emailTemplate = signingFlowMeta?.emailTemplate ?? previousMeta?.emailTemplate;
        const originalSenderEmail = signingFlowMeta?.originalSenderEmail ?? previousMeta?.originalSenderEmail;
        const completionToEmails = signingFlowMeta?.completionToEmails ?? previousMeta?.completionToEmails;
        const completionCcEmails = signingFlowMeta?.completionCcEmails ?? previousMeta?.completionCcEmails;
        const completionBccEmails = signingFlowMeta?.completionBccEmails ?? previousMeta?.completionBccEmails;
        const lockedSignatureFields = signingFlowMeta?.lockedSignatureFields ?? previousMeta?.lockedSignatureFields;
        const lockedFormFields = signingFlowMeta?.lockedFormFields ?? previousMeta?.lockedFormFields;
        const documentStage = signingFlowMeta?.documentStage ?? previousMeta?.documentStage;
        const hashChain = signingFlowMeta?.hashChain ?? previousMeta?.hashChain;
        const signers = auditEntries.map((e) => ({ name: e.signerName || '', timestamp: e.timestamp || '' }));
        const keywordsPayload = buildSigningKeywords({ signers, expectedSigners, emailTemplate, originalSenderEmail, completionToEmails, completionCcEmails, completionBccEmails, lockedSignatureFields, lockedFormFields, documentStage, hashChain });
        const kwArray = toKeywordsArray(keywordsPayload);
        if (kwArray) outDoc.setKeywords(kwArray);

        outDoc.setModificationDate(new Date());
        outDoc.setProducer('Free PDF Editor');
        outDoc.setCreator('Free PDF Editor');

        // Save and return the modified PDF
        return await outDoc.save();
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

            case 'highlight':
            case 'whiteout':
            case 'rect':
                this.drawRect(page, obj, scaleFactor, pageHeight);
                break;

            case 'ellipse':
                // Export ellipse as an image for fidelity (stroke/fill/opacity)
                await this.drawObjectAsImage(pdfDoc, page, obj, scaleFactor, pageHeight);
                break;

            case 'underline':
            case 'strike':
            case 'arrow':
            case 'stamp':
            case 'note':
            case 'path':
            case 'draw':
                await this.drawObjectAsImage(pdfDoc, page, obj, scaleFactor, pageHeight);
                break;

            case 'signature':
            case 'image':
                await this.drawImage(pdfDoc, page, obj, scaleFactor, pageHeight);
                // Handle signature audit trail (signerName stored in metadata only, not drawn on page)
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

            case 'signature-field':
                // Render signature field as a dashed box with label text
                await this.drawSignatureField(pdfDoc, page, obj, scaleFactor, pageHeight);
                break;

            case 'textfield':
                await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'text');
                break;

            case 'checkbox':
                await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'checkbox');
                break;

            case 'dropdown':
                await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'dropdown');
                break;

            case 'radio':
                await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'radio');
                break;

            case 'date':
                await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'text');
                break;

            case 'group':
                // Handle grouped objects
                if (obj._annotationType === 'textfield') {
                    await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'text');
                } else if (obj._annotationType === 'checkbox') {
                    await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'checkbox');
                } else if (obj._annotationType === 'dropdown') {
                    await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'dropdown');
                } else if (obj._annotationType === 'radio') {
                    await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'radio');
                } else if (obj._annotationType === 'date') {
                    await this.createFormField(pdfDoc, page, obj, scaleFactor, pageHeight, 'text');
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
    async drawObjectAsImage(pdfDoc, page, obj, scaleFactor, pageHeight) {
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
            console.warn('Export object as image failed:', e);
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
     * Draw signature field (empty placeholder box with label).
     * Also creates a read-only text form field so it appears in the sidebar when the PDF is reopened.
     */
    async drawSignatureField(pdfDoc, page, obj, scaleFactor, pageHeight) {
        const { rgb } = PDFLib;
        const bounds = obj.getBoundingRect();
        const left = bounds.left * scaleFactor;
        const width = bounds.width * scaleFactor;
        const height = bounds.height * scaleFactor;
        const top = pageHeight - bounds.top * scaleFactor - height;
        const label = (obj._signatureFieldLabel || 'Signature').trim();

        // Draw dashed rectangle border
        page.drawRectangle({
            x: left,
            y: top,
            width: width,
            height: height,
            borderColor: rgb(0.6, 0.6, 0.6),
            borderWidth: 1.5,
            borderDashArray: [3, 3],
            color: rgb(1, 1, 0.9),
            opacity: 0.3
        });

        // Draw label text (two lines, centered) - use font and proper line spacing
        const fontSize = Math.min(12, height / 4);
        const font = this.fonts.helvetica;
        const lineHeight = fontSize * 1.2;
        const lines = [label || 'Signature', '(Sign here)'];
        const totalTextHeight = lines.length * lineHeight - (lineHeight - fontSize);
        let y = top + (height - totalTextHeight) / 2 + fontSize;

        for (const line of lines) {
            const lineWidth = line.length * fontSize * 0.55; // Approximate width for Helvetica
            const x = left + (width - lineWidth) / 2;
            page.drawText(line, {
                x,
                y,
                size: fontSize,
                font,
                color: rgb(0.4, 0.4, 0.4)
            });
            y -= lineHeight;
        }

        // Create a read-only text form field so it appears in form field lists when PDF is reopened.
        // pdf-lib has no createSignature(); using a text field with sig_ prefix for our metadata.
        const baseName = (label || 'Signature').replace(/[^a-zA-Z0-9_-]/g, '_') || 'Signature';
        const fieldName = `sig_${baseName}_${Math.round(left)}_${Math.round(top)}`;
        try {
            const textField = pdfDoc.form.createTextField(fieldName);
            textField.enableReadOnly();
            textField.addToPage(page, {
                x: left,
                y: top,
                width: width,
                height: height,
                borderWidth: 0,
                backgroundColor: rgb(1, 1, 1),
                borderColor: rgb(1, 1, 1)
            });
        } catch (e) {
            console.warn('Could not create signature form field, using visual only:', e);
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
            } else if (fieldType === 'dropdown') {
                const cache = (this._formFieldCache ||= { dropdown: new Map(), radio: new Map() });
                let dd = cache.dropdown.get(fieldName);
                if (!dd) {
                    dd = pdfDoc.form.createDropdown(fieldName);
                    const opts = (obj._options || []).map((s) => String(s));
                    if (opts.length) dd.addOptions(opts);
                    cache.dropdown.set(fieldName, dd);
                }
                dd.addToPage(page, {
                    x: left,
                    y: top,
                    width: width,
                    height: height,
                    borderColor: rgb(0.15, 0.39, 0.92),
                    borderWidth: 1,
                    backgroundColor: rgb(1, 1, 1)
                });
                if (obj._selectedOption) {
                    try {
                        dd.select(String(obj._selectedOption));
                    } catch {
                        // ignore invalid option
                    }
                }
            } else if (fieldType === 'radio') {
                const cache = (this._formFieldCache ||= { dropdown: new Map(), radio: new Map() });
                let rg = cache.radio.get(fieldName);
                if (!rg) {
                    rg = pdfDoc.form.createRadioGroup(fieldName);
                    cache.radio.set(fieldName, rg);
                }
                const value = String(obj._radioValue || `${fieldName}_${Math.round(left)}_${Math.round(top)}`);
                rg.addOptionToPage(value, page, {
                    x: left,
                    y: top,
                    width: width,
                    height: height,
                    borderColor: rgb(0.15, 0.39, 0.92),
                    borderWidth: 1,
                    backgroundColor: rgb(1, 1, 1)
                });
                if (obj._checked) {
                    try {
                        rg.select(value);
                    } catch {
                        // ignore
                    }
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
