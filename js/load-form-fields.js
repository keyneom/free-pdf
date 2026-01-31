/**
 * Load form fields from a PDF using pdf-lib.
 * Returns field descriptors for creating canvas annotations when opening a PDF.
 */

const { PDFDocument } = PDFLib;

/**
 * Extract form field info from PDF bytes.
 * @param {ArrayBuffer} pdfBytes
 * @returns {Promise<Array<{ type: string; name: string; pageIndex: number; rect: { left: number; top: number; width: number; height: number }; value?: string; checked?: boolean; options?: string[]; signatureLabel?: string }>>}
 */
export async function loadFormFieldsFromPdf(pdfBytes) {
    const results = [];
    try {
        const pdfDoc = await PDFDocument.load(pdfBytes);
        const form = pdfDoc.getForm();
        const fields = form.getFields();
        const pages = pdfDoc.getPages();
        const pageRefToIndex = new Map();
        const refKey = (r) => (r && (r.toString?.() ?? `${r.objectNumber ?? ''}_${r.generationNumber ?? ''}`)) || '';

        pages.forEach((p, i) => {
            const r = p?.ref ?? p?.node?.ref;
            if (r != null) pageRefToIndex.set(refKey(r), i);
        });

        for (const field of fields) {
            const fieldType = field.constructor?.name || '';
            const name = field.getName?.() || '';
            const widgets = field.acroField?.getWidgets?.() || [];

            for (const widget of widgets) {
                const pageRef = typeof widget.P === 'function' ? widget.P() : widget.P;
                let pageIndex = 0;
                if (pageRef != null) {
                    const key = refKey(pageRef);
                    pageIndex = pageRefToIndex.has(key) ? pageRefToIndex.get(key) : 0;
                }

                const rect = widget.getRectangle?.() || { x: 0, y: 0, width: 0, height: 0 };
                const page = pages[pageIndex];
                const pageHeight = page?.getHeight?.() || 792;
                const pageWidth = page?.getWidth?.() || 612;

                const left = rect.x ?? 0;
                const width = rect.width ?? 0;
                const height = rect.height ?? 0;
                const bottom = rect.y ?? 0;
                const top = pageHeight - bottom - height;

                const descriptor = {
                    type: mapFieldType(fieldType, name),
                    name,
                    pageIndex,
                    rect: { left, top, width, height },
                    pageHeight,
                    pageWidth
                };

                if (fieldType === 'PDFTextField') {
                    descriptor.value = field.getText?.() ?? '';
                    if (name.startsWith('sig_')) {
                        descriptor.signatureLabel = parseSignatureLabel(name);
                    }
                } else if (fieldType === 'PDFCheckBox') {
                    descriptor.checked = field.isChecked?.() ?? false;
                } else if (fieldType === 'PDFDropdown') {
                    try {
                        descriptor.options = field.getOptions?.()?.map((o) => String(o)) ?? [];
                        descriptor.value = field.getSelected?.()?.map((o) => String(o))?.[0] ?? '';
                    } catch {
                        descriptor.options = [];
                        descriptor.value = '';
                    }
                } else if (fieldType === 'PDFRadioGroup') {
                    try {
                        descriptor.options = field.getOptions?.()?.map((o) => String(o)) ?? [];
                        descriptor.value = field.getSelected?.() ?? '';
                    } catch {
                        descriptor.options = [];
                        descriptor.value = '';
                    }
                } else if (fieldType === 'PDFSignature') {
                    descriptor.signatureLabel = name;
                }

                results.push(descriptor);
            }
        }
    } catch (e) {
        console.warn('Could not load form fields from PDF:', e);
    }
    return results;
}

function mapFieldType(pdfType, name) {
    if (name.startsWith('sig_')) return 'signature-field';
    if (pdfType === 'PDFTextField') return 'textfield';
    if (pdfType === 'PDFCheckBox') return 'checkbox';
    if (pdfType === 'PDFDropdown') return 'dropdown';
    if (pdfType === 'PDFRadioGroup') return 'radio';
    if (pdfType === 'PDFSignature') return 'signature-field';
    return 'textfield';
}

function parseSignatureLabel(name) {
    const m = name.match(/^sig_(.+)_\d+_\d+$/);
    return m ? m[1].replace(/_/g, ' ') : 'Signature';
}
