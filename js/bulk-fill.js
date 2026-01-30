/**
 * Bulk Fill Handler - Handles CSV-based bulk PDF filling
 */

const { PDFDocument, rgb } = PDFLib;

export class BulkFillHandler {
    /**
     * Parse CSV file content
     * @param {string} csvText - CSV file content
     * @returns {Array<Object>} Array of row objects with column headers as keys
     */
    parseCSV(csvText) {
        const lines = csvText.split('\n').map(line => line.trim()).filter(line => line);
        if (lines.length === 0) return [];

        // Parse header row
        const headers = this.parseCSVLine(lines[0]);
        
        // Parse data rows
        const rows = [];
        for (let i = 1; i < lines.length; i++) {
            const values = this.parseCSVLine(lines[i]);
            const row = {};
            headers.forEach((header, index) => {
                row[header] = values[index] || '';
            });
            rows.push(row);
        }

        return rows;
    }

    /**
     * Parse a single CSV line handling quoted values
     */
    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    // Escaped quote
                    current += '"';
                    i++;
                } else {
                    // Toggle quote state
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                // End of field
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        // Add last field
        result.push(current.trim());
        
        return result;
    }

    /**
     * Get page count of a PDF
     * @param {ArrayBuffer} pdfBytes - PDF file bytes
     * @returns {Promise<number>} Number of pages
     */
    async getPageCount(pdfBytes) {
        try {
            const pdfDoc = await PDFDocument.load(pdfBytes);
            return pdfDoc.getPageCount();
        } catch (error) {
            console.error('Error getting page count:', error);
            return 0;
        }
    }

    /**
     * Extract form field names from a PDF
     * @param {ArrayBuffer} pdfBytes - PDF file bytes
     * @returns {Promise<Array<string>>} Array of form field names
     */
    async extractFormFieldNames(pdfBytes) {
        try {
            const pdfDoc = await PDFDocument.load(pdfBytes);
            const form = pdfDoc.getForm();
            const fields = form.getFields();
            
            const fieldNames = fields.map(field => field.getName());
            return fieldNames;
        } catch (error) {
            console.error('Error extracting form fields:', error);
            return [];
        }
    }

    /**
     * Fill a PDF template with data from a CSV row
     * @param {ArrayBuffer} templateBytes - Template PDF bytes
     * @param {Object} rowData - Data object with field names as keys
     * @param {Object} fieldMapping - Mapping from CSV column names to PDF field names
     * @returns {Promise<Uint8Array>} Filled PDF bytes
     */
    async fillPDF(templateBytes, rowData, fieldMapping) {
        const pdfDoc = await PDFDocument.load(templateBytes);
        const form = pdfDoc.getForm();
        const fields = form.getFields();

        // Create reverse mapping: PDF field name -> CSV column name
        const reverseMapping = {};
        Object.keys(fieldMapping).forEach(csvColumn => {
            const pdfFieldName = fieldMapping[csvColumn];
            if (pdfFieldName) {
                reverseMapping[pdfFieldName] = csvColumn;
            }
        });

        // Fill form fields
        fields.forEach(field => {
            const fieldName = field.getName();
            const csvColumn = reverseMapping[fieldName];
            
            if (csvColumn && rowData[csvColumn] !== undefined && rowData[csvColumn] !== '') {
                try {
                    const value = String(rowData[csvColumn]).trim();
                    
                    if (field.constructor.name === 'PDFTextField') {
                        field.setText(value);
                    } else if (field.constructor.name === 'PDFCheckBox') {
                        // For checkboxes, treat 'true', '1', 'yes', 'checked' as checked
                        const checked = ['true', '1', 'yes', 'checked', 'x'].includes(value.toLowerCase());
                        if (checked) {
                            field.check();
                        } else {
                            field.uncheck();
                        }
                    }
                } catch (error) {
                    console.warn(`Failed to fill field ${fieldName}:`, error);
                }
            }
        });

        // Flatten form to prevent further editing
        form.flatten();

        const filledBytes = await pdfDoc.save();
        return filledBytes;
    }

    /**
     * Generate filename for a filled PDF
     * @param {string} baseFilename - Base filename template
     * @param {Object} rowData - CSV row data
     * @param {number} index - Row index
     * @returns {string} Generated filename
     */
    generateFilename(baseFilename, rowData, index) {
        let filename = baseFilename || 'filled-document';
        
        // Replace placeholders like {{column_name}} or {{row}}
        filename = filename.replace(/\{\{(\w+)\}\}/g, (match, columnName) => {
            if (columnName === 'row') {
                return String(index + 1);
            }
            return rowData[columnName] || match;
        });
        
        // If no extension, add .pdf
        if (!filename.toLowerCase().endsWith('.pdf')) {
            filename += '.pdf';
        }
        
        // If filename is still generic, add index
        if (filename === 'filled-document.pdf' || filename === 'document.pdf' || filename === 'document-{{row}}.pdf') {
            filename = `filled-document-${index + 1}.pdf`;
        }
        
        return filename;
    }

    /**
     * Download a PDF file
     */
    downloadPDF(pdfBytes, filename) {
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

    /**
     * Process bulk fill: parse CSV, extract fields, map columns, fill and download PDFs
     * @param {ArrayBuffer} templateBytes - Template PDF bytes
     * @param {string} csvText - CSV file content
     * @param {Object} fieldMapping - Mapping from CSV columns to PDF field names
     * @param {string} filenameTemplate - Filename template
     * @param {Function} progressCallback - Callback for progress updates (current, total)
     * @returns {Promise<void>}
     */
    async processBulkFill(templateBytes, csvText, fieldMapping, filenameTemplate, progressCallback) {
        const rows = this.parseCSV(csvText);
        
        if (rows.length === 0) {
            throw new Error('CSV file contains no data rows');
        }

        for (let i = 0; i < rows.length; i++) {
            if (progressCallback) {
                progressCallback(i + 1, rows.length);
            }

            const rowData = rows[i];
            const filledBytes = await this.fillPDF(templateBytes, rowData, fieldMapping);
            const filename = this.generateFilename(filenameTemplate, rowData, i);
            
            // Download with a small delay to avoid browser blocking multiple downloads
            if (i === 0) {
                this.downloadPDF(filledBytes, filename);
            } else {
                // Use setTimeout to stagger downloads
                await new Promise(resolve => setTimeout(resolve, 100));
                this.downloadPDF(filledBytes, filename);
            }
        }
    }
}
