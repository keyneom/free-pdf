# Free PDF Editor

A fully client-side PDF editor that runs entirely in the browser. Built for GitHub Pages deployment - no server required.

## Project Objective

Create a **free, privacy-focused PDF editing tool** that:
- Works 100% in the browser (no file uploads to servers)
- Can be deployed as a static site on GitHub Pages
- Provides WYSIWYG editing capabilities for common PDF tasks
- Leverages open-source libraries for PDF rendering and manipulation

## Features

### Page Operations
- **Pages sidebar** – Thumbnails for all pages; drag-and-drop to **reorder**
- **Append** – Add pages from another PDF (append to end)
- **Delete** – Remove selected pages (with confirmation)
- **Extract** – Download selected pages as a new PDF
- **Split** – Enter page ranges (e.g. `1-3,4-6`) to split the document into multiple PDFs

### Core Editing Tools
- **Text** – Add text annotations; font, size, color, bold, italic, alignment
- **Whiteout** – Cover existing content with white rectangles
- **Freehand drawing** – Draw directly on the document (color, stroke width)
- **Eraser** – Click an annotation to remove it
- **Highlight** – Semi-transparent highlight rectangles (color, opacity)
- **Underline / Strikethrough** – Draw lines under or through text
- **Shapes** – Rectangle, ellipse, arrow (stroke/fill, color)
- **Sticky note** – Add note callouts (double-click to edit)
- **Stamp** – Place text stamps (e.g. APPROVED, DRAFT); editable presets
- **Insert image** – Place PNG/JPG images on the page
- **Digital signatures** – Create signatures (draw or type). Intent/consent, signer identity, timestamps, and an embedded audit trail. Document SHA-256 hash stored for association.

### Form Fields
- **Text field** – Fillable text inputs
- **Checkbox** – Interactive checkboxes
- **Radio** – Radio button groups (shared field name)
- **Dropdown** – Select-one dropdowns (configurable options)
- **Date** – Date-style fillable fields (YYYY-MM-DD)

Form fields support **field names** for **Bulk Fill from CSV** (see below).

### Document Navigation
- Multi-page PDF support with prev/next and page-number input
- **Pages sidebar** with thumbnails; click to jump to a page
- Zoom in/out and fit-to-width
- Page navigation reflects current view order (after reorder/append/delete)

### Undo / Redo
- **Undo** (Ctrl/Cmd+Z) and **Redo** (Ctrl/Cmd+Y or Ctrl/Cmd+Shift+Z) for all edits, including freehand drawing
- Signature-pad drawing has its own undo/redo when the signature modal is open

### Export
- **Download** – Save edited PDF with annotations flattened into the document
- **Send via email** – Download the PDF and open your email client with a template-filled subject and body. Manually attach the downloaded file and send. Uses **email templates** (below).
- **Bulk Fill from CSV** – Use the current PDF (or an uploaded template) plus a CSV. Map CSV columns to form field names, then generate one filled PDF per CSV row; each downloads automatically.

### Email Templates
- Stored locally (localStorage). **Templates** in the toolbar: add, edit, delete, set default.
- **Placeholders**: `{{filename}}`, `{{date}}`, `{{signatureSummary}}`, `{{signerNames}}`, `{{pageCount}}`, `{{documentHash}}`, `{{attachmentNote}}`.
- **Export** templates as JSON; **Import** to merge (optional **Replace existing**). Use to sync across devices.

## Future Considerations

Ideas to keep in sync as the app evolves:

- **Signatures** – Initials, multiple saved signatures, placement helpers (e.g. “place on all pages”), timestamp/reason metadata.
- **Export & interoperability** – Flatten annotations vs keep editable; PDF/A-style export; compress/optimize output size.
- **Search & navigation** – Text search (find in document), thumbnails/outline (TOC) where available.
- **Security** – Password protection, restrict editing, remove metadata, redaction that permanently removes text.

## Technical Architecture

### Libraries Used

| Library | Purpose | CDN |
|---------|---------|-----|
| [PDF.js](https://mozilla.github.io/pdf.js/) | PDF rendering in canvas | Mozilla's PDF rendering engine |
| [Fabric.js](http://fabricjs.com/) | Canvas-based WYSIWYG editing | Interactive canvas with object manipulation |
| [pdf-lib](https://pdf-lib.js.org/) | PDF modification and export | Create/modify PDFs in JavaScript |
| [fontkit](https://github.com/foliojs/fontkit) | Font embedding support | Required by pdf-lib for custom fonts |

### How It Works

1. **PDF loading** – PDF.js loads one or more PDFs; pages are tracked in a **view-order** model (supports reorder, append, delete).
2. **Annotation layer** – Fabric.js overlays a transparent canvas on each page for interactive editing.
3. **Object manipulation** – Users add, move, resize, and delete annotations. Canvases and history are keyed by stable page IDs so reordering does not break annotations.
4. **Export** – pdf-lib builds a new PDF from the view-order model (copying source pages in order, including from appended docs), draws annotations onto each output page, and saves. Bulk Fill uses the same pipeline once per CSV row.

### File Structure

```
free-pdf/
├── index.html          # Main HTML, toolbar, modals, pages sidebar
├── css/
│   └── styles.css      # Application styles
├── js/
│   ├── app.js          # Main application, UI, pages sidebar, modals
│   ├── pdf-handler.js  # PDF load/render, multi-doc, view-order model
│   ├── canvas-manager.js # Fabric overlays, tools, history, form fields
│   ├── signature-pad.js  # Signature draw/type, undo/redo
│   ├── export.js       # PDF export, form fields, audit trail
│   ├── email-templates.js # Email template storage, placeholders, import/export
│   └── bulk-fill.js    # CSV parse, form-field mapping, bulk PDF generation
└── README.md           # This file
```

## Development Roadmap

### Phase 1: Core Infrastructure ✅
- [x] Project structure, HTML/CSS, toolbar
- [x] PDF.js integration, multi-page support

### Phase 2: WYSIWYG Canvas ✅
- [x] Fabric.js overlay per page, selection/manipulation
- [x] Undo/redo (including freehand and signature pad)

### Phase 3: Annotation Tools ✅
- [x] Text, whiteout, freehand, eraser
- [x] Highlight, underline, strikethrough; shapes (rect, ellipse, arrow)
- [x] Sticky note, stamp, insert image
- [x] Signature pad (draw and type)

### Phase 4: Form Fields ✅
- [x] Text field, checkbox
- [x] Radio, dropdown, date
- [x] Field names for bulk fill

### Phase 5: Page Operations ✅
- [x] View-order page model, multi-doc support
- [x] Pages sidebar, thumbnails, drag-and-drop reorder
- [x] Append, delete, extract, split

### Phase 6: Export & Workflows ✅
- [x] PDF export (view-order, annotations, form fields)
- [x] Send via email, email templates
- [x] Bulk fill from CSV

### Phase 7: Polish ✅
- [x] Keyboard shortcuts (V, T, W, D, S, Delete, Ctrl+Z, Ctrl+Y, Escape)
- [x] Touch/mobile (signature drawing)
- [x] Error handling and validation
- [ ] Performance optimization (ongoing)

## Getting Started

### Local Development

1. Clone the repository.
2. Serve the project with a static file server:
   ```bash
   # Python
   python -m http.server 8000

   # Node.js
   npx serve

   # PHP
   php -S localhost:8000
   ```
3. Open `http://localhost:8000` (or the port you used) in your browser.

### GitHub Pages Deployment

1. Push the repo to GitHub.
2. Settings → Pages → choose source branch (e.g. `main`).
3. Site will be at `https://<username>.github.io/<repo-name>`.

## Privacy

**Your files never leave your computer.** All PDF and CSV processing runs in the browser. Nothing is uploaded to any server.

## Browser Support

- Chrome 80+
- Firefox 75+
- Safari 14+
- Edge 80+

## License

MIT License — free for personal and commercial use.
