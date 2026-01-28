# Free PDF Editor

A fully client-side PDF editor that runs entirely in the browser. Built for GitHub Pages deployment - no server required.

## Project Objective

Create a **free, privacy-focused PDF editing tool** that:
- Works 100% in the browser (no file uploads to servers)
- Can be deployed as a static site on GitHub Pages
- Provides WYSIWYG editing capabilities for common PDF tasks
- Leverages open-source libraries for PDF rendering and manipulation

## Features

### Core Editing Tools
- **Text Boxes** - Add text annotations anywhere on the PDF
- **Whiteout/Redaction** - Cover existing content with white rectangles
- **Freehand Drawing** - Draw signatures and annotations directly on the document
- **Digital Signatures** - Create and insert signatures (drawn or typed)

### Form Fields
- **Text Input Fields** - Add fillable text fields
- **Checkboxes** - Add interactive checkboxes

### Document Navigation
- Multi-page PDF support with page navigation
- Zoom in/out controls
- Fit-to-width view

### Export
- Download edited PDF with all annotations flattened into the document

## Technical Architecture

### Libraries Used

| Library | Purpose | CDN |
|---------|---------|-----|
| [PDF.js](https://mozilla.github.io/pdf.js/) | PDF rendering in canvas | Mozilla's PDF rendering engine |
| [Fabric.js](http://fabricjs.com/) | Canvas-based WYSIWYG editing | Interactive canvas with object manipulation |
| [pdf-lib](https://pdf-lib.js.org/) | PDF modification and export | Create/modify PDFs in JavaScript |
| [fontkit](https://github.com/foliojs/fontkit) | Font embedding support | Required by pdf-lib for custom fonts |

### How It Works

1. **PDF Loading**: PDF.js loads and renders each page to a canvas element
2. **Annotation Layer**: Fabric.js overlays a transparent canvas on each page for interactive editing
3. **Object Manipulation**: Users add/move/resize annotations using Fabric.js
4. **Export**: pdf-lib reads the original PDF, embeds annotations as native PDF elements, and outputs a new PDF

### File Structure

```
free-pdf/
├── index.html          # Main HTML file
├── css/
│   └── styles.css      # Application styles
├── js/
│   ├── app.js          # Main application entry point
│   ├── pdf-handler.js  # PDF loading and rendering logic
│   ├── canvas-manager.js # Fabric.js canvas management
│   ├── signature-pad.js  # Signature drawing/typing
│   └── export.js       # PDF export functionality
└── README.md           # This file
```

## Development Roadmap

### Phase 1: Core Infrastructure ✅
- [x] Project structure setup
- [x] HTML/CSS layout with toolbar
- [x] PDF.js integration for rendering
- [x] Multi-page support

### Phase 2: WYSIWYG Canvas ✅
- [x] Fabric.js overlay on PDF pages
- [x] Selection and manipulation of objects
- [x] Undo/redo functionality

### Phase 3: Annotation Tools ✅
- [x] Text box tool
- [x] Whiteout/redaction tool
- [x] Freehand drawing tool
- [x] Signature pad (draw and type modes)

### Phase 4: Form Fields ✅
- [x] Text input field tool
- [x] Checkbox tool

### Phase 5: Export ✅
- [x] PDF export with pdf-lib
- [x] Flatten annotations into PDF
- [x] Preserve original PDF quality

### Phase 6: Polish ✅
- [x] Keyboard shortcuts (V, T, W, D, S, Delete, Ctrl+Z, Ctrl+Y)
- [x] Touch/mobile support (signature drawing)
- [x] Error handling and validation
- [ ] Performance optimization (ongoing)

## Getting Started

### Local Development

1. Clone the repository
2. Serve the directory with any static file server:
   ```bash
   # Using Python
   python -m http.server 8000

   # Using Node.js
   npx serve

   # Using PHP
   php -S localhost:8000
   ```
3. Open `http://localhost:8000` in your browser

### GitHub Pages Deployment

1. Push code to a GitHub repository
2. Go to Settings → Pages
3. Select source branch (usually `main`)
4. Site will be available at `https://username.github.io/repo-name`

## Privacy

**Your files never leave your computer.** All PDF processing happens entirely in the browser using JavaScript. No data is uploaded to any server.

## Browser Support

- Chrome 80+
- Firefox 75+
- Safari 14+
- Edge 80+

## License

MIT License - Free for personal and commercial use.
