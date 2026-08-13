# PDF Table to Excel PWA — Version 2

## GitHub Pages deployment

This is a static GitHub Pages application. No Python, Node.js server, API, or backend is required.

Upload the files in this folder to the root of a GitHub repository.

### Required files

- `index.html`
- `app.js`
- `manifest.json`
- `service-worker.js`
- `.nojekyll`
- `icons/icon-192.png`
- `icons/icon-512.png`

### GitHub Pages

1. Create/open the repository.
2. Upload all files and folders while preserving the structure.
3. Open **Settings → Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Choose `main` and `/ (root)`.
6. Save.
7. Open the published HTTPS URL.

GitHub Pages can publish static files directly from a branch. The included `.nojekyll` prevents an unnecessary Jekyll build for this static app.

### Important: do not double-click index.html

Use the GitHub Pages HTTPS URL for normal testing. PDF.js requires a web server for its worker and browser module loading; `file://` is not a supported test environment.

### Version 2 changes

- CSS is embedded in `index.html`, eliminating the missing-style-sheet problem shown in Version 1.
- All application paths are relative (`./...`), so the app works under a GitHub project URL such as `/repository-name/`.
- Uses browser ES modules for PDF.js.
- Registers the service worker only when served over HTTP(S).
- Includes `.nojekyll`.
- Shows a library-loading status.
- Uses a more reliable header + `(1)(2)(3)...` anchor strategy for reconstructing tables.
- Keeps the original Python program's key heuristics where they can be reproduced in the browser.

## Limitation

The original Python program calls `pdfplumber.extract_tables()`, which has its own ruled-table detection. A browser implementation based on PDF.js receives text items and coordinates instead of the same table detector. Therefore Version 2 reconstructs columns from PDF text positions.

For scanned/image-only PDFs, OCR is required and is not included in this version.
