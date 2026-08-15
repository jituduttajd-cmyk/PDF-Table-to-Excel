# PDF Table to Excel — Version 3

## Why Version 3

The previous browser-only version reconstructed tables from PDF text positions. Your uploaded NPPA/Gazette PDF contains ruled tables, multi-line cells and page-break continuation rows. Your original Python program already uses `pdfplumber.extract_tables()` for this job.

Version 3 therefore keeps the PWA interface on GitHub Pages but moves the actual extraction to a Python API using the same pdfplumber + openpyxl approach.

## Project

```text
pdf-table-to-excel-v3/
├── frontend/        # GitHub Pages PWA
└── backend/         # Python API
```

## 1. Deploy backend

The easiest deployment is a Python web service such as Render.

The backend folder contains:
- `app.py`
- `requirements.txt`
- `Procfile`
- `render.yaml`

Create a web service from this repository, using the `backend` directory as the service root if your provider asks for a root directory.

After deployment, verify:

`https://YOUR-API-DOMAIN/health`

It should return JSON similar to:

`{"engine":"pdfplumber+openpyxl","status":"ok"}`

## 2. Configure GitHub Pages frontend

Open:

`frontend/config.js`

Change:

```js
window.APP_CONFIG = {
  API_BASE_URL: ""
};
```

to:

```js
window.APP_CONFIG = {
  API_BASE_URL: "https://YOUR-API-DOMAIN"
};
```

Do not add a trailing slash.

Then publish the contents of `frontend/` through GitHub Pages.

## 3. Test

Open the GitHub Pages URL, select a PDF and click **Convert PDF to Excel**.

The browser sends the PDF to the Python API. The API returns an `.xlsx` file.

## Extraction behavior

The backend follows the extraction rules from the supplied Python program:

1. Uses `pdfplumber` table detection.
2. Identifies an English header by ASCII ratio.
3. Uses the `(1)`, `(2)`, `(3)` column-index row as a strong header signal.
4. Keeps the logical table open across pages.
5. Merges rows whose first column is blank and other columns contain continuation text.
6. Uses `openpyxl` to create a formatted workbook.
7. Can insert line breaks after commas while preserving numeric thousands separators.

For the supplied S.O.4264(E) PDF, the source table has six columns and 23 numbered data rows; the extraction logic is designed to preserve that structure and merge the row that continues across the page boundary.

## Important

GitHub Pages cannot execute Python. The frontend and backend must therefore be deployed separately.

Also note that this service processes the PDF on the Python server. It is not local-only processing like Version 2.

## Version 4 extraction improvement

Version 4 uses pdfplumber's ruled-line table detector (`find_tables`) with explicit line-based settings. It only accepts an English table after detecting the English header followed by the `(1)..(n)` index row. Continuation pages are accepted only when their ruled grid has the same geometry and column count. This prevents running page headers, page numbers and notes below the table from being treated as table data.

For the supplied S.O.4264(E)-31-07-2027.pdf, the extractor detects 1 English table, 6 columns and 23 data rows. The row 18 continuation on page 8 is merged back into row 18. Page 9 Notes are excluded because page 9 contains no ruled table.
