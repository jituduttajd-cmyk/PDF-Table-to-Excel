# PDF → Excel Converter (PWA)

A fully client-side app: PDFs are parsed and converted to `.xlsx` entirely in
the browser using **pdf.js** (text extraction) and **ExcelJS** (styled
`.xlsx` generation). Nothing is uploaded anywhere.

## Files
- `index.html` — the app (UI + conversion logic)
- `manifest.json` — makes it installable as a PWA
- `service-worker.js` — caches the app shell so it works offline after first load
- `icon-192.png`, `icon-512.png` — app icons

## Running it locally
PWAs require being served over `http(s)://`, not opened directly as a
`file://` URL (service workers won't register otherwise). From this folder:

```bash
python3 -m http.server 8080
```

Then open `http://localhost:8080` in Chrome/Edge. You'll see an "Install
app" button once the browser fires its install prompt.

## Deploying it for real use
Any static host works, since there's no backend: GitHub Pages, Netlify/Vercel,
or any web server over HTTPS (required for service workers outside `localhost`).

## How table extraction works
Per page, the app tries two methods, in order:

1. **Ruling-line grid detection (primary)**. PDFs with visible table
   borders draw them as vector lines/rectangles in the page's content
   stream. The app walks the page's drawing operations (tracking the
   transform matrix through save/restore/transform), collects the
   straight horizontal and vertical strokes, clusters them into grid
   lines, and assigns each text fragment to the grid cell whose bounds
   contain it. This is the same source of truth pdfplumber uses, and it
   fixes the problems text-position clustering alone can't: running page
   headers/footers and surrounding paragraph text are automatically
   excluded, since they fall outside the ruled area entirely rather than
   getting picked up as extra "columns."
2. **Text-position clustering (fallback)**. If a page has no detectable
   grid lines (a borderless/whitespace-separated table), the app falls
   back to reading each fragment's `(x, y)` position, grouping into rows
   by `y`, and clustering `x` positions into column bins.

On top of per-page extraction, the app reconstructs **logical tables that
span multiple pages**:

- **Header detection**: for tables that follow the gazette/legal/government
  document convention, the header is only printed once, on the page the
  table starts, followed by a "(1) (2) (3)…" column-index row. That index
  row is located anywhere on a page and requires at least 2 matching
  bracket-digit cells (not just 1) — this avoids a single stray bracketed
  number elsewhere in a cell (e.g. a dosage like "(50)") being mistaken
  for the index row and incorrectly splitting the table mid-way through.
- **Row continuation / blank rows**: any row without its own value in the
  first (serial number) column never becomes a standalone row — it's
  always folded into the row before it (its non-empty cells appended with
  a line break), whether that's a genuine page-break continuation or just
  a stray/blank grid row from an extra ruling line.
- **Notes/footnote section = end of table**: gazette-style documents often
  end with a "Notes:" heading followed by a lettered list "(a)", "(b)",
  "(c)"… Either the heading itself, or (as a safety net, since the heading
  row often has only one populated cell and can get filtered out earlier)
  the first lettered marker, ends the table there. The letter check
  deliberately excludes every letter that's also a roman numeral (i, v,
  x, l, c, d, m), since those legitimately appear as list markers inside
  real table cells (e.g. a manufacturer column listing several companies)
  — only unambiguous letters like "(a)"/"(b)" trigger it, which is enough
  since a notes list always starts at "(a)".
- **Split running-header fragments**: "[PART II—SEC. 3(ii)]" is matched
  piece-by-piece (the "PART <roman>" part, the "SEC. <n>" part, stray
  brackets) rather than as one whole-cell bracket pair, because column
  detection can split it across two cells. A cell is only collapsed to
  blank if something was actually matched and stripped from it and the
  remainder is negligible — this avoids wiping out short legitimate cells
  ("(1)", "Sl.", "1 Tablet") that the pattern never touched in the first
  place, and avoids a bare "(ii)" pattern being treated as furniture
  everywhere (that would delete real roman-numeral list markers too).
- **Stray page-footer numbers**: if the table's first column behaves like
  an incrementing serial number, a "new row" whose number isn't a
  plausible next value (e.g. "8" appearing right after serial number 18)
  is treated as contamination — likely a page number that bled into that
  column — and folded into the previous row as a continuation instead of
  becoming a bogus standalone entry.
- **Printing/publishing colophon**: a line starting "Uploaded by" or
  "Published by" (the closing credit line in many government documents)
  is treated as furniture and excluded, independent of the notes-heading
  detection.
- **Furniture-fragment scrubbing**: running headers/section markers (e.g.
  "THE GAZETTE OF INDIA", "[PART II—SEC. 3(ii)]") are stripped out of
  cell text even when only partially glued onto an otherwise-real cell,
  not just when they fill an entire cell/row on their own.
- **Whole-page language filtering**: with "English only" on, a page whose
  text is predominantly non-Latin is skipped entirely — this is what
  keeps a bilingual document's other-language copy of the same table out,
  rather than leaving behind filtered fragments of it.

This works well for ruled, gridline-style tables (invoices, gazettes,
government notices, reports). It uses two passes:
1. **Header-aware pass**: looks for the index-row convention above. Best
   results on documents that use it (page-break continuations are merged
   correctly).
2. **Fallback pass** (used only if pass 1 finds nothing at all): treats
   each page's first row as a header candidate and merges a following
   page into the same table when its column count matches. Less precise
   about page-break continuations, but ensures ordinary tables without
   the index-row convention still get extracted instead of producing an
   empty result.

Irregular layouts or merged cells may need light cleanup afterward.
Scanned/image-only PDFs have no text layer, so pdf.js won't extract
anything — OCR would be needed first.

## Tuning
Constants in `index.html` control extraction sensitivity:
- `rowTolerance` (in `itemsToRows`) — vertical pixel tolerance for grouping
  text into the same row
- `gapThreshold` (in `rowsToTable`) — minimum horizontal gap treated as a
  column boundary
- the `0.5` ASCII-ratio threshold (in `buildLogicalTables` / `asciiRatio`)
  — how non-Latin a page's text must be before it's skipped under
  "English only"

## Options in the UI
- **One sheet per table** — each *logical* table (which may span several
  PDF pages) becomes its own sheet, vs. all tables concatenated into one
  sheet with a blank row between them
- **Trim empty rows/columns** — removes fully blank rows/columns
- **Tables only (skip paragraphs/headings)** — drops rows that don't look
  like table content. A row starting with content needs 2+ populated
  columns to count (filters out left-aligned paragraphs/headings); a row
  with a *blank* first column is still kept even with just one populated
  cell, since that's the shape of a page-break-wrapped continuation
  fragment the merge step needs to see. Also drops rows that are purely
  page furniture — a lone page number, a bracketed section marker like
  "[PART II—SEC. 3(ii)]", or a running "GAZETTE"/"EXTRAORDINARY"-style
  header — as an extra safeguard on top of ruling-line detection.
- **English only** — two effects: whole pages that are predominantly
  non-Latin are skipped (see above), and within surviving pages,
  individual text fragments that are mostly non-Latin are stripped of
  their non-Latin characters (or blanked out if that leaves less than
  half the original visible text).
- **Color heading row** — the header row of each table gets a blue fill,
  bold white text, centered alignment, and a frozen pane.
- **Line break after comma & colon** — inserts a line break after each
  comma or colon in a data cell (wrap text is always on), so values like
  "Contains: X, Y, Z" or comma-separated ingredient lists display on
  multiple lines. Purely numeric values (e.g. `1,234.56`) are left
  untouched. Note: a colon inside something like a time or ratio (e.g.
  "10:30") will also get broken — there's no reliable way to tell that
  apart from a "Label:" style cell using text alone.

If no logical table survives the current filters, the app tells you
instead of downloading an empty file — try unchecking "Tables only" or
"English only".

## Output formatting
Cell writing uses **ExcelJS**, which supports real cell styling (fills,
borders, wrap text) in the browser — the free build of SheetJS cannot
write formatting at all. Every cell gets a light border and wrapped,
top-aligned text; columns are auto-sized to content (capped at 45
characters wide).
