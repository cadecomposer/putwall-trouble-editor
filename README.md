# Putwall Trouble Spreadsheet Editor

A small web tool that turns an **OC Clearing** export into a Putwall-ready
spreadsheet in a few clicks:

1. **Filter** the report down to the `TRBL-PW-xx` locations you are working.
2. **Sort** by the first `Active Locn` column, Z-A.
3. **Add a `Label` column** to the right of `Tote Gid`, numbering the distinct
   tote numbers top-to-bottom (blanks ignored, repeated totes share a number).
4. **Generate scannable Code 128 barcodes** for every distinct tote, each shown
   next to its label number, on a dedicated **Barcodes** sheet.
5. **Mark duplicate rows red** (only the later copies, never the first).
6. **Highlight the `Need` cell yellow** when it is more than 1.
7. **Preview the barcodes** right in the page.
8. **Download** the edited `.xlsx`, or **print** a document with the troubles
   table (landscape, 10pt) followed by a page of barcode labels.

Everything runs in the browser. No file, order, item or tote data is ever
uploaded anywhere.

---

## Quick start

### Option A - just open the single file (no install, no internet)

Open **`dist/putwall-trouble-editor.html`** in Chrome, Edge or Firefox.
It is fully self-contained (all libraries inlined), so it works offline and
does not need the internet or a server. You can copy it to a USB stick, email
it to yourself, or keep it on the work PC.

### Option B - run the source folder (for editing / development)

```bash
cd "putwall trouble spreadsheet editor"
python3 -m http.server 8137
# then open http://localhost:8137/index.html
```

### Option C - rebuild the single file after editing the sources

```bash
python3 build.py        # writes dist/putwall-trouble-editor.html
```

### Run the logic test (optional)

```bash
node test/test-processor.js
# or: node test/test-processor.js "OC-CLEARING-02 - ....xlsx" TRBL-PW-01,TRBL-PW-05
```

---

## How to use it

1. Tick the `TRBL-PW-01` … `TRBL-PW-20` locations you are working today.
   (`Select all 20` and `Clear` are there if you need them.)
2. Drag the OC Clearing `.xlsx` onto the box, or click to choose the file.
3. Check the preview. Duplicates are shaded red, and the generated barcodes
   appear in the **Barcodes** section below the table.
4. Click **Download .xlsx**, or **Print sheet + barcodes**.

Changing the location ticks after a file is loaded re-runs the edit instantly,
so you can see the effect without re-uploading.

### Printing

**Print sheet + barcodes** opens your browser's print dialog (choose a printer
or *Save as PDF*) with two parts:

- **Page 1 — the troubles table**, landscape, 10pt, limited to the columns
  `Tote Gid` through the first `Reserve Locn` (i.e. `Tote Gid`, `Label`, `Item`,
  `Need`, `Active Locn`, `Active Locn 2`, `Active OH`, `Reserve Locn`).
  Duplicate rows keep their red shading, and `Need` cells above 1 are yellow.
- **Following pages — the barcode labels**, also landscape, one block per
  distinct tote with the label number above a Code 128 barcode and the tote
  number below.

The barcode and colours are always the same as what you download, so the printed
page and the spreadsheet agree.

### Appearance

The page follows your browser/OS colour-scheme setting automatically by
default. A button in the top-right corner switches between **Auto**, **Light**
and **Dark**, and the choice is remembered on that computer. The printed
document is always light so it prints cleanly either way.

---

## What the output contains

**Sheet "Troubles"** - the filtered and sorted data, with:

- a new `Label` column immediately to the right of `Tote Gid`;
- duplicate rows filled light red (fill `FFC7CE`, font `9C0006`);
- `Need` cells greater than 1 filled yellow (fill `FFEB3B`);
- a frozen header row.

**Sheet "Barcodes"** - one row per distinct tote, containing the label number,
the tote number, and a Code 128 barcode image encoding the tote number.

Example from the included sample (locations 01, 04, 05):

```
Label   Tote Gid   Barcode
1       T710356    [||||| Code128 |||||]   -> decodes to "T710356"
2       T184024    [||||| Code128 |||||]   -> decodes to "T184024"
```

---

## Decisions & assumptions

These are the choices the tool makes. They are all easy to change - see
"Customizing" below.

| Topic | Behaviour |
| --- | --- |
| Header detection | Finds the first row containing both `Tote Gid` and `Active Locn`, so the title rows at the top of the report are safe. |
| Filter column | The `Location` column (that is where `TRBL-PW-xx` appears). Matching ignores case and surrounding spaces. |
| Sort | First `Active Locn` column, descending (Z-A), case-insensitive. Blank values sort to the bottom. Ties keep their original order (stable). |
| Labels | Distinct tote values numbered `1, 2, 3, …` in the post-sort top-to-bottom order. Blank tote cells get a blank label. The same tote value always gets the same number. |
| Duplicates | The entire original row must match. Comparison is case-sensitive and includes every original column. The first occurrence is left alone; all later copies are marked. |
| Barcode | Code 128, encoding the **tote number**. The label number is printed in the cell beside it. |
| Output layout | Two fresh, clean sheets (`Troubles`, `Barcodes`). The original report's title rows and formatting are not carried over. |
| Duplicate colour | Light red - visible in Excel and Google Sheets, and still readable when printed. |

If any of these should behave differently (for example, sort A-Z, put barcodes
on the same sheet, or encode the label number in the barcode instead of the
tote), it is a small change - see below.

---

## Customizing

All the spreadsheet logic lives in **`src/processor.js`**; the UI and the
workbook writing live in **`src/app.js`**.

- **Duplicate colour:** `src/app.js`, search for `FFFFC7CE` (fill) and
  `FF9C0006` (font).
- **Barcode format / size:** `src/app.js`, the `makeBarcode()` function.
  Change `format: 'CODE128'` to `'CODE39'`, or adjust `width`/`height`.
- **What the barcode encodes:** in `src/app.js`, `makeBarcode(b.tote)` - pass
  `String(b.label)` instead if you want the label number encoded.
- **Barcodes on the main sheet:** add the images to `ws` instead of the separate
  `Barcodes` worksheet in `buildWorkbook()`.
- **Sort direction / filter column:** `processData()` in `src/processor.js`.

After editing, run `python3 build.py` to regenerate the single-file version.

---

## Project structure

```
index.html                  page shell (references the files below)
src/styles.css              styles
src/processor.js            pure logic: header detect, filter, sort, label, duplicates
src/app.js                  UI, file reading (SheetJS), workbook writing (ExcelJS)
vendor/xlsx.full.min.js     SheetJS - reads the uploaded workbook
vendor/exceljs.min.js       ExcelJS - writes the output workbook + barcode images
vendor/JsBarcode.all.min.js JsBarcode - draws the Code 128 barcodes
build.py                    inlines everything into dist/putwall-trouble-editor.html
dist/putwall-trouble-editor.html   the self-contained tool (build output)
test/test-processor.js      runs the core logic against a real report in Node
HOSTING.md                  options for putting this on the web
```

---

## Privacy

The page reads the file locally with the File System APIs and writes the result
locally. There are no network calls, no analytics, and no backend. That is
deliberate, because the report contains order/OLPN/item data. Any hosting
option you choose in `HOSTING.md` keeps this property as long as you serve the
static files only.

---

## Third-party libraries

Bundled in `vendor/` (all permissively licensed, redistributed unmodified):

- [SheetJS `xlsx`](https://sheetjs.com/) - reading the uploaded workbook (Apache-2.0).
- [ExcelJS](https://github.com/exceljs/exceljs) - writing the output workbook and embedding barcode images (MIT).
- [JsBarcode](https://github.com/lindell/JsBarcode) - drawing the Code 128 barcodes (MIT).
