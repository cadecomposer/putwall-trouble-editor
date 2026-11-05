# Putwall Trouble Spreadsheet Editor

A small web tool that turns an **OC Clearing** export into a Putwall-ready
spreadsheet in a few clicks:

1. **Filter** the report down to the `TRBL-PW-xx` locations you are working.
2. **Sort** by the first `Active Locn` column, Z-A.
3. **Add a `Label` column** to the right of `Tote Gid`, numbering the distinct
   tote numbers top-to-bottom (blanks ignored, repeated totes share a number).
4. **Sort every kept row into a section** (see below):
   - **Troubles** - has an active location. These are the printable ones.
   - **Alt form** - no active location, but a reserve location exists. Send an
     alt form so an EO can grab it.
   - **Split cancel** - no active and no reserve location, and no other item on
     the same order is findable. Safe to split cancel.
   - **Do NOT split cancel yet** - no active/reserve, but another item on the
     **same order** (same `oLPN`) is findable, so hold it until that item is found.
   - **Duplicates** - exact repeat rows (later copies of an earlier row).
5. **Generate scannable Code 128 barcodes** for the printable troubles' totes.
6. **Highlight the `Need` cell yellow** when it is more than 1.
7. **Preview the barcodes** right in the page.
8. **Download** the edited `.xlsx` (one sheet per section), or **print** the
   troubles table plus a page of barcode labels.

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
# or: node test/test-processor.js "OC-CLEARING-02.xlsx" TRBL-PW-01,TRBL-PW-05
```

---

## How to use it

1. Tick the `TRBL-PW-01` … `TRBL-PW-20` locations you are working today.
   (`Select all 20` and `Clear` are there if you need them.)
2. Drag the OC Clearing `.xlsx` onto the box, or click to choose the file.
3. Read the sections:
   - **Troubles** (with barcodes) is the printable list.
   - **Alt form**, **Split cancel**, **Do NOT split cancel yet** and
     **Duplicates** are the special-handling lists. None of them print.
4. By default the tables show only the useful columns, from `oLPN` through
   `Reserve Locn 2`. Click **Show all columns** to reveal the rest
   (`Time in OC`, `Order Date`, `Wave`, `iLPN`, `iLPN2`, `Staging Locn`), and
   click again to hide them.
5. Click **Download .xlsx**, or **Print sheet + barcodes**.

Changing the location ticks after a file is loaded re-runs the edit instantly,
so you can see the effect without re-uploading.

### Why "Do NOT split cancel yet"?

Anything with the same `oLPN` is the same order. If an order is missing an item
you can find (it has an active or reserve location) **and** an item you cannot
find, the unfindable item must not be split-cancelled until the findable one is
resolved. Those rows are listed under **Do NOT split cancel yet**. The check is
done against the whole file, so it also catches a findable item that sits at a
location you did not select.

### Printing

**Print sheet + barcodes** opens your browser's print dialog (choose a printer
or *Save as PDF*) with two parts, both landscape:

- **Page 1 — the troubles table**, 10pt, limited to the columns `Tote Gid`
  through the first `Reserve Locn` (i.e. `Tote Gid`, `Label`, `Item`, `Need`,
  `Active Locn`, `Active Locn 2`, `Active OH`, `Reserve Locn`). `Need` cells
  above 1 are yellow.
- **Following pages — the barcode labels**, one block per printable tote with
  the label number above a Code 128 barcode and the tote number below.

Only the printable **Troubles** print. The alt-form, split-cancel, hold and
duplicate sections are never printed.

### Appearance

The page follows your browser/OS colour-scheme setting automatically by
default. A button in the top-right corner switches between **Auto**, **Light**
and **Dark**, and the choice is remembered on that computer. The printed
document is always light so it prints cleanly either way.

---

## What the output contains

The downloaded `.xlsx` has one sheet per section:

- **Troubles** - printable rows (active location present).
- **Alt Form** - no active location, reserve location present.
- **Split Cancel** - no active or reserve location, order has nothing findable.
- **Split Cancel HOLD** - no active/reserve, but the order has a findable item.
- **Duplicates** - later exact copies of an earlier row (filled light red).
- **Barcodes** - one row per printable tote: the label number, the tote number,
  and a Code 128 barcode image.

Every data sheet also has:

- the new `Label` column immediately to the right of `Tote Gid`;
- `Need` cells greater than 1 filled yellow (fill `FFEB3B`);
- a frozen header row.

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
| Active location | The first `Active Locn` column being non-empty. |
| Reserve location | The first `Reserve Locn` column being non-empty. (`Reserve Locn 2` is only ever filled when the first is, so it is ignored.) |
| Sections | normal = active; alt = no active but reserve; split = neither; hold = split whose order has a findable item; duplicate = later exact copy. |
| Order grouping | Rows are grouped by `oLPN` across the whole file for the hold rule. |
| Duplicates | The entire original row must match (every column except the derived `Label`). The first occurrence keeps its section; later copies go to **Duplicates**. |
| Barcodes | Code 128, encoding the **tote number**, generated only for the printable **Troubles**. |
| Column toggle | Web tables default to `oLPN` … `Reserve Locn 2`; the button reveals every column. The `.xlsx` always contains every column. |
| Print columns | Always `Tote Gid` … first `Reserve Locn`, landscape, 10pt. |

If any of these should behave differently, it is a small change - see below.

---

## Customizing

All the spreadsheet logic lives in **`src/processor.js`**; the UI and the
workbook writing live in **`src/app.js`**.

- **Section rules:** `processData()` in `src/processor.js` (the `categories`
  mapping and the `orderHasFindable` map).
- **Need highlight colour:** search for `FFFFEB3B` / `#ffeb3b`.
- **Duplicate colour:** search for `FFFFC7CE` (fill) and `FF9C0006` (font).
- **Barcode format / size:** `src/app.js`, the `makeBarcode()` /
  `makeBarcodeSvg()` functions. Change `format: 'CODE128'` to `'CODE39'`, etc.
- **Default web columns:** `visibleColumnIndices()` in `src/app.js`.
- **Printed columns:** the `PRINT_COLUMNS` array in `src/app.js`.
- **Sheet names:** `buildWorkbook()` in `src/app.js`.

After editing, run `python3 build.py` to regenerate the single-file version.

---

## Project structure

```
index.html                  page shell (references the files below)
src/styles.css              styles (light + dark)
src/processor.js            pure logic: header detect, filter, sort, label,
                            classify sections, duplicates, barcodes
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
