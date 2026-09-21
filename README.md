# Parcel Smart Entry

Mobile-responsive Excel data-entry app for the parcel mapping proforma.
Search a record by P/S, CNIC, H/S Code or Name, tap it, edit it, and the change
is written back to the same row of your workbook.

### ▶ Live site: https://mrcrysta1.github.io/parcel-smart-entry-app/

The live site runs **entirely in your browser** — your workbook is never
uploaded anywhere. Open it on a phone, load your `.xlsx`, and start editing.

---

## Two ways to run it

| | Live site (`docs/`) | Local app (Flask) |
| --- | --- | --- |
| Excel engine | ExcelJS, in the browser | openpyxl, on the server |
| Hosting | GitHub Pages (static) | `python run.py`, or Vercel |
| Your data | never leaves the device | posted to your own server |
| `.xlsm` macros | **not** preserved (saves as `.xlsx`) | preserved |

Both share the same HTML, CSS and UI JavaScript, and both read their field
definitions from `shared/schema.json`, so they behave identically.

### Run locally
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python run.py
```
Open http://127.0.0.1:5000

### Rebuild the live site after changing the UI
```bash
python build_static.py    # regenerates docs/ ; commit and push to publish
```

## Behavior
- Upload an existing `.xlsx`/`.xlsm` parcel mapping file.
- The header row and its columns are detected automatically, including common
  spellings such as `Parcel number`, `P/S`, `H/S Code`, `CNIC of Family head`.
- Searches P/S, CNIC, H/S/House Code **and Name** as you type. Exact matches
  rank above prefix matches, which rank above partial matches.
- Click (or tap) a result to load the whole row into the form. Saving then
  writes back to **that same row** — the header shows `Editing row N`.
- "New record" clears the form; Save then appends a new row at the bottom.
- P/S + Name are mandatory. Duplicate P/S numbers are refused, with an offer to
  open the existing record instead.
- Blank optional fields become `Nill`; blank Property Type becomes `Other`.
- With no file loaded, the app creates a fresh workbook in the proforma layout.
- Saved/added rows are highlighted yellow in the exported workbook. Existing
  formatting — bold headers, column widths, freeze panes — is preserved.

## Project layout
```
shared/schema.json   fields, labels, header aliases   <- single source of truth
app/                 Flask app (routes.py = openpyxl engine)
app/templates/       one Jinja template, renders both builds
app/static/          style.css + app.js, shared by both builds
docs/                GENERATED static site for GitHub Pages (build_static.py)
docs/engine.js       the browser Excel engine (ExcelJS), replaces the Flask API
api/index.py         Vercel entry point
```
Edit `app/`, `shared/` or `docs/engine.js`; everything else in `docs/` is
generated — re-run `build_static.py` rather than editing it by hand.

## API (local/Vercel mode only)
| Endpoint | Purpose |
| --- | --- |
| `POST /api/load` | Parse an uploaded workbook; returns records, header row, detected columns. |
| `POST /api/search` | Server-side search. A fallback — the UI searches its local cache. |
| `POST /api/save` | Update row `row`, or append when `row` is `null`. |
| `POST /api/download` | Stream the current workbook back as a file. |

The workbook is parsed once on upload and cached in the browser, so searching
is local and instant; the workbook only goes back to the server on save.

## Limitations
- Edits live in the browser tab until you press **Download Excel** — there is no
  server-side persistence, so closing the tab loses unsaved work.
- Single-user by design. Two people editing the same file will overwrite each
  other; that needs a database, not a spreadsheet.
- On the live site, macro-enabled `.xlsm` files are saved as `.xlsx` (macros
  dropped). Use the local Flask app if you need to keep macros.
