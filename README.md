# Parcel Smart Entry

Excel data entry for the parcel mapping proforma. Search a record by P/S, CNIC,
H/S Code or Name, tap it, edit it, and the change is written back to the same
row of the workbook.

It comes in two builds from one codebase:

| | **Shared** (`web/` + Netlify) | **Offline** (`docs/` + GitHub Pages) |
| --- | --- | --- |
| Who can edit | your whole team, at the same time | one person, one device |
| Where the file lives | on the server, shared | in that browser tab only |
| Live updates | yes, within ~2.5s | n/a |
| Needs internet | yes | only to load the page once |
| Excel engine | ExcelJS, in a Netlify Function | ExcelJS, in the browser |
| `.xlsm` macros | not preserved (saves `.xlsx`) | not preserved (saves `.xlsx`) |

Offline link (single user, works in the field with no signal):
**https://mrcrysta1.github.io/parcel-smart-entry-app/**

Open it and press **Open the team file** - a full 865-row proforma loads with no
upload, so you can try searching, editing and downloading straight away.

> That bundled copy is **demo data**. It keeps the real file's shape - every
> parcel number, house code, property type, row count and the header on row 3 -
> but every name and CNIC is invented and every coordinate is blunted to about
> a kilometre. The site is public, so no real person's details are on it.
> Regenerate it from any workbook with:
>
>     python tools/make_demo_workbook.py "your-file.xlsx" docs/sample-parcel-mapping.xlsx

The shared build is ready to deploy but **not yet live** — it needs a Netlify
account, which only you can create. See *Deploying the shared version* below.

---

## The shared version

Ali uploads the Excel file; it is stored on the server. Zain opens the same
file from the list. Both search, edit and add records at the same time, and see
each other's changes within a couple of seconds.

- **Who you are** — you type your name once; it is shown against every record
  you change ("last saved by Zain, 2m ago").
- **Live** — new and changed records arrive automatically. No refresh, no
  re-upload. A tab in the background pauses polling and catches up when you
  return to it.
- **If two people edit the same parcel** — the second person is warned *while
  still typing* ("Ali just saved changes to row 5"), and if they save anyway
  they get both versions side by side with the differing fields highlighted,
  and choose **Use their version** or **Overwrite with mine**. Nothing is ever
  silently lost.
- **Full page view** (`sheet.html`) — every record as one live table. Click a
  cell to edit it in place; it saves when you leave the cell. Rows another
  person just changed flash yellow.
- **Download Excel** rebuilds the workbook from the stored original, so the
  proforma's headers, bold, column widths and freeze panes survive.

### How concurrent edits are kept safe

Two layers, because one is not reliable enough on its own:

1. **A short lease per sheet.** Every write takes a lease built on Netlify
   Blobs' `onlyIfNew`, which is atomic in every runtime. A lease older than 12s
   is treated as abandoned and taken over, so a crashed request cannot wedge a
   file.
2. **A revision on every record.** A save quotes the revision it was based on.
   If someone else changed that record meanwhile, the save is rejected with
   both versions rather than applied.

Netlify Blobs also offers ETag compare-and-swap (`onlyIfMatch`) and it is used
when available — but the ETag has to come back from a *read*, and some runtimes
(the local dev sandbox among them) return an empty one. `onlyIfMatch: undefined`
degrades silently to last-write-wins, which for a survey means an enumerator's
record quietly disappearing. Hence the lease.

Verified: five simultaneous saves all land on distinct rows with nothing lost,
and five simultaneous edits to the *same* row produce exactly one winner and
four conflict warnings.

### Deploying the shared version

Everything is committed and ready; the storage needs no second account and no
connection string.

1. Sign in at **netlify.com** with GitHub.
2. **Add new site → Import an existing project →** pick
   `mrcrysta1/parcel-smart-entry-app`.
3. Netlify reads `netlify.toml` — publish directory `web`, functions
   `netlify/functions`. Leave the defaults and press **Deploy**.
4. Open the URL it gives you.

### Turning on the team password

Off by default: with no environment variable set, anyone with the URL can read
and change every record. To lock it down, in Netlify go to **Site configuration
-> Environment variables -> Add**:

    Key:    TEAM_PASSWORD
    Value:  whatever you want the team to type

Redeploy (or just **Trigger deploy -> Clear cache and deploy site**). From then
on the sign-in card asks for the password alongside the name. To turn it off
again, delete the variable and redeploy.

The password is exchanged once for a signed token that lasts 30 days; the
password itself is never stored in the browser and never sent again. Change the
variable and every existing token stops working, which is how you remove
someone's access.

Free tier notes: functions sleep when idle (the first request after a quiet
spell takes a few seconds) and Netlify Blobs storage persists across deploys.

---

## Running locally

Shared version, exactly as it runs on Netlify (Blobs are emulated):
```bash
npm install
npx netlify dev          # http://localhost:8888
```

Offline version:
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python run.py            # http://127.0.0.1:5000
```

After changing anything in `app/`, regenerate both front-ends:
```bash
python build_static.py
```

---

## Behaviour (both builds)

- The header row and its columns are detected automatically, including
  spellings such as `Parcel number`, `P/S`, `H/S Code`, `CNIC of Family head`.
- Search covers P/S, CNIC, H/S/House Code **and Name**, ranking exact matches
  above prefix above partial.
- P/S + Name are required. Duplicate P/S numbers are refused, naming whoever
  owns the existing record.
- Blank optional fields become `Nill`; blank Property Type becomes `Other`.
- Saved and added rows are highlighted yellow in the exported workbook.

### Fields that repeat across a file
- **Block ID, First Name, Username, Designation** are read from the file you
  load. One distinct value fills in automatically; several become a dropdown of
  exactly those values. Whatever is set carries over to the next record, so
  they are entered once, not once per row. `Other…` allows a new value.
- **House Code** — the leading digits shared by every house code in the file
  are pre-filled, so only the differing digits are typed. (`31846856`,
  `31846857`, `31846858` in the file → a new record starts at `3184685`.)
- **Comment** is an **Add** / **Delete** dropdown. A comment already in the
  sheet that is neither is added to the dropdown when that record is opened, so
  existing rows are never silently rewritten.

---

## Project layout
```
shared/schema.json     fields, labels, header aliases   <- single source of truth
shared/engine-core.cjs the Excel engine (ExcelJS), used by the browser AND the
                       Netlify function; docs/engine.js is generated from it
netlify/functions/     the shared multi-user API, backed by Netlify Blobs
netlify.toml           publish = web/, functions = netlify/functions
app/                   Flask app + the templates, CSS and JS both builds share
build_static.py        renders app/ into docs/ (offline) and web/ (shared)
docs/  web/            GENERATED - do not edit by hand
api/index.py           Vercel entry point for the single-user Flask app
```

## Shared API
| Endpoint | Purpose |
| --- | --- |
| `GET /api/sheets` | List the files on the server. |
| `POST /api/sheets` | Upload a workbook (or create a blank one) as a shared file. |
| `GET /api/sheets/:id` | Full snapshot: records, revision, dropdown options. |
| `GET /api/sheets/:id/changes?since=N` | Only the records changed since revision N. |
| `POST /api/sheets/:id/records` | Add a record. 409 if the P/S already exists. |
| `PUT /api/sheets/:id/records/:row` | Update a record; 409 + both versions on a conflict. |
| `GET /api/sheets/:id/export` | Rebuild and download the .xlsx. |
| `DELETE /api/sheets/:id` | Remove a shared file. |

## Publishing data safely

`tools/make_demo_workbook.py` exists because a survey workbook carries real
names, CNICs and the GPS position of people's homes, and the GitHub Pages site
is public. It replaces names and CNICs, blunts coordinates, and clears the
places a value hides after its cell is overwritten:

- **Filter values.** A column filter stores the exact values it was filtering
  on in the sheet XML. A real CNIC survived there long after its cell was
  replaced.
- **Document properties**, cell notes, headers/footers, defined names, and any
  sheet after the first.

Coordinates appear in two encodings - `34.7712` and `302126607` (meaning
30.2126607). Only the first is helped by rounding; the integer form looks
already-rounded to a naive check and sailed through the first attempt at full
centimetre precision. Both are handled now.

None of this is obvious by eye, which is why the check is a script and not a
glance: it diffs the demo against its source and fails on any surviving name,
CNIC or precise coordinate, scanning every part of the .xlsx rather than just
the visible cells.

## Limitations
- **No per-person accounts.** The team password is one shared secret, so it
  cannot tell you who signed in, only who made each edit (by the name they
  typed). Change `TEAM_PASSWORD` to revoke access for everyone at once.
- `.xlsm` macros are dropped on save in both builds; use the local Flask app to
  keep them.
- The offline build's edits live in the tab until you press Download Excel.
- Sized for a team of about five on files of a few thousand rows. Much larger
  than that and the per-sheet blob should be split per record.
