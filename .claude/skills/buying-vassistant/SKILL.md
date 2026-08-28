---
name: buying-vassistant
description: >-
  Virtual assistant for a fashion buying/merchandising team that turns Excel exports
  from the buying system into presentation-ready PowerPoint decks. Two modes: (1) BUYER
  LINEGUIDE — one portrait slide per style for a buy meeting, from a buyer lineguide
  .xls; (2) MERCH BOARD — a 16:9 merchandising assortment board (cover, month dividers,
  fabric-grouped category cards with CAD images) from a buy-prep .xls with a "Line Guide
  View" sheet. Use whenever the user attaches or points to a buying/buy-prep/lineguide
  Excel file (.xls/.xlsx) and wants a deck, slides, board, or "buy meeting" PPTX — even
  if they only say "run the lineguide," "make the merch board," "buyer lineguide,"
  "assortment board," "buy meeting deck," "add these styles to the board," or name a
  brand + buy (e.g. "LF Nov", "NBD 2027"). Also use for incremental updates to an
  existing merch board. Do NOT use for generic .xlsx cleanup, charts, or Word/PDF work
  unrelated to these buy-meeting deck formats.
---

# Buying Virtual Assistant — buy-meeting deck builder

This skill reformats the buying team's Excel exports into two specific, approved PowerPoint
layouts. Both are already solved by bundled Python scripts. **Your job is to identify which
deck the user wants, run the matching script, and hand back the `.pptx` — not to rebuild the
layout from scratch.**

## The one rule that matters most: run the script, don't re-derive

Both of these decks encode dozens of hard-won layout decisions (off-slide image slots, exact
point line-spacing for Google Slides, pre-broken text lines, fabric-group row packing, image
sizing by height, category waterfalls). Every time someone re-implemented the layout from the
prose spec instead of running the script, it reintroduced known bugs — all-caps category
titles, 9pt brand text, `wrap="none"` headers that overflow, mis-cropped CADs. So the scripts
in `scripts/` are the source of truth. Run them. Only read the reference docs or edit a script
when a run fails or the user asks for a change the current config can't produce.

If a reference doc and a script ever disagree, **the script wins** (the docs are commentary on
what the script does).

## Step 0 — one-time environment setup

The scripts need LibreOffice (`soffice`, for `.xls → .xlsx` conversion), a few Python libs, and
the DejaVu/Liberation fonts (already present in this sandbox). Install the Python deps once:

```bash
pip install python-pptx openpyxl pillow numpy xlrd --break-system-packages
```

`xlrd` is only needed for the lineguide mode (it reads the original `.xls` cache); the rest are
shared. If `soffice` isn't on PATH, tell the user their environment needs LibreOffice and stop.

## Step 1 — pick the mode

Ask only if it's genuinely unclear. The signals:

| Signal | Mode |
|---|---|
| "buyer lineguide", "buy meeting deck", "one slide per style", portrait spec sheets, yellow brand band | **Buyer Lineguide** (Mode A) |
| "merch board", "assortment board", "line guide view", CAD flats, month dividers, fabric groups, cards | **Merch Board** (Mode B) |
| "add these to the board", "update the board", "keep everything where it is" + an existing `.pptx` | **Merch Board — incremental update** (Mode B, updater) |

A quick tell from the file itself: a **Merch Board** source has a sheet literally named
`Line Guide View` with embedded CAD images. A **Buyer Lineguide** source is one ~56-row block
per style with a yellow brand cell in column G. If you're unsure which file you have, open it
and check the sheet names before committing.

---

## Mode A — Buyer Lineguide deck

Full detail (layout, parsing rules, brand filename rules, edge cases) is in
`references/lineguide.md`. Read it if a run fails or the output looks off; otherwise the script
handles it.

**Inputs.** The script takes *two* paths: the original `.xls` (values are read from its cached
formula results via `xlrd` — this matters, because converting to `.xlsx` silently re-evaluates
some formulas and corrupts a few cells like Duty Rate) and an `.xlsx` companion (only used to
extract the embedded product images, which `xlrd` can't reach). You produce the companion
yourself with LibreOffice.

**Run it:**

```bash
# 1. make the .xlsx companion for image extraction
soffice --headless --convert-to xlsx --outdir /tmp "path/to/lineguide.xls"

# 2. build the deck (values from the .xls, images from the .xlsx)
python3 scripts/lineguide_deck_generator.py \
    "path/to/lineguide.xls" "/tmp/lineguide.xlsx" \
    --out-dir /mnt/user-data/outputs
```

- If the user only has an `.xlsx` (no `.xls`), pass it as **both** arguments and mention that
  formula-cached values may differ slightly from the original `.xls` source of truth.
- Meeting date: the filename uses today's date as `{M.DD}` (month no leading zero, day two
  digits). Override with `--meeting-date 6.02` if the user names a specific meeting date.

**Sample photos (from Google Drive).** Optional in general, but **when the user asks for the
lineguide "with sample photos" (or points you at a photo folder/Drive link), fetching them is part
of the deliverable — not an optional extra.** The user needs the Google Drive connector connected.

1. For each style, take the stylecode prefix (before the first `-`, e.g. `MJOW10024`) and search
   Drive for images whose title contains it. The search reaches every folder the user can access,
   so you do **not** need a folder link — a stylecode search finds photos wherever they live.
   **Only download the photos actually needed** — never bulk-download a folder, and never try to
   pull every photo in Drive. Not every style has photos; those render with the CAD only.
2. Decode/download them into a **sandbox temp dir** (`default_photos_tmp()`, e.g.
   `/tmp/lineguide_photos`), saving each file under its original Drive title (it already starts
   with the stylecode and carries the view word — `front/back/side/wr left/wr right`). NEVER write
   sample photos into the user's output/working folder or anywhere on their computer.
   - The Drive download tool returns each file as base64 (large files land in a tool-result file on
     disk, small ones inline); decode to bytes and write. If there are many photos, delegate the
     downloads to a subagent so the base64 stays out of the main context — but download **once**,
     in a single pass. Do not re-run the whole download after a partial success; fetch only the
     files still missing.
3. Pass that dir and run with cleanup so nothing persists:

```bash
python3 scripts/lineguide_deck_generator.py \
    "path/to/lineguide.xls" "/tmp/lineguide.xlsx" \
    --photos-dir /tmp/lineguide_photos --cleanup-photos \
    --out-dir /mnt/user-data/outputs
```

Placement (all off-slide, hidden in present/print): CAD top-right (right edge at `x=-0.20"`),
FRONT top-left, SIDE under front, BACK bottom-right; any OTHER views in a row **below** the slide
aligned to the slide's left edge. View mapping: `front`->front, `back`->back, `side` or `WR left`
->side, everything else->other. See `references/lineguide.md` -> "Sample photos".

**Verify photos populated, and report honestly (REQUIRED when photos were requested).** The photos
sit off-slide, so you cannot tell from a glance whether they made it in — you MUST check the built
deck programmatically before claiming success. A deck with zero photos still opens and looks fine,
so "it generated" is NOT evidence the photos are there. After building, count the pictures per
slide and compare to how many photos you downloaded:

```python
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
prs = Presentation("path/to/output.pptx")
for i, s in enumerate(prs.slides, 1):
    n = sum(1 for sh in s.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
    print(f"slide {i}: {n} pictures")  # 1 == CAD only (no sample photos); >1 == photos present
```

Then tell the user plainly: **"X of N styles have sample photos"**, and name the styles that came
back CAD-only. If the Google Drive connector is not connected, or a stylecode search returns
nothing, say so explicitly — **never present a CAD-only or photo-less deck as if the photos were
added.** If the user asked for photos and none populated, that is a failure to surface, not a quiet
success.

**Output filename:** `{M.DD} BUY MEETING - {BRAND_ABBR}.pptx`. `{BRAND_ABBR}` is the brand
abbreviation; a multi-brand lineguide joins each distinct brand's abbreviation with ` + ` in
first-seen order (e.g. `LF + MAJ + TULA`). Map: MAJORELLE->`MAJ`, Tularosa->`TULA`,
Lovers & Friends->`LF`; unmapped brand -> first 3 letters uppercased. The script does this — you
don't have to. (Extend `BRAND_ABBR` in the script for new brands.)

**Quickstrike background.** Any style whose `Quickstrike` field is the REVOLVE/NUULY program gets a
light-lavender (`#D9D2E9`) page background; every other style stays white. It's per-style and
automatic. Matching is order-insensitive, so `REVOLVE/NUULY` and `NUULY/REVOLVE` both trigger
(case/space-insensitive). See `references/lineguide.md` → "Quickstrike background" to change the
color, the trigger value, or the field label it keys on.

---

## Mode B — Merch Board

Full detail (category waterfall, image handling, exact layout constants, Google-Slides
correctness, brand config, first-run checks) is in `references/merch_board.md`. Read it before a
first run for a **new brand**, or whenever a run fails / looks wrong.

### Before running, confirm two things with the user

1. **Cover period.** The buy name (e.g. "LF Nov") is *not* the release month — the release month
   is read from the `R.Month` column. The cover shows the *period* the user wants, which may be a
   quarter (`2026 Q2`) or a buy label (`2026 Nov`). Confirm it; it becomes the cover subtitle and
   part of the filename.
2. **Brand.** Defaults to "Lovers and Friends". For any other brand set `MB_BRAND` (see below).
   The engine is brand-agnostic — same categories, layout, fonts, palette; only the label changes.

### Fresh build

```bash
# default brand (Lovers and Friends)
python3 scripts/merch_board.py "path/to/Buy_Meeting_Prep_List (LF Nov).xls" "2026 Nov"

# other brand
MB_BRAND="NBD" python3 scripts/merch_board.py "path/to/Buy_Meeting_Prep_List (NBD 2027).xls" "2027"
```

By default it writes `<Brand> MERCH BOARD <period>.pptx` in the current directory; pass a third
arg to control the output path, e.g. `"/mnt/user-data/outputs/Lovers and Friends MERCH BOARD 2026 Nov.pptx"`.
Do **not** prefix a date, use the buy code, or add a version suffix — the filename is
`<Brand> MERCH BOARD <Year> <Period>.pptx`.

### Incremental update (keep existing cards where they are)

When a revised buy file arrives and the ask is "add the new styles but leave everything already on
the board in place," use the updater instead of a rebuild. It matches each style against cards
already on the board (by wrapped name + colour + price), leaves matches untouched, and appends only
new styles as continuation slides within their month/category section (dropping `Status = DROPPED`
styles to a DROP page).

```bash
python3 scripts/merch_board_update.py \
    "path/to/new_source.xls" "path/to/existing MERCH BOARD.pptx" "2027" \
    "/mnt/user-data/outputs/Lovers and Friends MERCH BOARD 2027.pptx"
```

### QA before delivering (Merch Board only)

The layout is dense, so render a few slides to images and eyeball them: at least one page per
category type, plus a **recolor** card (should keep its colour-swatch image, not be sliced) and
the **DROP** page (grey background). Confirm no text overlaps or collides, and that swim /
accessory / excluded styles are absent. A quick way:

```bash
soffice --headless --convert-to pdf --outdir /tmp "path/to/output.pptx"
pdftoppm -png -r 80 /tmp/output.pdf /tmp/qa   # then view a few /tmp/qa-*.png
```

If a card is mis-cropped or mis-filed, prefer a surgical fix (re-crop one CAD, regenerate one
all-new section) over re-flowing a section that already contains approved positions.

---

## Step 2 — deliver

Copy the finished `.pptx` to `/mnt/user-data/outputs/` if it isn't already there, then present it
with `present_files` and a one-line summary (which mode, brand, period/date, slide count). Don't
paste a long description of the deck — the user opens the file. Deletion isn't available here, so
if you saved a wrong filename, save the correct one and ask the user to delete the stray file.

## Adding a new brand or segment (Merch Board)

Most brands just need `MB_BRAND` and run as-is. The first time a *new* brand's buy file appears,
run the four-point first-run config check in `references/merch_board.md` (stylecode→colour shape,
knit vendor codes, swim/exclusion convention, category coverage) — a mismatch silently mis-files a
card rather than erroring. Keep the engine identical; change only the segment config in the script.
