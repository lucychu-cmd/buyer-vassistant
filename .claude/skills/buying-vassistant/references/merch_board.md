# Buyer — Merch Board (Claude Project Context)

Paste this whole document into the Claude Project (Custom Instructions and/or Knowledge), and add the
bundled **`merch_board.py`** (full build) and **`merch_board_update.py`** (incremental update) scripts.
Playbook for generating the **Merchandising Assortment board** PowerPoint from a buy-meeting prep Excel
export.

The engine is **brand-agnostic** — the same categories, layout, fonts, and palette serve every brand;
only the brand label changes (via `MB_BRAND`). **Lovers + Friends — Femme (LF Femme)** is the reference
segment that the category model was tuned on. NBD, Tularosa, and Majorelle reuse it as-is. See "Adding a
new brand / segment" for the one-time config check to run the first time a brand's buy file appears.

---

## Preferred method — run the script, don't re-implement

The reliable way to reproduce the approved layout is to **run the bundled `merch_board.py`** (add it to
the Project knowledge and execute it), not to re-derive the build from these prose rules:

```
python merch_board.py "<source>.xls" "<cover period>"     # e.g. "2026 Oct"
```

It runs the whole pipeline (load → extract images → derive → crop/sharpen → categorize → build) with
the exact constants below and writes `Lovers and Friends MERCH BOARD <period>.pptx`. Re-writing the
build from scratch has repeatedly introduced bugs (all-caps category, 9 pt brand, `wrap="none"` headers
that overflow/squeeze the cards). The rules below document what the script does — change the script and
this doc together; never hand-roll a different build. Requires python-pptx, openpyxl, pillow, and
LibreOffice (`soffice`).

### Updating an existing board (keep positions) — `merch_board_update.py`

When a revised buy file arrives and the ask is "add the changes but keep everything already on the board
where it is," use the incremental updater instead of a full rebuild:

```
python merch_board_update.py "<new source>.xls" "<existing board>.pptx" "<period>"   # writes back to the board
```

It reuses the identical data/image/categorize/layout logic, then: (1) matches each style in the new file
against the cards already on the board (by wrapped name + colour + price) — **matched styles stay exactly
where they are**; (2) appends only the **new** styles as continuation slides within their month/category
section (calendar order for months, standard category order within a month); (3) routes any
`Status = DROPPED` style to a DROP page. Styles already on the board but absent from the new file are left
in place (not removed) unless dropped. New cards land on their own continuation pages after the existing
pages for that section, so reviewers find prior items untouched. Example: the LF 2027 board went 22 → 29
slides when 62 new styles (53 Jan, 9 Mar) were added with the 35 existing styles unmoved.

For a one-off image/layout correction on a board that's already approved (e.g. re-cropping a single
mis-handled CAD), patch the affected styles surgically rather than rebuilding: a section that is entirely
new styles can be regenerated and swapped in, and a single-card slide can have just its picture replaced
in place. Never re-flow a section that contains pinned/existing positions.

---

## Role & workflow

Generate a 16:9 merch board (PowerPoint) from a buy-prep Excel export, built with **Python +
python-pptx**, delivered as a `.pptx` the user uploads to **Google Slides**.

**At the start of every request:**
1. Confirm the **segment** (e.g. "LF Femme"). If new, ask for the differences before building.
2. Confirm the source has the **`Line Guide View`** sheet with the embedded `CAD Image` column.
3. Read **release month(s)** from `R.Month` (the buy name like "LF Oct" is the buy, not necessarily the
   release — read R.Month). Confirm the **cover period + filename** with the user.

**Procedure:** load → extract & crop CAD images → derive fields → categorize → flow layout → render →
QA (no overlap; excluded styles absent; check a recolor + a DROP page) → deliver.

---

## Data source

- A `Buy_Meeting_Prep_List (LF …).xls` export. Convert .xls → .xlsx (LibreOffice).
- Sheet **`Line Guide View`** — leading **`CAD Image`** column holds product flats as floating images.
- **Extract images:** parse `xl/drawings/drawing1.xml` anchors; each `<xdr:from><xdr:row>` is 0-based →
  Excel row *N+1*; resolve `r:embed` → `…/media/imageX` via `drawing1.xml.rels`.

### Fields used
| Field | Used for |
|---|---|
| `Style Name` | Product name + silhouette detection |
| `Stylecode` | Colour = text after season code (e.g. `…-H26NAVY` → NAVY). Swim flag = 3rd letter `X` |
| `Retail Price` | Price (`$XXX`; ≤ 0.01 → blank) |
| `Material Name` | Fabric group header |
| `R.Month` | Release month divider(s) — whatever months appear (may be one) |
| `Class` | Recolor handling + annotation tag (`*RECOLOR*`, `*NEW MULT*`, … for non-"New" class) |
| `Design Notes` / `Buyer Note` | Tier — parsed from free text `TIER A/B/C` |
| `Vendor Name` | Knits classification + swim exclusion |
| `SKU LeadTime` | Chase (`CHASE`) |
| `Status` | DROP (`DROPPED`) |

---

## Category model

**Display order per month:** `Dresses → Tops → Bottoms → Outerwear → Rompers & Jumpsuits → Knits → Chase → Long Lead → DROP`

**Assignment waterfall (first match wins):**
1. Exclude — Swim: stylecode 3rd letter `X` or vendor Sovereign Sky (`SOVSK`).
2. Exclude — Accessories: handbag, necklace, clutch, tote, earring, bracelet, anklet, belt, etc.
3. DROP — `Status = DROPPED`.
4. Chase — `SKU LeadTime = CHASE`.
5. Long Lead — `SKU LeadTime = LONG-LEAD` (also matches "Long Lead") → its own **"Long Lead"** group.
6. Knits — vendor EEFG (`ELEFG`) or New Vantage (`NEWVA`), **or any style whose name contains "sweater"**
   (unless it's a dress/gown — those stay in Dresses).
7. Outerwear — jacket, coat, blazer, cardigan, bolero, kimono, trench, parka, cape, poncho.
8. Dresses — DRESS or GOWN.
9. Rompers & Jumpsuits — romper, jumpsuit, one piece.
10. Tops — top, tank, tee, bralette, bodysuit, bustier, cami, corset, shirt, blouse, halter.
11. Bottoms — skirt, skort, short, pant, bottom, legging, trouser.
12. Fallback — name has MAXI/MIDI/MINI → Dresses, else Tops.

- White Label — no identifier in the data → omitted.

**Sort within a section:** Fabric (A→Z) → Style Name (A→Z).

**Keep a fabric group together on one row.** Pack the grid group by group: if a group won't fully fit in
the remaining row width, wrap the whole group to the next row (only split a group wider than a full row).

---

## Image handling

- **Crop:** front+back CADs → crop to left (front) figure, then whitespace-trim. A CAD is only treated as
  front+back when AR ≥ 0.45 **and the top strip contains two separated figures** (`n_top_figures ≥ 2`).
  This two-figure guard is essential: a single-figure CAD that carries a colour swatch also has a wide AR,
  and the old AR-only rule sliced the lone figure down the middle (the MAE GOWN bug — fixed June 2026).
  Single-figure-with-swatch and accessories are never cropped; Recolor (`Class = Recolor`) also keeps the
  full image incl. swatch.
- **Size ALL images by HEIGHT, never width:** full-body **1.39 in**, half-body **1.20 in** (wide CADs,
  AR ≥ 0.55); width = height × AR. Flex cell `max(0.59, image_width + 0.10)`; centered, bottom-aligned.
- **Sharpness:** CADs are low-res (≤ 200 px). Trim (threshold ≈ 10), upscale ~4× (LANCZOS) +
  UnsharpMask, save JPEG q95. (Ceiling: to fully match a crisp reference, source full-res CADs from
  PLM/web — Excel thumbnails stay slightly soft.)

---

## Layout & visual settings (LF Femme)

**Canvas = 10 × 5.625 in** (Google default; do NOT use 13.33 × 7.5). Exact constants (inches):

```
CW=10.0  CH=5.625   MX=0.48
GRID_TOP=0.60   ROWPITCH=2.20   ROWS_PER=2
HDR_H=0.28   IMG_H=1.39 (HALF_H=1.20)   GAP_IMG=0.02   TXT_H=0.45   GAP_AFTER=0.06
CARD=0.59   GIN=0.04   GGRP=0.30   CELL_PAD=0.10   HALF_AR=0.55
```

**Fonts (native Google Slides fonts):**
- **Playfair Display** — Category 25 pt; **Brand top-right 20 pt**; Fabric header 5.5 pt; Cover title
  44 pt, subtitle/year 35 pt; Month divider 48 pt (Playfair SemiBold). Colours: brand/cover `3A3A3A`,
  category `666666`, fabric headers black.
- **Lato** — all card text 4 pt, **regular (NOT bold)**: Style name, Colour, Price, Tier, Annotation.

**Text CASE:** render strings in natural case — do NOT uppercase. Category titles are mixed case
("Dresses", "Tops", "Rompers & Jumpsuits", "Knits", "Chase", "DROP" — DROP is the only caps one). Only
the month divider is ALL CAPS ("OCTOBER"). (A previous build wrongly rendered "DRESSES".)

**Alignment & brand:** Card text center-aligned, top-anchored, `wrap="square"`. Fabric headers centered.
Category top-left (Playfair 25, `666666`, mixed case). Brand top-right = "Lovers and Friends", Playfair
Display, **20 pt**, `3A3A3A`, right-aligned, one row. (A previous build wrongly used 9 pt.)

**Colours:** Category `666666` · Fabric header black · Card text (all lines) black · Cover/brand
`3A3A3A` · Divider bg (sage) `D9E8CD` · DROP bg (grey) `ECEBE9` · else white. No taupe headers, no
orange/italic annotation — headers and annotations are plain black.

**Card content (top→bottom):** fabric header · CAD image · STYLE NAME · COLOUR · $PRICE · TIER x ·
`*ANNOTATION*` (all Lato 4, black, centered).

**Slide furniture:** Cover white — centered brand (44) / "Merchandising Assortment" (35) / period (35);
brand also top-right. Month divider sage-green, month ALL CAPS centered (48), brand top-right. Category
page white, category top-left, brand top-right. DROP page = category page on grey.

---

## Google Slides correctness (mandatory)

1. **Exact line spacing, never a percentage** (`<a:spcPts>`, not `<a:spcPct>`) — percentages collapse
   wrapped lines on top of each other in Google Slides.
2. **Pre-break wrapping text into explicit lines, in ONE text box** with `wrap="square"`. Google does
   not advance for soft-wrapped lines, so a wrapping label overlaps the one below; long names also
   overflow sideways and squeeze neighbours. Measure each line (Lato/sans ≈ DejaVu Sans; Playfair/serif
   ≈ DejaVu Serif) and emit one paragraph per visual line inside a single shape. **Never `wrap="none"`**
   and never one text box per line.
3. **No stray borders** (python-pptx text boxes have none by default).

---

## Output & naming

- **Filename = `Lovers and Friends MERCH BOARD <Year> <Period>.pptx`** — brand spelled out, then
  `MERCH BOARD`, year, period. Example: **`Lovers and Friends MERCH BOARD 2026 Oct.pptx`**. Do NOT
  prefix a date, use the buy code, or add a version suffix (not "6.25 MERCH BOARD - LF NOV_1.pptx").
  `<Period>` is confirmed with the user; the cover subtitle matches it; the month divider shows R.Month.
- Always present the final file. File deletion isn't available — if misnamed, save the right name and
  ask the user to delete the old one.

### Boards produced (LF Femme)
| Buy | Release month(s) | Cover period | File |
|---|---|---|---|
| LF Q2  | April / May / June | 2026 Q2  | `Lovers and Friends MERCH BOARD 2026 Q2.pptx` |
| LF Jul | October            | 2026 Oct | `Lovers and Friends MERCH BOARD 2026 Oct.pptx` |
| LF Oct | October (175 styles)| 2026 Oct | `Lovers and Friends MERCH BOARD 2026 Oct.pptx` |
| LF Nov | November           | 2026 Nov | `Lovers and Friends MERCH BOARD 2026 Nov.pptx` |
| LF Q4  | Oct / Nov / Dec (334 of 352 styles) | 2026 Q4 | `Lovers and Friends MERCH BOARD 2026 Q4.pptx` |
| LF 2027 | Jan–Oct 2027 | 2027 | `Lovers and Friends MERCH BOARD 2027.pptx` (22→29 slides via incremental update; +62 new styles) |

---

## Adding a new brand / segment

Keep the engine identical (load → extract → derive → crop → categorize → flow → render). Change only the
segment config: brand strings & filename, period, category list/order, exclusion rules, knit vendors,
palette/fonts, image rules.

**Brand string is parameterized** — set it via the `MB_BRAND` env var (default `Lovers and Friends`); it
drives the cover title, the top-right brand, and the default output filename. Same engine, same categories
and format, just a different brand label:

```
MB_BRAND="NBD" python merch_board.py "Buy_Meeting_Prep_List (NBD 2027).xls" "2027"
```

When a new brand needs different categories/exclusions/knit-vendors (not just a different label), change
those config lists in the script too. NBD reused the LF category model and format unchanged.

### First-run config check for a new brand
Most brands (NBD, Tularosa, Majorelle) just need `MB_BRAND` and run as-is. The first time a brand's buy
file appears, sanity-check four things against that file — they're the only parts of the engine that are
data-dependent, and a mismatch silently mis-files a card rather than erroring:

1. **Stylecode → colour.** Colour is the text after the season code on the last `-` segment, with a
   leading `^[A-Z]\d{2}` stripped (e.g. `…-S27CHEETAH` → `CHEETAH`). Confirm the brand's codes follow this
   shape; if not, adjust `color_of`.
2. **Knit vendors.** Knits are detected by vendor codes `EEFG / NEWVA / NEW VANTAGE / ELEVEN ELEVEN` (or
   a name containing "sweater"). If the brand sources knits from other mills, add those vendor codes, or
   knit styles will fall into Tops/Dresses.
3. **Swim / exclusions.** Swim is dropped via stylecode 3rd letter `X` or vendor Sovereign Sky (`SOVSK`);
   accessories via the keyword list. Verify the brand uses the same swim convention; widen the rules if not.
4. **Categories.** The category list/order (Dresses → Tops → Bottoms → Outerwear → Rompers & Jumpsuits →
   Knits → Chase → Long Lead → DROP) fits contemporary womenswear. If a brand carries silhouettes the
   keyword map doesn't cover, extend `categorize`.

Run, then eyeball one page per category plus a recolor and the DROP page (the standard QA) before sharing.

### Segment registry
| # | Brand | Segment | Status |
|---|---|---|---|
| 1 | Lovers + Friends | Femme | Done — reference segment |
| 2 | NBD | LF category model + format; `MB_BRAND="NBD"` | Done — `NBD MERCH BOARD 2027.pptx`, 12 slides, 38 styles |
| 3 | Tularosa | LF category model + format; `MB_BRAND="Tularosa"` | Ready — run on first buy file; do the first-run config check |
| 4 | Majorelle | LF category model + format; `MB_BRAND="Majorelle"` | Ready — run on first buy file; do the first-run config check |
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 