# Buyer Lineguide (web) — Business Context

## Overview

The Buyer Lineguide deck generator converts a buying team's Excel lineguide — the master product workbook used by Merchandising and Design — into a presentation-ready PowerPoint deck for buy meetings. One slide per style, branded consistently, with all product specs, notes, and costing tiers laid out the way they appear on the printed Excel template.

In this **web project**, the workflow is simple: a user attaches a `.xls` (or `.xlsx`) lineguide to a Claude conversation, and Claude returns the generated `.pptx` deck in the same conversation.

The output PPTX is designed to be imported into Google Slides without re-formatting.

## Why this exists

Buy meetings move quickly across dozens of styles per brand. The Excel lineguide is the single source of truth — but it isn't a presentation format. The historical workarounds (printing the workbook, screen-sharing the spreadsheet, manually rebuilding decks in Slides) are wasteful, hard to read, or slow and inconsistent. This generator removes the manual step: the same Excel that already exists in the buying workflow is converted to a deck automatically, in a layout that mirrors what the buyers already know from print.

## Audience

| Role | How they use the deck |
|---|---|
| Buyer | Lead the buy meeting; reference per-style retail, margin, and notes |
| Lead Designer | Confirm design intent and any pending fit changes |
| Product Developer | Track stylecode, vendor, fabric, and dates |
| Tech Designer | Reference fit notes and design notes |
| Vendor partners | Receive a clean export of the buy decisions for their styles |

## Inputs

A single buyer lineguide workbook attached to the chat. The canonical format is **legacy Excel `.xls`** as exported from the buying system; `.xlsx` is also accepted.

The workbook contains:

- A file-level brand name in column G near the top of the sheet (typically row 1)
- One block per style (approximately 56 rows each), with the standard field layout
- A per-product brand-band cell in column G exactly 9 rows above each Style Name row
- Embedded product images, one per style, anchored to that style's row range
- A Costing Tiers sub-section per style with 1 to 5 quantity rows

### Why `.xls` is preferred

The generator reads cell values **directly from the original `.xls` via xlrd**, not from a `.xls → .xlsx` conversion. Headless converters (LibreOffice and similar) silently re-evaluate formulas during conversion and overwrite the cached values, which corrupts a small number of cells in practice. The clearest observed case is a Duty Rate field cached as `0.0` in the source `.xls` (displaying "0%" in the buyer's printed page) that becomes `0.28` after conversion when the formula `=28/100` is re-evaluated, which would render "28%" in the deck and disagree with the buyer's source of truth. Reading the `.xls` cache preserves the displayed values exactly.

Because xlrd cannot extract embedded `.xls` drawings, the script also requires a `.xlsx` companion (produced on the fly with headless LibreOffice) **only** to pull out the product images and their row anchors. All textual values still come from the original `.xls`.

If the user uploads `.xlsx` directly, the generator will still run, but be aware that cached formulas may differ from the source `.xls`.

## Outputs

A single `.pptx` file, US Letter portrait orientation, named:

**`{M.DD} BUY MEETING - {Brand}.pptx`**

- `{M.DD}` — month with no leading zero, day with two digits (e.g. `6.02`, `12.07`). Uses today's local date unless overridden via `--meeting-date`.
- `{Brand}` — see "Filename brand rule" below.

Examples:
- `6.02 BUY MEETING - NBD.pptx`
- `6.02 BUY MEETING - Lovers + Friends.pptx`
- `6.02 BUY MEETING - FEMME.pptx` (multi-brand fan-out, see below)

The deck contains one slide per style in source order.

### Filename brand rule

- **Single-brand lineguide** — every product carries the same brand → use that brand verbatim.
- **Multi-brand lineguide containing "Lovers + Friends"** → `FEMME`.
- **Multi-brand lineguide containing "All the ways" / "superdown" / "more to come"** (any one of these substrings, case-insensitive) → `LPP`.
- **Multi-brand lineguide otherwise** → fall back to the file-level brand (column G, row 1).

## Slide layout

Every slide is 8.5 × 11 inches, portrait, white background by default. One exception: styles flagged as **Quickstrike = REVOLVE/NUULY** get a light-lavender page background instead (see "Quickstrike background" below).

### Header

- **Product image — off-slide, on the left side of the canvas.** The image is intentionally placed *outside* the slide bounds (in the gray "scratch" area to the left of the slide in Google Slides / PowerPoint edit mode). It is visible while editing but NOT visible during presentation, print, or PDF export. The slot's right edge is anchored at `x = -0.20"` (0.20" off the slide's left edge), and the slot extends leftward by whatever width the category requires:
  - **Garment slot (default)** — `2.85" × 4.89"`. Slot extends from `x = -3.05"` to `x = -0.20"`.
  - **Shoe slot** — `2.00" × 7.00"`. Slot extends from `x = -2.20"` to `x = -0.20"`.

  Top edge is fixed at `y = 0.80"` for both slots (aligned with the yellow band's top edge). Image fits inside the slot with aspect ratio preserved; height auto-adjusts to the source photo's aspect (shoes are wide → short image, garments are tall → tall image). The image's top is anchored to `IMG_TOP` — no vertical centering inside the slot.

  **Category detection.** The slot is chosen per-product from the `Category` field in the lineguide. If `Category == "SHOES"` (case-insensitive) → shoe slot. Anything else → garment slot. Mixed-category lineguides get per-slide sizing automatically.

- **Yellow brand band** — a solid rectangle right-anchored against the slide margin: position `(2.60, 0.80)`, size `5.50 × 0.91`, fill `#FFEB00`, no border.

- **Brand text** — placed in a separate text box at `(2.60, 1.00)` so it survives a Google Slides import (Slides ignores PPT's vertical-anchor-middle). Bold Arial, black, right-aligned, fixed **26pt** regardless of brand length.

### Field grid
A 4-column key/value grid in 9pt Arial begins at `y = 1.85`. Each row uses two label/value pairs side by side, in Excel row order. Labels are bold; values are plain. **No colons after labels.**

Column x-positions (fixed):

| | Left labels | Left values | Right labels | Right values |
|---|---|---|---|---|
| x | 0.45 | 1.85 | 4.45 | 5.95 |
| width | 1.40 | 2.55 | 1.40 | 2.55 |

Row pitch is `0.20"`. A **blank source row** (where both A/B and D/E are empty) becomes a `0.12"` vertical gap — this is what produces the visible breaks between the identity / calendar / commercial / construction groups.

**Auto-shrink for long values.** Style Name, Stylecode, Color, and any other value that would otherwise overflow its 2.55" column at 9pt is automatically shrunk in half-point steps down to a minimum of 6pt so it still fits on one line.

### Notes box
A single bordered rectangle at position `(0.40, 5.75)`, size `7.70 × 3.40`, solid white fill, 0.75pt black border. Inside:
- "Buyer Notes" header in **bold Arial 10pt** at `(0.55, 5.83)`
- Buyer notes body at `(0.55, 6.11)`, width `7.40`, height `1.40`, Arial 9pt
- "Design Notes" header in **bold Arial 10pt** at `(0.55, 7.60)`
- Design notes body at `(0.55, 7.88)`, width `7.40`, height `1.20`, Arial 9pt

Both bodies auto-shrink (minimum 6.5pt) to keep long notes inside their allocated boxes; the two regions never share space, so a long Buyer Note can never push the Design Notes header.

### Quickstrike background
Styles whose **quickstrike** field reads `REVOLVE/NUULY` get a light-lavender page background (`#D9D2E9`) instead of the default white; all other styles stay white. The rule is per-style, so a deck can mix white and lavender pages depending on each style's quickstrike value.

Implementation notes (in `lineguide_deck_generator.py`):
- The background is a full-slide rectangle painted as the slide's **first** shape, so it sits behind everything. It's sized to the slide (0,0 → 8.5×11), so the off-slide product image is untouched, and the yellow band, the white notes card, and the white costing table all render on top and stay readable. Only the visible page and the field-grid area take the lavender tint.
- Detection reads the field labelled **`Quickstrike`** (constant `QUICKSTRIKE_LABEL`) and matches its value **order-insensitively** on the slash-separated tokens: both `REVOLVE/NUULY` and `NUULY/REVOLVE` trigger (real lineguides store the pair in either order), and matching ignores case and spacing. A single token (just `NUULY`) does **not** match — both tokens must be present.
- Pinning to the `Quickstrike` label means the same value appearing in some other field won't tint a page. To go label-agnostic, set `QUICKSTRIKE_LABEL = None`. The color and target program are the constants `QUICKSTRIKE_BG` (`#D9D2E9`) and `QUICKSTRIKE_VALUE` (`REVOLVE/NUULY`); to require an exact ordered string, compare `_norm_qs(value) == _norm_qs(QUICKSTRIKE_VALUE)` instead of the token sets.

### Costing Tiers
- "Costing Tiers" label in **bold Arial 10pt** at `(0.40, 9.30)` — **outside** the table, above-left.
- Table at `(0.40, 9.60)`, total width `7.70`, 7 evenly-distributed columns, all cells with thin black borders and **white fill** (no banded shading).
- Header row: Arial 8pt bold, centered.
- Data rows: Arial 8pt, centered. Row height is `0.28"` for ≤2 data rows and `0.24"` for 3+ rows, with further shrinkage if needed to keep the table on-slide.

## Parsing rules

### Brand
The file-level brand appears at column G near the top of the workbook (typically G1). Each product block additionally carries a per-product brand cell at exactly **`G(Style Name row − 9)`**. If that cell is populated with a brand-looking string (non-numeric, not one of the known costing-tier column headers like `EstTotalCost`), it overrides the file-level brand for that product. Otherwise the file-level brand is used.

This narrow rule is important: scanning column G top-to-bottom would pick up the previous block's costing-tier header (e.g. "EstTotalCost") before reaching the actual brand cell.

### Image-to-product mapping
Product images are embedded in the workbook and anchored to specific rows. The generator reads the drawing-anchor XML (from the .xlsx companion) in document order and pairs each anchor with the product whose Style Name row is the smallest one greater than or equal to the anchor row. This correctly handles:
- Products with no embedded image (slot shows a "Product Image" placeholder text)
- Two products sharing the same source image file (each slide gets its own copy of that image)

Image fit: aspect ratio preserved. The image is fit inside the category-specific slot by height first, falling back to width if it would overflow. Horizontally the image is centered inside its slot; vertically the image's top is anchored to `IMG_TOP = 0.80"` (no vertical centering, because the slot is much taller than most images and the desired alignment is "top of image = top of yellow band").

### Image compression
Embedded images are resized so the longest side is at most **468 px** (= 1.56″ × 300 DPI) and re-encoded as JPEG at quality 88. This yields ~300 DPI in the rendered slide, well above the print-quality threshold, while keeping deck file sizes small.

### Field capture
Each product block contributes two columns of key/value pairs (A/B and D/E) from the Style Name row up to the Buyer Notes row. Both columns are read in row order so the deck reproduces the Excel layout. Blank source rows produce a vertical gap (see Field grid above).

### Value display rules
| Field | Display format |
|---|---|
| DDP / FOB / Estimated Total / Retail / Est Total Duty | `$X` for whole numbers, `$X.XX` for fractional |
| Margin / Duty Rate | Integer percent like `63%` (source values in `[0,1]` are multiplied by 100) |
| Brand text size | Fixed 26pt |
| Buyer / Design Notes | Auto-shrunk to fit allocated height; minimum 6.5pt, maximum 9pt |
| Field value cells | Auto-shrunk to one line; minimum 6pt, maximum 9pt |

**NA suppression.** Values of `NA` or `N/A` (case-insensitive) are treated as empty — the label is still drawn, but no value text box is emitted. This matches what buyers see on the printed Excel template, where formula-driven `NA` placeholders render as empty cells. (E.g. `Fabric Code`, `Main Construction`, and similar fields commonly arrive as `NA`.)

### Notes capture
Buyer Notes content sits between the "Buyer Notes" label row and the "Design Notes" label row. Design Notes content sits between "Design Notes" and "Costing Tiers". All non-empty cells across the row are joined as a line of text; cells containing embedded newlines (`\r\n` in the .xls cache) are split into separate paragraphs. Triple-or-greater blank lines are collapsed to a single blank line to avoid runaway whitespace.

### Costing Tiers capture
Tier rows are captured starting at the "Quantity" header row and continuing while the first cell of each subsequent row parses as a number. The generator supports any tier count from 1 to 5; with more tiers the row-height auto-fit accommodates them too.

**Costing-tier display rules.** Tier-cell numbers are rendered **without a `$` prefix** (e.g. `66.26`, `0`, `65.01`) — this matches the printed Excel template, where the tier table is purely numeric. Quantities are integer (`50`, not `50.00`). Rows are always sorted ascending by Quantity.

## Design decisions and rationale

| Decision | Rationale |
|---|---|
| Portrait US Letter (8.5 × 11) | Matches the Excel print template the team already uses |
| White background | Matches print; keeps focus on product and data |
| Lavender background (#D9D2E9) for Quickstrike = REVOLVE/NUULY | Gives buyers an at-a-glance visual flag for REVOLVE/NUULY quickstrike styles without adding a separate callout; per-style so a deck can mix white and lavender pages |
| Yellow brand band (#FFEB00) | High-contrast brand identifier; same color across all decks |
| Arial throughout | Renders identically in PowerPoint, Google Slides, and PDF |
| Brand text positioned manually (not via vertical-anchor) | Google Slides doesn't honor PPT's vertical-anchor-middle; manual placement keeps it visually centered in both renderers |
| Fixed 26pt brand font | Band is sized to fit any brand name at this size; simpler than length-conditional scaling |
| No colons after field labels | Cleaner look; matches the printed Excel template |
| 4-column field grid with fixed x-positions | Predictable label/value alignment across all slides |
| Read .xls directly (not converted .xlsx) for values | Headless converters silently re-evaluate formulas and corrupt some cached cells |
| Read .xlsx companion for images only | xlrd can't expose embedded .xls drawings |
| Suppress "NA" values | Matches what buyers see on the printed page; avoids visual noise |
| No `$` in costing-tier numbers | Matches printed Excel template |
| Costing-tier table: white fill, no banded shading | Cleaner print look; default banded styles distract from the numbers |
| Notes auto-shrink to fit | Notes length varies wildly between styles; locked positions prevent overflow into the costing table |
| Value cells auto-shrink to one line | Long Style Names / Stylecodes / Fabric Codes otherwise wrap into the row below |
| Image compression to ~300 DPI | Print-quality at slide-display size; keeps file size small enough to share via chat |
| Image positioned off-slide (right edge at x=−0.20") | Keeps the slide content area focused on the data; buyers use the off-slide product photo as a visual reference while editing, but it's hidden during presentation/export |
| Separate slots for shoes vs garments | Shoe photos are wide / short; garment photos are tall / narrow. A single slot can't accommodate both well, so the generator picks the slot per-product from the `Category` field |
| Top-aligned image (no vertical centering) | Aligns the top of the product image with the top of the yellow brand band, giving a consistent visual anchor across mixed-category slides |
| One slide per product | Buy meetings discuss one style at a time |
| Filename `{M.DD} BUY MEETING - {Brand}.pptx` | Standardized convention for the buy-meeting folder; FEMME / LPP umbrellas for multi-brand cases |

## Edge cases supported

- **Empty product blocks** — styles with no buyer or design notes render with empty space inside the notes box.
- **Long notes** — both notes sections auto-shrink to fit.
- **Missing images** — products lacking an embedded image render with a "Product Image" text placeholder.
- **Shared images** — two products referencing the same source image file each receive their own copy in the deck.
- **Duplicate product names** — handled as separate slides (e.g., two `WILLOW PUMP` rows in the same brand).
- **Out-of-order tier rows** — source data may store tier rows in any order; the deck sorts them ascending by Quantity.
- **Variable tier count** — 1, 2, 3, 4, or 5 tier rows all render correctly; row height auto-adjusts to fit.
- **`NA` placeholder values** — suppressed from the value column (label still shown).
- **Per-product brand override** — the cell at `G(Style Name row − 9)` overrides the file-level brand; bogus column-G content from the previous block's costing-tier header is correctly ignored.
- **Long values** — Style Name / Stylecode / Color / Fabric Code auto-shrink so they stay on one line.
- **Non-Latin characters in fields** — Chinese characters in Fabric Code render cleanly (Arial fallback).
- **Quickstrike = REVOLVE/NUULY** — that style's slide gets a lavender (`#D9D2E9`) page background; all other styles stay white. Matching reads the `Quickstrike` field and is order-insensitive, so `NUULY/REVOLVE` triggers too (case/space-insensitive).

## Operational notes

**Google Slides compatibility.** The deck is built specifically to survive a Google Drive import. Brand text uses manual positioning rather than vertical-anchor-middle; table cell borders are written via direct XML; image positioning is in EMU. All techniques round-trip without visual shifts.

**Source of truth.** The Excel workbook remains the canonical record. The deck is a snapshot for discussion. Edits made during the meeting go back into the Excel, not the deck.

**Stateless per conversation.** Unlike the Drive-based scheduled workflow, the web project does not track which lineguides have been processed. Every upload is processed fresh; if the user uploads the same lineguide twice, two decks come back.

## Out of scope

- Editing source data — the Excel remains the source of truth.
- Cross-brand consolidation — one workbook per deck (the FEMME / LPP umbrella names handle multi-brand-within-one-file, not multi-file consolidation).
- Slide-by-slide approvals or comments — the deck is a snapshot for discussion, not a workflow tool.
- Persistent state — no Drive upload, no `.done` markers, no cross-conversation memory. Each conversation processes what was attached and returns the deck.

## Running the generator manually

The attached `lineguide_deck_generator.py` is a standalone CLI:

```
python3 lineguide_deck_generator.py <input.xls> <input.xlsx> [--out-dir DIR] [--meeting-date M.DD]
```

It prints `wrote: <path>` on success.


## Form photos (Google Drive) — v2

The embedded image in the `.xls` is the **CAD**, kept in the off-slide top-right slot. In
addition, vendor **form photos** can be placed around it. Photos live in the shared Drive under
the brand folder, named by stylecode + view (e.g. `MJOW10024-H26 front.jpg`,
`LFOW10127.25JB283.SIDE.jpg`).

Rules:
- **Only fetch the photos needed** — match each style's stylecode prefix (before the first `-`)
  against Drive filenames. Styles with no match render with the CAD only.
- Photos are **transient inputs, never deliverables.** Download/decode into a sandbox temp dir
  (`default_photos_tmp()`), pass `--photos-dir`, and run with `--cleanup-photos` so they are
  deleted after the deck is built. Never write form photos into the user's folder or computer.
- View -> slot mapping (case-insensitive substring on the filename): `front`->front, `back`->back,
  `side` **or** `WR left`->side (these vendors label the left profile "WR left"); anything else
  (`WR right`, detail shots) -> **other**.

Off-slide grid (each cell = the category image slot, `0.15"` gap; everything off the left edge /
below the slide so it's hidden in present/print):

```
[ FRONT ] [ CAD  ] | SLIDE
[ SIDE  ] [ BACK ] |
[ OTHER ][ OTHER ] ...  (below the slide, aligned to the slide's LEFT edge, filling rightward)
```

- CAD: top-right, right edge `x=-0.20"`, top `y=0.80"`.
- Front: top-left (one cell + gap left of CAD). Side: under front. Back: bottom-right (under CAD).
- Other: `top = SLIDE_H + 0.10"` (below the slide), left aligned to `x=0`, each next image `+= slot_w + 0.15`.

## Filename convention — v2 (supersedes FEMME/LPP)

`{M.DD} BUY MEETING - {BRAND_ABBR}.pptx`. `{BRAND_ABBR}` is the brand abbreviation; a multi-brand
lineguide joins distinct brands' abbreviations with ` + ` in first-seen order. Map: MAJORELLE->`MAJ`,
Tularosa->`TULA`, Lovers & Friends / Lovers + Friends->`LF`; unmapped -> first 3 letters uppercased.
Extend `BRAND_ABBR` in `scripts/lineguide_deck_generator.py`. This replaces the earlier FEMME / LPP
umbrella-collapse rule for filenames.
