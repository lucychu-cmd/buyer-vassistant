"""Buyer Lineguide → PPTX deck generator (v2, layout matched to reference).

- Reads cell values from the original .xls via xlrd (preserves cached
  formula values; LibreOffice's conversion to .xlsx silently re-evaluates
  some formulas and corrupts a few cells, e.g. Block 20 Duty Rate).
- Reads embedded images from the converted .xlsx via openpyxl (xlrd can't
  expose .xls drawings).
- Builds the PPTX with the exact layout used in the reference deck.
"""

from __future__ import annotations

import argparse
import glob
import io
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, List, Optional, Tuple

import xlrd
from openpyxl import load_workbook
try:
    from PIL import ImageFont
    _ARIAL_REG = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 9)
except Exception:
    ImageFont = None
    _ARIAL_REG = None
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree


# ---------------------------------------------------------------------------
# Layout constants (matched to reference deck via shape inspection)
# ---------------------------------------------------------------------------

SLIDE_W = Inches(8.5)
SLIDE_H = Inches(11.0)

# Header
# Image slot is anchored OFF the left edge of the slide. The slot's right edge
# always sits IMG_RIGHT_GAP off the slide so the image is fully invisible during
# presentation/export, regardless of slot width.
IMG_RIGHT_GAP = Inches(0.20)
IMG_RIGHT_ANCHOR = Inches(0) - IMG_RIGHT_GAP    # = -0.20"
IMG_TOP    = Inches(0.80)
# Slot dimensions. SHOES use the original 2.00 x 7.00 slot; everything else
# uses the larger 2.85 x 4.89 garment slot. See _img_slot_for_category().
IMG_W      = Inches(2.00)
IMG_H      = Inches(7.00)
IMG_W_GARMENT = Inches(2.85)
IMG_H_GARMENT = Inches(4.89)

BAND_LEFT  = Inches(2.60)
BAND_TOP   = Inches(0.80)
BAND_W     = Inches(5.50)
BAND_H     = Inches(0.91)
BAND_FILL  = RGBColor(0xFF, 0xEB, 0x00)

BRAND_TEXT_LEFT = Inches(2.60)
BRAND_TEXT_TOP  = Inches(1.00)
BRAND_TEXT_W    = Inches(5.50)
BRAND_TEXT_H    = Inches(0.52)

# Field grid columns
COL1_LABEL_LEFT = Inches(0.45)
COL1_LABEL_W    = Inches(1.40)
COL1_VAL_LEFT   = Inches(1.85)
COL1_VAL_W      = Inches(2.55)

COL2_LABEL_LEFT = Inches(4.45)
COL2_LABEL_W    = Inches(1.40)
COL2_VAL_LEFT   = Inches(5.95)
COL2_VAL_W      = Inches(2.55)

FIELD_TOP       = Inches(1.85)
FIELD_ROW_H     = Inches(0.20)
FIELD_BLANK_GAP = Inches(0.12)   # extra gap added by a blank source row

# Notes box
NOTES_BORDER_LEFT = Inches(0.40)
NOTES_BORDER_TOP  = Inches(5.75)
NOTES_BORDER_W    = Inches(7.70)
NOTES_BORDER_H    = Inches(3.40)

BUYER_HDR_LEFT  = Inches(0.55)
BUYER_HDR_TOP   = Inches(5.83)
BUYER_HDR_W     = Inches(2.00)
BUYER_HDR_H     = Inches(0.25)

BUYER_BODY_LEFT = Inches(0.55)
BUYER_BODY_TOP  = Inches(6.11)
BUYER_BODY_W    = Inches(7.40)
BUYER_BODY_H    = Inches(1.40)

DESIGN_HDR_LEFT = Inches(0.55)
DESIGN_HDR_TOP  = Inches(7.60)
DESIGN_HDR_W    = Inches(2.00)
DESIGN_HDR_H    = Inches(0.25)

DESIGN_BODY_LEFT = Inches(0.55)
DESIGN_BODY_TOP  = Inches(7.88)
DESIGN_BODY_W    = Inches(7.40)
DESIGN_BODY_H    = Inches(1.20)

# Costing Tiers
TIERS_LABEL_LEFT = Inches(0.40)
TIERS_LABEL_TOP  = Inches(9.30)
TIERS_LABEL_W    = Inches(2.00)
TIERS_LABEL_H    = Inches(0.25)

TIERS_TABLE_LEFT = Inches(0.40)
TIERS_TABLE_TOP  = Inches(9.60)
TIERS_TABLE_W    = Inches(7.70)
TIERS_ROW_H      = Inches(0.28)   # used for 2-data-row case (3 rows × 0.28")
TIERS_ROW_H_SMALL = Inches(0.24)  # used for ≥3 data rows (auto-shrink)


# ---------------------------------------------------------------------------
# Display formatting
# ---------------------------------------------------------------------------

CURRENCY_FIELDS = {"DDP", "FOB", "Estimated Total", "Retail", "Est Total Duty"}
PERCENT_FIELDS  = {"Margin", "Duty Rate"}

# --- Quickstrike page-background rule -------------------------------------
# When a style's "Quickstrike" field is the REVOLVE/NUULY program, that style's
# slide gets a light-lavender page background instead of the default white.
#
# Matching is order-insensitive on the slash-separated tokens, so both
# "REVOLVE/NUULY" and "NUULY/REVOLVE" trigger, and it's case/space-insensitive.
# It is pinned to the field labelled "Quickstrike" (QUICKSTRIKE_LABEL), so the
# same value appearing in some other field won't tint a page. To go
# label-agnostic, set QUICKSTRIKE_LABEL = None. To require an exact ordered
# string instead of a token set, compare
# _norm_qs(value) == _norm_qs(QUICKSTRIKE_VALUE) directly.
QUICKSTRIKE_BG    = RGBColor(0xD9, 0xD2, 0xE9)  # #D9D2E9 light lavender
QUICKSTRIKE_VALUE = "REVOLVE/NUULY"
QUICKSTRIKE_LABEL = "QUICKSTRIKE"  # normalized label to match; None = any field


def _norm_qs(v) -> str:
    """Normalize a cell for quickstrike comparison: uppercase, no spaces."""
    return "".join(str(v).split()).upper() if v is not None else ""


def _qs_tokens(v) -> frozenset:
    """Slash-separated tokens of a normalized cell, as an order-insensitive set."""
    return frozenset(t for t in _norm_qs(v).split("/") if t)


def _is_quickstrike(product: "Product") -> bool:
    """True if this style's Quickstrike field is the NUULY/REVOLVE program.

    Order-insensitive on the slash tokens; pinned to QUICKSTRIKE_LABEL when set.
    """
    target = _qs_tokens(QUICKSTRIKE_VALUE)
    want_label = _norm_qs(QUICKSTRIKE_LABEL) if QUICKSTRIKE_LABEL else None
    for kL, vL, kR, vR, _blank in product.field_rows:
        for label, value in ((kL, vL), (kR, vR)):
            if want_label is not None and _norm_qs(label) != want_label:
                continue
            if _qs_tokens(value) == target:
                return True
    return False


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def fmt_value(key: str, value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip() if not isinstance(value, (int, float)) else None
    if s == "":
        return ""
    # The reference deck treats "NA" placeholders as empty (no value shown)
    if s is not None and s.upper() in {"NA", "N/A"}:
        return ""
    if key in CURRENCY_FIELDS:
        try:
            f = float(value)
            if f == int(f):
                return f"${int(f)}"
            return f"${f:.2f}"
        except (TypeError, ValueError):
            return str(value)
    if key in PERCENT_FIELDS:
        try:
            f = float(value)
            if 0 < f <= 1:
                f *= 100
            return f"{int(round(f))}%"
        except (TypeError, ValueError):
            return str(value)
    # Generic number: drop .0 suffix for ints
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value)


def fmt_tier_cell(header: str, value: Any) -> str:
    """Tier cells are NOT $-prefixed in the reference."""
    if value is None or _norm(value) == "":
        return ""
    h_lower = header.lower()
    if h_lower == "quantity":
        try:
            return str(int(float(value)))
        except (TypeError, ValueError):
            return str(value)
    # Numeric: show as int if integer, else 2-decimal
    try:
        f = float(value)
        if f == int(f):
            return str(int(f))
        return f"{f:.2f}"
    except (TypeError, ValueError):
        return str(value)



LPP_TRIGGERS = ("all the ways", "superdown", "more to come")


def effective_filename_brand(file_brand: str, product_brands) -> str:
    """Return the brand label to use in the output filename.

    Rules:
    - Single brand throughout the lineguide -> that brand
    - Multi-brand AND contains "Lovers + Friends"   -> "FEMME"
    - Multi-brand AND contains any of LPP_TRIGGERS  -> "LPP"
    - Multi-brand otherwise                          -> file-level brand
    """
    distinct = {b for b in product_brands if b}
    if len(distinct) <= 1:
        return file_brand or (next(iter(distinct), "") if distinct else "")
    blower = {b.lower() for b in distinct}
    if any("lovers + friends" in b for b in blower):
        return "FEMME"
    if any(any(t in b for t in LPP_TRIGGERS) for b in blower):
        return "LPP"
    return file_brand


# --- Brand abbreviation filename convention ---------------------------------
# Filename uses the brand ABBREVIATION; a multi-brand lineguide joins each distinct
# brand's abbreviation with " + " in first-seen order. Extend BRAND_ABBR as needed;
# unmapped brands fall back to first 3 letters uppercased. (Supersedes the FEMME/LPP
# umbrella-collapse rule for filenames.)
BRAND_ABBR = {
    "majorelle": "MAJ",
    "tularosa": "TULA",
    "lovers & friends": "LF",
    "lovers + friends": "LF",
}


def brand_abbr(brand: str) -> str:
    key = (brand or "").strip().lower()
    if key in BRAND_ABBR:
        return BRAND_ABBR[key]
    return re.sub(r"[^A-Za-z]", "", brand or "")[:3].upper()


def filename_brand_label(file_brand, products) -> str:
    seen = []
    for p in products:
        b = p.brand or file_brand
        if b and b not in seen:
            seen.append(b)
    out = []
    for b in seen:
        a = brand_abbr(b)
        if a not in out:
            out.append(a)
    return " + ".join(out)


def brand_font_size(brand: str) -> int:
    # Fixed at 26pt; the band is sized to fit any brand name at this size.
    return 26


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Product:
    brand: str
    style_row: int  # 1-indexed
    # field grid rows (in source order) as (label_L, val_L, label_R, val_R, is_blank)
    field_rows: List[Tuple[Optional[str], Any, Optional[str], Any, bool]] = field(default_factory=list)
    buyer_notes: str = ""
    design_notes: str = ""
    tier_headers: List[str] = field(default_factory=list)
    tier_rows: List[List[Any]] = field(default_factory=list)
    image_bytes: Optional[bytes] = None
    image_ext: Optional[str] = None


# ---------------------------------------------------------------------------
# .xls parsing via xlrd
# ---------------------------------------------------------------------------

def _xls_cell_text(sheet, row: int, col: int) -> str:
    """0-indexed; return stripped string or ''. Dates kept as text from cache."""
    try:
        c = sheet.cell(row, col)
    except IndexError:
        return ""
    v = c.value
    if v is None or v == "":
        return ""
    if isinstance(v, float):
        return v
    if isinstance(v, str):
        return v.replace("\r\n", "\n").replace("\r", "\n").rstrip()
    return v


def _xls_value(sheet, row: int, col: int):
    """0-indexed; raw value (float, str, etc.) or None."""
    try:
        c = sheet.cell(row, col)
    except IndexError:
        return None
    v = c.value
    if v == "":
        return None
    if isinstance(v, str):
        v = v.replace("\r\n", "\n").replace("\r", "\n")
    return v


def parse_xls(xls_path: Path) -> Tuple[str, List[Product]]:
    book = xlrd.open_workbook(str(xls_path), formatting_info=False)
    sheet = book.sheet_by_index(0)

    # File-level brand at column G (index 6) near top
    file_brand = ""
    for r in range(min(20, sheet.nrows)):
        v = _xls_value(sheet, r, 6)
        if v:
            file_brand = str(v).strip()
            break

    # Find all "Style Name" rows (col A = index 0)
    style_rows: List[int] = []
    for r in range(sheet.nrows):
        if _norm(_xls_value(sheet, r, 0)) == "Style Name":
            style_rows.append(r)

    products: List[Product] = []
    for i, sr in enumerate(style_rows):
        next_sr = style_rows[i + 1] if i + 1 < len(style_rows) else sheet.nrows

        # Locate section labels inside this block
        bn = dn = ct = qh = None
        for r in range(sr, next_sr):
            a = _norm(_xls_value(sheet, r, 0))
            if a == "Buyer Notes" and bn is None:
                bn = r
            elif a == "Design Notes" and dn is None:
                dn = r
            elif a == "Costing Tiers" and ct is None:
                ct = r
            elif a == "Quantity" and qh is None and ct is not None:
                qh = r
        if bn is None: bn = sr + 21
        if dn is None: dn = sr + 26
        if ct is None: ct = sr + 38
        if qh is None: qh = sr + 40

        # Per-product brand override: cell at (sr-9, col G)
        per_brand = ""
        cand = max(0, sr - 9)
        v = _xls_value(sheet, cand, 6)
        if v is not None:
            s = str(v).strip()
            blocked = {"esttotalcost", "esttotalfob", "quantity"}
            try:
                float(s); is_num = True
            except (TypeError, ValueError):
                is_num = False
            if s and not is_num and s.lower() not in blocked:
                per_brand = s
        brand = per_brand or file_brand

        # Field grid rows: sr..bn-1, columns (A,B) and (D,E)
        field_rows = []
        for r in range(sr, bn):
            kL = _norm(_xls_value(sheet, r, 0))
            vL = _xls_value(sheet, r, 1)
            kR = _norm(_xls_value(sheet, r, 3))
            vR = _xls_value(sheet, r, 4)
            is_blank = not (kL or _norm(vL) or kR or _norm(vR))
            field_rows.append((kL or None, vL, kR or None, vR, is_blank))

        # Notes
        buyer_notes = _collect_notes_xls(sheet, bn + 1, dn - 1)
        design_notes = _collect_notes_xls(sheet, dn + 1, ct - 1)

        # Costing tiers
        # Header row qh: enumerate non-empty columns
        headers: List[Tuple[int, str]] = []
        for c in range(min(sheet.ncols, 12)):
            v = _xls_value(sheet, qh, c)
            if v is None or _norm(v) == "":
                continue
            headers.append((c, str(v).strip()))
        header_cols = [c for c, _ in headers]
        header_labels = [lbl for _, lbl in headers]

        tier_rows: List[List[Any]] = []
        for r in range(qh + 1, min(qh + 8, next_sr)):
            a = _xls_value(sheet, r, 0)
            if isinstance(a, (int, float)) or (isinstance(a, str) and a.replace(".", "", 1).isdigit()):
                row_vals = [_xls_value(sheet, r, c) for c in header_cols]
                tier_rows.append(row_vals)
            else:
                break

        def _qty(row):
            try:
                return float(row[0])
            except (TypeError, ValueError):
                return float("inf")
        tier_rows.sort(key=_qty)

        products.append(Product(
            brand=brand,
            style_row=sr + 1,   # store 1-indexed for clarity
            field_rows=field_rows,
            buyer_notes=buyer_notes,
            design_notes=design_notes,
            tier_headers=header_labels,
            tier_rows=tier_rows,
        ))

    return file_brand, products


def _collect_notes_xls(sheet, r_from: int, r_to: int) -> str:
    """r_from/r_to are 0-indexed inclusive. Returns the notes text with
    multi-line cell contents preserved and 3+ blank lines collapsed to 1."""
    if r_from > r_to:
        return ""
    lines: List[str] = []
    for r in range(r_from, r_to + 1):
        parts: List[str] = []
        for c in range(7):
            v = _xls_value(sheet, r, c)
            if v is None:
                continue
            s = str(v).strip()
            if not s:
                continue
            parts.append(s)
        joined = " ".join(parts).strip()
        if "\n" in joined:
            for sub in joined.split("\n"):
                lines.append(sub.rstrip())
        else:
            lines.append(joined)
    text = "\n".join(lines)
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)
    return text.strip("\n")


# ---------------------------------------------------------------------------
# Image extraction via openpyxl (.xlsx)
# ---------------------------------------------------------------------------

def attach_images(xlsx_path: Path, products: List[Product]) -> None:
    wb = load_workbook(str(xlsx_path), data_only=True)
    ws = wb.active
    images = list(ws._images)
    if not images:
        return
    # Compute anchor row (1-indexed) for each image
    anchors: List[Tuple[int, Any]] = []
    for img in images:
        a = img.anchor
        if hasattr(a, "_from") and a._from is not None:
            row = a._from.row + 1
        else:
            row = 1
        anchors.append((row, img))
    anchors.sort(key=lambda x: x[0])

    sn_rows = sorted(p.style_row for p in products)
    by_sr = {p.style_row: p for p in products}

    for anchor_row, img in anchors:
        # Find the smallest SN row >= anchor_row
        target = None
        for sr in sn_rows:
            if sr >= anchor_row:
                target = sr
                break
        if target is None:
            target = sn_rows[-1]
        prod = by_sr[target]
        if prod.image_bytes is not None:
            continue
        try:
            data = img._data() if callable(getattr(img, "_data", None)) else None
            if data is None and hasattr(img, "ref"):
                ref = img.ref
                if hasattr(ref, "read"):
                    ref.seek(0)
                    data = ref.read()
                elif isinstance(ref, (bytes, bytearray)):
                    data = bytes(ref)
            if data is None:
                continue
        except Exception:
            continue
        prod.image_bytes = data
        prod.image_ext = (getattr(img, "format", None) or "png").lower()


# ---------------------------------------------------------------------------
# Image extraction DIRECTLY from the .xls (LibreOffice-free fallback)
# ---------------------------------------------------------------------------
# The normal path reads embedded CADs from an .xlsx companion produced by
# LibreOffice. When `soffice` is unavailable or fails (some sandboxes have a
# broken LibreOffice that can't load any file), we extract the CADs straight
# from the .xls BIFF stream instead, so the deck still gets its CAD images.
#
# .xls (BIFF8) stores drawings as Escher/MSODRAWING records inside the
# 'Workbook' stream: image bytes live in the drawing-group BLIP store
# (MSODRAWINGGROUP, record 0xEB) and each shape's cell anchor + BLIP index
# live in a per-sheet MSODRAWING record (0xEC). We reassemble those records
# (merging CONTINUE, 0x3C), pull each BLIP's bytes, read each shape's anchor
# row + BLIP index, and map image -> product exactly like attach_images does.

_ESCHER_BLIP_EXT = {
    0xF01D: "jpg", 0xF02A: "jpg", 0xF01E: "png",
    0xF01F: "dib", 0xF029: "tif", 0xF01A: "emf", 0xF01B: "wmf", 0xF01C: "pict",
}
_ESCHER_BLIP_TWO_UID = {0x46B, 0x6E1, 0x6E3, 0x6E5, 0x7A9}  # recInstance -> 2 rgbUids


def _biff_logical_records(wb: bytes):
    """Yield (record_type, payload) with CONTINUE (0x3C) merged into the prior record."""
    import struct
    logical = []
    pos, n = 0, len(wb)
    while pos + 4 <= n:
        rt, rl = struct.unpack("<HH", wb[pos:pos + 4])
        body = wb[pos + 4:pos + 4 + rl]
        if rt == 0x3C and logical:
            logical[-1][1] += body
        else:
            logical.append([rt, bytearray(body)])
        pos += 4 + rl
    return [(rt, bytes(b)) for rt, b in logical]


def _escher_walk(buf):
    """Yield (fbt, ver, inst, body) for every Escher record, descending into containers."""
    import struct
    p, n = 0, len(buf)
    while p + 8 <= n:
        ver_inst, fbt, length = struct.unpack("<HHI", buf[p:p + 8])
        body = buf[p + 8:p + 8 + length]
        yield fbt, ver_inst & 0x0F, ver_inst >> 4, body
        if (ver_inst & 0x0F) == 0xF:  # container
            yield from _escher_walk(body)
        p += 8 + length


def _escher_blips(dgg_blob: bytes):
    """Return [(ext, image_bytes)] for each BLIP in the drawing-group BLIP store, in order.

    Parses the BStoreContainer -> FBSE -> BLIP structure to get exact byte
    boundaries (robust against JPEGs that carry an embedded EXIF thumbnail).
    """
    import struct
    blips = []
    for fbt, _ver, _inst, body in _escher_walk(dgg_blob):
        if fbt != 0xF007:  # OfficeArtFBSE
            continue
        if len(body) < 36:
            continue
        cb_name = body[33]
        off = 36 + cb_name  # FBSE header (36) + optional name
        if off + 8 > len(body):
            continue  # blip stored out-of-line (foDelay) — skip
        vi, rt, rl = struct.unpack("<HHI", body[off:off + 8])
        rec_instance = vi >> 4
        ext = _ESCHER_BLIP_EXT.get(rt)
        data = body[off + 8:off + 8 + rl]
        if ext in ("jpg", "png", "dib", "tif"):
            n_uid = 2 if rec_instance in _ESCHER_BLIP_TWO_UID else 1
            prefix = 16 * n_uid + 1  # rgbUid(s) + 1-byte tag
            blips.append((ext, data[prefix:]))
        elif ext:
            blips.append((ext, data))  # vector formats, rarely used as CADs
    return blips


def _carve_blips(dgg_blob: bytes):
    """Fallback: carve JPEG/PNG images from the drawing blob by file signature."""
    imgs = []
    i, n = 0, len(dgg_blob)
    while i < n:
        j = dgg_blob.find(b"\xff\xd8\xff", i)
        p = dgg_blob.find(b"\x89PNG\r\n\x1a\n", i)
        cands = [x for x in (j, p) if x != -1]
        if not cands:
            break
        s = min(cands)
        if s == j:
            e = dgg_blob.find(b"\xff\xd9", s + 3)
            if e == -1:
                break
            e += 2
            imgs.append(("jpg", dgg_blob[s:e]))
        else:
            e = dgg_blob.find(b"IEND", s)
            if e == -1:
                break
            e += 8  # IEND + 4-byte CRC
            imgs.append(("png", dgg_blob[s:e]))
        i = e
    return imgs


def _valid_images(imgs):
    """Keep only entries whose bytes actually open as an image."""
    out = []
    try:
        from PIL import Image
        import io as _io
        for ext, data in imgs:
            try:
                Image.open(_io.BytesIO(data)).verify()
                out.append((ext, data))
            except Exception:
                out.append((ext, data))  # keep; PPTX can still embed it
    except Exception:
        return imgs
    return out


def attach_images_from_xls(xls_path: Path, products: List[Product]) -> bool:
    """Extract embedded CADs straight from the .xls (no LibreOffice). Returns True
    if at least one product got an image. Best-effort: any failure returns False
    so the caller can fall back to placeholders."""
    import struct
    try:
        import olefile
    except Exception:
        return False
    try:
        ole = olefile.OleFileIO(str(xls_path))
    except Exception:
        return False
    try:
        stream = "Workbook" if ole.exists("Workbook") else ("Book" if ole.exists("Book") else None)
        if stream is None:
            return False
        wb = ole.openstream(stream).read()
    except Exception:
        return False
    finally:
        try:
            ole.close()
        except Exception:
            pass

    recs = _biff_logical_records(wb)
    dgg_blob = b"".join(body for rt, body in recs if rt == 0xEB)   # MSODRAWINGGROUP
    draws = [body for rt, body in recs if rt == 0xEC]              # per-shape MSODRAWING
    if not dgg_blob or not draws:
        return False

    imgs = _escher_blips(dgg_blob)
    if not any(ext in ("jpg", "png") for ext, _ in imgs):
        imgs = _carve_blips(dgg_blob)  # fallback if the structured parse found nothing usable
    imgs = _valid_images(imgs)
    if not imgs:
        return False

    # Per-shape anchor row (0-indexed row1) + BLIP index (pib), in document order.
    anchors = []
    for d in draws:
        row1 = pib = None
        for fbt, _ver, _inst, body in _escher_walk(d):
            if fbt == 0xF010 and len(body) >= 8:  # OfficeArtClientAnchor (Excel): flag,col1,dx1,row1,...
                row1 = struct.unpack("<HHHH", body[:8])[3]
            elif fbt == 0xF00B:  # OfficeArtFOPT
                q = 0
                while q + 6 <= len(body):
                    pid, val = struct.unpack("<HI", body[q:q + 6])
                    q += 6
                    if (pid & 0x3FFF) == 0x0104:  # pib = BLIP index (1-based)
                        pib = val
        anchors.append((row1 if row1 is not None else 0, pib))

    sn_rows = sorted(p.style_row for p in products)
    by_sr = {p.style_row: p for p in products}
    got = False
    for idx, (row1, pib) in enumerate(anchors):
        if pib and 1 <= pib <= len(imgs):
            ext, data = imgs[pib - 1]
        elif idx < len(imgs):
            ext, data = imgs[idx]
        else:
            continue
        anchor_row = row1 + 1  # products store style_row 1-indexed; drawing row1 is 0-indexed
        target = next((sr for sr in sn_rows if sr >= anchor_row), sn_rows[-1])
        prod = by_sr[target]
        if prod.image_bytes is None:
            prod.image_bytes = data
            prod.image_ext = ext
            got = True
    return got


# ---------------------------------------------------------------------------
# Border helper for table cells
# ---------------------------------------------------------------------------

def _set_cell_border(cell, color_hex="000000", weight=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
        ln = etree.SubElement(tcPr, qn(tag))
        ln.set("w", str(weight))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        solid = etree.SubElement(ln, qn("a:solidFill"))
        srgb = etree.SubElement(solid, qn("a:srgbClr"))
        srgb.set("val", color_hex)
        prstDash = etree.SubElement(ln, qn("a:prstDash"))
        prstDash.set("val", "solid")


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------


def _compress_image_bytes(data: bytes, max_dim: int = 468, jpeg_quality: int = 88) -> bytes:
    """Resize and re-encode an embedded product image so the resulting PPTX
    is small enough to upload through MCP create_file (which carries the
    payload as base64 inline, with an effective ~1 MB cap).

    - Resizes so the longest side is <= max_dim px (the image slot in the
      deck is 1.56" x 1.35" — anything beyond ~400 px is invisible.)
    - Re-encodes as JPEG at the given quality.
    - Falls back to the original bytes if PIL isn't available or anything
      goes wrong.
    """
    try:
        from PIL import Image, ImageOps
        import io as _io
        img = Image.open(_io.BytesIO(data))
        # Honor EXIF orientation: bake any camera/phone rotation into the pixels
        # and drop the orientation tag. Re-encoding to JPEG otherwise strips the
        # tag without applying it, so a photo shot in portrait-with-rotation-flag
        # would render sideways (a no-op for photos that carry no such flag).
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass
        if img.mode in ("RGBA", "P"):
            # JPEG can't do alpha — flatten onto white
            bg = Image.new("RGB", img.size, (255, 255, 255))
            try:
                bg.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
            except Exception:
                bg.paste(img)
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
        w, h = img.size
        scale = min(1.0, max_dim / max(w, h))
        if scale < 1.0:
            img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
        out = _io.BytesIO()
        img.save(out, format="JPEG", quality=jpeg_quality, optimize=True, progressive=False)
        return out.getvalue()
    except Exception:
        return data



def _img_slot_for_category(product):
    """Return the (width, height) slot for a product image.

    SHOES keep the original 2.00 x 7.00 slot. Everything else uses the 2.85 x 4.89
    garment slot. Aspect ratio is preserved when the image is placed inside the slot.
    """
    cat = ""
    for kL, vL, _kR, _vR, _is_blank in product.field_rows:
        if kL and kL.strip().lower() == "category":
            cat = (str(vL) if vL is not None else "").strip().lower()
            break
    if cat == "shoes":
        return IMG_W, IMG_H
    return IMG_W_GARMENT, IMG_H_GARMENT


def _add_image(slide, product: Product):
    slot_w, slot_h = _img_slot_for_category(product)
    slot_left = IMG_RIGHT_ANCHOR - slot_w   # slot's RIGHT edge is fixed off-slide; left edge follows the slot width
    if product.image_bytes:
        try:
            compressed = _compress_image_bytes(product.image_bytes)
            stream = io.BytesIO(compressed)
            pic = slide.shapes.add_picture(stream, slot_left, IMG_TOP)
            # Fit inside the image slot, preserving aspect ratio
            ratio = pic.width / pic.height if pic.height else 1.0
            new_h = slot_h
            new_w = int(new_h * ratio)
            if new_w > slot_w:
                new_w = slot_w
                new_h = int(new_w / ratio) if ratio else slot_h
            pic.width = new_w
            pic.height = new_h
            pic.left = slot_left + (slot_w - pic.width) // 2
            pic.top = IMG_TOP
            return
        except Exception:
            pass
    # Placeholder text
    tb = slide.shapes.add_textbox(slot_left, IMG_TOP, slot_w, slot_h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Product Image"
    r.font.name = "Arial"
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)


def _add_band_and_brand(slide, brand: str):
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, BAND_LEFT, BAND_TOP, BAND_W, BAND_H)
    band.fill.solid()
    band.fill.fore_color.rgb = BAND_FILL
    band.line.fill.background()
    band.shadow.inherit = False
    band.text_frame.text = ""

    pt = brand_font_size(brand or "")
    tb = slide.shapes.add_textbox(BRAND_TEXT_LEFT, BRAND_TEXT_TOP, BRAND_TEXT_W, BRAND_TEXT_H)
    tf = tb.text_frame
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = brand or ""
    r.font.name = "Arial"
    r.font.bold = True
    r.font.size = Pt(pt)
    r.font.color.rgb = RGBColor(0, 0, 0)


def _measure_text_pt(text: str) -> float:
    """Approximate width in points of `text` rendered at Arial 9pt."""
    if not text:
        return 0
    if _ARIAL_REG is not None:
        try:
            bbox = _ARIAL_REG.getbbox(text)
            return float(bbox[2] - bbox[0])
        except Exception:
            pass
    # Fallback: rough monospace estimate
    return len(text) * 5.0


def _fit_pt(text: str, max_width_emu: int, default_pt: float = 9.0, min_pt: float = 6.0) -> float:
    """Largest font size <= default_pt and >= min_pt at which `text` fits in max_width_emu.

    PIL measurements tend to underestimate vs LibreOffice/PowerPoint rendering of
    Arial; we use a 5% safety margin and always round DOWN to the nearest half-pt
    so the chosen size is guaranteed to fit (and never gets rounded back up to a
    size that would overflow, which was the SUZI bug).
    """
    import math
    if not text:
        return default_pt
    max_width_pt = max_width_emu / 914400 * 72
    base_w_pt = _measure_text_pt(text)
    # Apply safety margin to the threshold, not just the scaling: even when
    # PIL says it fits at 9pt, leave some breathing room.
    if base_w_pt <= max_width_pt * 0.97:
        return default_pt
    scaled = default_pt * (max_width_pt / base_w_pt) * 0.95
    return max(min_pt, math.floor(scaled * 2) / 2)


def _add_text_cell(slide, left, top, width, height, text: str, *, bold: bool, size_pt: float = 9):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.name = "Arial"
    r.font.bold = bold
    r.font.size = Pt(size_pt)
    r.font.color.rgb = RGBColor(0, 0, 0)


def _add_field_grid(slide, product: Product):
    cursor = FIELD_TOP
    for kL, vL, kR, vR, is_blank in product.field_rows:
        if is_blank:
            cursor += FIELD_BLANK_GAP
            continue
        if kL:
            _add_text_cell(slide, COL1_LABEL_LEFT, cursor, COL1_LABEL_W, FIELD_ROW_H, kL, bold=True)
            txt = fmt_value(kL, vL)
            if txt:
                _add_text_cell(slide, COL1_VAL_LEFT, cursor, COL1_VAL_W, FIELD_ROW_H, txt, bold=False,
                               size_pt=_fit_pt(txt, COL1_VAL_W))
        if kR:
            _add_text_cell(slide, COL2_LABEL_LEFT, cursor, COL2_LABEL_W, FIELD_ROW_H, kR, bold=True)
            txt = fmt_value(kR, vR)
            if txt:
                _add_text_cell(slide, COL2_VAL_LEFT, cursor, COL2_VAL_W, FIELD_ROW_H, txt, bold=False,
                               size_pt=_fit_pt(txt, COL2_VAL_W))
        cursor += FIELD_ROW_H


def _add_notes_box(slide, product: Product):
    # Border rectangle (no fill so it doesn't cover anything underneath; but
    # reference uses solid white fill — match that)
    box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, NOTES_BORDER_LEFT, NOTES_BORDER_TOP, NOTES_BORDER_W, NOTES_BORDER_H
    )
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    box.line.color.rgb = RGBColor(0, 0, 0)
    box.line.width = Pt(0.75)
    box.shadow.inherit = False
    box.text_frame.text = ""

    # Buyer Notes header
    _add_text_cell(slide, BUYER_HDR_LEFT, BUYER_HDR_TOP, BUYER_HDR_W, BUYER_HDR_H,
                   "Buyer Notes", bold=True, size_pt=10)
    _add_notes_body(slide, BUYER_BODY_LEFT, BUYER_BODY_TOP, BUYER_BODY_W, BUYER_BODY_H,
                    product.buyer_notes or "")

    # Design Notes header
    _add_text_cell(slide, DESIGN_HDR_LEFT, DESIGN_HDR_TOP, DESIGN_HDR_W, DESIGN_HDR_H,
                   "Design Notes", bold=True, size_pt=10)
    _add_notes_body(slide, DESIGN_BODY_LEFT, DESIGN_BODY_TOP, DESIGN_BODY_W, DESIGN_BODY_H,
                    product.design_notes or "")


def _add_notes_body(slide, left, top, width, height, text: str):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.word_wrap = True

    # Determine font size with auto-shrink (max 9pt, min 6.5pt)
    lines = text.split("\n") if text else [""]
    # Estimate fit: each pt ≈ 1.2 line-height in pt → in EMU
    h_in = height / 914400
    def lines_at(pt):
        line_h_in = (pt * 1.2) / 72.0
        return max(1, int(h_in / line_h_in))
    target = 9.0
    while target > 6.5 and len(lines) > lines_at(target):
        target -= 0.5
    target = max(6.5, target)

    # First paragraph is the existing first paragraph
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        # remove any default runs
        for r in list(p.runs):
            r.text = ""
        r = p.add_run()
        r.text = line
        r.font.name = "Arial"
        r.font.size = Pt(target)
        r.font.color.rgb = RGBColor(0, 0, 0)


def _add_costing_tiers(slide, product: Product):
    _add_text_cell(slide, TIERS_LABEL_LEFT, TIERS_LABEL_TOP, TIERS_LABEL_W, TIERS_LABEL_H,
                   "Costing Tiers", bold=True, size_pt=10)

    headers = product.tier_headers or [
        "Quantity", "Costing Type", "ActualFOB", "estTotalFreight",
        "Est Total Duty", "estTotalFob", "EstTotalCost",
    ]
    rows = product.tier_rows
    n_cols = len(headers) if headers else 7
    n_data = len(rows)
    n_table_rows = 1 + max(n_data, 1)

    # Row height: 0.28" for ≤2 data rows, 0.24" for 3+, scales down to fit slide
    row_h = TIERS_ROW_H if n_data <= 2 else TIERS_ROW_H_SMALL
    # If too tall, shrink to fit available bottom space (slide height - top - margin)
    available = SLIDE_H - TIERS_TABLE_TOP - Inches(0.05)
    max_h = available // n_table_rows
    if row_h > max_h:
        row_h = max_h
    total_h = row_h * n_table_rows

    tbl_shape = slide.shapes.add_table(n_table_rows, n_cols, TIERS_TABLE_LEFT, TIERS_TABLE_TOP,
                                       TIERS_TABLE_W, total_h)
    tbl = tbl_shape.table
    # Disable the default banded-row style so data rows don't get blue shading
    tbl.first_row = False
    tbl.horz_banding = False
    tbl.vert_banding = False

    # Even column widths
    col_w = TIERS_TABLE_W // n_cols
    for c in range(n_cols):
        tbl.columns[c].width = col_w
    for ri in range(n_table_rows):
        tbl.rows[ri].height = row_h

    # Header row
    for ci, label in enumerate(headers):
        cell = tbl.cell(0, ci)
        _fill_cell(cell, label, bold=True, size_pt=8, bg=RGBColor(0xFF, 0xFF, 0xFF))

    # Data rows
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            _fill_cell(cell, fmt_tier_cell(headers[ci], val), bold=False, size_pt=8)

    # If no data rows, blank-but-bordered single row
    if n_data == 0:
        for ci in range(n_cols):
            cell = tbl.cell(1, ci)
            _fill_cell(cell, "", bold=False, size_pt=8)


def _fill_cell(cell, text: str, *, bold: bool, size_pt: float, bg: RGBColor | None = None):
    if bg is None:
        bg = RGBColor(0xFF, 0xFF, 0xFF)
    cell.text = ""
    tf = cell.text_frame
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = "Arial"
    r.font.bold = bold
    r.font.size = Pt(size_pt)
    r.font.color.rgb = RGBColor(0, 0, 0)
    if bg is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
    _set_cell_border(cell)


# ---------------------------------------------------------------------------
# Sample-photo grid (Google Drive photos) — off-slide
# ---------------------------------------------------------------------------
# The embedded CAD stays in the top-right off-slide slot. Optional sample photos
# (fetched from Drive by the assistant, keyed by stylecode) are placed around it:
#     [ FRONT ] [ CAD  ] | SLIDE      other views -> a row BELOW the slide,
#     [ SIDE  ] [ BACK ] |            aligned to the slide's LEFT edge.
# Photos are TRANSIENT inputs: download them to a sandbox temp dir (default_photos_tmp),
# pass --photos-dir, and run with --cleanup-photos so nothing persists in the user's folder.

PHOTO_GAP = Inches(0.15)


def default_photos_tmp() -> Path:
    import tempfile
    return Path(tempfile.gettempdir()) / "lineguide_photos"


def _photo_stylecode(product) -> str:
    for kL, vL, kR, vR, _ in product.field_rows:
        if kL and kL.strip().lower() == "stylecode":
            return str(vL or "")
        if kR and kR.strip().lower() == "stylecode":
            return str(vR or "")
    return ""


def _classify_view(name: str) -> str:
    n = name.lower()
    if "front" in n:
        return "front"
    if "back" in n:
        return "back"
    if "side" in n or "wr left" in n or "wr_left" in n:
        return "side"
    return "other"


def gather_views(photos_dir, prefix):
    res = {"front": None, "side": None, "back": None, "other": []}
    if not photos_dir or not prefix:
        return res
    for f in sorted(glob.glob(os.path.join(str(photos_dir), prefix + "*"))):
        if not os.path.isfile(f):
            continue
        role = _classify_view(os.path.basename(f))
        data = open(f, "rb").read()
        if role == "other":
            res["other"].append(data)
        elif res[role] is None:
            res[role] = data
        else:
            res["other"].append(data)
    return res


def _place_photo(slide, data, left, top, slot_w, slot_h):
    comp = _compress_image_bytes(data)
    pic = slide.shapes.add_picture(io.BytesIO(comp), left, top)
    ratio = pic.width / pic.height if pic.height else 1.0
    new_h = slot_h
    new_w = int(new_h * ratio)
    if new_w > slot_w:
        new_w = slot_w
        new_h = int(new_w / ratio) if ratio else slot_h
    pic.width = new_w
    pic.height = new_h
    pic.left = left + (slot_w - pic.width) // 2
    pic.top = top


def _cad_placeholder(slide, left, top, w, h):
    tb = slide.shapes.add_textbox(left, top, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "CAD"
    r.font.name = "Arial"
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)


def add_images_with_photos(slide, product, photos_dir):
    """CAD (embedded) top-right off-slide; front/side/back grid to its left;
    other views in a row BELOW the slide aligned to the slide's left edge."""
    slot_w, slot_h = _img_slot_for_category(product)
    cad_left = IMG_RIGHT_ANCHOR - slot_w
    top1 = IMG_TOP
    top2 = top1 + slot_h + PHOTO_GAP
    right_col = cad_left
    left_col = cad_left - PHOTO_GAP - slot_w

    if product.image_bytes:
        _place_photo(slide, product.image_bytes, cad_left, top1, slot_w, slot_h)
    else:
        _cad_placeholder(slide, cad_left, top1, slot_w, slot_h)

    v = gather_views(photos_dir, _photo_stylecode(product).split("-")[0])
    if v["front"]:
        _place_photo(slide, v["front"], left_col, top1, slot_w, slot_h)
    if v["side"]:
        _place_photo(slide, v["side"], left_col, top2, slot_w, slot_h)
    if v["back"]:
        _place_photo(slide, v["back"], right_col, top2, slot_w, slot_h)

    other_top = SLIDE_H + Inches(0.10)
    x = Inches(0)
    for data in v["other"]:
        _place_photo(slide, data, x, other_top, slot_w, slot_h)
        x = x + PHOTO_GAP + slot_w


# ---------------------------------------------------------------------------
# Top-level builder
# ---------------------------------------------------------------------------

def _add_page_background(slide, color: RGBColor):
    """Paint a full-slide rectangle as the page background.

    Added as the slide's FIRST shape so it sits behind everything else. Sized to
    the slide only (0,0 → W×H), so the off-slide product image is untouched. The
    yellow band, notes card, and costing table are drawn afterward and remain on
    top, so only the visible page and the field-grid area take the color.
    """
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    r.fill.solid()
    r.fill.fore_color.rgb = color
    r.line.fill.background()
    r.shadow.inherit = False


def build_deck(xls_path: Path, xlsx_path: Path, out_dir: Path,
               meeting_date: Optional[str] = None, photos_dir=None,
               cleanup: bool = False) -> Path:
    file_brand, products = parse_xls(xls_path)

    # CAD images: prefer the .xlsx companion (LibreOffice conversion); if it's
    # missing or yields nothing (e.g. no/broken LibreOffice), fall back to
    # extracting the CADs straight from the .xls BIFF stream. Either way the
    # deck gets its CADs; if both fail, slides show the "Product Image" placeholder.
    if xlsx_path and Path(xlsx_path).exists():
        try:
            attach_images(xlsx_path, products)
        except Exception:
            pass
    if not any(p.image_bytes for p in products):
        try:
            attach_images_from_xls(xls_path, products)
        except Exception:
            pass

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    for prod in products:
        slide = prs.slides.add_slide(blank)
        if _is_quickstrike(prod):
            _add_page_background(slide, QUICKSTRIKE_BG)
        add_images_with_photos(slide, prod, photos_dir)
        _add_band_and_brand(slide, prod.brand or file_brand)
        _add_field_grid(slide, prod)
        _add_notes_box(slide, prod)
        _add_costing_tiers(slide, prod)

    # Date for the filename: "M.DD" — month with no leading zero, day with leading zero (e.g. "6.02").
    # If the caller passed --meeting-date, accept it verbatim (the caller controls the format).
    if meeting_date is None:
        _now = datetime.now()
        meeting_date = f"{_now.month}.{_now.day:02d}"

    # Multi-brand rule: if a single lineguide contains products under several brands, collapse to
    # the umbrella label (FEMME / LPP) instead of the file-level brand.
    brand_label = filename_brand_label(file_brand, products)
    fname = f"{meeting_date} BUY MEETING - {brand_label}.pptx"
    out_dir.mkdir(parents=True, exist_ok=True)
    final = out_dir / fname
    prs.save(str(final))

    # Photos are transient inputs: never leave them behind.
    if cleanup and photos_dir:
        import shutil
        shutil.rmtree(str(photos_dir), ignore_errors=True)
    return final


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xls", help="original .xls path")
    ap.add_argument("xlsx", nargs="?", default=None,
                    help="OPTIONAL .xlsx companion (LibreOffice conversion) for CAD extraction. "
                         "If omitted or unreadable, CADs are extracted directly from the .xls.")
    ap.add_argument("--out-dir", default=".")
    ap.add_argument("--meeting-date", default=None)
    ap.add_argument("--photos-dir", default=None,
                    help="TEMP folder of Drive sample photos named <STYLECODE>... front/back/side/etc. "
                         "Use a sandbox temp dir (default_photos_tmp), NOT the user output folder.")
    ap.add_argument("--cleanup-photos", action="store_true",
                    help="delete --photos-dir after the deck is built")
    args = ap.parse_args()
    xlsx = Path(args.xlsx) if args.xlsx else None
    final = build_deck(Path(args.xls), xlsx, Path(args.out_dir), args.meeting_date,
                       args.photos_dir, cleanup=args.cleanup_photos)
    print("wrote:", final)


if __name__ == "__main__":
    main()
