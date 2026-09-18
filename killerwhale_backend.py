import math
import re
import sys
import subprocess
import tempfile
from pathlib import Path

try:
    import fitz  # PyMuPDF
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image
except ImportError as e:
    missing = str(e).split("'")[-2] if "'" in str(e) else str(e)
    raise SystemExit(f"Missing package: {missing}. Run: pip install pymupdf openpyxl pillow")


# ============================================================
# MANUAL LEGEND MASK
# ============================================================
# When True, each RE / CE graph opens a preview before it is saved.
# Drag the red vertical bar left/right. Only the detected legend text row
# is whitened from the red bar to the right edge.
MANUAL_LEGEND_MASK = False  # False = KillerWhale Way; True = Manual Way
LEGEND_MASK_X_PERCENT = 70.0
WEB_MASK_PREVIEW_MODE = False
WEB_MASK_PREVIEWS = []
WEB_MANUAL_MASK_FRACTIONS = []
WEB_MANUAL_MASK_INDEX = 0


def _manual_legend_mask_picker(pix, default_x, y0, y1, title="Mask Legend"):
    """Web version of the old draggable Manual Way."""
    global WEB_MANUAL_MASK_INDEX

    if not MANUAL_LEGEND_MASK:
        return int(default_x)

    if WEB_MASK_PREVIEW_MODE:
        import io
        n = pix.n
        mode = "RGB" if n == 3 else "RGBA"
        raw = Image.frombytes(mode, (pix.width, pix.height), bytes(pix.samples))
        if mode == "RGBA":
            raw = raw.convert("RGB")

        max_w = 1000
        if raw.width > max_w:
            ratio = max_w / raw.width
            preview = raw.resize(
                (max_w, max(1, int(round(raw.height * ratio)))),
                Image.Resampling.LANCZOS
            )
        else:
            preview = raw

        bio = io.BytesIO()
        preview.save(bio, format="PNG")
        WEB_MASK_PREVIEWS.append({
            "title": str(title),
            "png": bio.getvalue(),
            "display_width": preview.width,
            "display_height": preview.height,
            "default_fraction": float(default_x) / max(1, pix.width),
        })
        return int(default_x)

    if WEB_MANUAL_MASK_INDEX < len(WEB_MANUAL_MASK_FRACTIONS):
        fraction = WEB_MANUAL_MASK_FRACTIONS[WEB_MANUAL_MASK_INDEX]
        WEB_MANUAL_MASK_INDEX += 1
        if fraction is None:
            return None
        try:
            fraction = max(0.0, min(1.0, float(fraction)))
            return int(round(pix.width * fraction))
        except (TypeError, ValueError):
            pass

    return int(default_x)

def _manual_or_default_mask(buf, pix, default_x, y0, y1, title):
    """Ask for the X boundary and white-out only the supplied legend row."""
    chosen_x = _manual_legend_mask_picker(pix, default_x, y0, y1, title)
    if chosen_x is not None:
        _white_out_pix_rect(buf, pix, chosen_x, y0, pix.width, y1)
    return chosen_x


def format_frequency(value):
    """Format frequency with a space as the thousands separator and 3 decimals."""
    if isinstance(value, (int, float)):
        return f"{value:,.3f}".replace(",", " ")
    return value


def _detect_standard_from_name(name):
    """Return a normalized EMC standard name detected from a filename/title."""
    name = str(name or "").upper()

    if "EN55032" in name or re.search(r"\bEN\b", name):
        return "EN"
    if "FCC" in name:
        return "FCC"
    if "ICES" in name:
        return "ICES"
    if "VCCI" in name:
        return "VCCI"
    if re.search(r"\bKC\b", name):
        return "KC"

    return "OTHER"


def _detect_range_rank(name):
    """Sort common RE frequency ranges from low to high."""
    name = str(name or "").upper()

    if "30M-1G" in name:
        return 0
    if "1G-6G" in name:
        return 1
    if "1G-18G" in name:
        return 2
    if "18G-40G" in name:
        return 3

    return 99


def _detect_voltage_rank(name):
    """Put common AC test voltages in a predictable order."""
    name = str(name or "").upper().replace(" ", "")

    # DC-powered RE files (for example 13.5VDC) sort before AC variants
    # within the same standard/frequency range.
    if re.search(r"\d+(?:\.\d+)?VDC", name):
        return -1

    voltage_order = {
        "100VAC": 0,
        "110VAC": 1,
        "115VAC": 2,
        "120VAC": 3,
        "220VAC": 4,
        "230VAC": 5,
        "240VAC": 6,
    }

    for token, rank in voltage_order.items():
        if token in name:
            return rank

    return 99


def _result_range_rank(result):
    """Determine RE range from extracted data first, filename second.

    This keeps RE sections in the correct frequency order even if filenames or
    selected files are out of order.
    """
    rows = result.get("rows") or []
    freqs = []
    for row in rows:
        try:
            freqs.append(float(row[1]))
        except (TypeError, ValueError, IndexError):
            pass

    if freqs:
        lo, hi = min(freqs), max(freqs)
        if hi <= 1000.5:
            return 0  # 30M-1G
        if lo >= 999.0 and hi <= 6000.5:
            return 1  # 1G-6G
        if lo >= 999.0 and hi <= 18000.5:
            return 2  # 1G-18G
        if lo >= 17999.0 and hi <= 40000.5:
            return 3  # 18G-40G

    return _detect_range_rank(result.get("source_name", ""))


def re_sort_key(result):
    """
    RE ordering.

    Special FCC/ICES pairing:
      FCC 30M-1G 120VAC
      ICES 30M-1G 120VAC
      FCC 30M-1G 240VAC
      ICES 30M-1G 240VAC

    The same FCC-before-ICES pairing is used for other frequency ranges too.
    """
    name = str(result.get("source_name", ""))
    standard = _detect_standard_from_name(name)
    range_rank = _result_range_rank(result)
    voltage_rank = _detect_voltage_rank(name)
    upper_name = name.upper()

    # EN first.
    if standard == "EN":
        return (0, range_rank, voltage_rank, 0, upper_name)

    # FCC + ICES are treated as one paired group.
    # Within each range/voltage: FCC first, then ICES.
    if standard in ("FCC", "ICES"):
        standard_pair_rank = 0 if standard == "FCC" else 1
        return (1, range_rank, voltage_rank, standard_pair_rank, upper_name)

    # Remaining standards.
    remaining_standard_order = {
        "VCCI": 0,
        "KC": 1,
        "OTHER": 99,
    }
    return (
        2,
        remaining_standard_order.get(standard, 99),
        range_rank,
        voltage_rank,
        upper_name,
    )


KC_HEADERS = [
    "No.", "Frequency (MHz)", "Polarization", "Reading AV (dBµV)",
    "Reading PK (dBµV)", "Corr. (dB(1/m))", "Level AV (dBµV/m)",
    "Level PK (dBµV/m)", "Limit AV (dBµV/m)", "Limit PK (dBµV/m)",
    "Margin AV (dB)", "Margin PK (dB)", "Height (cm)", "Angle (deg)"
]

EN_HEADERS = [
    "No.", "Frequency (MHz)", "Polarization", "Reading QP (dBµV)",
    "Corr. (dB(1/m))", "Level QP (dBµV/m)", "Limit QP (dBµV/m)",
    "Margin QP (dB)", "Height (cm)", "Angle (deg)"
]


def _re_float(value):
    """Convert RE numeric text to float, including chamber values like 1,015.900."""
    return float(str(value).replace(",", "").strip())


def parse_numeric_row(line):
    """Parse RE final-result rows, including alternate PDFs whose table columns were rearranged."""
    parts = line.split()
    if not parts or not parts[0].isdigit():
        return None, None

    # Some PDFs split >=1000 MHz as "1" "440.003".
    if len(parts) >= 4:
        try:
            if parts[1].isdigit() and len(parts[1]) <= 3 and "." in parts[2] and parts[3].upper() in ("H", "V"):
                parts = [parts[0], parts[1] + parts[2]] + parts[3:]
        except (ValueError, IndexError):
            pass

    try:
        pol = parts[2].upper() if len(parts) > 2 else ""

        # Normal 30M-1G QP/QPK row.
        if len(parts) == 10 and pol in ("H", "V"):
            row = [int(parts[0]), _re_float(parts[1]), pol] + [_re_float(x) for x in parts[3:]]
            return "EN_QP", row

        # Normal 1G+ AV/PK or CAV/PK+ row.
        if len(parts) == 14 and pol in ("H", "V"):
            row = [int(parts[0]), _re_float(parts[1]), pol] + [_re_float(x) for x in parts[3:]]
            return "KC_AVPK", row

        # FALLBACK: alternate 1G+ report where the PDF table columns were
        # deliberately rearranged. Rebuild the row by EMC relationships rather
        # than trusting the printed left-to-right column order:
        #   Level = Raw + Correction
        #   Margin = Limit - Level
        # This lets the RE Excel output stay in the normal RE column order.
        pol_positions = [i for i, x in enumerate(parts) if x.upper() in ("H", "V")]
        if len(parts) == 14 and pol_positions:
            pol_i = pol_positions[0]
            pol = parts[pol_i].upper()
            no = int(parts[0])
            freq = _re_float(parts[1])
            vals = [_re_float(x) for i, x in enumerate(parts[2:], start=2) if i != pol_i]

            # Correction is the negative factor in these RE final-result rows.
            corr_candidates = [v for v in vals if v < 0]
            if len(corr_candidates) == 1:
                corr = corr_candidates[0]
                remaining = list(vals)
                remaining.remove(corr)

                # Reconstruct the two measurement chains. RE limits in these
                # 1G+ reports are the recognizable 56/60 (CAV) and 76/80 (PK+)
                # values. For each limit, solve the Raw/Level/Margin relationship.
                limit_candidates = [v for v in remaining if any(abs(v-k) <= 0.08 for k in (56.0, 60.0, 76.0, 80.0))]
                chains = []
                for limit in limit_candidates:
                    best = None
                    best_err = 999.0
                    for raw in remaining:
                        if raw == limit:
                            continue
                        for level in remaining:
                            if level in (limit, raw):
                                continue
                            for margin in remaining:
                                if margin in (limit, raw, level):
                                    continue
                                err = abs((raw + corr) - level) + abs((limit - level) - margin)
                                if err < best_err:
                                    best_err = err
                                    best = (raw, level, limit, margin)
                    if best is not None and best_err <= 0.16:
                        chains.append(best)

                if len(chains) >= 2:
                    cav_candidates = [ch for ch in chains if ch[2] < 70]
                    pk_candidates = [ch for ch in chains if ch[2] >= 70]
                    if cav_candidates and pk_candidates:
                        cav = cav_candidates[0]
                        pk = pk_candidates[0]
                        used = [corr] + list(cav) + list(pk)
                        leftovers = list(vals)
                        for u in used:
                            # remove one approximately matching occurrence
                            for j, v in enumerate(leftovers):
                                if abs(v - u) <= 1e-9:
                                    leftovers.pop(j)
                                    break

                        if len(leftovers) == 2:
                            # Height is normally 1-4 m or 100-400 cm; angle is 0-360 deg.
                            a, b = leftovers
                            if 0.5 <= a <= 4.5 and 0 <= b <= 360:
                                height, angle = a * 100.0, b
                            elif 0.5 <= b <= 4.5 and 0 <= a <= 360:
                                height, angle = b * 100.0, a
                            elif 80 <= a <= 450 and 0 <= b <= 360:
                                height, angle = a, b
                            elif 80 <= b <= 450 and 0 <= a <= 360:
                                height, angle = b, a
                            else:
                                height, angle = a, b

                            row = [
                                no, freq, pol,
                                cav[0], pk[0], corr,
                                cav[1], pk[1], cav[2], pk[2],
                                cav[3], pk[3], height, angle,
                            ]
                            return "KC_AVPK", row

    except (ValueError, IndexError):
        pass

    return None, None


def _clean_re_header(value):
    return re.sub(r"\s+", " ", str(value or "").replace("\n", " ")).strip().upper()


def _extract_new_chamber_table_rows(page):
    """Extract the new Rohde & Schwarz RE tables by header name.

    The new chamber changes column order between 1G-18G and 18G-40G, so
    mapping by header is safer than assuming a fixed printed order.
    """
    text = page.get_text("text", sort=True)
    if "EMI Final Results" not in text:
        return None, []

    try:
        tables = page.find_tables().tables
    except Exception:
        return None, []

    for table in tables:
        data = table.extract()
        if not data or len(data) < 2:
            continue
        headers = [_clean_re_header(h) for h in data[0]]
        if not any("FREQUENCY" in h for h in headers) or not any("POLARIZATION" in h for h in headers):
            continue

        def idx_contains(*terms):
            for i, h in enumerate(headers):
                if all(term in h for term in terms):
                    return i
            return None

        freq_i = idx_contains("FREQUENCY")
        pol_i = idx_contains("POLARIZATION")
        height_i = idx_contains("ANTENNA", "HEIGHT")
        angle_i = idx_contains("AZIMUTH")

        qpk_raw_i = idx_contains("QPK", "RAW")
        qpk_level_i = idx_contains("QPK", "LEVEL")
        qpk_limit_i = idx_contains("QPK", "LIMIT")
        qpk_margin_i = idx_contains("QPK", "MARGIN")
        corr_i = idx_contains("CORRECTION")

        if None not in (freq_i, pol_i, qpk_raw_i, corr_i, qpk_level_i, qpk_limit_i, qpk_margin_i, height_i, angle_i):
            rows = []
            for n, cells in enumerate(data[1:], 1):
                try:
                    rows.append([
                        n, _re_float(cells[freq_i]), str(cells[pol_i]).strip().upper(),
                        _re_float(cells[qpk_raw_i]), _re_float(cells[corr_i]),
                        _re_float(cells[qpk_level_i]), _re_float(cells[qpk_limit_i]),
                        _re_float(cells[qpk_margin_i]), _re_float(cells[height_i]),
                        _re_float(cells[angle_i]),
                    ])
                except (TypeError, ValueError, IndexError):
                    continue
            if rows:
                return "EN_QP", rows

        pk_raw_i = idx_contains("PK+", "RAW")
        cav_raw_i = idx_contains("CAV", "RAW")
        pk_level_i = idx_contains("PK+", "LEVEL")
        cav_level_i = idx_contains("CAV", "LEVEL")
        pk_limit_i = idx_contains("PK+", "LIMIT")
        cav_limit_i = idx_contains("CAV", "LIMIT")
        pk_margin_i = idx_contains("PK+", "MARGIN")
        cav_margin_i = idx_contains("CAV", "MARGIN")

        needed = (freq_i, pol_i, cav_raw_i, pk_raw_i, corr_i, cav_level_i, pk_level_i,
                  cav_limit_i, pk_limit_i, cav_margin_i, pk_margin_i, height_i, angle_i)
        if None not in needed:
            rows = []
            for n, cells in enumerate(data[1:], 1):
                try:
                    rows.append([
                        n, _re_float(cells[freq_i]), str(cells[pol_i]).strip().upper(),
                        _re_float(cells[cav_raw_i]), _re_float(cells[pk_raw_i]),
                        _re_float(cells[corr_i]), _re_float(cells[cav_level_i]),
                        _re_float(cells[pk_level_i]), _re_float(cells[cav_limit_i]),
                        _re_float(cells[pk_limit_i]), _re_float(cells[cav_margin_i]),
                        _re_float(cells[pk_margin_i]), _re_float(cells[height_i]),
                        _re_float(cells[angle_i]),
                    ])
                except (TypeError, ValueError, IndexError):
                    continue
            if rows:
                return "KC_AVPK", rows

    return None, []


def extract_rows(page):
    """Extract RE rows only from a page that actually contains a Final Results table.

    RE reports may include extra Hardware Setup / information pages, so do not
    assume a fixed page number and do not try to parse those extra pages as data.
    """
    text = page.get_text("text", sort=True)
    if "EMI Final Results" not in text and "Final Result" not in text:
        return None, []

    # Prefer header-based extraction for the newer chamber tables. This supports
    # 30M-1G, 1G-18G, and 18G-40G even when the printed column order changes.
    table_type, table_rows = _extract_new_chamber_table_rows(page)
    if table_rows:
        return table_type, table_rows

    detected_type = None
    rows = []

    for line in text.splitlines():
        row_type, row = parse_numeric_row(line)
        if row is not None:
            if detected_type is None:
                detected_type = row_type
            if row_type == detected_type:
                rows.append(row)

    return detected_type, rows


# ============================================================
# NEW-CHAMBER RE MATH CHECK
# ============================================================
# The alternate chamber can occasionally print a wrong calculated Level or
# Margin value even though Raw Level, Correction, and Limit are correct.
#
# For NEW chamber PDFs only, verify:
#
#   30M-1G / QP:
#       Level  = Raw Level + Correction
#       Margin = Limit - Level
#
#   1G+ / CAV + PK+:
#       CAV Level  = CAV Raw Level + Correction
#       PK+ Level  = PK+ Raw Level + Correction
#       CAV Margin = CAV Limit - CAV Level
#       PK+ Margin = PK+ Limit - PK+ Level
#
# If the printed value differs by more than this tolerance, the program uses
# the recalculated value in the output Excel. Original-chamber data is untouched.
NEW_RE_MATH_TOLERANCE = 0.06


def _is_new_re_chamber(page):
    """Return True for the alternate/new Rohde & Schwarz RE result-table pages."""
    text = page.get_text("text", sort=True)
    return (
        "EMI Final Results" in text
        and "Frequency" in text
        and ("Raw Lvl" in text or "Raw Lvl" in text.replace("\n", " "))
    )


def _math_differs(actual, expected, tolerance=NEW_RE_MATH_TOLERANCE):
    """True when a printed calculated value is outside the rounding tolerance."""
    try:
        return abs(float(actual) - float(expected)) > float(tolerance)
    except (TypeError, ValueError):
        return True


def check_and_fix_new_re_math(data_type, rows):
    """
    Validate and, when necessary, correct calculated RE values for the new chamber.

    Returns:
        corrected_rows, correction_count, correction_notes

    correction_notes is one string per data row. A blank string means the row
    needed no software correction. When a value is corrected, the note records
    the field name and the original PDF value -> corrected value.

    Raw readings, correction factors, limits, frequency, polarization,
    antenna height, and angle are NEVER changed.
    """
    fixed_rows = []
    correction_count = 0
    correction_notes = []

    for source_row in rows:
        row = list(source_row)
        notes = []

        # 30M-1G / QP normalized order:
        # 0 No, 1 Freq, 2 Pol, 3 RawQP, 4 Corr, 5 LevelQP,
        # 6 LimitQP, 7 MarginQP, 8 Height, 9 Angle
        if data_type == "EN_QP" and len(row) >= 10:
            expected_level = round(float(row[3]) + float(row[4]), 2)
            expected_margin = round(float(row[6]) - expected_level, 2)

            if _math_differs(row[5], expected_level):
                old_value = float(row[5])
                row[5] = expected_level
                correction_count += 1
                notes.append(f"Level QP: {old_value:.2f} -> {expected_level:.2f}")

            if _math_differs(row[7], expected_margin):
                old_value = float(row[7])
                row[7] = expected_margin
                correction_count += 1
                notes.append(f"Margin QP: {old_value:.2f} -> {expected_margin:.2f}")

        # 1G+ / CAV + PK+ normalized order:
        # 0 No, 1 Freq, 2 Pol, 3 RawCAV, 4 RawPK, 5 Corr,
        # 6 CAVLevel, 7 PKLevel, 8 CAVLimit, 9 PKLimit,
        # 10 CAVMargin, 11 PKMargin, 12 Height, 13 Angle
        elif data_type == "KC_AVPK" and len(row) >= 14:
            expected_cav_level = round(float(row[3]) + float(row[5]), 2)
            expected_pk_level = round(float(row[4]) + float(row[5]), 2)
            expected_cav_margin = round(float(row[8]) - expected_cav_level, 2)
            expected_pk_margin = round(float(row[9]) - expected_pk_level, 2)

            checks = (
                (6, "CAV Level", expected_cav_level),
                (7, "PK+ Level", expected_pk_level),
                (10, "CAV Margin", expected_cav_margin),
                (11, "PK+ Margin", expected_pk_margin),
            )

            for idx, field_name, expected in checks:
                if _math_differs(row[idx], expected):
                    old_value = float(row[idx])
                    row[idx] = expected
                    correction_count += 1
                    notes.append(f"{field_name}: {old_value:.2f} -> {expected:.2f}")

        fixed_rows.append(row)
        correction_notes.append("; ".join(notes))

    return fixed_rows, correction_count, correction_notes


# ============================================================
# ALTERNATE / NEW CHAMBER RE GRAPH CROP SETTINGS
# ============================================================
# These settings are ONLY for the new chamber PDFs that contain:
#   "Radiated Emission Test Result" and "EMI Final Results"
#
# LEFT:   larger value = crop more from LEFT toward RIGHT
# RIGHT:  smaller value = crop more from RIGHT toward LEFT
# TOP_ADJUST:    larger value = crop farther DOWN from the TOP
# BOTTOM_ADJUST: larger value = crop farther UP from the BOTTOM
#
# Example fine tuning:
#   crop more from left  -> 0.075 to 0.085
#   crop more from right -> 0.900 to 0.880
#   crop more from top   -> 18 to 25
#   crop more from bottom-> 20 to 30
# ============================================================
NEW_RE_CROP_LEFT = 0.09
NEW_RE_CROP_RIGHT = 0.82
NEW_RE_CROP_TOP_ADJUST = 16
NEW_RE_CROP_BOTTOM_ADJUST = 19


def graph_clip(page):
    """Return an RE graph crop for either the original or alternate chamber PDF."""
    w, h = page.rect.width, page.rect.height
    text = page.get_text("text", sort=True)

    # Alternate chamber layout: large "Radiated Emission Test Result" heading,
    # graph in the upper half, and "EMI Final Results" directly below it.
    # Detect by content instead of filename so selection/order does not matter.
    if "Radiated Emission Test Result" in text and "EMI Final Results" in text:
        title_hits = page.search_for("Radiated Emission Test Result")
        table_hits = page.search_for("EMI Final Results")

        base_top = (max(r.y1 for r in title_hits) + 2) if title_hits else h * 0.145
        base_bottom = (min(r.y0 for r in table_hits) - 10) if table_hits else h * 0.515

        # New-chamber crop is intentionally separate from the original chamber.
        # Change the four NEW_RE_CROP_* values above if you want to fine-tune it.
        left = w * NEW_RE_CROP_LEFT
        right = w * NEW_RE_CROP_RIGHT
        top = base_top + NEW_RE_CROP_TOP_ADJUST
        bottom = base_bottom - NEW_RE_CROP_BOTTOM_ADJUST

        # Safety guards in case a value is adjusted too far.
        left = max(0, min(left, w - 2))
        right = max(left + 1, min(right, w))
        top = max(0, min(top, h - 2))
        bottom = max(top + 1, min(bottom, h))
        return fitz.Rect(left, top, right, bottom)

    # Multi-page RE layout: the graph is on one page and the EMI Final Results
    # table is on a later page. Treat the graph page like the upper half of the
    # normal one-page new-chamber report. In other words, use the SAME crop
    # geometry as the one-page layout; only the table itself lives elsewhere.
    if "Radiated Emission Test Result" in text and "EMI Final Results" not in text:
        title_hits = page.search_for("Radiated Emission Test Result")
        base_top = (max(r.y1 for r in title_hits) + 2) if title_hits else h * 0.145

        # On a normal one-page report the EMI Final Results heading starts at
        # about 51.5% of the page height. Use that virtual boundary here so an
        # extra blank lower half does not change the graph snip.
        base_bottom = h * 0.515

        left = w * NEW_RE_CROP_LEFT
        right = w * NEW_RE_CROP_RIGHT
        top = base_top + NEW_RE_CROP_TOP_ADJUST
        bottom = base_bottom - NEW_RE_CROP_BOTTOM_ADJUST

        left = max(0, min(left, w - 2))
        right = max(left + 1, min(right, w))
        top = max(0, min(top, h - 2))
        bottom = max(top + 1, min(bottom, h))
        return fitz.Rect(left, top, right, bottom)

    # Original chamber layout. Accept both ASCII-u and micro-symbol variants.
    db_hits = []
    for label in ("[dB(uV/m)]", "[dB(µV/m)]", "[dB(μV/m)]"):
        db_hits.extend(page.search_for(label))

    if db_hits:
        top_hit = min(db_hits, key=lambda r: r.y0)
        top = max(0, top_hit.y0 - 3)
    else:
        top = h * 0.17

    left = w * 0.070
    right = w * 0.990

    final_hits = page.search_for("Final Result")
    if final_hits:
        bottom = final_hits[0].y0 - 20
    else:
        bottom = h * 0.635

    return fitz.Rect(left, top, right, bottom)


def save_graph(page, out_png, dpi=220):
    """
    Save the cropped graph and remove anything appearing AFTER the
    frequency range ('30M-1G', '1G-6G', '1G-18G', or '18G-40G') on the Class A legend line.

    This fixes partial leftovers such as:
        <EN55032 Class A 1G-6G 23
        <FCC Class A 1G-18G 24

    so the PNG keeps only the clean range text:
        <EN55032 Class A 1G-6G
        <FCC Class A 1G-18G

    It works for KC, EN, and FCC.
    The source PDF is not changed.
    """
    clip = graph_clip(page)
    scale = dpi / 72.0

    pix = page.get_pixmap(
        matrix=fitz.Matrix(scale, scale),
        clip=clip,
        alpha=False
    )

    # Find the frequency-range anchor in the legend.
    # 30M-1G is used by VCCI / ICES / FCC / EN55032 low-frequency RE graphs.
    # Higher-frequency graphs commonly use 1G-6G, 1G-18G, or 18G-40G.
    # Erase EVERYTHING to the right of the detected range on that line,
    # removing AC-power text such as 100VAC / 240VAC and anything after it.
    anchors = []
    for range_text in ("30M-1G", "1G-6G", "1G-18G", "18G-40G"):
        anchors.extend(page.search_for(range_text))

    # Work on the rendered RE image so we can remove the right-side graph border.
    n = pix.n
    stride = pix.stride
    buf = bytearray(pix.samples)

    # Use the right-side frequency-range legend occurrence as the row anchor.
    valid_anchors = [a for a in anchors if a.intersects(clip)]
    if valid_anchors:
        anchor = max(valid_anchors, key=lambda r: r.x0)
        default_x = max(0, int((anchor.x1 - clip.x0) * scale) + 3)
        py0 = max(0, int((anchor.y0 - clip.y0) * scale) - 4)
        py1 = min(pix.height, int((anchor.y1 - clip.y0) * scale) + 5)
        _manual_or_default_mask(
            buf, pix, default_x, py0, py1,
            "RE Legend Mask — drag red bar, then Apply"
        )

    # Remove the RE graph's vertical black border on the right, same idea as CE cleanup.
    # The crop now extends past the border (right = 0.990), so find the strong vertical
    # dark line in the rightmost part of the image and white it out without cropping it.
    search_x0 = int(pix.width * 0.82)
    best_x = None
    best_dark = 0
    for x in range(search_x0, pix.width):
        dark = 0
        for y in range(pix.height):
            p = y * stride + x * n
            if buf[p] < 90 and buf[p + 1] < 90 and buf[p + 2] < 90:
                dark += 1
        if dark > best_dark:
            best_dark = dark
            best_x = x

    # Require a long vertical line so text/strokes are not mistaken for the border.
    if best_x is not None and best_dark >= int(pix.height * 0.45):
        for x in range(max(0, best_x - 2), min(pix.width, best_x + 3)):
            for y in range(pix.height):
                p = y * stride + x * n
                buf[p] = 255
                buf[p + 1] = 255
                buf[p + 2] = 255

    pix = fitz.Pixmap(
        fitz.csRGB,
        pix.width,
        pix.height,
        bytes(buf),
        False
    )

    pix.save(str(out_png))

def save_excel(rows, out_xlsx, source_name, data_type):
    headers = EN_HEADERS if data_type == "EN_QP" else KC_HEADERS

    wb = Workbook()
    ws = wb.active
    ws.title = "Test Data"

    ws.merge_cells(
        start_row=1, start_column=1,
        end_row=1, end_column=len(headers)
    )
    ws.cell(1, 1, f"EMC Test Data — {source_name}")
    ws.cell(1, 1).font = Font(name="Calibri", size=9, bold=True)
    ws.cell(1, 1).alignment = Alignment(horizontal="center")

    fill = PatternFill("solid", fgColor="D9EAF7")
    no_fill = PatternFill("solid", fgColor="D9D9D9")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, header in enumerate(headers, 1):
        c = ws.cell(3, col, header)
        c.font = Font(name="Calibri", size=9, bold=True)
        c.fill = fill
        c.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        c.border = border

    for r_idx, row in enumerate(rows, 4):
        for c_idx, value in enumerate(row, 1):
            # Force Frequency to exactly 3 visible decimal places.
            excel_value = format_frequency(value) if c_idx == 2 else value
            c = ws.cell(r_idx, c_idx, excel_value)
            c.font = Font(name="Calibri", size=9)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border

            # No. stays integer. All measurement values display to 0.1.
            if isinstance(value, (int, float)):
                if c_idx == 1:
                    c.number_format = "0"       # No.
                elif c_idx == 2:
                    pass  # Frequency is already written as exact 3-decimal text
                else:
                    c.number_format = "0.0"     # All other numeric data: 1 decimal place

    for i, header in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(
            12, min(22, len(header) + 2)
        )

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = (
        f"A3:{get_column_letter(len(headers))}{3 + len(rows)}"
    )

    wb.save(out_xlsx)



CE_HEADERS = [
    "Frequency (MHz)", "Line", "Reading QP (dBµV)",
    "Reading AV (dBµV)", "Factor (dB)", "Level QP (dBµV)",
    "Level AV (dBµV)", "Limit QP (dBµV)", "Limit AV (dBµV)",
    "Margin QP (dB)", "Margin AV (dB)"
]


def read_ce_excel(excel_path, allowed_lines=("L1", "N")):
    """
    Read CE final data from the KillerWhale 'Final Data List(Peak Search)' tab.

    allowed_lines controls which values from source column D are accepted.
    Examples:
      AC line: ("L1", "N")
      DC line: ("+", "-")
    """
    wb = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if "Final Data List(Peak Search)" not in wb.sheetnames:
            raise ValueError(
                "CE Excel does not contain the 'Final Data List(Peak Search)' sheet."
            )

        ws = wb["Final Data List(Peak Search)"]
        rows = []
        allowed = {str(x).strip().upper() for x in allowed_lines}

        # The source layout stores useful CE columns in C:M.
        # Detect data rows by Frequency in C and the measurement line in D.
        for r in range(1, ws.max_row + 1):
            line = str(ws.cell(r, 4).value or "").strip()
            freq = ws.cell(r, 3).value

            if line.upper() not in allowed:
                continue
            if not isinstance(freq, (int, float)):
                continue

            # Keep Frequency through Margin AV (C:M).
            # Drop Range (B), Pass/Fail (N), and Remark (O).
            values = [ws.cell(r, c).value for c in range(3, 14)]
            rows.append(values)

        if not rows:
            expected = ", ".join(str(x) for x in allowed_lines)
            raise ValueError(
                f"No CE final-data rows were found for line(s): {expected}."
            )

        return rows
    finally:
        wb.close()


def _detect_dc_line(page, pdf_path=None):
    """
    Detect DC-line labels in this priority order:
      1) DC Line 1..20
      2) bare DC Line

    PDF text is checked first, then the filename as a fallback.
    """
    sources = [page.get_text("text", sort=True)]
    if pdf_path is not None:
        sources.append(Path(pdf_path).stem)

    for source in sources:
        text = str(source or "")

        match = re.search(
            r"\bDC\s+Line\s+(20|1[0-9]|[1-9])\b",
            text,
            flags=re.IGNORECASE,
        )
        if match:
            return f"DC Line {match.group(1)}"

        # Bare DC Line only. Do not let this fallback steal DC Line 1..20.
        match = re.search(
            r"\bDC\s+Line\b(?!\s*(?:20|1[0-9]|[1-9])\b)",
            text,
            flags=re.IGNORECASE,
        )
        if match:
            return "DC Line"

    return None


def read_tel_ce_excel(excel_path):
    """
    Read Tel-line CE final data from the KillerWhale
    'Final Data List(Peak Search)' tab.

    Output columns:
      Frequency,
      Reading QP, Reading AV,
      Factor,
      Level QP, Level AV,
      Limit QP, Limit AV,
      Margin QP, Margin AV

    Range, Line, Pass/Fail, and Remark are intentionally removed.
    """
    wb = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if "Final Data List(Peak Search)" not in wb.sheetnames:
            raise ValueError(
                "CE Excel does not contain the 'Final Data List(Peak Search)' sheet."
            )

        ws = wb["Final Data List(Peak Search)"]
        rows = []

        for r in range(1, ws.max_row + 1):
            freq = ws.cell(r, 3).value  # C = Frequency
            line_value = ws.cell(r, 4).value  # D = ISN / Tel measurement line

            if not isinstance(freq, (int, float)):
                continue

            # Tel-line files use an ISN measurement line such as "ISN T800".
            # Keep this check loose enough for future ISN names.
            if "ISN" not in str(line_value or "").upper():
                continue

            # C, E:M
            # Drop:
            #   B = Range
            #   D = Line
            #   N = Pass/Fail
            #   O = Remark
            values = [ws.cell(r, 3).value] + [
                ws.cell(r, c).value for c in range(5, 14)
            ]
            rows.append(values)

        if not rows:
            raise ValueError(
                "No Tel-line CE final-data rows were found in the CE Excel file."
            )

        return rows
    finally:
        wb.close()


def _detect_tel_port(page, pdf_path=None):
    """
    Detect Tel-line labels in this priority order:
      1) IPMI0..IPMI20
      2) bare IPMI
      3) ETH0..ETH20
      4) bare Tel Line

    PDF text is checked first, then the filename as a fallback.
    """
    sources = [page.get_text("text", sort=True)]
    if pdf_path is not None:
        sources.append(Path(pdf_path).stem)

    for source in sources:
        text = str(source or "")

        # 1) IPMI0..IPMI20
        match = re.search(r"\bIPMI(?:20|1[0-9]|[0-9])\b", text, flags=re.IGNORECASE)
        if match:
            return match.group(0).upper()

        # 2) Bare IPMI only (do not let IPMI match IPMI0..IPMI20).
        match = re.search(r"\bIPMI\b(?!\d)", text, flags=re.IGNORECASE)
        if match:
            return "IPMI"

        # 3) ETH0..ETH20
        match = re.search(r"\bETH(?:20|1[0-9]|[0-9])\b", text, flags=re.IGNORECASE)
        if match:
            return match.group(0).upper()

        # 4) Generic Tel Line with no recognized IPMI/ETH suffix.
        if re.search(r"\bTel\s+Line\b", text, flags=re.IGNORECASE):
            return "TEL LINE"

    return None


# ============================================================
# TEL LINE GRAPH CROP SETTINGS
# Change these values ONLY for ETH/IPMI Tel-line graphs.
# AC L1/N crop settings below are separate and unchanged.
#
# RIGHT: smaller = crop more from the right
#        larger  = show more on the right
# Example: 0.90 crops more than 0.94; 0.98 shows more.
# ============================================================
TEL_CROP_LEFT = 0.030
TEL_CROP_RIGHT = 0.978
TEL_CROP_TOP_ADJUST = 1
TEL_CROP_BOTTOM_ADJUST = 12


def tel_graph_clip(page):
    """Return the graph+legend rectangle for a Tel-line CE graph."""
    db_hits = page.search_for("[dB(μV)]")
    freq_hits = page.search_for("Frequency")
    scan_hits = page.search_for("Scan(ISN T800, PK)")

    # Fallback for a future ISN name: find a legend word beginning with Scan(
    if not scan_hits:
        words = page.get_text("words")
        scan_words = [
            w for w in words
            if str(w[4]).startswith("Scan(")
        ]
        if scan_words:
            x0, y0, x1, y1 = scan_words[0][:4]
            scan_hits = [fitz.Rect(x0, y0, x1, y1)]

    if not scan_hits:
        return None

    scan = scan_hits[0]

    top_candidates = [r for r in db_hits if r.y0 < scan.y0]
    top_hit = max(top_candidates, key=lambda r: r.y0) if top_candidates else None

    bottom_candidates = [r for r in freq_hits if r.y0 > scan.y0]
    bottom_hit = min(bottom_candidates, key=lambda r: r.y0) if bottom_candidates else None

    top = max(0, (top_hit.y0 - 4) if top_hit else scan.y0 - 50)
    bottom = min(
        page.rect.height,
        (bottom_hit.y1 + 16) if bottom_hit else scan.y1 + 220
    )

    # TEL-LINE crop only. Adjust TEL_CROP_RIGHT above as needed.
    left = max(0, page.rect.width * TEL_CROP_LEFT)
    right = min(page.rect.width, page.rect.width * TEL_CROP_RIGHT)
    bottom -= TEL_CROP_BOTTOM_ADJUST
    top += TEL_CROP_TOP_ADJUST
    return fitz.Rect(left, top, right, bottom)


def save_tel_ce_graph(page, port, out_png, dpi=220):
    """
    Save one Tel-line CE graph and clean the standard legend line.

    Examples:
      <KC Class A Tel Line ETH0 220VAC-60Hz>  -> keep through ETH0
      <KC Class A Tel Line ETH20 220VAC-60Hz> -> keep through ETH20
      <KC Class A Tel Line IPMI 220VAC-60Hz>  -> keep through IPMI
      <KC Class A Tel Line IPMI20 ...>        -> keep through IPMI20
    """
    clip = tel_graph_clip(page)
    if clip is None:
        raise ValueError("Could not find the Tel-line CE graph in the PDF.")

    scale = dpi / 72.0
    pix = page.get_pixmap(
        matrix=fitz.Matrix(scale, scale),
        clip=clip,
        alpha=False
    )
    buf = bytearray(pix.samples)

    # Search only the detected label. This avoids ETH1 accidentally matching ETH10/ETH11/etc.,
    # or IPMI1 accidentally matching IPMI10/IPMI11/etc.. For the generic fallback, mask after
    # the bare "Tel Line" label.
    search_label = "Tel Line" if port == "TEL LINE" else port
    port_hits = [
        r for r in page.search_for(search_label)
        if r.intersects(clip) and r.x0 > page.rect.width * 0.70
    ]

    if port_hits:
        # Use the right-side legend occurrence.
        anchor = min(port_hits, key=lambda r: r.y0)
        default_x = int((anchor.x1 - clip.x0) * scale) + 2
        py0 = int((anchor.y0 - clip.y0) * scale) - 4
        py1 = int((anchor.y1 - clip.y0) * scale) + 5
        _manual_or_default_mask(
            buf, pix, default_x, py0, py1,
            f"CE Tel Line Legend Mask — {port}"
        )

    cleaned = fitz.Pixmap(
        fitz.csRGB, pix.width, pix.height, bytes(buf), False
    )
    cleaned.save(str(out_png))


def ce_graph_clip(page, line):
    """Return the graph+legend rectangle for AC (L1/N) or DC (+/-) CE graphs."""
    if line not in ("L1", "N", "+", "-"):
        raise ValueError("CE line must be 'L1', 'N', '+', or '-'.")

    db_hits = page.search_for("[dB(μV)]")
    freq_hits = page.search_for("Frequency")
    scan_hits = page.search_for(f"Scan({line}, PK)")

    if not scan_hits:
        return None

    scan = scan_hits[0]

    # Choose the dB label above the matching scan legend, and the matching
    # Frequency label below that graph. This works with stacked L1/N graphs.
    top_candidates = [r for r in db_hits if r.y0 < scan.y0]
    top_hit = max(top_candidates, key=lambda r: r.y0) if top_candidates else None

    bottom_candidates = [r for r in freq_hits if r.y0 > scan.y0]
    bottom_hit = min(bottom_candidates, key=lambda r: r.y0) if bottom_candidates else None

    top = max(0, (top_hit.y0 - 4) if top_hit else scan.y0 - 50)
    bottom = min(page.rect.height, (bottom_hit.y1 + 16) if bottom_hit else scan.y1 + 220)

    # Include plot and right-side legend, while dropping most page header whitespace.
    left = max(0, page.rect.width * 0.030)
    right = min(page.rect.width, page.rect.width * 0.94)
    bottom = bottom - 12
    top = top + 1
    return fitz.Rect(left, top, right, bottom)


def _white_out_pix_rect(buf, pix, x0, y0, x1, y1):
    """Paint a pixel rectangle white in a pixmap bytearray."""
    n = pix.n
    stride = pix.stride
    x0 = max(0, min(pix.width, int(x0)))
    x1 = max(0, min(pix.width, int(x1)))
    y0 = max(0, min(pix.height, int(y0)))
    y1 = max(0, min(pix.height, int(y1)))

    for y in range(y0, y1):
        row_start = y * stride
        for x in range(x0, x1):
            p = row_start + x * n
            buf[p] = 255
            buf[p + 1] = 255
            buf[p + 2] = 255


def save_ce_graph(page, line, out_png, dpi=220, ce_kind="AC", dc_label=None):
    """
    Save one AC or DC conducted-emission graph.

    AC cleanup masks after AC Line 1..20, then bare AC Line.
    DC cleanup masks after DC Line 1..20, then bare DC Line.
    """
    clip = ce_graph_clip(page, line)
    if clip is None:
        raise ValueError(f"Could not find the CE {line} graph in the PDF.")

    scale = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
    buf = bytearray(pix.samples)

    ce_kind = str(ce_kind or "AC").upper()
    anchor = None
    mask_title = f"CE {ce_kind} Legend Mask — {line}"

    if ce_kind == "DC":
        # Prefer the exact detected DC label, so DC Line 1 does not accidentally
        # match DC Line 10/11/etc. If no numbered label exists, use bare DC Line.
        labels = []
        if dc_label:
            labels.append(str(dc_label))
        labels.extend([f"DC Line {i}" for i in range(20, 0, -1)])
        labels.append("DC Line")

        seen = set()
        labels = [x for x in labels if not (x.upper() in seen or seen.add(x.upper()))]

        for label in labels:
            hits = [
                r for r in page.search_for(label)
                if r.intersects(clip) and r.x0 > page.rect.width * 0.55
            ]
            if hits:
                # For bare DC Line, skip it if the next token is a number 1..20.
                if label.upper() == "DC LINE":
                    words = page.get_text("words")
                    valid = []
                    for candidate in hits:
                        same_row_after = [
                            w for w in words
                            if w[0] >= candidate.x1 - 1
                            and abs(w[1] - candidate.y0) <= max(3, candidate.height * 0.6)
                            and abs(w[3] - candidate.y1) <= max(3, candidate.height * 0.6)
                        ]
                        same_row_after.sort(key=lambda w: w[0])
                        next_word = str(same_row_after[0][4]).strip() if same_row_after else ""
                        if re.fullmatch(r"(?:20|1[0-9]|[1-9])", next_word):
                            continue
                        valid.append(candidate)
                    hits = valid
                if hits:
                    anchor = min(hits, key=lambda r: r.y0)
                    break

    else:
        # Existing AC behavior: numbered AC Line 1..20 first, then bare AC Line.
        for line_label in [f"Line {i}" for i in range(20, 0, -1)]:
            line_hits = [
                r for r in page.search_for(line_label)
                if r.intersects(clip) and r.x0 > page.rect.width * 0.70
            ]
            if line_hits:
                anchor = min(line_hits, key=lambda r: r.y0)
                break

        if anchor is None:
            words = page.get_text("words")
            ac_line_hits = [
                r for r in page.search_for("AC Line")
                if r.intersects(clip) and r.x0 > page.rect.width * 0.70
            ]
            for candidate in ac_line_hits:
                same_row_after = [
                    w for w in words
                    if w[0] >= candidate.x1 - 1
                    and abs(w[1] - candidate.y0) <= max(3, candidate.height * 0.6)
                    and abs(w[3] - candidate.y1) <= max(3, candidate.height * 0.6)
                ]
                same_row_after.sort(key=lambda w: w[0])
                next_word = str(same_row_after[0][4]).strip() if same_row_after else ""
                if re.fullmatch(r"(?:20|1[0-9]|[1-9])", next_word):
                    continue
                anchor = candidate
                break

    if anchor is not None:
        default_x = int((anchor.x1 - clip.x0) * scale) + 2
        py0 = int((anchor.y0 - clip.y0) * scale) - 4
        py1 = int((anchor.y1 - clip.y0) * scale) + 5
        _manual_or_default_mask(buf, pix, default_x, py0, py1, mask_title)

    cleaned = fitz.Pixmap(
        fitz.csRGB, pix.width, pix.height, bytes(buf), False
    )
    cleaned.save(str(out_png))


def build_ce_results(pdf_path, excel_path, temp_dir):
    """
    Create CE section data by pairing PDF graphs with Excel final-data rows.

    Auto-detects:
      - AC line CE: L1 / N
      - DC line CE: + / -, with DC Line 1..20 then bare DC Line detection
      - Tel line CE: IPMI0..IPMI20, bare IPMI, ETH0..ETH20, and bare Tel Line
    """
    pdf_path = Path(pdf_path)
    excel_path = Path(excel_path)
    temp_dir = Path(temp_dir)

    doc = fitz.open(pdf_path)
    try:
        # Detect Tel-line CE from PDF content.
        tel_page = None
        tel_port = None

        for page in doc:
            page_text = page.get_text("text", sort=True)
            if "Tel Line" in page_text or "Scan(ISN" in page_text:
                port = _detect_tel_port(page, pdf_path)
                if port:
                    tel_page = page
                    tel_port = port
                    break

        if tel_page is not None:
            rows = read_tel_ce_excel(excel_path)

            safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", pdf_path.stem)
            graph_path = temp_dir / f"ce_tel_graph_{safe_stem}_{tel_port}.png"
            save_tel_ce_graph(tel_page, tel_port, graph_path)

            return {
                "ce_type": "TEL",
                "pdf_name": pdf_path.name,
                "excel_name": excel_path.name,
                "sections": [{
                    "line": tel_port,
                    "rows": rows,
                    "graph_path": graph_path,
                }],
            }

        # Detect DC-line CE before falling back to AC.
        dc_page = None
        dc_label = None
        for page in doc:
            page_text = page.get_text("text", sort=True)
            if re.search(r"\bDC\s+Line\b", page_text, flags=re.IGNORECASE) \
                    or page.search_for("Scan(+, PK)") \
                    or page.search_for("Scan(-, PK)"):
                label = _detect_dc_line(page, pdf_path)
                if label:
                    dc_page = page
                    dc_label = label
                    break

        if dc_page is not None:
            rows = read_ce_excel(excel_path, allowed_lines=("+", "-"))
            by_line = {
                "+": [r for r in rows if str(r[1]).strip() == "+"],
                "-": [r for r in rows if str(r[1]).strip() == "-"],
            }

            sections = []
            safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", pdf_path.stem)

            for line in ("+", "-"):
                matching_page = None
                for page in doc:
                    if page.search_for(f"Scan({line}, PK)"):
                        matching_page = page
                        break

                if matching_page is None:
                    continue

                graph_path = temp_dir / f"ce_dc_graph_{safe_stem}_{'plus' if line == '+' else 'minus'}.png"
                save_ce_graph(
                    matching_page, line, graph_path,
                    ce_kind="DC", dc_label=dc_label
                )
                sections.append({
                    "line": line,
                    "rows": by_line[line],
                    "graph_path": graph_path,
                })

            if not sections:
                raise ValueError(
                    "DC Line was detected, but no Scan(+, PK) or Scan(-, PK) graph was found."
                )

            return {
                "ce_type": "DC",
                "dc_label": dc_label,
                "pdf_name": pdf_path.name,
                "excel_name": excel_path.name,
                "sections": sections,
            }

        # Otherwise process as the existing AC-line L1/N format.
        rows = read_ce_excel(excel_path, allowed_lines=("L1", "N"))
        by_line = {
            "L1": [r for r in rows if str(r[1]).strip() == "L1"],
            "N": [r for r in rows if str(r[1]).strip() == "N"],
        }

        sections = []

        for line in ("L1", "N"):
            matching_page = None
            for page in doc:
                if page.search_for(f"Scan({line}, PK)"):
                    matching_page = page
                    break

            if matching_page is None:
                continue

            safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", pdf_path.stem)
            graph_path = temp_dir / f"ce_graph_{safe_stem}_{line}.png"
            save_ce_graph(matching_page, line, graph_path, ce_kind="AC")
            sections.append({
                "line": line,
                "rows": by_line[line],
                "graph_path": graph_path,
            })

        if not sections:
            raise ValueError(
                "No supported CE graph was found. Expected AC L1/N, DC +/-, "
                "or ETH/IPMI Tel line."
            )

        return {
            "ce_type": "AC",
            "pdf_name": pdf_path.name,
            "excel_name": excel_path.name,
            "sections": sections,
        }

    finally:
        doc.close()


def write_ce_table(ws, start_row, start_col, rows, fill, border):
    """Write one AC/DC CE table with a gray No. column.

    CE rows are sorted by Frequency from lowest to highest before writing.
    The complete row moves together, so Line/Reading/Factor/Level/Limit/Margin
    always stay attached to the correct frequency.
    """
    rows = sorted(rows, key=lambda row: float(row[0]))
    # 12 columns: No. + original 11 CE columns.
    group_row = start_row
    unit_row = start_row + 1
    sub_row = start_row + 2
    no_fill = PatternFill("solid", fgColor="D9D9D9")

    def cell(row, rel_col, value=None):
        c = ws.cell(row, start_col + rel_col - 1)
        if value is not None:
            c.value = value
        return c

    cell(group_row, 1, "No.")
    cell(group_row, 2, "Frequency")
    cell(group_row, 3, "Line")
    cell(group_row, 4, "Reading")
    cell(group_row, 6, "Factor")
    cell(group_row, 7, "Level")
    cell(group_row, 9, "Limit")
    cell(group_row, 11, "Margin")

    cell(unit_row, 2, "MHz")
    cell(unit_row, 4, "dB(µV)")
    cell(unit_row, 6, "dB")
    cell(unit_row, 7, "dB(µV)")
    cell(unit_row, 9, "dB(µV)")
    cell(unit_row, 11, "dB")

    for rel_col, value in ((4, "QP"), (5, "AV"), (7, "QP"), (8, "AV"),
                           (9, "QP"), (10, "AV"), (11, "QP"), (12, "AV")):
        cell(sub_row, rel_col, value)

    for rel_col in (1, 2, 3, 6):
        col = start_col + rel_col - 1
        ws.merge_cells(start_row=group_row, start_column=col,
                       end_row=sub_row, end_column=col)

    for rel_c1, rel_c2 in ((4, 5), (7, 8), (9, 10), (11, 12)):
        c1 = start_col + rel_c1 - 1
        c2 = start_col + rel_c2 - 1
        ws.merge_cells(start_row=group_row, start_column=c1,
                       end_row=group_row, end_column=c2)
        ws.merge_cells(start_row=unit_row, start_column=c1,
                       end_row=unit_row, end_column=c2)

    for r in range(group_row, sub_row + 1):
        for c in range(start_col, start_col + 12):
            x = ws.cell(r, c)
            x.font = Font(name="Calibri", size=9, bold=True)
            x.fill = no_fill if c == start_col else fill
            x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            x.border = border

    ws.row_dimensions[group_row].height = 22
    ws.row_dimensions[unit_row].height = 20
    ws.row_dimensions[sub_row].height = 20

    data_start = sub_row + 1
    for offset, row in enumerate(rows):
        r_idx = data_start + offset
        no_cell = ws.cell(r_idx, start_col, offset + 1)
        no_cell.font = Font(name="Calibri", size=9)
        no_cell.alignment = Alignment(horizontal="center", vertical="center")
        no_cell.border = border
        no_cell.fill = no_fill
        no_cell.number_format = "0"

        for rel_idx, value in enumerate(row, 1):
            excel_value = format_frequency(value) if rel_idx == 1 else value
            x = ws.cell(r_idx, start_col + rel_idx, excel_value)
            x.font = Font(name="Calibri", size=9)
            x.alignment = Alignment(horizontal="center", vertical="center")
            x.border = border
            if isinstance(value, (int, float)) and rel_idx != 1:
                x.number_format = "0.0"

    return data_start + len(rows)


def write_tel_ce_table(ws, start_row, start_col, rows, fill, border):
    """Write one Tel-line CE table with a gray No. column.

    Tel-line CE rows are sorted by Frequency from lowest to highest before writing.
    The complete measurement row moves together.
    """
    rows = sorted(rows, key=lambda row: float(row[0]))
    # 11 output columns: No. + original 10 Tel-line columns.
    group_row = start_row
    unit_row = start_row + 1
    sub_row = start_row + 2
    no_fill = PatternFill("solid", fgColor="D9D9D9")

    def cell(row, rel_col, value=None):
        c = ws.cell(row, start_col + rel_col - 1)
        if value is not None:
            c.value = value
        return c

    cell(group_row, 1, "No.")
    cell(group_row, 2, "Frequency")
    cell(group_row, 3, "Reading")
    cell(group_row, 5, "Factor")
    cell(group_row, 6, "Level")
    cell(group_row, 8, "Limit")
    cell(group_row, 10, "Margin")

    cell(unit_row, 2, "MHz")
    cell(unit_row, 3, "dB(µV)")
    cell(unit_row, 5, "dB")
    cell(unit_row, 6, "dB(µV)")
    cell(unit_row, 8, "dB(µV)")
    cell(unit_row, 10, "dB")

    for rel_col, value in ((3, "QP"), (4, "AV"), (6, "QP"), (7, "AV"),
                           (8, "QP"), (9, "AV"), (10, "QP"), (11, "AV")):
        cell(sub_row, rel_col, value)

    for rel_col in (1, 2, 5):
        col = start_col + rel_col - 1
        ws.merge_cells(start_row=group_row, start_column=col,
                       end_row=sub_row, end_column=col)

    for rel_c1, rel_c2 in ((3, 4), (6, 7), (8, 9), (10, 11)):
        c1 = start_col + rel_c1 - 1
        c2 = start_col + rel_c2 - 1
        ws.merge_cells(start_row=group_row, start_column=c1,
                       end_row=group_row, end_column=c2)
        ws.merge_cells(start_row=unit_row, start_column=c1,
                       end_row=unit_row, end_column=c2)

    for r in range(group_row, sub_row + 1):
        for c in range(start_col, start_col + 11):
            x = ws.cell(r, c)
            x.font = Font(name="Calibri", size=9, bold=True)
            x.fill = no_fill if c == start_col else fill
            x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            x.border = border

    ws.row_dimensions[group_row].height = 22
    ws.row_dimensions[unit_row].height = 20
    ws.row_dimensions[sub_row].height = 20

    data_start = sub_row + 1
    for offset, row in enumerate(rows):
        r_idx = data_start + offset
        no_cell = ws.cell(r_idx, start_col, offset + 1)
        no_cell.font = Font(name="Calibri", size=9)
        no_cell.alignment = Alignment(horizontal="center", vertical="center")
        no_cell.border = border
        no_cell.fill = no_fill
        no_cell.number_format = "0"

        for rel_idx, value in enumerate(row, 1):
            excel_value = format_frequency(value) if rel_idx == 1 else value
            x = ws.cell(r_idx, start_col + rel_idx, excel_value)
            x.font = Font(name="Calibri", size=9)
            x.alignment = Alignment(horizontal="center", vertical="center")
            x.border = border
            if isinstance(value, (int, float)) and rel_idx != 1:
                x.number_format = "0.0"

    return data_start + len(rows)


def ce_sort_key(ce_results):
    """
    CE ordering is grouped by STANDARD first, then measurement type.

    Example:
      EN AC Line
      EN Tel Line
      FCC AC Line
      FCC Tel Line
      VCCI AC Line
      VCCI Tel Line
      ICES AC Line
      ICES Tel Line
      KC AC Line
      KC Tel Line

    Tel ports are ordered IPMI0..IPMI20, bare IPMI, ETH0..ETH20, then bare Tel Line.
    """
    ce_type = str(ce_results.get("ce_type", "AC")).upper()
    pdf_name = str(ce_results.get("pdf_name", ""))
    standard = _detect_standard_from_name(pdf_name)

    standard_order = {
        "EN": 0,
        "FCC": 1,
        "VCCI": 2,
        "ICES": 3,
        "KC": 4,
        "OTHER": 99,
    }
    standard_rank = standard_order.get(standard, 99)

    # AC first, then DC, then Tel within each standard.
    if ce_type == "AC":
        return (
            standard_rank,
            0,
            _detect_voltage_rank(pdf_name),
            pdf_name.upper(),
        )

    if ce_type == "DC":
        dc_label = str(ce_results.get("dc_label", "DC Line"))
        m = re.fullmatch(r"DC Line (\d+)", dc_label, flags=re.IGNORECASE)
        dc_rank = int(m.group(1)) if m else 99
        return (
            standard_rank,
            1,
            dc_rank,
            pdf_name.upper(),
        )

    if ce_type == "TEL":
        sections = ce_results.get("sections", [])
        port = str(
            sections[0].get("line", "") if sections else ""
        ).upper()

        # Tel-line order: IPMI0..IPMI20, bare IPMI, ETH0..ETH20, bare Tel Line.
        m = re.fullmatch(r"IPMI(\d+)", port)
        if m:
            port_rank = (0, int(m.group(1)))
        elif port == "IPMI":
            port_rank = (1, 0)
        else:
            m = re.fullmatch(r"ETH(\d+)", port)
            if m:
                port_rank = (2, int(m.group(1)))
            elif port == "TEL LINE":
                port_rank = (3, 0)
            else:
                port_rank = (4, 999)

        return (
            standard_rank,
            2,
            _detect_voltage_rank(pdf_name),
            port_rank[0],
            port_rank[1],
            pdf_name.upper(),
        )

    return (standard_rank, 2, 99, pdf_name.upper())


def write_ce_sheet(wb, ce_results_list):
    # Keep CE sections grouped: all AC Line first, then all Tel Line.
    ce_results_list = sorted(ce_results_list, key=ce_sort_key)

    """
    Add CE datasets to one CE Results sheet.

    AC datasets:
      L1 and N stay side-by-side.

    Tel datasets:
      One ETH/IPMI graph and its 10-column final-data table are stacked
      vertically. Range, Line, Pass/Fail, and Remark are not included.
    """
    if "CE Results" in wb.sheetnames:
        del wb["CE Results"]

    ws = wb.create_sheet("CE Results")
    ws.sheet_view.zoomScale = 75
    ws.sheet_view.zoomScaleNormal = 75

    # Match the RE Results table formatting: light-blue headers and thin gray borders.
    fill = PatternFill("solid", fgColor="D9EAF7")
    no_fill = PatternFill("solid", fgColor="D9D9D9")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    left_start_col = 1        # A
    right_start_col = 27      # AA

    ac_widths = [5, 12, 8, 10, 10, 9, 10, 10, 10, 10, 10, 10]
    tel_widths = [5, 12, 10, 10, 9, 10, 10, 10, 10, 10, 10]

    for base_col in (left_start_col, right_start_col):
        for offset, width in enumerate(ac_widths):
            ws.column_dimensions[get_column_letter(base_col + offset)].width = width

    # Tel uses A:J, so these also receive appropriate widths.
    for offset, width in enumerate(tel_widths):
        ws.column_dimensions[get_column_letter(left_start_col + offset)].width = width

    for col in range(left_start_col + 12, right_start_col):
        ws.column_dimensions[get_column_letter(col)].width = 3

    current_row = 1

    for ce_results in ce_results_list:
        ce_type = ce_results.get("ce_type", "AC")

        ws.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=right_start_col + 10
        )
        c = ws.cell(
            current_row,
            1,
            ce_results["pdf_name"]
        )
        c.font = Font(name="Calibri", size=20, bold=True)
        c.alignment = Alignment(horizontal="left")
        current_row += 2

        # ------------------------------------------------------------
        # TEL LINE: one graph + one table per PDF/Excel pair
        # ------------------------------------------------------------
        if ce_type == "TEL":
            section = ce_results["sections"][0]
            port = section["line"]

            ws.merge_cells(
                start_row=current_row,
                start_column=left_start_col,
                end_row=current_row,
                end_column=left_start_col + 10
            )
            h = ws.cell(current_row, left_start_col, f"CE Tel Line — {port}")
            h.font = Font(name="Calibri", size=20, bold=True)

            current_row += 1
            graph_start_row = current_row

            graph_path = section["graph_path"]
            if graph_path.exists():
                img = XLImage(str(graph_path))
                max_width = 1050
                if img.width > max_width:
                    ratio = max_width / img.width
                    img.width = int(img.width * ratio)
                    img.height = int(img.height * ratio)

                ws.add_image(
                    img,
                    f"{get_column_letter(left_start_col)}{graph_start_row}"
                )
                graph_rows_used = max(24, int(img.height / 20) + 2)
            else:
                graph_rows_used = 2

            table_start_row = graph_start_row + graph_rows_used
            rows = section["rows"]

            if rows:
                end_row = write_tel_ce_table(
                    ws,
                    table_start_row,
                    left_start_col,
                    rows,
                    fill,
                    border,
                )
            else:
                ws.cell(
                    table_start_row,
                    left_start_col,
                    f"No {port} Tel-line data rows found in CE Excel."
                )
                end_row = table_start_row + 2

            current_row = end_row + 5
            continue

        # ------------------------------------------------------------
        # AC LINE or DC LINE: two graphs/tables side-by-side
        # AC uses L1 / N. DC uses + / -.
        # ------------------------------------------------------------
        sections = {
            section["line"]: section
            for section in ce_results["sections"]
        }
        if ce_type == "DC":
            pair = (("+", left_start_col), ("-", right_start_col))
            dc_label = ce_results.get("dc_label", "DC Line")
            heading_prefix = f"CE {dc_label}"
        else:
            pair = (("L1", left_start_col), ("N", right_start_col))
            heading_prefix = "CE"

        for line, start_col in pair:
            ws.merge_cells(
                start_row=current_row,
                start_column=start_col,
                end_row=current_row,
                end_column=start_col + 11
            )
            h = ws.cell(current_row, start_col, f"{heading_prefix} — {line}")
            h.font = Font(name="Calibri", size=20, bold=True)

        current_row += 1
        graph_start_row = current_row
        graph_rows_used = []

        for line, start_col in pair:
            section = sections.get(line)
            if section and section["graph_path"].exists():
                img = XLImage(str(section["graph_path"]))
                max_width = 1050
                if img.width > max_width:
                    ratio = max_width / img.width
                    img.width = int(img.width * ratio)
                    img.height = int(img.height * ratio)

                ws.add_image(
                    img,
                    f"{get_column_letter(start_col)}{graph_start_row}"
                )
                graph_rows_used.append(
                    max(24, int(img.height / 20) + 2)
                )
            else:
                graph_rows_used.append(2)

        table_start_row = graph_start_row + max(graph_rows_used)
        table_end_rows = []

        for line, start_col in pair:
            section = sections.get(line)
            rows = section["rows"] if section else []

            if rows:
                end_row = write_ce_table(
                    ws, table_start_row, start_col, rows, fill, border
                )
            else:
                ws.cell(
                    table_start_row,
                    start_col,
                    f"No {line} data rows found in CE Excel."
                )
                end_row = table_start_row + 2

            table_end_rows.append(end_row)

        current_row = max(table_end_rows) + 5

    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"


# ============================================================
# OATS conversion calibration data
# Copied from OATS_Data_Sheet_v6.xlsm.
# The Python math below reproduces the Excel formulas used by
# the KN_10m sheet for 30 MHz to 1 GHz.
# ============================================================
ACF_POINTS = [(30.0, 26.9), (40.0, 19.9), (50.0, 14.0), (60.0, 13.8), (70.0, 14.0), (80.0, 13.2), (90.0, 14.1), (100.0, 16.0),
 (110.0, 18.6), (120.0, 19.1), (130.0, 19.5), (140.0, 19.2), (150.0, 18.9), (160.0, 18.5), (170.0, 17.9), (180.0, 17.1),
 (190.0, 17.3), (200.0, 18.6), (210.0, 16.4), (220.0, 16.7), (230.0, 17.1), (240.0, 17.4), (250.0, 17.5), (260.0, 17.9),
 (270.0, 19.2), (280.0, 19.4), (290.0, 19.5), (300.0, 19.6), (310.0, 19.8), (320.0, 19.9), (330.0, 20.1), (340.0, 20.3),
 (350.0, 20.5), (360.0, 20.8), (370.0, 20.9), (380.0, 21.1), (390.0, 21.3), (400.0, 21.6), (410.0, 21.9), (420.0, 22.1),
 (430.0, 22.4), (440.0, 22.8), (450.0, 23.1), (460.0, 23.3), (470.0, 23.6), (480.0, 23.6), (490.0, 23.9), (500.0, 23.9),
 (510.0, 24.0), (520.0, 24.2), (530.0, 24.2), (540.0, 24.1), (550.0, 24.2), (560.0, 24.4), (570.0, 24.6), (580.0, 24.8),
 (590.0, 24.9), (600.0, 24.9), (610.0, 25.0), (620.0, 25.3), (630.0, 25.5), (640.0, 25.7), (650.0, 25.9), (660.0, 26.0),
 (670.0, 26.0), (680.0, 26.2), (690.0, 26.2), (700.0, 26.3), (710.0, 26.5), (720.0, 26.6), (730.0, 26.8), (740.0, 26.8),
 (750.0, 27.0), (760.0, 27.0), (770.0, 26.9), (780.0, 27.0), (790.0, 27.2), (800.0, 27.6), (810.0, 27.8), (820.0, 27.8),
 (830.0, 27.8), (840.0, 27.9), (850.0, 28.0), (860.0, 28.2), (870.0, 28.2), (880.0, 28.1), (890.0, 28.1), (900.0, 28.3),
 (910.0, 28.4), (920.0, 28.5), (930.0, 28.7), (940.0, 28.7), (950.0, 28.6), (960.0, 28.7), (970.0, 28.7), (980.0, 28.7),
 (990.0, 29.0), (1000.0, 29.0)]
LNA_POINTS = [(30.0, 27.32), (40.0, 27.31), (50.0, 27.29), (60.0, 27.28), (70.0, 27.27), (80.0, 27.26), (90.0, 27.26),
 (100.0, 27.25), (110.0, 27.23), (120.0, 27.22), (130.0, 27.21), (140.0, 27.21), (150.0, 27.21), (160.0, 27.21),
 (170.0, 27.21), (180.0, 27.2), (190.0, 27.19), (200.0, 27.18), (210.0, 27.18), (220.0, 27.18), (230.0, 27.18),
 (240.0, 27.19), (250.0, 27.19), (260.0, 27.18), (270.0, 27.17), (280.0, 27.17), (290.0, 27.16), (300.0, 27.17),
 (310.0, 27.18), (320.0, 27.19), (330.0, 27.2), (340.0, 27.21), (350.0, 27.2), (360.0, 27.2), (370.0, 27.2),
 (380.0, 27.2), (390.0, 27.21), (400.0, 27.22), (410.0, 27.23), (420.0, 27.24), (430.0, 27.24), (440.0, 27.23),
 (450.0, 27.22), (460.0, 27.22), (470.0, 27.21), (480.0, 27.21), (490.0, 27.2), (500.0, 27.19), (510.0, 27.17),
 (520.0, 27.15), (530.0, 27.12), (540.0, 27.08), (550.0, 27.05), (560.0, 27.01), (570.0, 26.98), (580.0, 26.95),
 (590.0, 26.91), (600.0, 26.87), (610.0, 26.82), (620.0, 26.77), (630.0, 26.71), (640.0, 26.66), (650.0, 26.62),
 (660.0, 26.58), (670.0, 26.55), (680.0, 26.53), (690.0, 26.5), (700.0, 26.45), (710.0, 26.4), (720.0, 26.35),
 (730.0, 26.31), (740.0, 26.29), (750.0, 26.28), (760.0, 26.28), (770.0, 26.28), (780.0, 26.27), (790.0, 26.24),
 (800.0, 26.2), (810.0, 26.15), (820.0, 26.12), (830.0, 26.11), (840.0, 26.12), (850.0, 26.16), (860.0, 26.2),
 (870.0, 26.21), (880.0, 26.21), (890.0, 26.2), (900.0, 26.19), (910.0, 26.18), (920.0, 26.18), (930.0, 26.19),
 (940.0, 26.22), (950.0, 26.26), (960.0, 26.32), (970.0, 26.36), (980.0, 26.38), (990.0, 26.4), (1000.0, 26.41)]
CABLE_POINTS = [(30.0, 0.6199999999999974), (35.0, 0.6599999999999966), (40.0, 0.7100000000000009), (45.0, 0.759999999999998),
 (50.0, 0.7999999999999972), (60.0, 0.8500000000000014), (70.0, 0.8999999999999986), (80.0, 0.9500000000000028),
 (90.0, 1.0), (100.0, 1.0399999999999991), (120.0, 1.0700000000000003), (140.0, 1.1000000000000014),
 (160.0, 1.1700000000000017), (180.0, 1.2899999999999991), (200.0, 1.4100000000000037), (250.0, 1.5600000000000023),
 (300.0, 1.75), (400.0, 2.0600000000000023), (500.0, 2.289999999999999), (600.0, 2.539999999999999),
 (700.0, 2.680000000000007), (800.0, 3.0500000000000007), (900.0, 3.1900000000000013), (1000.0, 3.4600000000000044)]
DELTA_H_POINTS = [(30.0, 14.0), (40.0, 13.6), (50.0, 13.3), (60.0, 13.0), (70.0, 12.7), (80.0, 12.4), (90.0, 12.1), (100.0, 11.7),
 (110.0, 11.45), (120.0, 11.2), (130.0, 11.0), (140.0, 10.8), (150.0, 10.6), (160.0, 10.4), (170.0, 10.3),
 (180.0, 10.2), (190.0, 10.2), (200.0, 10.2), (210.0, 10.18), (220.0, 10.16), (230.0, 10.14), (240.0, 10.12),
 (250.0, 10.1), (260.0, 9.98), (270.0, 9.86), (280.0, 9.74), (290.0, 9.62), (300.0, 9.5), (310.0, 9.44), (320.0, 9.38),
 (330.0, 9.32), (340.0, 9.26), (350.0, 9.2), (360.0, 9.14), (370.0, 9.08), (380.0, 9.02), (390.0, 8.96), (400.0, 8.9),
 (410.0, 8.95), (420.0, 9.0), (430.0, 9.05), (440.0, 9.1), (450.0, 9.15), (460.0, 9.2), (470.0, 9.25), (480.0, 9.3),
 (490.0, 9.35), (500.0, 9.4), (510.0, 9.42), (520.0, 9.44), (530.0, 9.46), (540.0, 9.48), (550.0, 9.5), (560.0, 9.52),
 (570.0, 9.54), (580.0, 9.56), (590.0, 9.58), (600.0, 9.6), (610.0, 9.62), (620.0, 9.64), (630.0, 9.66), (640.0, 9.68),
 (650.0, 9.7), (660.0, 9.72), (670.0, 9.74), (680.0, 9.76), (690.0, 9.78), (700.0, 9.8), (710.0, 9.75), (720.0, 9.7),
 (730.0, 9.65), (740.0, 9.6), (750.0, 9.55), (760.0, 9.5), (770.0, 9.45), (780.0, 9.4), (790.0, 9.35), (800.0, 9.3),
 (810.0, 9.34), (820.0, 9.38), (830.0, 9.42), (840.0, 9.46), (850.0, 9.5), (860.0, 9.54), (870.0, 9.58), (880.0, 9.62),
 (890.0, 9.66), (900.0, 9.7), (910.0, 9.7), (920.0, 9.7), (930.0, 9.7), (940.0, 9.7), (950.0, 9.7), (960.0, 9.7),
 (970.0, 9.7), (980.0, 9.7), (990.0, 9.7), (1000.0, 9.7)]
DELTA_V_POINTS = [(30.0, 8.5), (40.0, 8.399999999999999), (50.0, 8.3), (60.0, 8.1), (70.0, 7.9), (80.0, 7.700000000000001),
 (90.0, 7.3999999999999995), (100.0, 7.1000000000000005), (110.0, 6.75), (120.0, 6.4), (130.0, 5.95), (140.0, 5.5),
 (150.0, 4.9), (160.0, 4.3), (170.0, 3.65), (180.0, 3.0), (190.0, 3.8), (200.0, 4.6), (210.0, 5.12), (220.0, 5.64),
 (230.0, 6.16), (240.0, 6.68), (250.0, 7.2), (260.0, 7.56), (270.0, 7.92), (280.0, 8.28), (290.0, 8.64), (300.0, 9.0),
 (310.0, 9.09), (320.0, 9.18), (330.0, 9.27), (340.0, 9.36), (350.0, 9.45), (360.0, 9.54), (370.0, 9.63), (380.0, 9.72),
 (390.0, 9.81), (400.0, 9.9), (410.0, 9.88), (420.0, 9.86), (430.0, 9.84), (440.0, 9.82), (450.0, 9.8), (460.0, 9.78),
 (470.0, 9.76), (480.0, 9.74), (490.0, 9.72), (500.0, 9.7), (510.0, 9.49), (520.0, 9.28), (530.0, 9.07), (540.0, 8.86),
 (550.0, 8.65), (560.0, 8.44), (570.0, 8.23), (580.0, 8.02), (590.0, 7.81), (600.0, 7.600000000000001), (610.0, 7.66),
 (620.0, 7.72), (630.0, 7.78), (640.0, 7.84), (650.0, 7.9), (660.0, 7.96), (670.0, 8.02), (680.0, 8.08), (690.0, 8.14),
 (700.0, 8.2), (710.0, 8.23), (720.0, 8.26), (730.0, 8.29), (740.0, 8.32), (750.0, 8.35), (760.0, 8.38), (770.0, 8.41),
 (780.0, 8.44), (790.0, 8.47), (800.0, 8.5), (810.0, 8.64), (820.0, 8.78), (830.0, 8.92), (840.0, 9.06), (850.0, 9.2),
 (860.0, 9.34), (870.0, 9.48), (880.0, 9.62), (890.0, 9.76), (900.0, 9.9), (910.0, 9.9), (920.0, 9.9), (930.0, 9.9),
 (940.0, 9.9), (950.0, 9.9), (960.0, 9.9), (970.0, 9.9), (980.0, 9.9), (990.0, 9.9), (1000.0, 9.9)]


def excel_int_2(value):
    """Match Excel INT(value*100)/100 (round DOWN to 0.01)."""
    return math.floor(float(value) * 100.0 + 1e-10) / 100.0


def interpolate(points, frequency_mhz):
    """Linear interpolation matching the template's approximate VLOOKUP math."""
    f = float(frequency_mhz)
    if not points or f < points[0][0] or f > points[-1][0]:
        raise ValueError(f"Frequency {f:g} MHz is outside the OATS calibration range.")

    if f == points[-1][0]:
        return points[-1][1]

    for i in range(len(points) - 1):
        x0, y0 = points[i]
        x1, y1 = points[i + 1]
        if x0 <= f <= x1:
            if x1 == x0:
                return y0
            return y0 + (y1 - y0) * (f - x0) / (x1 - x0)

    raise ValueError(f"Could not interpolate {f:g} MHz.")


def oats_math(frequency, polarity, level_3m, antenna_height, azimuth, class_type="A"):
    """Convert one 3 m final-result row to the OATS/10 m table values."""
    f = float(frequency)
    pol = str(polarity).strip().upper()
    if pol not in ("H", "V"):
        raise ValueError(f"Unsupported polarization: {pol}")

    # These formulas mirror KN_10m columns N, O, P, Q, R, then F/G/E/H/I.
    antenna_factor = excel_int_2(interpolate(ACF_POINTS, f))
    lna_gain = -excel_int_2(interpolate(LNA_POINTS, f))
    cable_loss = excel_int_2(interpolate(CABLE_POINTS, f))
    correction = antenna_factor + lna_gain + cable_loss

    delta_points = DELTA_V_POINTS if pol == "V" else DELTA_H_POINTS
    delta_3m_10m = excel_int_2(interpolate(delta_points, f))

    # In the supplied OATS_Data_Sheet_v6 template, ACF Delta is 0.00 dB
    # throughout 30 MHz to 1 GHz.
    acf_delta = 0.0

    level_10m = float(level_3m) - delta_3m_10m - acf_delta
    reading_10m = level_10m - correction

    cls = str(class_type).strip().upper()
    if cls == "B":
        limit = 30.0 if f < 230.0 else 37.0
    else:
        limit = 40.0 if f < 230.0 else 47.0

    margin = limit - level_10m

    return {
        "frequency": f,
        "polarity": pol,
        "reading": reading_10m,
        "correction": correction,
        "level": level_10m,
        "limit": limit,
        "margin": margin,
        "antenna_height": max(100.0, float(antenna_height)),
        "azimuth": float(azimuth),
        # kept internally for troubleshooting / future expansion
        "antenna_factor": antenna_factor,
        "lna_gain": lna_gain,
        "cable_loss": cable_loss,
        "delta_3m_10m": delta_3m_10m,
        "acf_delta": acf_delta,
    }


def _oats_num(text):
    return float(str(text).replace(",", "").strip())


def parse_pdf_numeric_row(line):
    """Parse a 30M-1G EMI Final Results row from the selected PDF."""
    parts = line.split()
    if not parts or not parts[0].isdigit():
        return None

    # New chamber 30M-1G layout:
    # Rg, Frequency, Polarity, Raw QP, Correction, QP Level, QP Limit,
    # Margin, Height, Azimuth
    try:
        if len(parts) == 10 and parts[2].upper() in ("H", "V"):
            return {
                "frequency": _oats_num(parts[1]),
                "polarity": parts[2].upper(),
                "level_3m": _oats_num(parts[5]),
                "antenna_height": _oats_num(parts[8]),
                "azimuth": _oats_num(parts[9]),
            }
    except (ValueError, IndexError):
        return None
    return None


def extract_oats_source_rows(pdf_path):
    """Extract Frequency, Polarity, 3 m Level, Height and Azimuth from the PDF."""
    doc = fitz.open(pdf_path)
    rows = []
    all_text = []
    try:
        for page in doc:
            text = page.get_text("text", sort=True)
            all_text.append(text)
            for line in text.splitlines():
                row = parse_pdf_numeric_row(line)
                if row is not None:
                    rows.append(row)
    finally:
        doc.close()

    if not rows:
        raise ValueError(
            f"No 30M-1G EMI Final Results rows were found in:\n{Path(pdf_path).name}"
        )

    return rows


def build_oats_rows(pdf_path, class_type):
    source_rows = extract_oats_source_rows(pdf_path)
    converted = [
        oats_math(
            r["frequency"], r["polarity"], r["level_3m"],
            r["antenna_height"], r["azimuth"], class_type
        )
        for r in source_rows
    ]

    # OATS: lowest -> highest frequency. Each dictionary is one complete
    # measurement row, so all associated values move together.
    converted.sort(key=lambda row: float(row["frequency"]))
    return converted


def _write_oats_section(ws, rows, class_type, source_name, start_row):
    # FINAL OATS WRITE-TIME SORT: always enforce lowest -> highest Frequency
    # immediately before writing. Each dict is a complete measurement row, so
    # polarity, reading, correction, level, limit, margin, height, angle and
    # Pass status all stay attached to the correct frequency.
    rows = sorted(rows, key=lambda row: float(row["frequency"]))

    # Keep the OATS table visually consistent with the RE table.
    headers = [
        "No.",
        "Frequency\n(MHz)",
        "Polarization",
        "Reading\n(dBµV/m)",
        "Corr.\n(dB)",
        "Level\n(dBµV/m)",
        "Limit\n(dBµV/m)",
        "Margin\n(dB)",
        "Height\n(cm)",
        "Angle\n(deg)",
        "Pass",
    ]

    title_row = start_row
    header_row = start_row + 2
    data_start = start_row + 3

    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=11)
    ws.cell(title_row, 1, f"{source_name} — Class {class_type}")
    ws.cell(title_row, 1).font = Font(name="Calibri", size=9, bold=True)
    ws.cell(title_row, 1).alignment = Alignment(horizontal="center")

    # Same blue header style used by RE.
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    no_fill = PatternFill("solid", fgColor="D9D9D9")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, h in enumerate(headers, 1):
        cell = ws.cell(header_row, c, h)
        cell.font = Font(name="Calibri", size=9, bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    # No. uses gray instead of the blue measurement-header fill.
    ws.cell(header_row, 1).fill = PatternFill("solid", fgColor="D9D9D9")

    for offset, r in enumerate(rows):
        r_idx = data_start + offset
        values = [
            offset + 1,
            r["frequency"], r["polarity"], r["reading"], r["correction"],
            r["level"], r["limit"], r["margin"], r["antenna_height"], r["azimuth"]
        ]
        for c_idx, value in enumerate(values, 1):
            cell = ws.cell(r_idx, c_idx, value)
            cell.font = Font(name="Calibri", size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            if c_idx == 1:
                cell.fill = PatternFill("solid", fgColor="D9D9D9")

        # Match RE display formatting: No. integer, frequency 3 decimals,
        # and all remaining numeric measurements (including Height/Angle) 1 decimal.
        ws.cell(r_idx, 1).number_format = "0"
        ws.cell(r_idx, 2).number_format = "0.000"
        for c_idx in range(4, 11):
            ws.cell(r_idx, c_idx).number_format = "0.0"

        # OATS pass-status column based on margin.
        margin = float(r["margin"])
        if margin < 0:
            status_text = "Failed"
            status_fill = PatternFill("solid", fgColor="FF0000")
        elif margin <= 3:
            status_text = "No Margin"
            status_fill = PatternFill("solid", fgColor="FFFF00")
        else:
            status_text = "Passed"
            status_fill = PatternFill("solid", fgColor="00B050")

        status_cell = ws.cell(r_idx, 11, status_text)
        status_cell.font = Font(name="Calibri", size=9, bold=True)
        status_cell.alignment = Alignment(horizontal="center", vertical="center")
        status_cell.border = border
        status_cell.fill = status_fill

    ws.row_dimensions[header_row].height = 34
    return data_start + len(rows) + 2


def write_oats_sheet(wb, all_results, class_type):
    """Add OATS conversion as its own tab in the same combined workbook."""
    ws = wb.create_sheet("Recalculation")
    ws.sheet_view.zoomScale = 75
    ws.sheet_view.zoomScaleNormal = 75
    ws.merge_cells("A1:K1")
    ws["A1"] = f"Recalculation Converted Data — Class {class_type}"
    ws["A1"].font = Font(name="Calibri", size=20, bold=True)
    ws["A1"].alignment = Alignment(horizontal="left")
    widths = [8, 15, 14, 18, 12, 18, 18, 14, 16, 14, 14]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    next_row = 3
    for source_name, rows in all_results:
        next_row = _write_oats_section(ws, rows, class_type, source_name, next_row)
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"


def save_combined_excel(results, out_xlsx, ce_results_list=None, oats_results=None, oats_class="A"):
    # Keep RE sections grouped in a consistent standard/range order.
    results = sorted(results, key=re_sort_key)

    """
    Create ONE Excel worksheet containing every PDF result vertically:

        PDF 1 filename
        PDF 1 cleaned graph
        PDF 1 data table

        PDF 2 filename
        PDF 2 cleaned graph
        PDF 2 data table

        ...

    This makes it possible to find everything by scrolling only up/down.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "RE Results"

    # Open workbook at 75% zoom.
    ws.sheet_view.zoomScale = 75
    ws.sheet_view.zoomScaleNormal = 75

    fill = PatternFill("solid", fgColor="D9EAF7")
    no_fill = PatternFill("solid", fgColor="D9D9D9")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Compact widths sized to keep the tables close to graph width.
    kc_widths = [5, 10, 8, 8, 8, 7, 8, 8, 8, 8, 8, 8, 7, 7]
    en_widths = [5, 10, 8, 9, 7, 9, 9, 9, 7, 7]

    # Use the KC widths as the overall sheet widths because it has more columns.
    for i, width in enumerate(kc_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    current_row = 1

    for index, result in enumerate(results, 1):
        source_name = result["source_name"]
        data_type = result["data_type"]
        rows = result["rows"]
        graph_path = result["graph_path"]
        math_notes = result.get("math_notes") or [""] * len(rows)

        # FINAL RE WRITE-TIME SORT: always enforce lowest -> highest Frequency
        # immediately before writing the table. Keep the complete measurement row
        # and its software-math note paired together, then regenerate No. 1,2,3...
        paired = list(zip(rows, math_notes))
        paired.sort(key=lambda item: float(item[0][1]))
        rows = [list(item[0]) for item in paired]
        math_notes = [item[1] for item in paired]
        for new_no, row in enumerate(rows, 1):
            row[0] = new_no

        headers = EN_HEADERS if data_type == "EN_QP" else KC_HEADERS

        # Place the Software Math Check farther to the right for 30M-1G / QP
        # new-chamber tables. Column P = 16.
        # Higher-frequency AV/PK tables keep the existing one-spacer-column layout.
        if data_type == "EN_QP":
            note_col = 16  # Column P
        else:
            note_col = len(headers) + 2

        # ------------------------------------------------------------
        # PDF / section title
        # ------------------------------------------------------------
        ws.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=len(KC_HEADERS)
        )

        title_cell = ws.cell(current_row, 1, source_name)
        title_cell.font = Font(name="Calibri", size=20, bold=True)
        title_cell.alignment = Alignment(horizontal="left", vertical="center")

        current_row += 2

        # ------------------------------------------------------------
        # Cleaned graph
        # ------------------------------------------------------------
        graph_start_row = current_row

        if graph_path.exists():
            img = XLImage(str(graph_path))

            # Keep graph at the same practical size used previously.
            max_width = 1050
            if img.width > max_width:
                ratio = max_width / img.width
                img.width = int(img.width * ratio)
                img.height = int(img.height * ratio)

            ws.add_image(img, f"A{graph_start_row}")

            # Approximate how many Excel rows the image occupies.
            image_rows = max(24, int(img.height / 20) + 2)
            current_row += image_rows
        else:
            current_row += 2

        # ------------------------------------------------------------
        # Data table directly below its graph
        # ------------------------------------------------------------
        table_start = current_row

        for col, header in enumerate(headers, 1):
            cell = ws.cell(table_start, col, header)
            cell.font = Font(name="Calibri", size=9, bold=True)
            cell.fill = fill
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        # Shade the No. header gray so it stands out from the measurement columns.
        ws.cell(table_start, 1).fill = no_fill

        ws.row_dimensions[table_start].height = 36

        # New-chamber math audit column. It stays blank for rows that required
        # no correction, and records exactly what the software changed otherwise.
        note_header = ws.cell(table_start, note_col, "Software Math Check")
        note_header.font = Font(name="Calibri", size=9, bold=True)
        note_header.fill = PatternFill("solid", fgColor="FFF2CC")
        note_header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        note_header.border = border
        ws.column_dimensions[get_column_letter(note_col)].width = 42

        for row_offset, row in enumerate(rows):
            r_idx = table_start + 1 + row_offset
            for c_idx, value in enumerate(row, 1):

                # Frequency must visibly match PDF precision: exactly 3 decimals.
                excel_value = (
                    format_frequency(value)
                    if c_idx == 2
                    else value
                )

                cell = ws.cell(r_idx, c_idx, excel_value)
                cell.font = Font(name="Calibri", size=9)
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                cell.border = border
                if c_idx == 1:
                    cell.fill = no_fill

                if isinstance(value, (int, float)):
                    if c_idx == 1:
                        cell.number_format = "0"
                    elif c_idx == 2:
                        # Already written as exact 3-decimal text.
                        pass
                    else:
                        cell.number_format = "0.0"

            note_text = math_notes[row_offset] if row_offset < len(math_notes) else ""
            note_cell = ws.cell(r_idx, note_col, note_text)
            note_cell.font = Font(name="Calibri", size=9, bold=bool(note_text))
            note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            note_cell.border = border
            if note_text:
                note_cell.fill = PatternFill("solid", fgColor="FFF2CC")

        # Move below the data table.
        current_row = table_start + len(rows) + 3

        # Add a little blank space between PDF sections.
        current_row += 2

    # Start at the top-left whenever the workbook opens.
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"

    # Add CE as a second tab when CE source files were selected.
    if ce_results_list:
        write_ce_sheet(wb, ce_results_list)

    # Add OATS as a third tab when Recalculation PDFs were selected.
    if oats_results:
        write_oats_sheet(wb, oats_results, oats_class)

    wb.save(out_xlsx)


def _pair_ce_files(ce_pdf_paths, ce_excel_paths):
    """Pair CE PDFs and Excel files ONLY when their filename stems match exactly.

    Selection order does not matter. If any PDF or Excel does not have an
    exact same-name counterpart, stop and report the unmatched files.
    """
    pdfs = [Path(p) for p in ce_pdf_paths]
    excels = [Path(p) for p in ce_excel_paths]

    if not pdfs:
        return []

    # Match case-insensitively by filename without extension.
    pdf_by_stem = {}
    excel_by_stem = {}

    for p in pdfs:
        pdf_by_stem.setdefault(p.stem.lower(), []).append(p)
    for x in excels:
        excel_by_stem.setdefault(x.stem.lower(), []).append(x)

    # Duplicate stems are ambiguous, so do not guess.
    duplicate_pdfs = [items for items in pdf_by_stem.values() if len(items) > 1]
    duplicate_excels = [items for items in excel_by_stem.values() if len(items) > 1]
    if duplicate_pdfs or duplicate_excels:
        details = []
        if duplicate_pdfs:
            details.append("Duplicate CE PDF names:\n" + "\n".join(
                f"  - {p.name}" for group in duplicate_pdfs for p in group
            ))
        if duplicate_excels:
            details.append("Duplicate CE Excel names:\n" + "\n".join(
                f"  - {x.name}" for group in duplicate_excels for x in group
            ))
        raise ValueError(
            "CE files could not be paired safely.\n\n" + "\n\n".join(details)
        )

    pdf_stems = set(pdf_by_stem)
    excel_stems = set(excel_by_stem)

    unmatched_pdfs = [pdf_by_stem[k][0] for k in sorted(pdf_stems - excel_stems)]
    unmatched_excels = [excel_by_stem[k][0] for k in sorted(excel_stems - pdf_stems)]

    if unmatched_pdfs or unmatched_excels:
        details = []
        if unmatched_pdfs:
            details.append(
                "PDF(s) without matching Excel file:\n"
                + "\n".join(f"  - {p.name}" for p in unmatched_pdfs)
            )
        if unmatched_excels:
            details.append(
                "Excel file(s) without matching PDF:\n"
                + "\n".join(f"  - {x.name}" for x in unmatched_excels)
            )

        raise ValueError(
            "CE PDF and Excel filenames must match exactly (except extension).\n\n"
            + "\n\n".join(details)
            + "\n\nExample: Test_220V.pdf must pair with Test_220V.xlsx"
        )

    # All names match. Preserve the user's PDF selection order.
    return [(pdf, excel_by_stem[pdf.stem.lower()][0]) for pdf in pdfs]

def extract_multiple_pdfs(pdf_paths, ce_pdf_paths=None, ce_excel_paths=None, oats_pdf_paths=None, oats_class="A", progress_callback=None):
    """
    Batch mode:
      - Multiple RE PDFs supported.
      - Multiple CE PDF + Excel pairs supported.
      - Each CE dataset is stacked vertically in the CE Results tab.
      - L1 and N remain side-by-side for each CE dataset.
    """
    pdf_paths = [Path(p) for p in pdf_paths]
    ce_pdf_paths = [Path(p) for p in (ce_pdf_paths or [])]
    ce_excel_paths = [Path(p) for p in (ce_excel_paths or [])]
    oats_pdf_paths = [Path(p) for p in (oats_pdf_paths or [])]
    oats_class = str(oats_class).strip().upper()
    if oats_class not in ("A", "B"):
        raise ValueError("OATS Class must be A or B.")

    if bool(ce_pdf_paths) != bool(ce_excel_paths):
        raise ValueError("For CE, select BOTH CE PDF file(s) and CE Excel file(s).")

    ce_pairs = _pair_ce_files(ce_pdf_paths, ce_excel_paths)

    if not pdf_paths and not ce_pairs and not oats_pdf_paths:
        raise ValueError("Select RE PDF(s), CE PDF + Excel pair(s), or Recalculation PDF(s).")

    base_path = pdf_paths[0] if pdf_paths else (ce_pairs[0][0] if ce_pairs else oats_pdf_paths[0])
    out_dir = base_path.parent / "EMC_Batch_Extracted"
    out_dir.mkdir(exist_ok=True)

    results = []
    total_rows = 0

    # Progress is based on each RE PDF, each CE PDF/Excel pair, and the final Excel save.
    total_steps = len(pdf_paths) + len(ce_pairs) + len(oats_pdf_paths) + 1
    completed_steps = 0

    def report_progress(message):
        if progress_callback is not None:
            percent = int(round((completed_steps / max(1, total_steps)) * 100))
            progress_callback(percent, message)

    report_progress("Preparing extraction...")

    with tempfile.TemporaryDirectory(prefix="emc_graphs_") as temp_dir:
        temp_dir = Path(temp_dir)
        graph_count = 0

        for pdf_index, pdf_path in enumerate(pdf_paths, 1):
            report_progress(
                f"Extracting RE PDF {pdf_index}/{len(pdf_paths)}: {pdf_path.name}"
            )
            doc = fitz.open(pdf_path)

            for page_no, page in enumerate(doc, 1):
                data_type, rows = extract_rows(page)
                page_text = page.get_text("text")

                # New-chamber RE reports can occasionally contain incorrect
                # calculated Level/Margin values. Verify the math before the
                # rows are written to Excel. Original-chamber reports are not changed.
                math_corrections = 0
                math_notes = []
                if rows and data_type and _is_new_re_chamber(page):
                    rows, math_corrections, math_notes = check_and_fix_new_re_math(data_type, rows)
                elif rows:
                    math_notes = [""] * len(rows)

                if not rows and "Final Result" not in page_text:
                    continue

                if rows and data_type:
                    # RE: sort lowest -> highest frequency while keeping every
                    # measurement value (and any math-correction note) with its row.
                    paired = list(zip(rows, math_notes))
                    paired.sort(key=lambda item: float(item[0][1]))
                    rows = [item[0] for item in paired]
                    math_notes = [item[1] for item in paired]

                    graph_count += 1
                    graph_path = temp_dir / f"graph_{graph_count}.png"

                    # New-chamber 1G+ reports place the graph on page 1 and the
                    # EMI Final Results table on page 2. Use the nearest earlier
                    # Radiated Emission Test Result page as the graph source.
                    graph_page = page
                    if "Radiated Emission Test Result" not in page_text:
                        for prior_index in range(page_no - 2, -1, -1):
                            prior_page = doc[prior_index]
                            prior_text = prior_page.get_text("text", sort=True)
                            if "Radiated Emission Test Result" in prior_text:
                                graph_page = prior_page
                                break

                    save_graph(graph_page, graph_path)

                    results.append({
                        "source_name": pdf_path.name,
                        "data_type": data_type,
                        "rows": rows,
                        "graph_path": graph_path,
                        "math_corrections": math_corrections,
                        "math_notes": math_notes,
                    })
                    total_rows += len(rows)

            doc.close()
            completed_steps += 1
            report_progress(f"Finished RE PDF {pdf_index}/{len(pdf_paths)}")

        # If RE PDFs were selected, do not silently create a blank RE sheet.
        if pdf_paths and not results:
            raise ValueError(
                "RE PDF(s) were selected, but no supported RE Final Result rows "
                "could be extracted. Check that the PDFs contain EN/KC Final Result tables."
            )

        ce_results_list = []

        for ce_index, (ce_pdf_path, ce_excel_path) in enumerate(ce_pairs, 1):
            report_progress(
                f"Extracting CE file {ce_index}/{len(ce_pairs)}: {ce_pdf_path.name}"
            )
            ce_result = build_ce_results(
                ce_pdf_path,
                ce_excel_path,
                temp_dir
            )
            ce_results_list.append(ce_result)
            total_rows += sum(
                len(section["rows"]) for section in ce_result["sections"]
            )
            completed_steps += 1
            report_progress(f"Finished CE file {ce_index}/{len(ce_pairs)}")

        oats_results = []
        for oats_index, oats_pdf_path in enumerate(oats_pdf_paths, 1):
            report_progress(
                f"Extracting Recalculation PDF {oats_index}/{len(oats_pdf_paths)}: {oats_pdf_path.name}"
            )
            oats_rows = build_oats_rows(oats_pdf_path, oats_class)
            oats_results.append((oats_pdf_path.name, oats_rows))
            total_rows += len(oats_rows)
            completed_steps += 1
            report_progress(f"Finished Recalculation PDF {oats_index}/{len(oats_pdf_paths)}")

        if not results and not ce_results_list and not oats_results:
            raise ValueError("No supported EMC result data were detected.")

        combined_xlsx = out_dir / "EMC_All_Results_RE_CE_OATS.xlsx"
        report_progress("Creating combined Excel file...")
        save_combined_excel(
            results,
            combined_xlsx,
            ce_results_list=ce_results_list,
            oats_results=oats_results,
            oats_class=oats_class
        )
        completed_steps += 1
        report_progress("Extraction complete")

    return out_dir, combined_xlsx, total_rows


def extract_pdf(pdf_path):
    # Backward-compatible single-file wrapper.
    out_dir, xlsx, total_rows = extract_multiple_pdfs([pdf_path])
    return out_dir, [xlsx], total_rows

def open_folder(path):
    path = str(path)
    if sys.platform.startswith("win"):
        import os
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


HELP_PDF_NAME = "KillerWhale Extractor 3ThouWow.pdf"


def find_help_pdf():
    """Find the Help PDF next to the app/script, including PyInstaller builds."""
    candidates = []

    # PyInstaller one-file/one-folder bundled resource location.
    if getattr(sys, "frozen", False):
        bundle_dir = getattr(sys, "_MEIPASS", None)
        if bundle_dir:
            candidates.append(Path(bundle_dir) / HELP_PDF_NAME)

        # Also allow the PDF to sit next to the generated .exe.
        candidates.append(Path(sys.executable).resolve().parent / HELP_PDF_NAME)

    # Normal Python script location.
    try:
        candidates.append(Path(__file__).resolve().parent / HELP_PDF_NAME)
    except NameError:
        pass

    # Last fallback: current working directory.
    candidates.append(Path.cwd() / HELP_PDF_NAME)

    for candidate in candidates:
        if candidate.exists():
            return candidate

    return None


def open_file_default(path):
    """Open a file with the operating system's default application."""
    path = str(path)
    if sys.platform.startswith("win"):
        import os
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])
