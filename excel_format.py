#!/usr/bin/env python3
"""
Format an Excel workbook professionally and apply static status coloring
(GREEN/RED) directly to cells (no conditional formatting rules).

Works best for "report-like" sheets that contain:
- a top title line
- section break lines
- table header rows
- data rows (including status-like values such as OK/NOK/Yes/No/Pass/Fail)

USAGE:
  python format_excel.py input.xlsx output.xlsx
"""

import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ----------------------------
# Palette (edit if you want)
# ----------------------------
DARK_FILL    = PatternFill("solid", fgColor="1F4E78")  # dark blue (title)
SECTION_FILL = PatternFill("solid", fgColor="4472C4")  # medium blue (section)
HEADER_FILL  = PatternFill("solid", fgColor="D9E1F2")  # light blue (table header)
ALT_FILL     = PatternFill("solid", fgColor="F2F2F2")  # zebra (very light gray)

GOOD_FILL    = PatternFill("solid", fgColor="C6EFCE")  # green for OK/YES...
BAD_FILL     = PatternFill("solid", fgColor="FFC7CE")  # red for NOK/NO...

TITLE_FONT   = Font(bold=True, color="FFFFFF", size=14)
SECTION_FONT = Font(bold=True, color="FFFFFF", size=12)
HEADER_FONT  = Font(bold=True, color="1F2937")
NORMAL_FONT  = Font(color="111827")

CENTER_WRAP  = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP    = Alignment(horizontal="left",   vertical="center", wrap_text=True)

THIN = Side(style="thin",   color="9CA3AF")
MED  = Side(style="medium", color="6B7280")


GOOD_VALUES = {"OK", "YES", "TRUE", "PASS", "GRANTED", "ALLOW"}
BAD_VALUES  = {"NOK", "NO", "FALSE", "FAIL", "DENIED", "BLOCK"}


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def norm_status(v):
    """Return normalized status token or None."""
    if not isinstance(v, str):
        return None
    return v.strip().upper()


def used_max_col(ws):
    """Find the last used column index (by non-empty cells)."""
    max_row = ws.max_row
    max_col = ws.max_column
    used = set()
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if not is_blank(cell.value):
                used.add(cell.column)
    return max(used) if used else 0


def delete_blank_rows(ws):
    """Remove rows that are entirely blank (keeps row order)."""
    rows = list(ws.iter_rows(values_only=True))
    ws.delete_rows(1, ws.max_row)
    for r in rows:
        if any(not is_blank(v) for v in r):
            ws.append(list(r))


def clear_merges(ws):
    """Unmerge everything so we can re-merge safely."""
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))


def clear_conditional_formatting(ws):
    """Remove all conditional formatting rules."""
    try:
        ws.conditional_formatting._cf_rules.clear()
    except Exception:
        pass


def detect_section_rows(ws, max_row, max_col):
    """
    Section row heuristic:
    - Only first column has text, the rest of the row is blank.
    """
    section = set()
    for r in range(2, max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        nonblank = [v for v in vals if not is_blank(v)]
        if len(nonblank) == 1 and not is_blank(vals[0]):
            section.add(r)
    return section


def detect_header_rows(ws, max_row, max_col, section_rows):
    """
    Table header heuristic:
    - only consider row 1 or the first row after a detected section row
    - require several label-like strings
    - reject rows that look like actual data (numbers, emails, filenames, dates)
    """
    def looks_like_label(v):
        if not isinstance(v, str):
            return False

        s = v.strip()
        if not s:
            return False

        # Data-ish values that should not make a row look like a header.
        if "@" in s:
            return False
        if any(ch.isdigit() for ch in s):
            return False
        if "." in s or "/" in s or "\\" in s:
            return False

        return True

    headers = []
    status_tokens = GOOD_VALUES | BAD_VALUES | {"HIGH", "MEDIUM", "LOW"}

    candidate_rows = {1}
    for section_row in section_rows:
        if section_row + 1 <= max_row:
            candidate_rows.add(section_row + 1)

    for r in range(1, max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        nonblank = [v for v in vals if not is_blank(v)]
        if len(nonblank) < 2:
            continue

        str_count = sum(isinstance(v, str) and not is_blank(v) for v in vals)
        label_count = sum(looks_like_label(v) for v in vals)
        number_count = sum(isinstance(v, (int, float)) for v in vals)
        status_count = sum(
            isinstance(v, str) and v.strip().upper() in status_tokens
            for v in vals
        )

        # Rows with an empty first cell and labels in later columns are
        # common in permission matrices: treat them as stacked table headers.
        blank_first_col_header = (
            is_blank(vals[0])
            and str_count == len(nonblank)
            and number_count == 0
            and status_count == 0
            and len(nonblank) >= 2
        )

        # Also allow the original "top of section" style header rows.
        top_of_section_header = (
            r in candidate_rows
            and str_count >= 3
            and label_count >= 3
            and label_count >= max(3, int(len(nonblank) * 0.6))
            and number_count == 0
            and status_count <= 1
        )

        if blank_first_col_header or top_of_section_header:
            headers.append(r)
    return headers


def apply_title_row(ws, max_col):
    """Create/format a merged title row in row 1, only when A1 is the sole non-empty cell.
    Returns True if row 1 was treated as a header instead of a title banner."""
    last_col = get_column_letter(max_col)

    # Only treat row 1 as a title banner when no other cell in the row has content
    row1_other = [ws.cell(1, c).value for c in range(2, max_col + 1)]
    if any(not is_blank(v) for v in row1_other):
        apply_header_row(ws, 1, max_col)
        return True

    title = ws.cell(1, 1).value or ws.title

    # wipe row 1 values except A1
    ws.cell(1, 1).value = str(title)
    for c in range(2, max_col + 1):
        ws.cell(1, c).value = None

    ws.merge_cells(f"A1:{last_col}1")

    cell = ws["A1"]
    cell.fill = DARK_FILL
    cell.font = TITLE_FONT
    cell.alignment = CENTER_WRAP
    ws.row_dimensions[1].height = 24


def apply_section_row(ws, r, max_col):
    """Merge across + style section row."""
    last_col = get_column_letter(max_col)
    # wipe row values except col1
    for c in range(2, max_col + 1):
        ws.cell(r, c).value = None

    ws.merge_cells(f"A{r}:{last_col}{r}")
    cell = ws.cell(r, 1)
    cell.fill = SECTION_FILL
    cell.font = SECTION_FONT
    cell.alignment = LEFT_WRAP
    ws.row_dimensions[r].height = 20


def apply_header_row(ws, r, max_col):
    """Style a table header row."""
    for c in range(1, max_col + 1):
        cell = ws.cell(r, c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_WRAP


def adjust_header_row_heights(ws, max_col, header_rows):
    """Estimate wrapped header height from final column widths."""
    for r in header_rows:
        max_lines = 1
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            if is_blank(cell.value):
                continue

            width = ws.column_dimensions[get_column_letter(c)].width or 10
            usable_width = max(1, int(width - 2))
            text_len = len(str(cell.value))
            estimated_lines = max(1, (text_len + usable_width - 1) // usable_width)
            max_lines = max(max_lines, estimated_lines)

        ws.row_dimensions[r].height = max(18, min(15 * max_lines, 90))


def apply_default_fonts_alignment(ws, max_row, max_col):
    """Set base styling: left align first column, center others."""
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.font = NORMAL_FONT
            cell.alignment = LEFT_WRAP if c == 1 else CENTER_WRAP


def apply_borders(ws, max_row, max_col):
    """Apply thin grid and thicker outside border for the used range."""
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.border = Border(
                left=MED if c == 1 else THIN,
                right=MED if c == max_col else THIN,
                top=MED if r == 1 else THIN,
                bottom=MED if r == max_row else THIN,
            )


def apply_zebra_fill(ws, max_row, max_col, section_rows, header_rows):
    """
    Add alternating row fill for data rows, but DO NOT overwrite
    any existing fill (keeps other background colors).
    """
    header_set = set(header_rows)
    last_header = 1
    for r in range(2, max_row + 1):
        if r in header_set:
            last_header = r
            continue
        if r in section_rows:
            continue

        # apply zebra to every other data row, only if cell has no fill
        if ((r - last_header) % 2) == 0:
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                # keep existing fill
                if cell.fill is None or cell.fill.patternType is None:
                    cell.fill = ALT_FILL


def autofit_columns(ws, max_row, max_col):
    """Basic autosize with sane caps."""
    for c in range(1, max_col + 1):
        max_len = 0
        for r in range(1, max_row + 1):
            v = ws.cell(r, c).value
            if is_blank(v):
                continue
            max_len = max(max_len, len(str(v)))

        width = min(max(10, max_len + 2), 45)
        if c == 1:
            width = min(max(18, max_len + 2), 55)
        ws.column_dimensions[get_column_letter(c)].width = width


def freeze_and_filter(ws, max_row, max_col, header_rows):
    """Freeze panes and enable filter on first detected header row."""
    ws.freeze_panes = "B2"

    if header_rows:
        hr = header_rows[0]
        last_col = get_column_letter(max_col)
        ws.auto_filter.ref = f"A{hr}:{last_col}{max_row}"


def apply_static_status_colors(ws, max_row, max_col, good_fill, bad_fill):
    """
    Apply static GREEN/RED fills based on cell value.
    IMPORTANT: Only overwrite fills for those statuses.
    Everything else keeps its existing background color.
    """
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            token = norm_status(cell.value)
            if token in GOOD_VALUES:
                cell.fill = good_fill
            elif token in BAD_VALUES:
                cell.fill = bad_fill
            # else: keep existing fill


def format_workbook(input_path, output_path):
    wb = load_workbook(input_path)

    for ws in wb.worksheets:
        # 1) Cleanup
        clear_merges(ws)
        clear_conditional_formatting(ws)
        delete_blank_rows(ws)

        max_row = ws.max_row
        max_col = used_max_col(ws)
        if max_row < 1 or max_col == 0:
            continue

        # 2) Base formatting
        row1_is_header = apply_title_row(ws, max_col)

        # Detect structure after cleanup
        max_row = ws.max_row  # in case delete_blank_rows changed it
        section_rows = detect_section_rows(ws, max_row, max_col)
        header_rows = detect_header_rows(ws, max_row, max_col, section_rows)

        # If row 1 was treated as a header (multi-column), prepend it
        if row1_is_header and 1 not in header_rows:
            header_rows = [1] + header_rows

        apply_default_fonts_alignment(ws, max_row, max_col)

        # Apply section/header styling (these set their own fills)
        for r in sorted(section_rows):
            apply_section_row(ws, r, max_col)
        for r in header_rows:
            apply_header_row(ws, r, max_col)

        # Borders, widths, zebra, freeze/filter
        apply_borders(ws, max_row, max_col)
        apply_zebra_fill(ws, max_row, max_col, section_rows, header_rows)
        autofit_columns(ws, max_row, max_col)
        adjust_header_row_heights(ws, max_col, header_rows)
        freeze_and_filter(ws, max_row, max_col, header_rows)

        # 3) FINAL STEP: Static status coloring (green/red only)
        # Keeps all other colors intact.
        apply_static_status_colors(ws, max_row, max_col, GOOD_FILL, BAD_FILL)

    wb.save(output_path)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python format_excel.py input.xlsx output.xlsx")
        sys.exit(1)

    import os
    input_path = os.path.realpath(sys.argv[1])
    output_path = os.path.realpath(sys.argv[2])

    if input_path == output_path:
        answer = input(f"Input and output are the same file ({sys.argv[1]}). Overwrite? [y/N]: ").strip().lower()
        if answer != "y":
            print("Aborted.")
            sys.exit(0)
    elif os.path.exists(output_path):
        answer = input(f"Output file already exists: {sys.argv[2]}. Overwrite? [y/N]: ").strip().lower()
        if answer != "y":
            print("Aborted.")
            sys.exit(0)

    format_workbook(sys.argv[1], sys.argv[2])
    print(f"Saved: {sys.argv[2]}")
