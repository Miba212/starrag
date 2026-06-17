# -*- coding: utf-8 -*-
"""
pdf_to_excel.py
Converts 'Checkliste Große Wartung.pdf' to 'Checkliste Große Wartung.xlsx',
preserving the visual layout of each page as closely as possible.

Usage:
    python pdf_to_excel.py
"""

import io
import os

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------------------------------------------------------------------------
# Grid constants (per A4 page: 595 × 842 pts)
# ---------------------------------------------------------------------------
GRID_COLS = 100          # number of Excel columns per page width
GRID_ROWS = 150          # number of Excel rows per page height

# Excel cell sizing (must result in roughly A4 dimensions when printed)
# A4 = 210 mm wide  → each column = 210/100 = 2.1 mm  ≈ 0.79 char units
# A4 = 297 mm tall  → each row   = 297/150 = 1.98 mm ≈ 5.61 pts
COL_WIDTH = 1.9          # character units per column
ROW_HEIGHT = 5.65        # points per row

# Word-grouping tolerances (in PDF points)
LINE_TOL = 3.0           # words whose 'top' values differ by ≤ LINE_TOL are on the same line
BLOCK_V_GAP = 6.0        # maximum vertical gap between consecutive lines in one block
BLOCK_H_OVERLAP = -10.0  # minimum horizontal overlap (negative = gap allowed) between lines


def pts_to_col(x: float, page_width: float) -> int:
    """Map an x-coordinate (pts) to a 0-based column index."""
    return max(0, min(GRID_COLS - 1, int(x / page_width * GRID_COLS)))


def pts_to_row(y: float, page_height: float) -> int:
    """Map a y-coordinate (pts) to a 0-based row index."""
    return max(0, min(GRID_ROWS - 1, int(y / page_height * GRID_ROWS)))


# ---------------------------------------------------------------------------
# Text-block extraction
# ---------------------------------------------------------------------------

def group_words_into_lines(words: list) -> list:
    """Group word dicts (from pdfplumber) into lines sorted top→bottom."""
    if not words:
        return []

    lines: list = []
    # Sort by top then x0
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))

    current_line: list = [sorted_words[0]]
    current_top: float = sorted_words[0]["top"]

    for word in sorted_words[1:]:
        if abs(word["top"] - current_top) <= LINE_TOL:
            current_line.append(word)
        else:
            lines.append(current_line)
            current_line = [word]
            current_top = word["top"]
    lines.append(current_line)
    return lines


def group_lines_into_blocks(lines: list) -> list:
    """Merge consecutive lines into text blocks based on vertical/horizontal proximity."""
    if not lines:
        return []

    blocks: list = []

    def line_bbox(line):
        x0 = min(w["x0"] for w in line)
        x1 = max(w["x1"] for w in line)
        top = min(w["top"] for w in line)
        bottom = max(w["bottom"] for w in line)
        return x0, x1, top, bottom

    def lines_are_adjacent(line_a, line_b):
        ax0, ax1, _, ab = line_bbox(line_a)
        bx0, bx1, bt, _ = line_bbox(line_b)
        v_gap = bt - ab
        h_overlap = min(ax1, bx1) - max(ax0, bx0)
        return 0 <= v_gap <= BLOCK_V_GAP and h_overlap >= BLOCK_H_OVERLAP

    current_block = [lines[0]]
    for line in lines[1:]:
        if lines_are_adjacent(current_block[-1], line):
            current_block.append(line)
        else:
            blocks.append(current_block)
            current_block = [line]
    blocks.append(current_block)
    return blocks


def block_text_and_bbox(block: list):
    """Return (text, x0, y0, x1, y1) for a block of lines."""
    all_words = [w for line in block for w in line]
    x0 = min(w["x0"] for w in all_words)
    x1 = max(w["x1"] for w in all_words)
    top = min(w["top"] for w in all_words)
    bottom = max(w["bottom"] for w in all_words)

    # Reconstruct text line by line, preserving order
    line_texts = []
    for line in block:
        sorted_line = sorted(line, key=lambda w: w["x0"])
        line_texts.append(" ".join(w["text"] for w in sorted_line))
    text = "\n".join(line_texts)
    return text, x0, top, x1, bottom


# ---------------------------------------------------------------------------
# Occupation grid to avoid cell conflicts
# ---------------------------------------------------------------------------

class OccupancyGrid:
    """Tracks which (row, col) cells are already occupied by merged regions."""

    def __init__(self):
        self._occupied: set = set()

    def is_free(self, r0: int, c0: int, r1: int, c1: int) -> bool:
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) in self._occupied:
                    return False
        return True

    def mark(self, r0: int, c0: int, r1: int, c1: int):
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                self._occupied.add((r, c))


# ---------------------------------------------------------------------------
# Main conversion
# ---------------------------------------------------------------------------

def convert_page(pdf_page, ws, tmp_dir: str, page_idx: int):
    """Convert a single pdfplumber page to an openpyxl worksheet."""
    pw = pdf_page.width
    ph = pdf_page.height

    # Render the page once (for image extraction)
    scale = 150 / 72  # 150 DPI
    rendered: PILImage.Image = pdf_page.to_image(resolution=150).original

    # ---- Set column widths and row heights --------------------------------
    for col_idx in range(1, GRID_COLS + 1):
        col_letter = ws.column_dimensions[
            __import__("openpyxl").utils.get_column_letter(col_idx)
        ]
        col_letter.width = COL_WIDTH

    for row_idx in range(1, GRID_ROWS + 1):
        ws.row_dimensions[row_idx].height = ROW_HEIGHT

    occupancy = OccupancyGrid()

    # ---- Place images first (text overlaid on top) ------------------------
    for img_meta in pdf_page.images:
        ix0 = img_meta["x0"]
        iy0 = img_meta["top"]   # pdfplumber uses top-down coords for 'top'
        ix1 = img_meta["x1"]
        iy1 = img_meta["bottom"]

        # Skip tiny images (likely invisible artefacts)
        if (ix1 - ix0) < 5 or (iy1 - iy0) < 5:
            continue

        # Crop from rendered page image (coords are already top-down at 72dpi;
        # rendered is at 150dpi so we scale by scale factor)
        crop_box = (
            int(ix0 * scale),
            int(iy0 * scale),
            int(ix1 * scale),
            int(iy1 * scale),
        )
        crop_box = (
            max(0, crop_box[0]),
            max(0, crop_box[1]),
            min(rendered.width, crop_box[2]),
            min(rendered.height, crop_box[3]),
        )
        if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
            continue

        cropped = rendered.crop(crop_box)

        # Save cropped image to a temp file
        img_path = os.path.join(tmp_dir, f"page{page_idx}_img_{ix0:.0f}_{iy0:.0f}.png")
        cropped.save(img_path, format="PNG")

        # Excel anchor: top-left cell of the image
        col_anchor = pts_to_col(ix0, pw) + 1  # 1-based
        row_anchor = pts_to_row(iy0, ph) + 1

        xl_img = XLImage(img_path)
        # Size in EMUs: 1 pt = 12700 EMU; image dimensions in pts
        img_w_pts = ix1 - ix0
        img_h_pts = iy1 - iy0
        xl_img.width = int(img_w_pts * 12700 / 12700 * 0.75 * 1.333)   # pts → px at 96dpi
        xl_img.height = int(img_h_pts * 12700 / 12700 * 0.75 * 1.333)

        from openpyxl.utils import get_column_letter
        cell_ref = f"{get_column_letter(col_anchor)}{row_anchor}"
        ws.add_image(xl_img, cell_ref)

        # Mark cells occupied
        col_end = pts_to_col(ix1, pw)
        row_end = pts_to_row(iy1, ph)
        occupancy.mark(row_anchor - 1, col_anchor - 1, row_end, col_end)

    # ---- Place text blocks ------------------------------------------------
    words = pdf_page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
    lines = group_words_into_lines(words)
    blocks = group_lines_into_blocks(lines)

    from openpyxl.utils import get_column_letter

    for block in blocks:
        text, bx0, by0, bx1, by1 = block_text_and_bbox(block)
        if not text.strip():
            continue

        r0 = pts_to_row(by0, ph)
        r1 = pts_to_row(by1, ph)
        c0 = pts_to_col(bx0, pw)
        c1 = pts_to_col(bx1, pw)

        # Ensure at least 1 cell
        if r1 < r0:
            r1 = r0
        if c1 < c0:
            c1 = c0

        # 1-based Excel indices
        er0, er1, ec0, ec1 = r0 + 1, r1 + 1, c0 + 1, c1 + 1

        # If top-left cell is occupied, shift slightly down
        if not occupancy.is_free(r0, c0, r1, c1):
            # Try the cell as-is anyway (text > image for occupancy)
            pass

        # Merge cells if the block spans more than one row/col
        if er0 != er1 or ec0 != ec1:
            try:
                ws.merge_cells(
                    start_row=er0, start_column=ec0,
                    end_row=er1, end_column=ec1,
                )
            except Exception:
                pass  # overlap with existing merge; write to anchor cell only

        # ws.cell() returns a MergedCell (read-only) when the coordinate falls
        # inside an existing merged range but is NOT the top-left anchor.
        # In that case we skip writing to avoid the AttributeError.
        from openpyxl.cell.cell import MergedCell
        cell = ws.cell(row=er0, column=ec0)
        if isinstance(cell, MergedCell):
            continue

        cell.value = text
        cell.alignment = Alignment(wrap_text=True, vertical="top")

        occupancy.mark(r0, c0, r1, c1)


def main():
    pdf_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "Checkliste Große Wartung.pdf")
    xlsx_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "Checkliste Große Wartung.xlsx")
    tmp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_tmp_imgs")
    os.makedirs(tmp_dir, exist_ok=True)

    wb = Workbook()
    # Remove the default empty sheet
    wb.remove(wb.active)

    print(f"Opening {pdf_path!r} …")
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            sheet_name = f"Page {i + 1}"
            print(f"  Converting {sheet_name}/{total} …", end="\r", flush=True)
            ws = wb.create_sheet(title=sheet_name)
            convert_page(page, ws, tmp_dir, i + 1)

    print(f"\nSaving {xlsx_path!r} …")
    wb.save(xlsx_path)
    print("Done.")

    # Clean up temp images
    import shutil
    shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
