"""
Парсер Excel с вычислением формул, извлечением изображений и метаданных ячеек
=============================================================================

Использует:
- xlwings: вычисление формул через Excel (поддержка русских функций, внешних ссылок)
- openpyxl: извлечение изображений, формул, форматирования, объединённых ячеек

Автор: Чураев Вадим Эдуардович || РЭУ ИМ.ПЛЕХАНОВА
"""

import os
import re
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl_image_loader import SheetImageLoader


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_letter(n: int) -> str:
    return get_column_letter(n)


def _resolve_merged(ws_op) -> Dict[str, str]:
    """Return mapping coord -> top-left coord for every cell inside a merged range."""
    mapping: Dict[str, str] = {}
    for merged_range in ws_op.merged_cells.ranges:
        top_left = f"{_col_letter(merged_range.min_col)}{merged_range.min_row}"
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                coord = f"{_col_letter(col)}{row}"
                if coord != top_left:
                    mapping[coord] = top_left
    return mapping


def _detect_real_bounds(ws_op, max_scan_col: int = 500) -> Tuple[int, int]:
    """
    Detect actual last row/col with data, capped at max_scan_col columns.
    Avoids the bloated used_range that openpyxl reports when there are
    stray styles far to the right (e.g. A1:XCA20323 when data ends at BN).
    """
    last_row = 0
    last_col = 0
    for row in ws_op.iter_rows(max_col=max_scan_col):
        for cell in row:
            if cell.value is not None:
                if cell.row > last_row:
                    last_row = cell.row
                if cell.column > last_col:
                    last_col = cell.column
    return last_row, last_col


def _cell_color(cell) -> Optional[str]:
    """Return hex fill color of a cell, or None if no fill / transparent."""
    try:
        fg = cell.fill.fgColor
        if fg.type == "rgb" and fg.rgb not in ("00000000", "FFFFFFFF", "00FFFFFF"):
            return fg.rgb
        if fg.type == "theme":
            return f"theme:{fg.theme}"
    except Exception:
        pass
    return None


def _cell_formatting(cell) -> Dict[str, Any]:
    fmt: Dict[str, Any] = {}
    try:
        fmt["bold"] = bool(cell.font.bold)
        fmt["font_name"] = cell.font.name
        fmt["font_size"] = cell.font.size
    except Exception:
        pass
    color = _cell_color(cell)
    if color:
        fmt["fill_color"] = color
    try:
        fmt["number_format"] = cell.number_format
    except Exception:
        pass
    return fmt


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_excel(
    filepath: str,
    image_dir: str = "images",
    sheet_name: Optional[str] = None,
    close_app: bool = True,
    skip_empty: bool = True,
    include_formula: bool = True,
    include_formatting: bool = False,
    max_col: Optional[int] = None,
    max_row: Optional[int] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    Parse an Excel file and extract computed values, formulas, images,
    merged-cell info, and optionally cell formatting.

    Parameters
    ----------
    filepath : str
        Path to the .xlsx file.
    image_dir : str
        Directory to save extracted images.
    sheet_name : str, optional
        Sheet name or None for the active sheet.
    close_app : bool
        Quit Excel after reading (default True).
    skip_empty : bool
        Omit cells where value, formula, and image are all None/empty.
    include_formula : bool
        Include raw formula text in the output (default True).
    include_formatting : bool
        Include cell formatting metadata (bold, fill color, etc.).
    max_col : int, optional
        Stop scanning after this column number. Useful for files with
        thousands of bloat columns (e.g. stray styles reaching XCA).
        If None, auto-detects by scanning up to column 500.
    max_row : int, optional
        Stop scanning after this row. If None, uses all rows.

    Returns
    -------
    dict
        Keyed by Excel coordinate (e.g. "A1"):
        - "value"      : computed cell value (str/int/float/datetime/None)
        - "formula"    : raw formula string or None  (if include_formula)
        - "image_path" : path to saved image or None
        - "merged_into": coordinate of top-left merge cell, if applicable
        - "formatting" : dict with bold/fill_color/etc.  (if include_formatting)

    Raises
    ------
    FileNotFoundError
        If the Excel file does not exist.
    RuntimeError
        If Excel fails to open the file.
    """
    filepath = os.path.abspath(filepath)
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"Excel file not found: {filepath}")

    Path(image_dir).mkdir(parents=True, exist_ok=True)
    logger.info(f"Parsing: {filepath}")

    # ── Step 1: computed values via xlwings ──────────────────────────────────
    computed_values: Dict[str, Any] = {}
    app = None
    wb_xw = None
    try:
        logger.debug("Opening Excel via xlwings...")
        app = xw.App(visible=False, add_book=False)
        wb_xw = app.books.open(filepath)
        ws_xw = wb_xw.sheets[sheet_name] if sheet_name else wb_xw.sheets.active

        used = ws_xw.used_range
        if used is None:
            logger.warning("xlwings: no used range detected.")
        else:
            raw = used.value
            if raw is None:
                logger.warning("xlwings: used range is empty.")
            else:
                if not isinstance(raw, list):
                    raw = [[raw]]
                elif not isinstance(raw[0], list):
                    raw = [raw]

                origin_row = used.row
                origin_col = used.column
                for i, row in enumerate(raw):
                    for j, val in enumerate(row):
                        r = origin_row + i
                        c = origin_col + j
                        # Respect max_col / max_row limits
                        if max_col and c > max_col:
                            break
                        if max_row and r > max_row:
                            break
                        coord = f"{_col_letter(c)}{r}"
                        computed_values[coord] = val

    except Exception as e:
        logger.exception("xlwings failed — computed values unavailable")
        raise RuntimeError(f"xlwings error: {e}") from e
    finally:
        if wb_xw:
            try:
                wb_xw.close()
            except Exception:
                pass
        if app and close_app:
            try:
                app.quit()
            except Exception:
                pass

    # ── Step 2: formulas, images, formatting via openpyxl ───────────────────
    logger.debug("Loading workbook via openpyxl...")
    wb_op = load_workbook(filepath, data_only=False)
    ws_op = wb_op[sheet_name] if sheet_name else wb_op.active

    # Auto-detect real column/row bounds to skip bloat
    if max_col is None:
        logger.debug("Auto-detecting real column bounds (max 500 cols)...")
        _, detected_col = _detect_real_bounds(ws_op, max_scan_col=500)
        effective_max_col = max(detected_col, 1) if detected_col else 500
    else:
        effective_max_col = max_col

    effective_max_row = max_row  # None = no limit

    # Merged cells mapping
    merged_map = _resolve_merged(ws_op)

    # Image loader
    try:
        image_loader = SheetImageLoader(ws_op)
    except (ValueError, Exception) as e:
        logger.warning(f"Image loader unavailable: {e}")
        image_loader = None

    result: Dict[str, Dict[str, Any]] = {}

    for row in ws_op.iter_rows(max_col=effective_max_col, max_row=effective_max_row):
        for cell in row:
            coord = cell.coordinate

            # Formula
            formula: Optional[str] = None
            if include_formula and isinstance(cell.value, str) and cell.value.startswith("="):
                formula = cell.value

            # Image
            img_path: Optional[str] = None
            if image_loader:
                try:
                    if image_loader.image_in(coord):
                        img = image_loader.get(coord)
                        img_path = os.path.join(image_dir, f"{coord}.png")
                        img.save(img_path)
                        logger.debug(f"Image saved: {coord} -> {img_path}")
                except Exception as img_err:
                    logger.error(f"Image error at {coord}: {img_err}")

            value = computed_values.get(coord)

            # Skip truly empty cells
            if skip_empty and value is None and formula is None and img_path is None:
                continue

            entry: Dict[str, Any] = {
                "value": value,
                "image_path": img_path,
            }

            if include_formula:
                entry["formula"] = formula

            merged_into = merged_map.get(coord)
            if merged_into:
                entry["merged_into"] = merged_into

            if include_formatting:
                entry["formatting"] = _cell_formatting(cell)

            result[coord] = entry

    wb_op.close()
    logger.info(f"Done. Parsed {len(result)} cells.")
    return result


# ---------------------------------------------------------------------------
# Header extraction helper
# ---------------------------------------------------------------------------

def extract_headers(
    result: Dict[str, Dict[str, Any]],
    header_rows: List[int],
) -> Dict[str, str]:
    """
    Build a column -> header label mapping from one or more header rows.
    Concatenates non-empty values from each header row for each column.

    Example: header_rows=[2, 3] mimics the two-row group/sub-header pattern
    in Практика Python.xlsx.
    """
    headers: Dict[str, List[str]] = {}
    for coord, info in result.items():
        match = re.match(r"([A-Z]+)(\d+)", coord)
        if not match:
            continue
        col, row = match.group(1), int(match.group(2))
        if row not in header_rows:
            continue
        val = info.get("value")
        if val is None:
            val = info.get("formula") or ""
        val = str(val).strip()
        if val:
            headers.setdefault(col, []).append(val)

    return {col: " / ".join(parts) for col, parts in headers.items()}


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python \"parser (1).py\" <file.xlsx> [image_dir] [sheet_name]")
        sys.exit(1)

    excel_path = sys.argv[1]
    img_dir    = sys.argv[2] if len(sys.argv) > 2 else "images"
    sname      = sys.argv[3] if len(sys.argv) > 3 else None

    try:
        data = parse_excel(
            excel_path,
            image_dir=img_dir,
            sheet_name=sname,
            include_formula=True,
            include_formatting=True,
        )

        # Print summary
        has_image = sum(1 for v in data.values() if v.get("image_path"))
        has_formula = sum(1 for v in data.values() if v.get("formula"))
        print(f"Total cells: {len(data)}")
        print(f"  with values : {sum(1 for v in data.values() if v['value'] is not None)}")
        print(f"  with formulas: {has_formula}")
        print(f"  with images  : {has_image}")

        # Print first 30 non-empty cells
        print("\nFirst 30 cells:")
        for coord, info in list(sorted(data.items()))[:30]:
            print(f"  {coord}: value={repr(info['value'])[:50]}  formula={info.get('formula','')[:40]}")

    except Exception:
        logger.exception("Parsing failed")
        sys.exit(1)
