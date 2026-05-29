"""
Парсер Excel с функцией вычисления формул и извлечения изображений
=========================================================

В этом модуле используются:
- xlwings: для вычисления формул (включая русские функции и внешние ссылки)
- openpyxl + openpyxl-image-loader: для извлечения изображений, связанных с ячейками

Автор: Чураев Вадим Эдуардович
"""

import os
import logging
from pathlib import Path
from typing import Dict, Any, Optional

import xlwings as xw
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader


# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def parse_excel(
    filepath: str,
    image_dir: str = "images",
    sheet_name: Optional[str] = None,
    close_app: bool = True
) -> Dict[str, Dict[str, Any]]:
    """
    Parses an Excel file to extract:
    - Computed cell values (as seen in Excel, including results of formulas like ЕСЛИ, ВПР, etc.)
    - Images embedded over cells

    Parameters
    ----------
    filepath : str
        Path to the .xlsx file.
    image_dir : str, optional
        Directory to save extracted images (default: "images").
    sheet_name : str, optional
        Name of the sheet to parse. If None, uses the active sheet.
    close_app : bool, optional
        Whether to close the Excel application after parsing (default: True).

    Returns
    -------
    dict
        Dictionary keyed by Excel coordinate (e.g., "A1") with:
        - "value": computed value (str, int, float, datetime, or None)
        - "image_path": path to saved image or None

    Raises
    ------
    FileNotFoundError
        If the Excel file does not exist.
    RuntimeError
        If Excel application fails to start or file fails to open.
    """
    filepath = os.path.abspath(filepath)
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"Excel file not found: {filepath}")

    Path(image_dir).mkdir(parents=True, exist_ok=True)
    logger.info(f"Parsing Excel file: {filepath}")
    logger.info(f"Saving images to: {os.path.abspath(image_dir)}")

    # === Шаг 1: Получение вычисленных значений через xlwings ===
    app = None
    wb_xw = None
    try:
        logger.debug("Starting Excel via xlwings...")
        app = xw.App(visible=False, add_book=False)
        wb_xw = app.books.open(filepath)
        ws_xw = wb_xw.sheets[sheet_name] if sheet_name else wb_xw.sheets.active

        used_range = ws_xw.used_range
        computed_values = {}

        if used_range is not None:
            raw_data = used_range.value
            if raw_data is None:
                logger.warning("Used range is empty.")
            else:
                # Приведение к 2D списку
                if not isinstance(raw_data, list):
                    raw_data = [[raw_data]]
                elif not isinstance(raw_data[0], list):
                    raw_data = [raw_data]

                for i, row in enumerate(raw_data):
                    for j, value in enumerate(row):
                        coord = ws_xw.range(i + 1, j + 1).address.replace("$", "")
                        computed_values[coord] = value
        else:
            logger.warning("No used range detected in worksheet.")

    except Exception as e:
        logger.exception("Failed to read computed values via xlwings")
        raise RuntimeError(f"xlwings error: {e}") from e
    finally:
        if wb_xw:
            wb_xw.close()
        if app and close_app:
            app.quit()

    # === Шаг 2: Извлечение изображений через openpyxl ===
    logger.debug("Loading workbook with openpyxl for image extraction...")
    wb_op = load_workbook(filepath, data_only=False)
    ws_op = wb_op[wb_op.sheetnames[0]] if not sheet_name else wb_op[sheet_name]

    try:
        image_loader = SheetImageLoader(ws_op)
    except ValueError as e:
        # Возникает, если в файле нет изображений
        logger.warning(f"No images found or unsupported format: {e}")
        image_loader = None

    result = {}

    for row in ws_op.iter_rows():
        for cell in row:
            coord = cell.coordinate
            img_path = None

            if image_loader and image_loader.image_in(coord):
                try:
                    img = image_loader.get(coord)
                    img_path = os.path.join(image_dir, f"{coord}.png")
                    img.save(img_path)
                    logger.debug(f"Saved image for {coord} -> {img_path}")
                except Exception as img_err:
                    logger.error(f"Failed to save image for {coord}: {img_err}")
                    img_path = f"#IMAGE_ERROR: {img_err}"

            result[coord] = {
                "value": computed_values.get(coord),
                "image_path": img_path,
            }

    wb_op.close()
    logger.info(f"Parsed {len(result)} cells. Done.")
    return result


# Пример использования (запуск как скрипт)
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python excel_parser.py <path_to_excel_file.xlsx> [image_dir]")
        sys.exit(1)

    excel_path = sys.argv[1]
    img_dir = sys.argv[2] if len(sys.argv) > 2 else "images"

    try:
        data = parse_excel(excel_path, image_dir=img_dir)
        for coord, info in sorted(data.items()):
            print(f"{coord}: value={info['value']}, image={info['image_path']}")
    except Exception as e:
        logger.exception("Parsing failed")
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
