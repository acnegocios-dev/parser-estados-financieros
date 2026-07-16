"""Extract the versioned ER visual contract from the approved manual workbook."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "sample-inputs" / "EEFF_202602_AL_Serv_Prueba.xlsx"
TARGET = ROOT / "src" / "er_style_spec.json"
TARGETS = {
    "ER": ROOT / "src" / "er_style_spec.json",
    "BG": ROOT / "src" / "bg_style_spec.json",
    "BAL": ROOT / "src" / "bal_style_spec.json",
}


def _color(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    return {
        "type": value.type,
        "rgb": value.rgb if value.type == "rgb" else None,
        "indexed": value.indexed if value.type == "indexed" else None,
        "theme": value.theme if value.type == "theme" else None,
        "tint": value.tint,
    }


def _side(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    return {"style": value.style, "color": _color(value.color)}


def _style(cell: Any) -> dict[str, Any]:
    font = cell.font
    alignment = cell.alignment
    fill = cell.fill
    border = cell.border
    protection = cell.protection
    return {
        "font": {
            "name": font.name,
            "size": font.sz,
            "bold": font.bold,
            "italic": font.italic,
            "underline": font.underline,
            "strike": font.strike,
            "color": _color(font.color),
            "vertAlign": font.vertAlign,
            "charset": font.charset,
            "family": font.family,
            "scheme": font.scheme,
            "outline": font.outline,
            "shadow": font.shadow,
            "condense": font.condense,
            "extend": font.extend,
        },
        "alignment": {
            "horizontal": alignment.horizontal,
            "vertical": alignment.vertical,
            "textRotation": alignment.textRotation,
            "wrapText": alignment.wrap_text,
            "shrinkToFit": alignment.shrink_to_fit,
            "indent": alignment.indent,
            "relativeIndent": alignment.relativeIndent,
            "justifyLastLine": alignment.justifyLastLine,
            "readingOrder": alignment.readingOrder,
        },
        "fill": {
            "fillType": fill.fill_type,
            "fgColor": _color(fill.fgColor),
            "bgColor": _color(fill.bgColor),
        },
        "border": {
            "left": _side(border.left),
            "right": _side(border.right),
            "top": _side(border.top),
            "bottom": _side(border.bottom),
            "diagonal": _side(border.diagonal),
            "diagonalUp": border.diagonalUp,
            "diagonalDown": border.diagonalDown,
            "outline": border.outline,
            "vertical": _side(border.vertical),
            "horizontal": _side(border.horizontal),
        },
        "numberFormat": cell.number_format,
        "protection": {"locked": protection.locked, "hidden": protection.hidden},
    }


def extract(sheet: str = "ER") -> dict[str, Any]:
    """Offline extractor only; generated specs are the runtime dependency."""
    workbook = load_workbook(SOURCE, data_only=False, keep_links=True)
    worksheet = workbook[sheet]
    styles: list[dict[str, Any]] = []
    style_ids: dict[str, int] = {}
    cells: dict[str, int] = {}

    bounds = {
        "ER": (9, 70, 2, 10),
        "BG": (7, 47, 2, 12),
        "BAL": (1, 185, 3, 7),
    }[sheet]
    first_row, last_row, first_column, last_column = bounds
    for row in range(first_row, last_row + 1):
        for column in range(first_column, last_column + 1):
            cell = worksheet.cell(row=row, column=column)
            style = _style(cell)
            key = json.dumps(style, sort_keys=True, ensure_ascii=False)
            if key not in style_ids:
                style_ids[key] = len(styles)
                styles.append(style)
            cells[cell.coordinate] = style_ids[key]

    row_heights = {
        str(row): worksheet.row_dimensions[row].height
        for row in range(1, 88)
        if worksheet.row_dimensions[row].height is not None
    }
    hidden_rows = [
        row for row in range(1, 88) if worksheet.row_dimensions[row].hidden
    ]

    return {
        "version": f"2026-07-16.{sheet.lower()}-v1" if sheet != "ER" else "2026-07-13.manual-er-v1",
        "source": SOURCE.name,
        "source_sha256": hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
        "sheet": sheet,
        "visible_range": f"{worksheet.cell(first_row, first_column).coordinate}:{worksheet.cell(last_row, last_column).coordinate}",
        "geometry": {
            "default_row_height": worksheet.sheet_format.defaultRowHeight,
            "default_column_width": worksheet.sheet_format.defaultColWidth,
            "base_column_width": worksheet.sheet_format.baseColWidth,
            "column_widths": {
                column: worksheet.column_dimensions[column].width
                for column in "ABCDEFGHIJ"
            },
            "row_heights": row_heights,
            "hidden_rows": hidden_rows,
            "merged_ranges": sorted(str(value) for value in worksheet.merged_cells.ranges),
            "page_margins": {
                key: getattr(worksheet.page_margins, key)
                for key in ("left", "right", "top", "bottom", "header", "footer")
            },
            "page_setup": {
                "orientation": worksheet.page_setup.orientation,
                "horizontalDpi": worksheet.page_setup.horizontalDpi,
                "verticalDpi": worksheet.page_setup.verticalDpi,
            },
            "show_grid_lines": worksheet.sheet_view.showGridLines,
            "top_left_cell": worksheet.sheet_view.topLeftCell,
            "freeze_panes": worksheet.freeze_panes,
        },
        "styles": styles,
        "cells": cells,
    }


if __name__ == "__main__":
    for sheet, target in TARGETS.items():
        target.write_text(
            json.dumps(extract(sheet), indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
        print(target)
