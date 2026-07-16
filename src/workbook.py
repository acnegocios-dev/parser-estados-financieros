"""Workbook generation for the local financial statements prototype.

This module owns the generated ER sheet only. It accepts loose dict/dataclass
inputs so it can be wired to parser and ER-dataset modules created elsewhere.
"""

from __future__ import annotations

import calendar
from dataclasses import dataclass, field, is_dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import unicodedata
from typing import Any, Iterable, Mapping

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_DIR = ROOT / "sample-outputs"
ER_STYLE_SPEC_PATH = Path(__file__).with_name("er_style_spec.json")
BG_STYLE_SPEC_PATH = Path(__file__).with_name("bg_style_spec.json")
BAL_STYLE_SPEC_PATH = Path(__file__).with_name("bal_style_spec.json")

ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
MONTHS_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


@dataclass
class WorkbookBuildResult:
    workbook: Workbook
    output_path: Path | None
    company: str
    period: str
    source_path: str | None
    formula_cells: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    missing_accounts: list[str] = field(default_factory=list)
    generated_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))


ER_LAYOUT = [
    {"row": 17, "label": "I N G R E S O S", "kind": "section"},
    {"row": 18, "label": "Ingresos por servicios", "key": "ingresos_por_servicios"},
    {"row": 19, "label": "Descuentos o bonificaciones", "key": "descuentos_o_bonificaciones"},
    {"row": 20, "label": "INGRESOS NETOS", "kind": "subtotal", "formula": "SUM({col}18:{col}19)"},
    {"row": 22, "label": "C O S T O S", "kind": "section"},
    {"row": 23, "label": "Costo de Ventas", "key": "costo_de_ventas"},
    {"row": 25, "label": "UTILIDAD BRUTA", "kind": "subtotal", "formula": "{col}20-{col}23"},
    {"row": 27, "label": "Gastos de Operacion", "kind": "section"},
    {"row": 28, "label": "Sueldos y Salarios", "key": "sueldos_y_salarios"},
    {"row": 29, "label": "Impuestos y Derechos", "key": "impuestos_y_derechos"},
    {"row": 30, "label": "Honorarios", "key": "honorarios"},
    {"row": 31, "label": "Arrendamiento", "key": "arrendamiento"},
    {"row": 32, "label": "Seguros y Fianzas", "key": "seguros_y_fianzas"},
    {"row": 33, "label": "Servicios", "key": "servicios"},
    {"row": 34, "label": "Capacitacion al Personal", "key": "capacitacion_al_personal"},
    {"row": 35, "label": "Fletes y/o Mensajeria", "key": "fletes_y_o_mensajeria"},
    {"row": 36, "label": "Seguridad e higiene", "key": "seguridad_e_higiene"},
    {"row": 37, "label": "Mantenimiento", "key": "mantenimiento"},
    {"row": 38, "label": "Combustibles", "key": "combustibles"},
    {"row": 39, "label": "Propaganda y Publicidad", "key": "propaganda_y_publicidad"},
    {"row": 40, "label": "Cuotas y Suscripciones", "key": "cuotas_y_suscripciones"},
    {"row": 41, "label": "Gastos de Viaje", "key": "gastos_de_viaje"},
    {"row": 42, "label": "Herrajes y Herramientas", "key": "herrajes_y_herramientas"},
    {"row": 43, "label": "Papeleria y Art. de Oficina", "key": "papeleria_y_art_de_oficina"},
    {"row": 44, "label": "Depreciaciones", "key": "depreciaciones"},
    {"row": 45, "label": "Recargos", "key": "recargos"},
    {"row": 46, "label": "Varios", "key": "varios"},
    {"row": 47, "label": "Uniformes", "key": "uniformes"},
    {"row": 50, "label": "No deducibles", "key": "no_deducibles"},
    {"row": 51, "label": "GASTOS DE OPERACION", "kind": "subtotal", "formula": "SUM({col}28:{col}50)"},
    {"row": 53, "label": "UTILIDAD O (PERDIDA) DE OPERACION", "kind": "subtotal", "formula": "{col}25-{col}51"},
    {"row": 55, "label": "OTROS INGRESOS Y GASTOS", "kind": "section"},
    {"row": 56, "label": "Otros Productos", "key": "otros_productos"},
    {"row": 57, "label": "Otros Gastos", "key": "otros_gastos"},
    {"row": 58, "label": "TOTAL OTROS INGRESOS", "kind": "subtotal", "formula": "SUM({col}56:{col}57)"},
    {"row": 60, "label": "RES. INT. DE FINANCIAMIENTO", "kind": "section"},
    {"row": 61, "label": "Productos Financieros", "key": "productos_financieros"},
    {"row": 62, "label": "Gastos Financieros", "key": "gastos_financieros"},
    {"row": 63, "label": "TOTAL R. I. F.", "kind": "subtotal", "formula": "SUM({col}61:{col}62)"},
    {"row": 65, "label": "RESULTADO ANTES DE IMPUESTOS", "kind": "subtotal", "formula": "{col}53+{col}58+{col}63"},
    {"row": 67, "label": "ISR DEL EJERCICIO", "key": "isr_del_ejercicio"},
    {"row": 68, "label": "PTU DEL EJERCICIO", "key": "ptu_del_ejercicio"},
    {"row": 70, "label": "RESULTADO DEL EJERCICIO", "kind": "subtotal", "formula": "{col}65-{col}67-{col}68"},
]


def build_er_workbook(
    dataset: Any | None = None,
    *,
    metadata: Any | None = None,
    source_path: str | Path | None = None,
) -> WorkbookBuildResult:
    """Compatibility wrapper for the full financial-statements workbook."""

    bg_dataset, bal_dataset, metadata = _derive_supporting_datasets(
        dataset, metadata=metadata, source_path=source_path
    )
    return build_financial_statements_workbook(
        dataset,
        bg_dataset=bg_dataset,
        bal_dataset=bal_dataset,
        metadata=metadata,
        source_path=source_path,
    )


def build_financial_statements_workbook(
    er_dataset: Any | None = None,
    *,
    bg_dataset: Any | None = None,
    bal_dataset: Any | None = None,
    metadata: Any | None = None,
    source_path: str | Path | None = None,
) -> WorkbookBuildResult:
    """Build the self-contained, ordered ``BG`` + ``ER`` + ``BAL`` workbook."""

    metadata_values = _metadata_from_inputs(er_dataset, metadata, source_path)
    company = metadata_values["company"]
    period = metadata_values["period"]
    warnings = list(metadata_values["warnings"])
    missing_accounts = _coerce_list(_read_field(er_dataset, ("missing_accounts", "cuentas_no_encontradas")))
    missing_accounts.extend(_coerce_list(_read_field(er_dataset, ("missing", "not_found"))))

    wb = Workbook()
    _configure_recalculation_on_open(wb)
    bg = wb.active
    bg.title = "BG"
    er = wb.create_sheet("ER")
    bal = wb.create_sheet("BAL")
    for worksheet in wb.worksheets:
        worksheet.sheet_state = "visible"

    formula_cells: list[str] = []
    _write_bg_sheet(bg, company, period, bg_dataset or {}, formula_cells)
    _write_header(er, company, period)
    line_values = _index_dataset_lines(er_dataset, warnings)
    formula_cells.extend(_write_er_body(er, line_values, warnings))
    _style_er_sheet(er)
    _write_bal_sheet(bal, company, period, bal_dataset or {}, metadata, formula_cells)

    return WorkbookBuildResult(
        workbook=wb,
        output_path=None,
        company=company,
        period=period,
        source_path=str(source_path) if source_path is not None else metadata_values["source_path"],
        formula_cells=formula_cells,
        warnings=warnings,
        missing_accounts=missing_accounts,
    )


def _derive_supporting_datasets(
    er_dataset: Any,
    *,
    metadata: Any | None,
    source_path: str | Path | None,
) -> tuple[dict[str, Any], dict[str, Any], Any | None]:
    """Derive BG/BAL only from the loaded input, never from a manual workbook."""

    if source_path is None:
        return {}, {}, metadata
    try:
        from .engine import build_bal_dataset, build_bg_dataset, build_er_dataset
        from .parser import parse_balanza
    except ImportError:  # pragma: no cover - supports direct execution.
        from engine import build_bal_dataset, build_bg_dataset, build_er_dataset
        from parser import parse_balanza

    parsed = parse_balanza(source_path)
    effective_er = er_dataset or build_er_dataset(
        parsed.rows,
        company=parsed.company_name,
        period=parsed.period.period_ym,
        source_path=parsed.source_path,
    )
    merged_metadata = dict(_to_dict(metadata))
    merged_metadata.update({
        "company": parsed.company_name,
        "period": parsed.period.period_ym,
        "rfc": parsed.period.rfc,
        "source_path": parsed.source_path,
    })
    raw_amounts = _read_field(effective_er, ("raw_amounts",)) or {}
    return (
        build_bg_dataset(parsed.rows, result_ejercicio=raw_amounts.get("resultado_ejercicio")),
        build_bal_dataset(parsed.rows),
        merged_metadata,
    )


def _write_bg_sheet(
    ws,
    company: str,
    period: str,
    dataset: Any,
    formula_cells: list[str],
) -> None:
    spec = _load_style_spec(BG_STYLE_SPEC_PATH)
    _apply_sheet_geometry(ws, spec["geometry"])
    for merged_range in ("B7:L7", "B8:L8", "B9:L9", "B10:L10"):
        ws.merge_cells(merged_range)
    ws["B7"] = company
    ws["B8"] = "Balance General"
    ws["B9"] = _bg_period_caption(period)
    ws["B10"] = "(Importes expresados en pesos)"

    lines = {str(line.get("key")): line for line in _coerce_list_of_dicts(_read_field(dataset, ("lines",)))}
    asset_keys = (
        "caja_chica", "bancos", "inversiones_y_valores", "cuentas_por_cobrar", "contribuciones",
        "inventarios", "deudores_diversos", "anticipo_a_proveedores", "pagos_anticipados",
        "mobiliario_y_equipo", "equipo_de_computo", "equipo_de_transporte", "maquinaria_y_equipo",
        "depreciacion_acumulada",
    )
    liability_keys = ("proveedores", "acreedores_diversos", "anticipos_de_clientes", "impuestos_por_pagar", "otros_pasivos_ptu")
    capital_keys = ("capital_social", "aportaciones_para_aumentos_de_capital", "resultados_de_ejercicios_anteriores")
    _write_bg_lines(ws, asset_keys, lines, label_column="B", amount_column="F", start_row=13)
    _write_bg_lines(ws, liability_keys, lines, label_column="H", amount_column="L", start_row=13)
    _write_bg_lines(ws, capital_keys, lines, label_column="H", amount_column="L", start_row=20)
    ws["H23"] = "Resultado del ejercicio"
    ws["L23"] = "=ER!H70"
    formula_cells.append("BG!L23")
    ws["B45"] = "TOTAL ACTIVO"
    ws["F45"] = "=SUM(F13:F26)"
    ws["H45"] = "TOTAL PASIVO Y CAPITAL"
    ws["L45"] = "=SUM(L13:L17,L20:L23)"
    ws["H47"] = "CUADRE"
    ws["L47"] = "=F45-L45"
    formula_cells.extend(("BG!F45", "BG!L45", "BG!L47"))
    _apply_bg_presentation(ws)
    _apply_white_area(ws, min_row=1, max_row=47, min_col=1, max_col=12)
    ws.sheet_view.showGridLines = False


def _write_bg_lines(ws, keys, lines, *, label_column: str, amount_column: str, start_row: int) -> None:
    for offset, key in enumerate(keys):
        row = start_row + offset
        line = lines.get(key, {})
        ws[f"{label_column}{row}"] = line.get("label", key.replace("_", " ").title())
        ws[f"{amount_column}{row}"] = line.get("amount", 0)


def _write_bal_sheet(
    ws,
    company: str,
    period: str,
    dataset: Any,
    metadata: Any,
    formula_cells: list[str],
) -> None:
    spec = _load_style_spec(BAL_STYLE_SPEC_PATH)
    _apply_sheet_geometry(ws, spec["geometry"])
    for merged_range in ("C1:G1", "C2:G2", "C3:G3", "C4:G4"):
        ws.merge_cells(merged_range)
    rfc = _read_field(metadata, ("rfc",)) or "RFC por confirmar"
    ws["C1"] = company
    ws["C2"] = "Balanza de Comprobacion"
    ws["C3"] = _bal_period_caption(period)
    ws["C4"] = f"RFC: {rfc}"
    headers = ("CUENTA", "SALDO INICIAL", "DEBE", "HABER", "SALDO FINAL")
    for column, value in enumerate(headers, start=3):
        ws.cell(6, column, value)

    rows = _coerce_list_of_dicts(_read_field(dataset, ("rows", "all_rows")))
    accumulator_rows: list[int] = []
    for index, row in enumerate(rows, start=7):
        ws.cell(index, 3, row.get("account_raw") or row.get("account_code"))
        ws.cell(index, 4, row.get("saldo_inicial", 0))
        ws.cell(index, 5, row.get("debe", 0))
        ws.cell(index, 6, row.get("haber", 0))
        ws.cell(index, 7, row.get("saldo_final", 0))
        if row.get("is_accumulator"):
            accumulator_rows.append(index)
    separator_row = 7 + len(rows)
    sum_row = separator_row + 1
    ws.cell(sum_row, 3, "SUMAS IGUALES")
    ws.cell(sum_row, 4, "=0")
    formula_cells.append(f"BAL!D{sum_row}")
    for column in (5, 6):
        formula = _sum_accumulator_formula(get_column_letter(column), accumulator_rows)
        ws.cell(sum_row, column, formula)
        formula_cells.append(f"BAL!{get_column_letter(column)}{sum_row}")
    ws.cell(sum_row, 7, f"=D{sum_row}+E{sum_row}-F{sum_row}")
    formula_cells.append(f"BAL!G{sum_row}")
    ws.print_area = f"C1:G{sum_row}"
    ws.freeze_panes = None  # The manual A137 pane is a historical residue.
    _apply_bal_presentation(ws, sum_row)


def _sum_accumulator_formula(column: str, rows: list[int]) -> str:
    if not rows:
        return "=0"
    return "=SUM(" + ",".join(f"{column}{row}" for row in rows) + ")"


def _apply_bg_presentation(ws) -> None:
    money = '#,##0.00;[Red]\\-#,##0.00'
    for row in range(7, 48):
        for column in range(2, 13):
            cell = ws.cell(row, column)
            cell.font = Font(name="Arial", size=8)
            cell.alignment = Alignment(vertical="center")
    for coordinate in ("B7", "B8", "B9", "B10"):
        ws[coordinate].font = Font(name="Arial", size=10, bold=True)
        ws[coordinate].alignment = Alignment(horizontal="center", vertical="center")
    for coordinate in ("F45", "L45", "L47"):
        ws[coordinate].font = Font(name="Arial", size=8, bold=True)
    for column in ("F", "L"):
        for row in range(13, 48):
            ws[f"{column}{row}"].number_format = money
            ws[f"{column}{row}"].alignment = Alignment(horizontal="right")


def _apply_bal_presentation(ws, sum_row: int) -> None:
    money = '"$"#,##0.00;[Red]\\-"$"#,##0.00'
    thin = Side(style="hair", color="FF000000")
    for row in range(1, sum_row + 1):
        for column in range(3, 8):
            cell = ws.cell(row, column)
            cell.font = Font(name="Arial", size=7, bold=row in (6, sum_row))
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if column >= 4:
                cell.number_format = money
                cell.alignment = Alignment(horizontal="right")
    for coordinate in ("C1", "C2", "C3", "C4"):
        ws[coordinate].font = Font(name="Arial", size=9, bold=True)
        ws[coordinate].alignment = Alignment(horizontal="center")


def _apply_white_area(ws, *, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    white = PatternFill(fill_type="solid", fgColor="FFFFFFFF", bgColor="FFFFFFFF")
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.fill = white


def _apply_sheet_geometry(ws, geometry: Mapping[str, Any]) -> None:
    ws.sheet_format.defaultRowHeight = geometry.get("default_row_height", 13.2)
    ws.sheet_format.defaultColWidth = geometry.get("default_column_width", 10)
    ws.sheet_format.baseColWidth = geometry.get("base_column_width", 10)
    for column, width in geometry.get("column_widths", {}).items():
        ws.column_dimensions[column].width = width
    for row, height in geometry.get("row_heights", {}).items():
        ws.row_dimensions[int(row)].height = height
    margins = geometry.get("page_margins", {})
    for name, value in margins.items():
        setattr(ws.page_margins, name, value)
    if geometry.get("page_setup", {}).get("orientation"):
        ws.page_setup.orientation = geometry["page_setup"]["orientation"]


def _load_style_spec(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _configure_recalculation_on_open(workbook: Workbook) -> None:
    """Ask Excel-compatible clients to fully recalculate formula cells."""

    calculation = getattr(workbook, "calculation", None)
    if calculation is None:
        return
    for attribute, value in (
        ("calcMode", "auto"),
        ("fullCalcOnLoad", True),
        ("forceFullCalc", True),
    ):
        if hasattr(calculation, attribute):
            setattr(calculation, attribute, value)


def save_er_workbook(
    dataset: Any | None = None,
    output_path: str | Path | None = None,
    *,
    metadata: Any | None = None,
    source_path: str | Path | None = None,
) -> WorkbookBuildResult:
    """Compatibility wrapper for :func:`save_financial_statements_workbook`."""

    return save_financial_statements_workbook(
        dataset, output_path, metadata=metadata, source_path=source_path
    )


def save_financial_statements_workbook(
    er_dataset: Any | None = None,
    output_path: str | Path | None = None,
    *,
    bg_dataset: Any | None = None,
    bal_dataset: Any | None = None,
    metadata: Any | None = None,
    source_path: str | Path | None = None,
) -> WorkbookBuildResult:
    """Build and save a self-contained financial-statements workbook."""

    if bg_dataset is None or bal_dataset is None:
        derived_bg, derived_bal, metadata = _derive_supporting_datasets(
            er_dataset, metadata=metadata, source_path=source_path
        )
        bg_dataset = bg_dataset or derived_bg
        bal_dataset = bal_dataset or derived_bal
    result = build_financial_statements_workbook(
        er_dataset,
        bg_dataset=bg_dataset,
        bal_dataset=bal_dataset,
        metadata=metadata,
        source_path=source_path,
    )
    target = Path(output_path) if output_path is not None else _default_output_path(result.company, result.period)
    target.parent.mkdir(parents=True, exist_ok=True)
    result.workbook.save(target)
    result.output_path = target
    return result


def detect_source_metadata(path: str | Path) -> dict[str, str | None]:
    """Read lightweight company/period metadata from a workbook-like source.

    This is intentionally not a full balanza parser; it only supports local
    evidence generation until the parser agent wires `parse_balanza`.
    """

    source = Path(path)
    with source.open("rb") as handle:
        wb = load_workbook(handle, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        company = None
        period = None
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 12), values_only=True):
            for value in row:
                if value is None:
                    continue
                text = str(value).strip()
                if not text:
                    continue
                lowered = text.lower()
                if company is None and "periodo" not in lowered and re.search(r"\b[A-Z0-9]{12,13}\b", text):
                    company = text
                if period is None:
                    match = re.search(r"(\d{4})[-/](\d{1,2})", text)
                    if match:
                        period = f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"
            if company and period:
                break
        wb.close()
    return {"company": company, "period": period, "source_path": str(source)}


def _write_header(ws, company: str, period: str) -> None:
    ws["B9"] = company
    ws["B10"] = "Estado de Resultados"
    ws["B11"] = _period_caption(period)
    ws["B12"] = "(Importes expresados en pesos)"
    ws["D15"] = "Del Período "
    ws["F15"] = "%"
    ws["H15"] = "Acumulado"
    ws["J15"] = "%"


def _write_er_body(ws, line_values: dict[str, dict[str, Any]], warnings: list[str]) -> list[str]:
    formula_cells: list[str] = []
    numeric_rows: list[int] = []

    for spec in ER_LAYOUT:
        row = int(spec["row"])
        label = str(spec["label"])
        kind = spec.get("kind", "line")
        ws.cell(row=row, column=2, value=label)

        if kind == "section":
            continue

        numeric_rows.append(row)
        if kind == "subtotal":
            formula = "=" + str(spec["formula"]).format(col="H")
            _set_formula(ws, f"H{row}", formula, formula_cells)
            continue

        values = _line_amounts(line_values.get(str(spec.get("key"))), warnings, label)
        ws[f"H{row}"] = values["accumulated"]

    for row in numeric_rows:
        _set_formula(ws, f"J{row}", f'=IF($H$18=0,0,H{row}/$H$18)', formula_cells)

    return formula_cells


def _set_formula(ws, cell: str, formula: str, formula_cells: list[str]) -> None:
    safe_formula = _sanitize_formula(formula)
    ws[cell] = safe_formula
    formula_cells.append(f"ER!{cell}")


def _sanitize_formula(formula: str) -> str:
    clean = formula.strip()
    if not clean.startswith("="):
        clean = "=" + clean
    upper = clean.upper()
    if any(token in upper for token in ERROR_TOKENS) or "[" in clean or "]" in clean:
        return "=0"
    return clean


def _style_er_sheet(ws) -> None:
    """Apply the versioned visual contract extracted from the manual ER."""

    spec = _load_er_style_spec()
    geometry = spec["geometry"]

    ws.sheet_format.defaultRowHeight = geometry["default_row_height"]
    ws.sheet_format.defaultColWidth = geometry["default_column_width"]
    ws.sheet_format.baseColWidth = geometry["base_column_width"]
    for column, width in geometry["column_widths"].items():
        ws.column_dimensions[column].width = width

    for row in range(1, 88):
        ws.row_dimensions[row].height = None
        ws.row_dimensions[row].hidden = False
    for row, height in geometry["row_heights"].items():
        ws.row_dimensions[int(row)].height = height
    for row in geometry["hidden_rows"]:
        ws.row_dimensions[int(row)].hidden = True

    for merged_range in geometry["merged_ranges"]:
        ws.merge_cells(merged_range)

    for name, value in geometry["page_margins"].items():
        setattr(ws.page_margins, name, value)
    for name, value in geometry["page_setup"].items():
        setattr(ws.page_setup, name, value)
    ws.sheet_view.showGridLines = geometry["show_grid_lines"]
    ws.sheet_view.topLeftCell = geometry["top_left_cell"]
    ws.freeze_panes = geometry["freeze_panes"]

    for coordinate, style_id in spec["cells"].items():
        cell = ws[coordinate]
        _apply_style(cell, spec["styles"][style_id])
    _apply_white_area(ws, min_row=1, max_row=70, min_col=1, max_col=10)
    ws.sheet_view.showGridLines = False


def _load_er_style_spec() -> dict[str, Any]:
    with ER_STYLE_SPEC_PATH.open(encoding="utf-8") as handle:
        return json.load(handle)


def _color_from_spec(value: dict[str, Any] | None) -> Color | None:
    if not value:
        return None
    kwargs: dict[str, Any] = {"type": value["type"]}
    for key in ("rgb", "indexed", "theme", "tint"):
        if value.get(key) is not None:
            kwargs[key] = value[key]
    return Color(**kwargs)


def _side_from_spec(value: dict[str, Any] | None) -> Side | None:
    if not value:
        return None
    return Side(style=value.get("style"), color=_color_from_spec(value.get("color")))


def _apply_style(cell: Any, spec: dict[str, Any]) -> None:
    font = spec["font"]
    cell.font = Font(
        name=font["name"],
        sz=font["size"],
        b=font["bold"],
        i=font["italic"],
        underline=font["underline"],
        strike=font["strike"],
        color=_color_from_spec(font["color"]),
        vertAlign=font["vertAlign"],
        charset=font["charset"],
        family=font["family"],
        scheme=font["scheme"],
        outline=font["outline"],
        shadow=font["shadow"],
        condense=font["condense"],
        extend=font["extend"],
    )
    alignment = spec["alignment"]
    cell.alignment = Alignment(
        horizontal=alignment["horizontal"],
        vertical=alignment["vertical"],
        textRotation=alignment["textRotation"],
        wrap_text=alignment["wrapText"],
        shrink_to_fit=alignment["shrinkToFit"],
        indent=alignment["indent"],
        relativeIndent=alignment["relativeIndent"],
        justifyLastLine=alignment["justifyLastLine"],
        readingOrder=alignment["readingOrder"],
    )
    fill = spec["fill"]
    cell.fill = PatternFill(
        fill_type=fill["fillType"],
        fgColor=_color_from_spec(fill["fgColor"]),
        bgColor=_color_from_spec(fill["bgColor"]),
    )
    border = spec["border"]
    cell.border = Border(
        left=_side_from_spec(border["left"]),
        right=_side_from_spec(border["right"]),
        top=_side_from_spec(border["top"]),
        bottom=_side_from_spec(border["bottom"]),
        diagonal=_side_from_spec(border["diagonal"]),
        diagonalUp=border["diagonalUp"],
        diagonalDown=border["diagonalDown"],
        outline=border["outline"],
        vertical=_side_from_spec(border["vertical"]),
        horizontal=_side_from_spec(border["horizontal"]),
    )
    protection = spec["protection"]
    cell.protection = Protection(
        locked=protection["locked"], hidden=protection["hidden"]
    )
    cell.number_format = spec["numberFormat"]


def _metadata_from_inputs(dataset: Any, metadata: Any, source_path: str | Path | None) -> dict[str, Any]:
    warnings: list[str] = []
    values: dict[str, Any] = {}
    for source in (metadata, dataset):
        if source is None:
            continue
        for key in ("company", "empresa", "company_name", "nombre_empresa"):
            _set_if_present(values, "company", _read_field(source, (key,)))
        for key in ("period", "periodo", "period_detected", "periodo_detectado"):
            _set_if_present(values, "period", _normalize_period(_read_field(source, (key,))))
        for key in ("source_path", "archivo_origen", "input_path"):
            _set_if_present(values, "source_path", _read_field(source, (key,)))

    if source_path is not None:
        values["source_path"] = str(source_path)
        try:
            detected = detect_source_metadata(source_path)
        except Exception as exc:  # pragma: no cover - defensive integration path
            warnings.append(f"Could not detect source metadata from {source_path}: {exc}")
        else:
            _set_if_present(values, "company", detected.get("company"))
            _set_if_present(values, "period", detected.get("period"))

    company = str(values.get("company") or "Empresa por confirmar").strip()
    period = str(values.get("period") or "Periodo por confirmar").strip()
    if company == "Empresa por confirmar":
        warnings.append("Company was not provided by parser/dataset metadata.")
    if period == "Periodo por confirmar":
        warnings.append("Period was not provided by parser/dataset metadata.")
    return {"company": company, "period": period, "source_path": values.get("source_path"), "warnings": warnings}


def _index_dataset_lines(dataset: Any, warnings: list[str]) -> dict[str, dict[str, Any]]:
    lines = _read_field(dataset, ("lines", "lineas", "items", "rows", "renglones"))
    indexed: dict[str, dict[str, Any]] = {}
    if lines is None:
        return indexed
    if isinstance(lines, dict):
        iterable: Iterable[Any] = lines.items()
    else:
        iterable = lines if isinstance(lines, Iterable) and not isinstance(lines, (str, bytes)) else []
    for raw_line in iterable:
        if isinstance(raw_line, tuple) and len(raw_line) == 2:
            line_key, raw_line = raw_line
        else:
            line_key = None
        line = _to_dict(raw_line)
        if not line and line_key is not None:
            line = {"key": line_key, "amount": raw_line}
        elif line_key is not None:
            line.setdefault("key", line_key)
        label = _first_present(line, ("key", "id", "label", "concept", "concepto", "name", "nombre"))
        if label is None:
            continue
        indexed[_slug(str(label))] = line
    return indexed


def _line_amounts(line: dict[str, Any] | None, warnings: list[str], label: str) -> dict[str, float]:
    if line is None:
        return {"period": 0.0, "accumulated": 0.0}
    period_formula = _first_present(line, ("period_formula", "formula_periodo", "formula"))
    accumulated_formula = _first_present(line, ("accumulated_formula", "formula_acumulado"))
    if period_formula or accumulated_formula:
        warnings.append(f"Dataset formulas for '{label}' were ignored; generated ER uses safe internal formulas only.")
    period = _number(_first_present(line, ("period_amount", "amount", "importe_periodo", "periodo", "value", "valor")))
    accumulated = _number(_first_present(line, ("accumulated_amount", "accumulated", "acumulado", "importe_acumulado")))
    if accumulated is None:
        accumulated = period
    return {"period": period or 0.0, "accumulated": accumulated or 0.0}


def _read_field(source: Any, keys: tuple[str, ...]) -> Any:
    if source is None:
        return None
    mapping = _to_dict(source)
    for key in keys:
        if key in mapping and mapping[key] not in (None, ""):
            return mapping[key]
    for key in keys:
        if hasattr(source, key):
            value = getattr(source, key)
            if value not in (None, ""):
                return value
    return None


def _to_dict(value: Any) -> dict[str, Any]:
    if value is None:
        return {}
    if isinstance(value, dict):
        return value
    if is_dataclass(value):
        return {field_name: getattr(value, field_name) for field_name in value.__dataclass_fields__}
    if hasattr(value, "__dict__"):
        return vars(value)
    return {}


def _first_present(mapping: dict[str, Any], keys: tuple[str, ...]) -> Any:
    for key in keys:
        value = mapping.get(key)
        if value not in (None, ""):
            return value
    return None


def _set_if_present(values: dict[str, Any], key: str, value: Any) -> None:
    if values.get(key) in (None, "") and value not in (None, ""):
        values[key] = value


def _number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text)
    except ValueError:
        return None


def _normalize_period(value: Any) -> str | None:
    if value in (None, ""):
        return None
    text = str(value).strip()
    match = re.search(r"(\d{4})[-/](\d{1,2})", text)
    if match:
        return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"
    return text


def _period_caption(period: str) -> str:
    match = re.search(r"(\d{4})-(\d{2})", period)
    if not match:
        return f"Periodo: {period}"
    year = int(match.group(1))
    month = int(match.group(2))
    last_day = calendar.monthrange(year, month)[1]
    month_name = MONTHS_ES.get(month, str(month))
    return f"Del 1ro de Enero al {last_day} de {month_name} {year}"


def _bg_period_caption(period: str) -> str:
    match = re.search(r"(\d{4})-(\d{2})", period)
    if not match:
        return f"Al cierre de {period}"
    year = int(match.group(1))
    month = int(match.group(2))
    return f"Al {calendar.monthrange(year, month)[1]} de {MONTHS_ES.get(month, month)} de {year}"


def _bal_period_caption(period: str) -> str:
    match = re.search(r"(\d{4})-(\d{2})", period)
    if not match:
        return period.upper()
    return f"{MONTHS_ES.get(int(match.group(2)), match.group(2)).upper()}.{match.group(1)}"


def _slug(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", ascii_text.lower()).strip("_")


def _coerce_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, Iterable):
        return [str(item) for item in value]
    return [str(value)]


def _coerce_list_of_dicts(value: Any) -> list[dict[str, Any]]:
    if value is None:
        return []
    if isinstance(value, Mapping):
        return [dict(value)]
    if not isinstance(value, Iterable) or isinstance(value, (str, bytes)):
        return []
    return [_to_dict(item) for item in value]


def _default_output_path(company: str, period: str) -> Path:
    company_slug = _slug(company)[:60] or "empresa"
    period_slug = _slug(period) or "periodo"
    return DEFAULT_OUTPUT_DIR / f"estados_financieros_{company_slug}_{period_slug}.xlsx"
