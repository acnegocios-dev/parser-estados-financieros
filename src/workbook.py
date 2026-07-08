"""Workbook generation for the local financial statements prototype.

This module owns the generated ER sheet only. It accepts loose dict/dataclass
inputs so it can be wired to parser and ER-dataset modules created elsewhere.
"""

from __future__ import annotations

import calendar
from dataclasses import dataclass, field, is_dataclass
from datetime import datetime, timezone
from pathlib import Path
import re
import unicodedata
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_DIR = ROOT / "sample-outputs"

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
    {"row": 58, "label": "TOTAL OTROS INGRESOS", "kind": "subtotal", "formula": "{col}56-{col}57"},
    {"row": 60, "label": "RES. INT. DE FINANCIAMIENTO", "kind": "section"},
    {"row": 61, "label": "Productos Financieros", "key": "productos_financieros"},
    {"row": 62, "label": "Gastos Financieros", "key": "gastos_financieros"},
    {"row": 63, "label": "TOTAL R. I. F.", "kind": "subtotal", "formula": "{col}61-{col}62"},
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
    """Build an ER workbook from a loose ER dataset.

    `dataset` may be a dict, dataclass, or object with `lines`/`lineas`.
    Base lines can provide values through `period_amount`, `amount`,
    `accumulated_amount`, or `acumulado`. Subtotals are generated as safe
    internal formulas.
    """

    metadata_values = _metadata_from_inputs(dataset, metadata, source_path)
    company = metadata_values["company"]
    period = metadata_values["period"]
    warnings = list(metadata_values["warnings"])
    missing_accounts = _coerce_list(_read_field(dataset, ("missing_accounts", "cuentas_no_encontradas")))
    missing_accounts.extend(_coerce_list(_read_field(dataset, ("missing", "not_found"))))

    wb = Workbook()
    ws = wb.active
    ws.title = "ER"
    ws.sheet_view.showGridLines = False

    _write_header(ws, company, period)
    line_values = _index_dataset_lines(dataset, warnings)
    formula_cells = _write_er_body(ws, line_values, warnings)
    _style_er_sheet(ws)

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


def save_er_workbook(
    dataset: Any | None = None,
    output_path: str | Path | None = None,
    *,
    metadata: Any | None = None,
    source_path: str | Path | None = None,
) -> WorkbookBuildResult:
    """Build and save the ER workbook.

    If `output_path` is omitted, the file is written under `sample-outputs`.
    """

    result = build_er_workbook(dataset, metadata=metadata, source_path=source_path)
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
    ws["D15"] = "Del Periodo"
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
            for column in ("D", "H"):
                formula = "=" + str(spec["formula"]).format(col=column)
                _set_formula(ws, f"{column}{row}", formula, formula_cells)
            continue

        values = _line_amounts(line_values.get(str(spec.get("key"))), warnings, label)
        ws[f"D{row}"] = values["period"]
        ws[f"H{row}"] = values["accumulated"]

    for row in numeric_rows:
        _set_formula(ws, f"F{row}", f'=IF($D$18=0,0,D{row}/$D$18)', formula_cells)
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
    widths = {"A": 3, "B": 36, "C": 3, "D": 15, "E": 3, "F": 10, "G": 3, "H": 15, "I": 3, "J": 10}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row in range(1, 85):
        ws.row_dimensions[row].height = 18

    title_font = Font(name="Calibri", size=12, bold=True)
    subtitle_font = Font(name="Calibri", size=11)
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    section_fill = PatternFill("solid", fgColor="EDEDED")
    subtotal_fill = PatternFill("solid", fgColor="D9EAD3")
    thin = Side(style="thin", color="B7B7B7")
    border = Border(bottom=thin)

    for cell in ("B9", "B10", "B11", "B12"):
        ws[cell].alignment = Alignment(horizontal="center")
    ws["B9"].font = title_font
    ws["B10"].font = title_font
    ws["B11"].font = subtitle_font
    ws["B12"].font = Font(name="Calibri", size=10, italic=True)
    for row in range(9, 13):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=10)

    for cell in ("D15", "F15", "H15", "J15"):
        ws[cell].font = Font(bold=True)
        ws[cell].alignment = Alignment(horizontal="center")
        ws[cell].fill = header_fill
        ws[cell].border = border

    subtotal_rows = {int(spec["row"]) for spec in ER_LAYOUT if spec.get("kind") == "subtotal"}
    section_rows = {int(spec["row"]) for spec in ER_LAYOUT if spec.get("kind") == "section"}
    for spec in ER_LAYOUT:
        row = int(spec["row"])
        ws[f"B{row}"].alignment = Alignment(horizontal="left")
        if row in section_rows:
            for col in range(2, 11):
                ws.cell(row=row, column=col).font = Font(bold=True)
                ws.cell(row=row, column=col).fill = section_fill
        if row in subtotal_rows:
            for col in range(2, 11):
                ws.cell(row=row, column=col).font = Font(bold=True)
                ws.cell(row=row, column=col).fill = subtotal_fill
                ws.cell(row=row, column=col).border = border

    for row in range(18, 71):
        for cell in (f"D{row}", f"H{row}"):
            ws[cell].number_format = '#,##0.00;[Red](#,##0.00);"-"'
        for cell in (f"F{row}", f"J{row}"):
            ws[cell].number_format = "0.00%"

    ws.freeze_panes = "A16"


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


def _default_output_path(company: str, period: str) -> Path:
    company_slug = _slug(company)[:60] or "empresa"
    period_slug = _slug(period) or "periodo"
    return DEFAULT_OUTPUT_DIR / f"estado_resultados_{company_slug}_{period_slug}.xlsx"
