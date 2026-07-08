from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    from .engine import build_income_statement
    from .parser import BalanzaRow, ParsedBalanza, parse_balanza
    from .validation import validate_generated_workbook
    from .workbook import save_er_workbook
except ImportError:  # pragma: no cover - supports direct script execution.
    from engine import build_income_statement
    from parser import BalanzaRow, ParsedBalanza, parse_balanza
    from validation import validate_generated_workbook
    from workbook import save_er_workbook


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"
OUTPUT_DIR = ROOT / "sample-outputs"

ENGINE_TO_ER_LAYOUT = {
    "ventas_ingresos_netos": "ingresos_por_servicios",
    "costo_ventas": "costo_de_ventas",
    "gastos_operacion": "varios",
    "otros_ingresos": "otros_productos",
    "otros_gastos": "otros_gastos",
    "productos_financieros": "productos_financieros",
    "gastos_financieros": "gastos_financieros",
    "impuestos": "isr_del_ejercicio",
}

ER_ACCOUNT_PREFIX_TO_LINE = (
    ("4110-9999", "otros_productos"),
    ("4110", "ingresos_por_servicios"),
    ("5110", "costo_de_ventas"),
    ("6110", "sueldos_y_salarios"),
    ("6120", "impuestos_y_derechos"),
    ("6125", "honorarios"),
    ("6130", "arrendamiento"),
    ("6135", "seguros_y_fianzas"),
    ("6140", "servicios"),
    ("6144", "capacitacion_al_personal"),
    ("6145", "fletes_y_o_mensajeria"),
    ("6146", "fletes_y_o_mensajeria"),
    ("6148", "seguridad_e_higiene"),
    ("6150", "mantenimiento"),
    ("6155", "propaganda_y_publicidad"),
    ("6160", "combustibles"),
    ("6165", "papeleria_y_art_de_oficina"),
    ("6170", "depreciaciones"),
    ("6175", "recargos"),
    ("6180", "cuotas_y_suscripciones"),
    ("6190", "varios"),
    ("6200", "no_deducibles"),
    ("700", "otros_productos"),
    ("701", "otros_productos"),
    ("702", "otros_productos"),
    ("71", "otros_productos"),
    ("703", "otros_gastos"),
    ("704", "otros_gastos"),
    ("705", "otros_gastos"),
    ("72", "otros_gastos"),
    ("730", "productos_financieros"),
    ("731", "productos_financieros"),
    ("732", "productos_financieros"),
    ("74", "productos_financieros"),
    ("733", "gastos_financieros"),
    ("734", "gastos_financieros"),
    ("735", "gastos_financieros"),
    ("75", "gastos_financieros"),
    ("76", "gastos_financieros"),
    ("77", "isr_del_ejercicio"),
    ("78", "isr_del_ejercicio"),
    ("79", "isr_del_ejercicio"),
)

EXPENSE_DETAIL_KEYS = {
    "sueldos_y_salarios",
    "impuestos_y_derechos",
    "honorarios",
    "arrendamiento",
    "seguros_y_fianzas",
    "servicios",
    "capacitacion_al_personal",
    "fletes_y_o_mensajeria",
    "seguridad_e_higiene",
    "mantenimiento",
    "propaganda_y_publicidad",
    "combustibles",
    "cuotas_y_suscripciones",
    "papeleria_y_art_de_oficina",
    "depreciaciones",
    "recargos",
    "varios",
    "uniformes",
    "no_deducibles",
}


def run_prototype(input_path: str | Path = DEFAULT_INPUT) -> dict[str, Any]:
    source = Path(input_path)
    parsed = parse_balanza(source)
    leaf_rows = _leaf_rows(parsed.rows)
    rows = [row.to_dict() for row in leaf_rows]
    engine_result = build_income_statement(rows)
    workbook_dataset = _workbook_dataset(parsed, engine_result, leaf_rows)

    workbook_result = save_er_workbook(
        workbook_dataset,
        metadata={
            "company": parsed.company_name,
            "period": parsed.period.period_ym,
            "source_path": str(source),
        },
        source_path=source,
    )
    balance_difference = calculate_balance_difference(parsed.rows)
    validation = validate_generated_workbook(
        workbook_result.output_path,
        balance_difference=balance_difference,
        tolerance=1.0,
    )

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source_path": str(source),
        "output_xlsx": str(workbook_result.output_path),
        "period": parsed.period.to_dict(),
        "company_name": parsed.company_name,
        "content_period_ym": parsed.content_period_ym,
        "parser": {
            "sheet_name": parsed.sheet_name,
            "header_row": parsed.header_row,
            "normalized_rows": len(parsed.rows),
            "leaf_rows_used_for_calculation": len(leaf_rows),
            "empty_rows": list(parsed.empty_rows),
            "structure_issues": [issue.to_dict() for issue in parsed.structure_issues],
        },
        "engine": {
            "statement_lines": engine_result["statement_lines"],
            "unmatched_accounts_count": len(engine_result["unmatched_accounts"]),
            "unmatched_accounts": engine_result["unmatched_accounts"],
            "formulas": engine_result["formulas"],
            "sign_policy": engine_result["sign_policy"],
        },
        "workbook": {
            "formula_cells": workbook_result.formula_cells,
            "formula_cells_count": len(workbook_result.formula_cells),
            "warnings": workbook_result.warnings,
            "missing_accounts": workbook_result.missing_accounts,
        },
        "validation": _validation_to_dict(validation),
        "balance_check": {
            "method": "leaf_account_debe_minus_haber",
            "difference": balance_difference,
            "tolerance": 1.0,
            "reference": validation.balance_reference,
        },
        "notes": [
            "El mapeo contable es provisional por prefijos hasta contar con catalogo definitivo.",
            "El XLSX generado no copia formulas #REF! ni referencias externas del manual.",
        ],
    }
    report_path = _report_path(parsed)
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    report["report_path"] = str(report_path)
    return report


def calculate_balance_difference(rows: tuple[BalanzaRow, ...]) -> float:
    leaf_codes = _leaf_account_codes(rows)
    difference = sum(
        float(row.debe - row.haber)
        for row in rows
        if row.account_code in leaf_codes
    )
    return round(difference, 2)


def _workbook_dataset(
    parsed: ParsedBalanza,
    engine_result: dict[str, Any],
    leaf_rows: tuple[BalanzaRow, ...],
) -> dict[str, Any]:
    amount_by_er_key = _er_amounts_from_leaf_rows(leaf_rows)
    amount_by_engine_key = {
        line["line_key"]: float(line["amount"])
        for line in engine_result["statement_lines"]
    }
    lines_by_key = dict(amount_by_er_key)
    for engine_key, er_key in ENGINE_TO_ER_LAYOUT.items():
        if er_key not in lines_by_key and not _engine_key_has_detail(engine_key, lines_by_key):
            lines_by_key[er_key] = amount_by_engine_key.get(engine_key, 0.0)
    lines = [
        {
            "key": er_key,
            "period_amount": amount,
            "accumulated_amount": amount,
        }
        for er_key, amount in lines_by_key.items()
    ]
    return {
        "company": parsed.company_name,
        "period": parsed.period.period_ym,
        "source_path": parsed.source_path,
        "lines": lines,
        "missing_accounts": [
            f"{account['account_code']} {account.get('account_name') or ''}".strip()
            for account in engine_result["unmatched_accounts"]
        ],
    }


def _er_amounts_from_leaf_rows(rows: tuple[BalanzaRow, ...]) -> dict[str, float]:
    amounts: dict[str, float] = {}
    for row in rows:
        key = _line_key_for_account(row.account_code)
        if key is None:
            continue
        amounts[key] = round(amounts.get(key, 0.0) + float(row.saldo_final), 2)
    return amounts


def _line_key_for_account(account_code: str) -> str | None:
    for prefix, line_key in ER_ACCOUNT_PREFIX_TO_LINE:
        if account_code.startswith(prefix):
            return line_key
    if account_code.startswith("6"):
        return "varios"
    if account_code.startswith("7"):
        return "otros_gastos"
    return None


def _engine_key_has_detail(engine_key: str, lines_by_key: dict[str, float]) -> bool:
    if engine_key == "gastos_operacion":
        return any(key in lines_by_key for key in EXPENSE_DETAIL_KEYS)
    return False


def _leaf_rows(rows: tuple[BalanzaRow, ...]) -> tuple[BalanzaRow, ...]:
    leaf_codes = _leaf_account_codes(rows)
    return tuple(row for row in rows if row.account_code in leaf_codes)


def _leaf_account_codes(rows: tuple[BalanzaRow, ...]) -> set[str]:
    codes = {row.account_code for row in rows}
    return {
        code
        for code in codes
        if not any(other != code and other.startswith(f"{code}-") for other in codes)
    }


def _validation_to_dict(validation: Any) -> dict[str, Any]:
    if is_dataclass(validation):
        data = asdict(validation)
    else:
        data = dict(validation)
    data["formula_issues"] = [
        asdict(issue) if is_dataclass(issue) else issue
        for issue in data.get("formula_issues", [])
    ]
    return data


def _report_path(parsed: ParsedBalanza) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    company = (parsed.company_name or "empresa").lower()
    slug = "".join(ch if ch.isalnum() else "_" for ch in company)
    slug = "_".join(part for part in slug.split("_") if part)[:60] or "empresa"
    return OUTPUT_DIR / f"validation_report_{slug}_{parsed.period.period_ym.replace('-', '_')}.json"


if __name__ == "__main__":
    result = run_prototype()
    print(json.dumps({
        "output_xlsx": result["output_xlsx"],
        "report_path": result["report_path"],
        "period_ym": result["period"]["period_ym"],
        "normalized_rows": result["parser"]["normalized_rows"],
        "formula_errors": result["validation"]["formula_error_count"],
        "balance_difference": result["balance_check"]["difference"],
        "validation_ok": result["validation"]["ok"],
    }, indent=2, ensure_ascii=False))
