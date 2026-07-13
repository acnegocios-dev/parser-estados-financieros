from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    from .engine import build_er_dataset
    from .parser import BalanzaRow, parse_balanza
    from .runtime_metadata import build_runtime_metadata, sha256_file
    from .validation import recalculate_workbook, validate_balance_sheet, validate_generated_workbook
    from .workbook import save_er_workbook
except ImportError:  # pragma: no cover - supports direct script execution.
    from engine import build_er_dataset
    from parser import BalanzaRow, parse_balanza
    from runtime_metadata import build_runtime_metadata, sha256_file
    from validation import recalculate_workbook, validate_balance_sheet, validate_generated_workbook
    from workbook import save_er_workbook


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"
OUTPUT_DIR = ROOT / "sample-outputs"


def run_prototype(input_path: str | Path = DEFAULT_INPUT) -> dict[str, Any]:
    source = Path(input_path)
    parsed = parse_balanza(source)
    leaf_rows = _leaf_rows(parsed.rows)
    rows = [row.to_dict() for row in leaf_rows]
    engine_result = build_er_dataset(
        rows,
        company=parsed.company_name,
        period=parsed.period.period_ym,
        source_path=parsed.source_path,
    )

    workbook_result = save_er_workbook(
        engine_result,
        metadata={
            "company": parsed.company_name,
            "period": parsed.period.period_ym,
            "source_path": str(source),
        },
        source_path=source,
    )
    result_ejercicio = engine_result["raw_amounts"]["resultado_ejercicio"]
    balance_check = validate_balance_sheet(
        parsed.rows,
        tolerance=1.0,
        result_ejercicio=result_ejercicio,
    )
    balance_difference = balance_check.diferencia_cuadre
    validation = validate_generated_workbook(
        workbook_result.output_path,
        balance_difference=balance_difference,
        tolerance=1.0,
    )
    recalculation = recalculate_workbook(
        workbook_result.output_path,
        balance_difference=balance_difference,
        tolerance=1.0,
    )
    if recalculation.blocked:
        evidence = f" Evidence: {recalculation.evidence_path}." if recalculation.evidence_path else ""
        raise RuntimeError(
            "Workbook download blocked because formula recalculation failed." + evidence
        )
    validation = replace(
        validation,
        formula_recalculation_performed=recalculation.performed,
        formula_recalculation_engine=recalculation.engine,
        formula_evaluated_error_count=recalculation.evaluated_error_count,
        formula_cached_values_available=recalculation.cached_values_available,
        warnings=[*validation.warnings, *recalculation.warnings],
    )
    generated_at = datetime.now(timezone.utc).isoformat()
    runtime = build_runtime_metadata(
        generated_at=generated_at,
        output_sha256=sha256_file(workbook_result.output_path),
        formula_static_validation=validation.formula_static_validation,
        formula_recalculation_performed=validation.formula_recalculation_performed,
        formula_evaluated_error_count=validation.formula_evaluated_error_count,
        formula_cached_values_available=validation.formula_cached_values_available,
    )

    report = {
        "generated_at": generated_at,
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
            "warnings": [warning.to_dict() for warning in parsed.warnings],
        },
        "engine": {
            "statement_lines": engine_result["statement_lines"],
            "unmatched_accounts_count": len(engine_result["unmatched_accounts"]),
            "unmatched_accounts": engine_result["unmatched_accounts"],
            "warnings": engine_result["warnings"],
            "formulas": engine_result["formulas"],
            "sign_policy": engine_result["sign_policy"],
        },
        "workbook": {
            "formula_cells": workbook_result.formula_cells,
            "formula_cells_count": len(workbook_result.formula_cells),
            "warnings": workbook_result.warnings,
            "missing_accounts": workbook_result.missing_accounts,
        },
        "balance_validation": _balance_check_to_dict(balance_check),
        "validation": _validation_to_dict(validation),
        "runtime": runtime,
        **runtime,
        "balance_check": {
            "method": "programmatic_balance_sheet",
            "difference_cuadre": balance_check.diferencia_cuadre,
            "tolerance": balance_check.tolerance,
            "cuadra": balance_check.cuadra,
            "balanza_no_cuadra": balance_check.balanza_no_cuadra,
            "total_activo": balance_check.total_activo,
            "total_pasivo": balance_check.total_pasivo,
            "capital_contable": balance_check.capital_contable,
            "componentes": _balance_components_to_dict(balance_check),
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
    for key in ("formula_static_issues", "formula_evaluated_issues"):
        data[key] = [
            asdict(issue) if is_dataclass(issue) else issue
            for issue in data.get(key, [])
        ]
    return data


def _report_path(parsed: ParsedBalanza) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    company = (parsed.company_name or "empresa").lower()
    slug = "".join(ch if ch.isalnum() else "_" for ch in company)
    slug = "_".join(part for part in slug.split("_") if part)[:60] or "empresa"
    return OUTPUT_DIR / f"validation_report_{slug}_{parsed.period.period_ym.replace('-', '_')}.json"


def _balance_check_to_dict(balance_check: Any) -> dict[str, Any]:
    return {
        "total_activo": balance_check.total_activo,
        "total_pasivo": balance_check.total_pasivo,
        "capital_contable": balance_check.capital_contable,
        "difference_cuadre": balance_check.diferencia_cuadre,
        "tolerance": balance_check.tolerance,
        "cuadra": balance_check.cuadra,
        "balanza_no_cuadra": balance_check.balanza_no_cuadra,
        "componentes": _balance_components_to_dict(balance_check),
    }


def _balance_components_to_dict(balance_check: Any) -> list[dict[str, Any]]:
    return [
        {
            "rubro": component.rubro,
            "total": component.total,
            "cuentas": component.cuentas,
        }
        for component in balance_check.componentes
    ]


if __name__ == "__main__":
    result = run_prototype()
    print(json.dumps({
        "output_xlsx": result["output_xlsx"],
        "report_path": result["report_path"],
        "period_ym": result["period"]["period_ym"],
        "normalized_rows": result["parser"]["normalized_rows"],
        "formula_static_validation": result["validation"]["formula_static_validation"],
        "formula_recalculation_performed": result["validation"]["formula_recalculation_performed"],
        "formula_recalculation_engine": result["validation"]["formula_recalculation_engine"],
        "formula_evaluated_error_count": result["validation"]["formula_evaluated_error_count"],
        "formula_cached_values_available": result["validation"]["formula_cached_values_available"],
        "difference_cuadre": result["balance_check"]["difference_cuadre"],
        "cuadra": result["balance_check"]["cuadra"],
        "balanza_no_cuadra": result["balance_check"]["balanza_no_cuadra"],
        "validation_ok": result["validation"]["ok"],
    }, indent=2, ensure_ascii=False))
