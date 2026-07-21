from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass, replace
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

try:
    from .account_catalog import BLOCKING, parse_account_catalog
    from .accounting_profiles import CatalogIdentity, AccountingProfile, load_accounting_profiles, select_profile_for_runtime
    from .engine import ProfileCoverageError, build_bal_dataset, build_bg_dataset, build_er_dataset
    from .parser import BalanzaRow, enrich_balanza_rows, parse_balanza
    from .runtime_metadata import build_runtime_metadata, sha256_file
    from .validation import recalculate_workbook, validate_balance_sheet, validate_generated_workbook
    from .workbook import save_financial_statements_workbook
except ImportError:  # pragma: no cover - supports direct script execution.
    from account_catalog import BLOCKING, parse_account_catalog
    from accounting_profiles import CatalogIdentity, AccountingProfile, load_accounting_profiles, select_profile_for_runtime
    from engine import ProfileCoverageError, build_bal_dataset, build_bg_dataset, build_er_dataset
    from parser import BalanzaRow, enrich_balanza_rows, parse_balanza
    from runtime_metadata import build_runtime_metadata, sha256_file
    from validation import recalculate_workbook, validate_balance_sheet, validate_generated_workbook
    from workbook import save_financial_statements_workbook


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"
OUTPUT_DIR = ROOT / "sample-outputs"
PROFILE_DIR = ROOT / "src" / "profiles"


class RuntimeGenerationError(ValueError):
    """A safe, machine-readable generation rejection for the HTTP boundary."""

    def __init__(self, code: str, *, details: list[dict[str, Any]] | None = None):
        self.code = code
        self.details = details or []
        super().__init__(code)


def run_prototype(
    input_path: str | Path = DEFAULT_INPUT,
    *,
    output_dir: str | Path | None = None,
    catalog_path: str | Path | None = None,
    profiles: tuple[AccountingProfile, ...] | None = None,
) -> dict[str, Any]:
    """Generate the three-sheet workbook from one uploaded balanza.

    ``output_dir`` lets request handlers and tests keep generated artifacts in
    their own temporary directory.  The legacy default remains available for
    the command-line prototype only.
    """
    source = Path(input_path)
    target_dir = Path(output_dir) if output_dir is not None else OUTPUT_DIR
    parsed = parse_balanza(source)
    profile: AccountingProfile | None = None
    catalog_identity: CatalogIdentity | None = None
    if catalog_path is not None:
        catalog = parse_account_catalog(catalog_path)
        if not catalog.is_valid or catalog.semantic_sha256 is None:
            raise RuntimeGenerationError(
                "catalog_validation_failed",
                details=[issue.to_dict() for issue in catalog.issues if issue.severity == BLOCKING],
            )
        catalog_identity = CatalogIdentity(catalog.source_sha256, catalog.semantic_sha256)
        try:
            profile = select_profile_for_runtime(
                profiles if profiles is not None else load_accounting_profiles(PROFILE_DIR),
                rfc=parsed.period.rfc,
                as_of=date(parsed.period.period_year, parsed.period.period_month, 1),
                catalog_identity=catalog_identity,
            )
        except Exception as exc:
            issues = getattr(exc, "issues", ())
            code = next((issue.code for issue in issues if issue.code in {
                "profile_not_found", "profile_not_approved", "catalog_hash_mismatch", "ambiguous_mapping",
            }), "profile_not_found")
            raise RuntimeGenerationError(code, details=[issue.to_dict() for issue in issues]) from exc
        parsed = replace(parsed, rows=enrich_balanza_rows(parsed.rows, catalog.rows))
    # The engine owns leaf selection so catalog parent_code can take
    # precedence over legacy code-prefix hierarchy.
    rows = [row.to_dict() for row in parsed.rows]
    try:
        engine_result = build_er_dataset(
            rows,
            company=parsed.company_name,
            period=parsed.period.period_ym,
            source_path=parsed.source_path,
            profile=profile,
            enforce_profile_coverage=True,
        )

        bg_dataset = build_bg_dataset(
            parsed.rows,
            result_ejercicio=engine_result["raw_amounts"]["resultado_ejercicio"],
            company=parsed.company_name,
            period=parsed.period.period_ym,
            source_path=parsed.source_path,
            profile=profile,
            enforce_profile_coverage=True,
        )
    except ProfileCoverageError as exc:
        blocker_codes = {item.get("code") for item in exc.blockers}
        code = (
            "unmapped_material_accounts" if "profile_mapping_unassigned_material" in blocker_codes
            else "ambiguous_mapping" if "profile_mapping_duplicate" in blocker_codes
            else "coverage_mismatch"
        )
        raise RuntimeGenerationError(code, details=list(exc.blockers)) from exc
    bal_dataset = build_bal_dataset(parsed.rows)
    workbook_result = save_financial_statements_workbook(
        engine_result,
        output_path=_output_path(parsed, target_dir),
        bg_dataset=bg_dataset,
        bal_dataset=bal_dataset,
        metadata={
            "company": parsed.company_name,
            "period": parsed.period.period_ym,
            "rfc": parsed.period.rfc,
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
            "leaf_rows_used_for_calculation": engine_result["source_rows_used"],
            "empty_rows": list(parsed.empty_rows),
            "structure_issues": [issue.to_dict() for issue in parsed.structure_issues],
            "warnings": [warning.to_dict() for warning in parsed.warnings],
        },
        "engine": {
            "statement_lines": engine_result["statement_lines"],
            "unmatched_accounts_count": len(engine_result["unmatched_accounts"]),
            "unmatched_accounts": engine_result["unmatched_accounts"],
            "warnings": engine_result["warnings"],
            "coverage": engine_result["coverage"],
            "formulas": engine_result["formulas"],
            "sign_policy": engine_result["sign_policy"],
            "profile": engine_result["profile"],
        },
        "workbook": {
            "sheet_names": list(workbook_result.workbook.sheetnames),
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
    report_path = _report_path(parsed, target_dir)
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


def _output_path(parsed: Any, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    company = (parsed.company_name or "empresa").lower()
    slug = "".join(ch if ch.isalnum() else "_" for ch in company)
    slug = "_".join(part for part in slug.split("_") if part)[:60] or "empresa"
    period = parsed.period.period_ym.replace("-", "_")
    return output_dir / f"estados_financieros_{slug}_{period}.xlsx"


def _report_path(parsed: Any, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    company = (parsed.company_name or "empresa").lower()
    slug = "".join(ch if ch.isalnum() else "_" for ch in company)
    slug = "_".join(part for part in slug.split("_") if part)[:60] or "empresa"
    return output_dir / f"validation_report_{slug}_{parsed.period.period_ym.replace('-', '_')}.json"


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
