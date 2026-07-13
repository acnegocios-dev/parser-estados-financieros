"""Validation helpers for generated local financial-statement workbooks."""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import Any, Iterable, Mapping

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

try:
    from .engine import canonical_account_code, leaf_account_rows, to_decimal
except ImportError:  # pragma: no cover - supports direct script execution.
    from engine import canonical_account_code, leaf_account_rows, to_decimal


FORMULA_ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
CONCEPTUAL_BALANCE_REFERENCE = "BG!L47"
BALANCE_TOLERANCE_DEFAULT = 1.0


@dataclass
class FormulaIssue:
    cell: str
    value: str
    reason: str


@dataclass
class ValidationResult:
    ok: bool
    formula_static_validation: bool
    formula_static_issues: list[FormulaIssue] = field(default_factory=list)
    formula_recalculation_performed: bool = False
    formula_recalculation_engine: str = "none"
    formula_evaluated_error_count: int | None = None
    formula_cached_values_available: bool = False
    formula_evaluated_issues: list[FormulaIssue] = field(default_factory=list)
    balance_difference: float | None = None
    balance_ok: bool | None = None
    balance_reference: str = CONCEPTUAL_BALANCE_REFERENCE
    warnings: list[str] = field(default_factory=list)


@dataclass
class FormulaRecalculationResult:
    performed: bool
    engine: str
    evaluated_error_count: int | None
    cached_values_available: bool
    evaluated_issues: list[FormulaIssue] = field(default_factory=list)
    evidence_path: str | None = None
    blocked: bool = False
    warnings: list[str] = field(default_factory=list)


@dataclass
class BalanceComponent:
    rubro: str
    total: float
    cuentas: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class BalanceCheckResult:
    total_activo: float
    total_pasivo: float
    capital_contable: float
    diferencia_cuadre: float
    tolerance: float
    cuadra: bool
    balanza_no_cuadra: bool
    componentes: list[BalanceComponent] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.cuadra


def validate_generated_workbook(
    workbook_or_path: Workbook | str | Path,
    *,
    balance_difference: float | int | str | None = None,
    tolerance: float = 1.0,
    formula_mode: str = "static_only",
    evaluated_workbook_or_path: Workbook | str | Path | None = None,
    formula_recalculation_engine: str | None = None,
) -> ValidationResult:
    """Validate formulas and the conceptual balance difference.

    The balance check mirrors the manual workbook concept at BG!L47 without
    copying or requiring that sheet. Pass `balance_difference` from the parser
    or balance-sheet integration when available.

    ``static_only`` inspects formula text only. ``recalculated_ok`` and
    ``recalculated_error`` require a separate data-only workbook containing
    evaluated formula results; this prevents a static token scan from being
    presented as proof that Excel formulas were calculated.
    """

    if formula_mode not in {"static_only", "recalculated_ok", "recalculated_error"}:
        raise ValueError(f"Unsupported formula validation mode: {formula_mode}")
    if formula_mode != "static_only" and evaluated_workbook_or_path is None:
        raise ValueError("An evaluated workbook is required for recalculated validation modes.")

    workbook, close_after = _load_workbook(workbook_or_path)
    evaluated_workbook = None
    evaluated_close_after = False
    try:
        formula_issues = find_formula_issues(workbook)
        evaluated_issues: list[FormulaIssue] = []
        cached_values_available = False
        if formula_mode != "static_only":
            evaluated_workbook, evaluated_close_after = _load_workbook(
                evaluated_workbook_or_path, data_only=True
            )
            evaluated_issues, cached_values_available = find_evaluated_formula_issues(
                workbook, evaluated_workbook
            )
        coerced_difference = _coerce_float(balance_difference)
        warnings: list[str] = []
        if balance_difference is None:
            warnings.append(
                "Balance difference was not provided; conceptual BG!L47 balance check was skipped."
            )
            balance_ok = None
        elif coerced_difference is None:
            warnings.append("Balance difference could not be parsed as a number.")
            balance_ok = False
        else:
            balance_ok = abs(coerced_difference) < tolerance
        static_validation = not formula_issues
        recalculation_performed = formula_mode != "static_only"
        evaluated_error_count = len(evaluated_issues) if recalculation_performed else None
        ok = (
            static_validation
            and (evaluated_error_count in (None, 0))
            and (balance_ok is not False)
        )
        return ValidationResult(
            ok=ok,
            formula_static_validation=static_validation,
            formula_static_issues=formula_issues,
            formula_recalculation_performed=recalculation_performed,
            formula_recalculation_engine=(
                "none" if not recalculation_performed else (formula_recalculation_engine or "unknown")
            ),
            formula_evaluated_error_count=evaluated_error_count,
            formula_cached_values_available=cached_values_available,
            formula_evaluated_issues=evaluated_issues,
            balance_difference=coerced_difference,
            balance_ok=balance_ok,
            warnings=warnings,
        )
    finally:
        if close_after:
            workbook.close()
        if evaluated_close_after and evaluated_workbook is not None:
            evaluated_workbook.close()


def validate_balance_sheet(
    rows: Iterable[Mapping[str, Any] | Any],
    *,
    tolerance: float = BALANCE_TOLERANCE_DEFAULT,
    result_ejercicio: float | int | str | Decimal | None = None,
) -> BalanceCheckResult:
    """Validate the balance equation derived from normalized balance rows.

    The calculation is programmatic and does not depend on BG!L47. Assets and
    liabilities use leaf/detail accounts. When `result_ejercicio` is provided,
    capital uses top-level balances plus the generated current result, matching
    the manual BG composition; without it, the legacy leaf-only behavior is
    preserved for standalone callers.
    """

    normalized_rows = tuple(_row_to_mapping(row) for row in rows)
    leaf_rows = leaf_account_rows(normalized_rows)
    totals: dict[str, Decimal] = {"1": Decimal("0"), "2": Decimal("0"), "3": Decimal("0")}
    details: dict[str, list[dict[str, Any]]] = {"1": [], "2": [], "3": []}

    balance_rows = {
        "1": leaf_rows,
        "2": leaf_rows,
        "3": _capital_balance_rows(normalized_rows) if result_ejercicio is not None else leaf_rows,
    }
    for class_prefix, class_rows in balance_rows.items():
        for row in class_rows:
            account_code = canonical_account_code(_read_field(row, ("account_code", "top_account", "cuenta")) or "")
            if not account_code or not account_code.startswith(class_prefix):
                continue
            saldo_final = to_decimal(_read_field(row, ("saldo_final", "saldoFinal", "saldo")))
            if class_prefix == "1" and "dep acum" in str(
                _read_field(row, ("account_name", "nombre")) or ""
            ).casefold():
                saldo_final = -saldo_final
            totals[class_prefix] += saldo_final
            details[class_prefix].append(
                {
                    "source_row": _read_field(row, ("source_row",)),
                    "account_code": account_code,
                    "account_name": _read_field(row, ("account_name", "nombre")) or "",
                    "saldo_final": _quantize_money(saldo_final),
                }
            )

    if result_ejercicio is not None:
        resultado = to_decimal(result_ejercicio)
        totals["3"] += resultado
        details["3"].append(
            {
                "source_row": None,
                "account_code": "resultado_ejercicio",
                "account_name": "Resultado del ejercicio generado por ER",
                "saldo_final": _quantize_money(resultado),
            }
        )

    total_activo = totals["1"]
    total_pasivo = totals["2"]
    capital_contable = totals["3"]
    diferencia = total_activo - (total_pasivo + capital_contable)
    cuadra = abs(diferencia) < Decimal(str(tolerance))

    return BalanceCheckResult(
        total_activo=_quantize_money(total_activo),
        total_pasivo=_quantize_money(total_pasivo),
        capital_contable=_quantize_money(capital_contable),
        diferencia_cuadre=_quantize_balance_difference(diferencia),
        tolerance=float(tolerance),
        cuadra=cuadra,
        balanza_no_cuadra=not cuadra,
        componentes=[
            BalanceComponent(rubro="activo", total=_quantize_money(total_activo), cuentas=details["1"]),
            BalanceComponent(rubro="pasivo", total=_quantize_money(total_pasivo), cuentas=details["2"]),
            BalanceComponent(
                rubro="capital_contable",
                total=_quantize_money(capital_contable),
                cuentas=details["3"],
            ),
        ],
    )


def _capital_balance_rows(rows: Iterable[Mapping[str, Any]]) -> tuple[Mapping[str, Any], ...]:
    """Select top-level capital balances, falling back to leaves when needed."""

    materialized = tuple(rows)
    capital_rows = []
    capital_codes = {
        canonical_account_code(_read_field(row, ("account_code", "top_account", "cuenta")) or "")
        for row in materialized
    }
    top_level_codes = {
        code for code in capital_codes if code.startswith("3") and "-" not in code
    }
    for row in materialized:
        code = canonical_account_code(_read_field(row, ("account_code", "top_account", "cuenta")) or "")
        if not code.startswith("3"):
            continue
        top_code = code.split("-", 1)[0]
        if top_code in top_level_codes:
            if code == top_code:
                capital_rows.append(row)
            continue
        if not any(
            other != code and other.startswith(f"{code}-")
            for other in capital_codes
        ):
            capital_rows.append(row)
    return tuple(capital_rows)


def find_formula_issues(workbook: Workbook) -> list[FormulaIssue]:
    """Return formula cells that contain known Excel error tokens or externals."""

    issues: list[FormulaIssue] = []
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                value = cell.value
                if not (isinstance(value, str) and value.startswith("=")):
                    continue
                upper = value.upper()
                for token in FORMULA_ERROR_TOKENS:
                    if token in upper:
                        issues.append(
                            FormulaIssue(
                                cell=f"{worksheet.title}!{cell.coordinate}",
                                value=value,
                                reason=f"Formula contains {token}",
                            )
                        )
                        break
                else:
                    if "[" in value or "]" in value:
                        issues.append(
                            FormulaIssue(
                                cell=f"{worksheet.title}!{cell.coordinate}",
                                value=value,
                                reason="Formula contains an external workbook reference",
                            )
                        )
    return issues


def find_evaluated_formula_issues(
    formula_workbook: Workbook,
    evaluated_workbook: Workbook,
) -> tuple[list[FormulaIssue], bool]:
    """Inspect cached/evaluated values for the formula cells in a workbook."""

    issues: list[FormulaIssue] = []
    formula_cells = []
    for worksheet in formula_workbook.worksheets:
        evaluated_sheet = evaluated_workbook[worksheet.title]
        for row in worksheet.iter_rows():
            for cell in row:
                if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                    continue
                formula_cells.append((worksheet.title, cell.coordinate))
                evaluated_value = evaluated_sheet[cell.coordinate].value
                if isinstance(evaluated_value, str):
                    upper = evaluated_value.upper()
                    for token in FORMULA_ERROR_TOKENS:
                        if token in upper:
                            issues.append(
                                FormulaIssue(
                                    cell=f"{worksheet.title}!{cell.coordinate}",
                                    value=evaluated_value,
                                    reason=f"Evaluated formula contains {token}",
                                )
                            )
                            break
    cached_values_available = bool(formula_cells) and all(
        evaluated_workbook[sheet][coordinate].value is not None
        for sheet, coordinate in formula_cells
    )
    return issues, cached_values_available


def recalculate_workbook(
    workbook_path: str | Path,
    *,
    balance_difference: float | int | str | None = None,
    tolerance: float = 1.0,
) -> FormulaRecalculationResult:
    """Recalculate a copy with LibreOffice when available.

    The final workbook is replaced atomically only after evaluated values are
    available and contain no formula errors. Failed recalculations leave a
    sibling evidence copy and mark the result as blocked.
    """

    source = Path(workbook_path)
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if executable is None:
        return FormulaRecalculationResult(
            performed=False,
            engine="none",
            evaluated_error_count=None,
            cached_values_available=False,
            warnings=[
                "Formula recalculation was not performed: libreoffice/soffice is unavailable."
            ],
        )

    with tempfile.TemporaryDirectory(prefix="estados_financieros_recalc_") as temp_dir:
        temp_root = Path(temp_dir)
        input_copy = temp_root / source.name
        output_dir = temp_root / "output"
        output_dir.mkdir()
        shutil.copy2(source, input_copy)
        completed = subprocess.run(
            [executable, "--headless", "--convert-to", "xlsx", "--outdir", str(output_dir), str(input_copy)],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        evaluated_path = output_dir / input_copy.name
        if completed.returncode != 0 or not evaluated_path.exists():
            return FormulaRecalculationResult(
                performed=False,
                engine=Path(executable).name,
                evaluated_error_count=None,
                cached_values_available=False,
                blocked=True,
                warnings=[
                    "LibreOffice recalculation failed before producing an evaluated workbook.",
                    completed.stderr.strip() or completed.stdout.strip(),
                ],
            )

        evaluated = validate_generated_workbook(
            source,
            balance_difference=balance_difference,
            tolerance=tolerance,
            formula_mode="recalculated_ok",
            evaluated_workbook_or_path=evaluated_path,
            formula_recalculation_engine=Path(executable).name,
        )
        if not evaluated.formula_cached_values_available or not evaluated.ok:
            evidence = source.with_name(
                f"{source.stem}.recalculation-failed-{os.getpid()}.xlsx"
            )
            shutil.copy2(evaluated_path, evidence)
            return FormulaRecalculationResult(
                performed=True,
                engine=Path(executable).name,
                evaluated_error_count=evaluated.formula_evaluated_error_count,
                cached_values_available=evaluated.formula_cached_values_available,
                evaluated_issues=evaluated.formula_evaluated_issues,
                evidence_path=str(evidence),
                blocked=True,
                warnings=[
                    "Evaluated workbook failed validation; final output was not replaced.",
                ],
            )

        os.replace(evaluated_path, source)
        return FormulaRecalculationResult(
            performed=True,
            engine=Path(executable).name,
            evaluated_error_count=0,
            cached_values_available=True,
        )


def validate_balance_difference(
    difference: float | int | str,
    *,
    tolerance: float = 1.0,
) -> bool:
    """Return True when abs(difference) is strictly less than tolerance."""

    numeric = _coerce_float(difference)
    if numeric is None:
        return False
    return abs(numeric) < tolerance


def assert_generated_workbook_valid(
    workbook_or_path: Workbook | str | Path,
    *,
    balance_difference: float | int | str | None = None,
    tolerance: float = 1.0,
) -> ValidationResult:
    """Validate and raise ValueError with compact details if checks fail."""

    result = validate_generated_workbook(
        workbook_or_path,
        balance_difference=balance_difference,
        tolerance=tolerance,
    )
    if result.ok:
        return result
    details = [f"{issue.cell}: {issue.reason}" for issue in result.formula_static_issues]
    if result.balance_ok is False:
        details.append(
            f"Balance difference {result.balance_difference!r} does not satisfy abs(diff) < {tolerance}"
        )
    raise ValueError("; ".join(details) or "Generated workbook validation failed")


def _load_workbook(
    workbook_or_path: Workbook | str | Path | None,
    *,
    data_only: bool = False,
) -> tuple[Workbook, bool]:
    if workbook_or_path is None:
        raise ValueError("Workbook path is required")
    if isinstance(workbook_or_path, Workbook):
        return workbook_or_path, False
    return load_workbook(Path(workbook_or_path), data_only=data_only), True


def _coerce_float(value: Any) -> float | None:
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


def _read_field(source: Any, keys: tuple[str, ...]) -> Any:
    if source is None:
        return None
    if isinstance(source, Mapping):
        for key in keys:
            if key in source and source[key] not in (None, ""):
                return source[key]
    for key in keys:
        if hasattr(source, key):
            value = getattr(source, key)
            if value not in (None, ""):
                return value
    return None


def _row_to_mapping(row: Any) -> dict[str, Any]:
    if isinstance(row, Mapping):
        return dict(row)
    if hasattr(row, "to_dict"):
        return dict(row.to_dict())
    if hasattr(row, "__dict__"):
        return dict(vars(row))
    return {}


def _quantize_money(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _quantize_balance_difference(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP))
