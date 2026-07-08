"""Validation helpers for generated local financial-statement workbooks."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


FORMULA_ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")
CONCEPTUAL_BALANCE_REFERENCE = "BG!L47"


@dataclass
class FormulaIssue:
    cell: str
    value: str
    reason: str


@dataclass
class ValidationResult:
    ok: bool
    formula_error_count: int
    formula_issues: list[FormulaIssue] = field(default_factory=list)
    balance_difference: float | None = None
    balance_ok: bool | None = None
    balance_reference: str = CONCEPTUAL_BALANCE_REFERENCE
    warnings: list[str] = field(default_factory=list)


def validate_generated_workbook(
    workbook_or_path: Workbook | str | Path,
    *,
    balance_difference: float | int | str | None = None,
    tolerance: float = 1.0,
) -> ValidationResult:
    """Validate formulas and the conceptual balance difference.

    The balance check mirrors the manual workbook concept at BG!L47 without
    copying or requiring that sheet. Pass `balance_difference` from the parser
    or balance-sheet integration when available.
    """

    workbook, close_after = _load_workbook(workbook_or_path)
    try:
        formula_issues = find_formula_issues(workbook)
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
        ok = not formula_issues and (balance_ok is not False)
        return ValidationResult(
            ok=ok,
            formula_error_count=len(formula_issues),
            formula_issues=formula_issues,
            balance_difference=coerced_difference,
            balance_ok=balance_ok,
            warnings=warnings,
        )
    finally:
        if close_after:
            workbook.close()


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
    details = [f"{issue.cell}: {issue.reason}" for issue in result.formula_issues]
    if result.balance_ok is False:
        details.append(
            f"Balance difference {result.balance_difference!r} does not satisfy abs(diff) < {tolerance}"
        )
    raise ValueError("; ".join(details) or "Generated workbook validation failed")


def _load_workbook(workbook_or_path: Workbook | str | Path) -> tuple[Workbook, bool]:
    if isinstance(workbook_or_path, Workbook):
        return workbook_or_path, False
    return load_workbook(Path(workbook_or_path), data_only=False), True


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
