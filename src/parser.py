from __future__ import annotations

import re
import zipfile
from dataclasses import asdict, dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook

try:
    from .period import PeriodVariables, extract_period_variables
except ImportError:  # pragma: no cover - supports PYTHONPATH=src usage.
    from period import PeriodVariables, extract_period_variables


REQUIRED_SHEET = "Balanza"
REQUIRED_COLUMNS = ("Cuenta", "Saldo Inicial", "Debe", "Haber", "SaldoFinal")

_ACCOUNT_RE = re.compile(
    r"^\s*(?P<code>\d+(?:-\s*[A-Za-z0-9]+)*)\s+(?P<name>.+?)\s*$"
)


@dataclass(frozen=True)
class BalanzaRow:
    source_row: int
    account_raw: str
    account_code: str
    account_name: str
    top_account: str
    saldo_inicial: Decimal
    debe: Decimal
    haber: Decimal
    saldo_final: Decimal

    def to_dict(self) -> dict[str, object]:
        data = asdict(self)
        for key in ("saldo_inicial", "debe", "haber", "saldo_final"):
            data[key] = str(data[key])
        return data


@dataclass(frozen=True)
class RowIssue:
    source_row: int
    message: str
    values: tuple[Any, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


@dataclass(frozen=True)
class ParsedBalanza:
    source_path: str
    sheet_name: str
    period: PeriodVariables
    company_name: str | None
    content_period_ym: str | None
    header_row: int
    rows: tuple[BalanzaRow, ...]
    empty_rows: tuple[int, ...]
    structure_issues: tuple[RowIssue, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "source_path": self.source_path,
            "sheet_name": self.sheet_name,
            "period": self.period.to_dict(),
            "company_name": self.company_name,
            "content_period_ym": self.content_period_ym,
            "header_row": self.header_row,
            "rows": [row.to_dict() for row in self.rows],
            "empty_rows": list(self.empty_rows),
            "structure_issues": [issue.to_dict() for issue in self.structure_issues],
        }


def load_ooxml_workbook(path: str | Path):
    workbook_path = Path(path)
    if not workbook_path.exists():
        raise FileNotFoundError(workbook_path)
    if not zipfile.is_zipfile(workbook_path):
        raise ValueError(f"Workbook is not an OOXML ZIP package: {workbook_path}")

    with workbook_path.open("rb") as fh:
        return load_workbook(fh, data_only=True, read_only=False)


def parse_balanza(path: str | Path, sheet_name: str = REQUIRED_SHEET) -> ParsedBalanza:
    workbook_path = Path(path)
    period = extract_period_variables(workbook_path)
    workbook = load_ooxml_workbook(workbook_path)
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Required sheet '{sheet_name}' not found. Sheets: {workbook.sheetnames}")

    worksheet = workbook[sheet_name]
    company_name, content_period_ym = _extract_sheet_metadata(worksheet)
    header_row, column_map = _find_header_row(worksheet.iter_rows(values_only=True))

    rows: list[BalanzaRow] = []
    empty_rows: list[int] = []
    issues: list[RowIssue] = []
    data_started = False

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        values = tuple(
            worksheet.cell(row=row_number, column=column_map[column]).value
            for column in REQUIRED_COLUMNS
        )
        if _is_empty(values):
            empty_rows.append(row_number)
            continue

        account_raw = _clean(values[0])
        parsed_account = _parse_account(account_raw)
        if parsed_account is None:
            if data_started:
                issues.append(RowIssue(row_number, "Missing or invalid account code.", values))
            continue

        data_started = True
        amounts = _parse_amounts(values[1:])
        if isinstance(amounts, RowIssue):
            issues.append(RowIssue(row_number, amounts.message, values))
            continue

        account_code, account_name = parsed_account
        rows.append(
            BalanzaRow(
                source_row=row_number,
                account_raw=account_raw,
                account_code=account_code,
                account_name=account_name,
                top_account=account_code.split("-", 1)[0],
                saldo_inicial=amounts[0],
                debe=amounts[1],
                haber=amounts[2],
                saldo_final=amounts[3],
            )
        )

    if not rows:
        issues.append(RowIssue(header_row, "No account rows were parsed.", ()))
    if content_period_ym and content_period_ym != period.period_ym:
        issues.append(
            RowIssue(
                header_row,
                f"Filename period {period.period_ym} differs from sheet period {content_period_ym}.",
                (),
            )
        )

    return ParsedBalanza(
        source_path=str(workbook_path),
        sheet_name=sheet_name,
        period=period,
        company_name=company_name,
        content_period_ym=content_period_ym,
        header_row=header_row,
        rows=tuple(rows),
        empty_rows=tuple(empty_rows),
        structure_issues=tuple(issues),
    )


def parse_balanza_dict(path: str | Path, sheet_name: str = REQUIRED_SHEET) -> dict[str, object]:
    return parse_balanza(path, sheet_name=sheet_name).to_dict()


def _find_header_row(rows: Iterable[tuple[Any, ...]]) -> tuple[int, dict[str, int]]:
    required = {_normalize_header(column): column for column in REQUIRED_COLUMNS}
    for row_number, row in enumerate(rows, start=1):
        headers = {_normalize_header(value): index for index, value in enumerate(row, start=1)}
        if all(normalized in headers for normalized in required):
            return row_number, {
                original: headers[normalized] for normalized, original in required.items()
            }
    raise ValueError(f"Required columns not found: {', '.join(REQUIRED_COLUMNS)}")


def _parse_account(value: str) -> tuple[str, str] | None:
    match = _ACCOUNT_RE.match(value)
    if not match:
        return None
    account_code = re.sub(r"\s+", "", match.group("code"))
    return account_code, match.group("name").strip()


def _parse_amounts(values: tuple[Any, ...]) -> tuple[Decimal, Decimal, Decimal, Decimal] | RowIssue:
    parsed: list[Decimal] = []
    for value in values:
        try:
            parsed.append(_to_decimal(value))
        except (InvalidOperation, TypeError, ValueError):
            return RowIssue(0, f"Invalid numeric value: {value!r}", values)
    return parsed[0], parsed[1], parsed[2], parsed[3]


def _to_decimal(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value).replace(",", "").strip())


def _is_empty(values: tuple[Any, ...]) -> bool:
    return all(value is None or _clean(value) == "" for value in values)


def _clean(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", _clean(value).casefold())


def _extract_sheet_metadata(worksheet) -> tuple[str | None, str | None]:
    company_name: str | None = None
    content_period_ym: str | None = None
    rfc_pattern = re.compile(r"\b[A-Z&Ñ]{3,4}\d{6}[A-Z0-9]{3}\b")
    period_pattern = re.compile(r"(\d{4})[-/](\d{1,2})")

    for row in worksheet.iter_rows(min_row=1, max_row=min(worksheet.max_row, 12), values_only=True):
        for value in row:
            text = _clean(value)
            if not text:
                continue
            if company_name is None and rfc_pattern.search(text) and "periodo" not in text.casefold():
                company_name = rfc_pattern.sub("", text).strip()
            if content_period_ym is None:
                match = period_pattern.search(text)
                if match:
                    content_period_ym = f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"
            if company_name and content_period_ym:
                return company_name, content_period_ym

    return company_name, content_period_ym
