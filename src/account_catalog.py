"""Parse and validate company account catalogs independently of a balanza.

The catalog contract is deliberately small: headerless CSV records with five
text fields.  This module does not choose a financial-statement line, mutate
catalog data, or infer an accounting profile.  It only retains the evidence
needed for those later steps.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import unicodedata
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable


CATALOG_FIELDS = (
    "account_code",
    "account_name",
    "parent_code",
    "nature",
    "sat_group_code",
)
UTF8_BOM = b"\xef\xbb\xbf"
UTF8_ENCODING = "utf-8"
UTF8_BOM_ENCODING = "utf-8-sig"
WINDOWS_1252_ENCODING = "windows-1252"
_ACCOUNT_CODE_RE = re.compile(r"^[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*$")
_SAT_GROUP_CODE_RE = re.compile(r"^\d{3}(?:\.\d{1,2})?$")

BLOCKING = "blocking"
WARNING = "warning"
OBSERVATION = "observation"


@dataclass(frozen=True)
class AccountCatalogRow:
    """A source row after CSV structural parsing, with every code as text."""

    source_row: int
    account_code: str
    account_name: str
    parent_code: str
    nature: str
    sat_group_code: str

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


@dataclass(frozen=True)
class CatalogIssue:
    """A non-mutating validation result with an explicit operational severity."""

    severity: str
    code: str
    message: str
    source_rows: tuple[int, ...] = ()
    account_codes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        data = asdict(self)
        data["source_rows"] = list(self.source_rows)
        data["account_codes"] = list(self.account_codes)
        return data


@dataclass(frozen=True)
class ParsedAccountCatalog:
    source_name: str
    source_sha256: str
    semantic_sha256: str | None
    encoding: str | None
    rows: tuple[AccountCatalogRow, ...]
    issues: tuple[CatalogIssue, ...]

    @property
    def is_valid(self) -> bool:
        return not any(issue.severity == BLOCKING for issue in self.issues)

    def issue_count(self, severity: str) -> int:
        return sum(issue.severity == severity for issue in self.issues)

    def to_dict(self) -> dict[str, object]:
        return {
            "source_name": self.source_name,
            "source_sha256": self.source_sha256,
            "semantic_sha256": self.semantic_sha256,
            "encoding": self.encoding,
            "row_count": len(self.rows),
            "is_valid": self.is_valid,
            "rows": [row.to_dict() for row in self.rows],
            "issues": [issue.to_dict() for issue in self.issues],
        }


def parse_account_catalog(path: str | Path) -> ParsedAccountCatalog:
    """Parse native catalog bytes from a private input path.

    Callers must keep raw customer CSV files outside the repository.  The
    returned source hash is calculated before decoding, so it is suitable for
    detecting byte-level changes to an approved catalog version.
    """

    catalog_path = Path(path)
    return parse_account_catalog_bytes(catalog_path.read_bytes(), source_name=catalog_path.name)


def parse_account_catalog_bytes(raw_bytes: bytes, *, source_name: str = "<memory>") -> ParsedAccountCatalog:
    """Parse a headerless catalog without coercing account or SAT codes."""

    source_sha256 = hashlib.sha256(raw_bytes).hexdigest()
    text, encoding, decoding_issue = _decode_catalog_bytes(raw_bytes)
    if decoding_issue is not None:
        return ParsedAccountCatalog(
            source_name=source_name,
            source_sha256=source_sha256,
            semantic_sha256=None,
            encoding=None,
            rows=(),
            issues=(decoding_issue,),
        )

    rows: list[AccountCatalogRow] = []
    issues: list[CatalogIssue] = [
        CatalogIssue(
            severity=OBSERVATION,
            code="encoding_detected",
            message=f"Catalog decoded as {encoding}.",
        )
    ]
    reader = csv.reader(io.StringIO(text, newline=""))
    for source_row, values in enumerate(reader, start=1):
        if not any(value.strip() for value in values):
            continue
        if len(values) != len(CATALOG_FIELDS):
            issues.append(
                CatalogIssue(
                    severity=BLOCKING,
                    code="invalid_column_count",
                    message=(
                        f"Expected {len(CATALOG_FIELDS)} headerless fields, found {len(values)}."
                    ),
                    source_rows=(source_row,),
                )
            )
            continue
        account_code, account_name, parent_code, nature, sat_group_code = (
            value.strip() for value in values
        )
        rows.append(
            AccountCatalogRow(
                source_row=source_row,
                account_code=account_code,
                account_name=account_name,
                parent_code=parent_code,
                nature=nature,
                sat_group_code=sat_group_code,
            )
        )

    if not rows:
        issues.append(
            CatalogIssue(
                severity=BLOCKING,
                code="empty_catalog",
                message="The catalog contains no account records.",
            )
        )
    else:
        issues.extend(_validate_rows(rows))

    return ParsedAccountCatalog(
        source_name=source_name,
        source_sha256=source_sha256,
        semantic_sha256=_semantic_sha256(rows),
        encoding=encoding,
        rows=tuple(rows),
        issues=tuple(issues),
    )


def build_catalog_matrix(catalogs: Iterable[ParsedAccountCatalog]) -> list[dict[str, object]]:
    """Return safe, aggregate evidence for a set of catalogs.

    Raw rows and customer account names are intentionally excluded.  This is
    suitable for a draft-profile intake report without placing CSV data in Git.
    """

    matrix: list[dict[str, object]] = []
    for catalog in catalogs:
        matrix.append(
            {
                "source_name": catalog.source_name,
                "source_sha256": catalog.source_sha256,
                "semantic_sha256": catalog.semantic_sha256,
                "encoding": catalog.encoding,
                "row_count": len(catalog.rows),
                "is_valid": catalog.is_valid,
                "blocking_count": catalog.issue_count(BLOCKING),
                "warning_count": catalog.issue_count(WARNING),
                "observation_count": catalog.issue_count(OBSERVATION),
            }
        )
    return matrix


def compare_catalog_semantics(
    previous: ParsedAccountCatalog, current: ParsedAccountCatalog
) -> tuple[CatalogIssue, ...]:
    """Describe safe account-code deltas without repairing or mapping either catalog.

    This comparison is deliberately limited to unique account codes.  A rename
    needs review, while a change to hierarchy, nature, or SAT classification is
    blocking because retaining an old mapping would be unsafe.  Account names
    are never copied into the issue payload.
    """

    previous_by_code = _unique_rows_by_code(previous.rows)
    current_by_code = _unique_rows_by_code(current.rows)
    issues: list[CatalogIssue] = []
    for code, row in sorted(current_by_code.items()):
        old = previous_by_code.get(code)
        if old is None:
            issues.append(
                CatalogIssue(
                    severity=OBSERVATION,
                    code="new_account_code",
                    message="New account code requires draft classification; no mapping was copied.",
                    source_rows=(row.source_row,),
                    account_codes=(code,),
                )
            )
            continue
        if old.account_name != row.account_name:
            issues.append(
                CatalogIssue(
                    severity=WARNING,
                    code="renamed_account_code",
                    message="Account name changed and requires review; no mapping was changed automatically.",
                    source_rows=(old.source_row, row.source_row),
                    account_codes=(code,),
                )
            )
        if (old.parent_code, old.nature, old.sat_group_code) != (
            row.parent_code,
            row.nature,
            row.sat_group_code,
        ):
            issues.append(
                CatalogIssue(
                    severity=BLOCKING,
                    code="account_code_semantics_changed",
                    message="Hierarchy, nature, or SAT evidence changed for an existing code.",
                    source_rows=(old.source_row, row.source_row),
                    account_codes=(code,),
                )
            )
    return tuple(issues)


def _unique_rows_by_code(rows: Iterable[AccountCatalogRow]) -> dict[str, AccountCatalogRow]:
    grouped: dict[str, list[AccountCatalogRow]] = {}
    for row in rows:
        grouped.setdefault(row.account_code, []).append(row)
    return {code: members[0] for code, members in grouped.items() if len(members) == 1}


def _decode_catalog_bytes(raw_bytes: bytes) -> tuple[str, str | None, CatalogIssue | None]:
    if raw_bytes.startswith(UTF8_BOM):
        try:
            return raw_bytes.decode(UTF8_BOM_ENCODING), UTF8_BOM_ENCODING, None
        except UnicodeDecodeError as exc:
            return "", None, _decoding_issue(exc)
    try:
        return raw_bytes.decode(UTF8_ENCODING), UTF8_ENCODING, None
    except UnicodeDecodeError:
        try:
            return raw_bytes.decode(WINDOWS_1252_ENCODING), WINDOWS_1252_ENCODING, None
        except UnicodeDecodeError as exc:
            return "", None, _decoding_issue(exc)


def _decoding_issue(exc: UnicodeDecodeError) -> CatalogIssue:
    return CatalogIssue(
        severity=BLOCKING,
        code="unsupported_encoding",
        message=f"Catalog could not be decoded as UTF-8 or Windows-1252: {exc.reason}.",
    )


def _validate_rows(rows: list[AccountCatalogRow]) -> list[CatalogIssue]:
    issues: list[CatalogIssue] = []
    by_code: dict[str, list[AccountCatalogRow]] = {}
    for row in rows:
        by_code.setdefault(row.account_code, []).append(row)
        issues.extend(_validate_row_fields(row))

    for account_code, duplicates in sorted(by_code.items()):
        if len(duplicates) > 1:
            issues.append(
                CatalogIssue(
                    severity=BLOCKING,
                    code="duplicate_account_code",
                    message=f"Account code {account_code} appears {len(duplicates)} times.",
                    source_rows=tuple(row.source_row for row in duplicates),
                    account_codes=(account_code,),
                )
            )

    known_codes = set(by_code)
    for row in rows:
        if row.parent_code and row.parent_code not in known_codes:
            issues.append(
                CatalogIssue(
                    severity=BLOCKING,
                    code="orphan_parent_code",
                    message=f"Parent code {row.parent_code} is not present in this catalog.",
                    source_rows=(row.source_row,),
                    account_codes=(row.account_code, row.parent_code),
                )
            )

    issues.extend(_find_parent_cycles(rows, by_code))
    return issues


def _validate_row_fields(row: AccountCatalogRow) -> list[CatalogIssue]:
    issues: list[CatalogIssue] = []
    for field_name, value in (
        ("account_code", row.account_code),
        ("account_name", row.account_name),
        ("nature", row.nature),
        ("sat_group_code", row.sat_group_code),
    ):
        if not value:
            issues.append(
                CatalogIssue(
                    severity=BLOCKING,
                    code=f"missing_{field_name}",
                    message=f"{field_name} is required.",
                    source_rows=(row.source_row,),
                    account_codes=(row.account_code,) if row.account_code else (),
                )
            )

    if row.account_code and not _ACCOUNT_CODE_RE.fullmatch(row.account_code):
        issues.append(
            CatalogIssue(
                severity=BLOCKING,
                code="invalid_account_code",
                message="account_code must use alphanumeric segments separated by hyphens.",
                source_rows=(row.source_row,),
                account_codes=(row.account_code,),
            )
        )
    if row.parent_code and not _ACCOUNT_CODE_RE.fullmatch(row.parent_code):
        issues.append(
            CatalogIssue(
                severity=BLOCKING,
                code="invalid_parent_code",
                message="parent_code must use alphanumeric segments separated by hyphens.",
                source_rows=(row.source_row,),
                account_codes=(row.account_code, row.parent_code),
            )
        )
    if row.nature and row.nature not in {"A", "D"}:
        issues.append(
            CatalogIssue(
                severity=BLOCKING,
                code="invalid_nature",
                message="nature must be exactly A or D; no value was normalized.",
                source_rows=(row.source_row,),
                account_codes=(row.account_code,),
            )
        )
    if row.sat_group_code and not _SAT_GROUP_CODE_RE.fullmatch(row.sat_group_code):
        issues.append(
            CatalogIssue(
                severity=BLOCKING,
                code="invalid_sat_group_code",
                message="sat_group_code must be a three-digit SAT group with an optional decimal suffix.",
                source_rows=(row.source_row,),
                account_codes=(row.account_code,),
            )
        )
    if (
        row.account_code[:1] in {"1", "2", "3"}
        and _SAT_GROUP_CODE_RE.fullmatch(row.sat_group_code)
        and row.sat_group_code[:1] != row.account_code[:1]
    ):
        issues.append(
            CatalogIssue(
                severity=WARNING,
                code="section_sat_mismatch",
                message=(
                    f"Account section {row.account_code[:1]} differs from SAT section "
                    f"{row.sat_group_code[:1]}; the source value was retained unchanged."
                ),
                source_rows=(row.source_row,),
                account_codes=(row.account_code,),
            )
        )
    return issues


def _find_parent_cycles(
    rows: list[AccountCatalogRow], by_code: dict[str, list[AccountCatalogRow]]
) -> list[CatalogIssue]:
    """Find cycles only when the parent relation is unambiguous."""

    parent_by_code = {
        code: members[0].parent_code
        for code, members in by_code.items()
        if len(members) == 1 and members[0].parent_code in by_code and len(by_code[members[0].parent_code]) == 1
    }
    source_row_by_code = {code: members[0].source_row for code, members in by_code.items() if len(members) == 1}
    visited: set[str] = set()
    cycles: set[tuple[str, ...]] = set()
    for start in sorted(parent_by_code):
        if start in visited:
            continue
        path: list[str] = []
        index_by_code: dict[str, int] = {}
        current = start
        while current in parent_by_code and current not in visited:
            if current in index_by_code:
                cycle = path[index_by_code[current] :]
                cycles.add(_canonical_cycle(cycle))
                break
            index_by_code[current] = len(path)
            path.append(current)
            current = parent_by_code[current]
        visited.update(path)

    issues: list[CatalogIssue] = []
    for cycle in sorted(cycles):
        issues.append(
            CatalogIssue(
                severity=BLOCKING,
                code="parent_cycle",
                message=f"Parent hierarchy contains a cycle: {' -> '.join((*cycle, cycle[0]))}.",
                source_rows=tuple(source_row_by_code[code] for code in cycle),
                account_codes=cycle,
            )
        )
    return issues


def _canonical_cycle(cycle: list[str]) -> tuple[str, ...]:
    rotations = [tuple(cycle[index:] + cycle[:index]) for index in range(len(cycle))]
    return min(rotations)


def _semantic_sha256(rows: Iterable[AccountCatalogRow]) -> str:
    """Hash canonical record values, independent of encoding, BOM and CSV order."""

    canonical_rows = sorted(
        [
            [
                unicodedata.normalize("NFC", row.account_code),
                unicodedata.normalize("NFC", row.account_name),
                unicodedata.normalize("NFC", row.parent_code),
                unicodedata.normalize("NFC", row.nature),
                unicodedata.normalize("NFC", row.sat_group_code),
            ]
            for row in rows
        ]
    )
    payload = json.dumps(
        {"schema": "account-catalog-semantic-v1", "rows": canonical_rows},
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()
