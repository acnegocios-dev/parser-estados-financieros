"""Offline intake reports for private company account catalogs.

This module is deliberately separate from runtime profile selection.  It reads
only local CSV paths supplied by an operator, never downloads data, and emits
redacted evidence suitable for preparing a *draft* accounting profile.  It
does not write raw CSV content, account names, or an approval decision.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

try:
    from .account_catalog import BLOCKING, AccountCatalogRow, ParsedAccountCatalog, parse_account_catalog
    from .accounting_profiles import AccountingProfile, MappingRule, load_accounting_profile, load_sat_rmf_2026_v1
except ImportError:  # pragma: no cover - supports PYTHONPATH=src usage.
    from account_catalog import BLOCKING, AccountCatalogRow, ParsedAccountCatalog, parse_account_catalog
    from accounting_profiles import AccountingProfile, MappingRule, load_accounting_profile, load_sat_rmf_2026_v1


DRAFT_REPORT_VERSION = "catalog-profile-draft-report-v1"
REJECTED = "rejected"
BLOCKED = "blocked"
DRAFT = "draft"
CONTROLLED_PARALLEL_MINIMUM = 3


@dataclass(frozen=True)
class DraftSuggestion:
    account_code: str
    statement: str | None
    line_key: str | None
    confidence: float
    status: str
    evidence: tuple[str, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "account_code": self.account_code,
            "statement": self.statement,
            "line_key": self.line_key,
            "confidence": self.confidence,
            "status": self.status,
            "evidence": list(self.evidence),
        }


@dataclass(frozen=True)
class ParallelCompanyEvidence:
    """Explicit non-financial readiness evidence for a controlled parallel pilot."""

    profile: AccountingProfile
    representative_balance_verified: bool = False
    accounting_review_completed: bool = False


def controlled_parallel_readiness(
    evidence: Iterable[ParallelCompanyEvidence], *, minimum: int = CONTROLLED_PARALLEL_MINIMUM
) -> dict[str, object]:
    """Require three independently evidenced companies before a parallel pilot.

    This is an offline governance gate.  It neither selects profiles at runtime
    nor claims that a fixture or a catalog has financial validation.
    """

    qualified: set[tuple[str, str]] = set()
    pending: list[str] = []
    for item in evidence:
        profile = item.profile
        identity = (profile.rfc or "", profile.profile_id)
        if profile.status != "approved" or not profile.rfc:
            pending.append("approved_profile_with_rfc_required")
            continue
        if not profile.catalog_identity.source_sha256 or not profile.catalog_identity.semantic_sha256:
            pending.append("approved_catalog_identity_required")
            continue
        if not item.representative_balance_verified:
            pending.append("representative_balance_review_required")
            continue
        if not item.accounting_review_completed:
            pending.append("accounting_review_required")
            continue
        qualified.add(identity)
    enough = len(qualified) >= minimum
    if not enough:
        pending.append("minimum_three_independently_qualified_companies_required")
    return {
        "allowed": enough,
        "minimum_companies": minimum,
        "qualified_companies": len(qualified),
        "pending": sorted(set(pending)),
        "financial_validation": "No se infiere de fixtures ni de catalogos sin balanza representativa.",
    }


def analyze_catalog_paths(
    paths: Iterable[str | Path],
    *,
    previous_profile: AccountingProfile | None = None,
    rfc_confirmed: str | None = None,
    representative_balance_path: str | Path | None = None,
) -> dict[str, object]:
    """Create one consolidated, redacted, offline draft report.

    ``rfc_confirmed`` and ``representative_balance_path`` are selection gates,
    not proof of financial validity.  The function never opens a balance: its
    existence only records whether an operator supplied representative evidence.
    """

    catalogs = tuple(parse_account_catalog(path) for path in paths)
    reports = [
        _catalog_report(
            catalog,
            previous_profile=previous_profile,
            rfc_confirmed=rfc_confirmed,
            representative_balance_path=representative_balance_path,
        )
        for catalog in catalogs
    ]
    totals = Counter()
    for report in reports:
        totals[report["status"]] += 1
        totals["records"] += int(report["row_count"])
        totals["blocking_issues"] += int(report["issue_counts"][BLOCKING])
        totals["unmapped"] += int(report["mapping_counts"]["unmapped"])
        totals["ambiguous"] += int(report["mapping_counts"]["ambiguous"])
        totals["collisions"] += int(report["mapping_counts"]["collisions"])
    return {
        "report_version": DRAFT_REPORT_VERSION,
        "mode": "offline",
        "approval_state": DRAFT,
        "runtime_selection": {
            "allowed": False,
            "reason": "Catalog intake only creates drafts; runtime requires approved profile, confirmed RFC and representative balance.",
        },
        "financial_validation": {
            "performed": False,
            "statement": "No se afirma validacion financiera real sin balanza representativa de prueba.",
        },
        "catalog_count": len(reports),
        "totals": dict(totals),
        "catalogs": reports,
    }


def write_redacted_report(report: Mapping[str, object], output_path: str | Path) -> Path:
    """Persist only the safe report; callers must keep raw CSVs outside Git."""

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")
    return path


def _catalog_report(
    catalog: ParsedAccountCatalog,
    *,
    previous_profile: AccountingProfile | None,
    rfc_confirmed: str | None,
    representative_balance_path: str | Path | None,
) -> dict[str, object]:
    # A company overlay is never borrowed merely because another catalog uses
    # the same numeric code.  It can participate only when its source hash
    # binds it to this exact catalog version.
    matching_previous = (
        previous_profile
        if previous_profile and previous_profile.catalog_identity.source_sha256 == catalog.source_sha256
        else None
    )
    taxonomy = load_sat_rmf_2026_v1()
    suggestions = tuple(_suggest_row(row, matching_previous, taxonomy.entries) for row in catalog.rows)
    collisions = _mapping_collisions(suggestions)
    unmapped = [item for item in suggestions if item.status == "unmapped"]
    ambiguous = [item for item in suggestions if item.status == "ambiguous"]
    blocking_issues = [issue for issue in catalog.issues if issue.severity == BLOCKING]
    status = REJECTED if not catalog.rows else BLOCKED if blocking_issues or collisions else DRAFT
    gates = _runtime_gates(rfc_confirmed, representative_balance_path)
    previous_diff = _previous_profile_diff(catalog, previous_profile, applicable=matching_previous is not None)
    return {
        "catalog_ref": _redacted_catalog_ref(catalog.source_name),
        "source_sha256": catalog.source_sha256,
        "semantic_sha256": catalog.semantic_sha256,
        "encoding": catalog.encoding,
        "row_count": len(catalog.rows),
        "status": status,
        "issue_counts": {
            BLOCKING: catalog.issue_count(BLOCKING),
            "warning": catalog.issue_count("warning"),
            "observation": catalog.issue_count("observation"),
        },
        "validations": [_redact_issue(issue.to_dict()) for issue in catalog.issues],
        "suggestions": [item.to_dict() for item in suggestions],
        "unmapped_accounts": [item.to_dict() for item in unmapped],
        "ambiguous_accounts": [item.to_dict() for item in ambiguous],
        "collisions": collisions,
        "mapping_counts": {
            "suggested": sum(item.status == "suggested" for item in suggestions),
            "name_suggestion_only": sum(item.status == "name_suggestion_only" for item in suggestions),
            "unmapped": len(unmapped),
            "ambiguous": len(ambiguous),
            "collisions": len(collisions),
        },
        "previous_profile_difference": previous_diff,
        "draft_profile": {
            "status": DRAFT,
            "rfc": rfc_confirmed if rfc_confirmed else None,
            "runtime_selectable": False,
            "runtime_gates": gates,
            "approval_required": True,
        },
        "notes": [
            "El agrupador SAT es la taxonomia base; el catalogo solo aporta evidencia de overlay.",
            "Los nombres de cuenta se redacted y solo pueden originar sugerencias, nunca resolucion ni aprobacion.",
        ],
    }


def _suggest_row(
    row: AccountCatalogRow,
    previous_profile: AccountingProfile | None,
    taxonomy_entries: Sequence[Any],
) -> DraftSuggestion:
    previous = _previous_exact_candidates(row, previous_profile)
    if len(previous) == 1:
        statement, line_key, rule_id = previous[0]
        return DraftSuggestion(row.account_code, statement, line_key, 0.99, "suggested", (f"previous_profile_exact:{rule_id}",))
    if len(previous) > 1:
        return DraftSuggestion(row.account_code, None, None, 0.0, "ambiguous", tuple(sorted(f"previous_profile:{rule_id}" for _, _, rule_id in previous)))

    exact = [entry for entry in taxonomy_entries if entry.sat_group_code == row.sat_group_code]
    if len(exact) == 1:
        entry = exact[0]
        return DraftSuggestion(row.account_code, entry.statement, entry.line_key, 0.90, "suggested", (f"sat_exact:{row.sat_group_code}", f"nature:{row.nature}", f"parent_present:{bool(row.parent_code)}"))
    if len(exact) > 1:
        return DraftSuggestion(row.account_code, None, None, 0.0, "ambiguous", tuple(sorted(f"sat_exact:{entry.statement}:{entry.line_key}" for entry in exact)))

    family = row.sat_group_code.split(".", 1)[0]
    family_entries = [entry for entry in taxonomy_entries if entry.sat_group_code == family]
    if len(family_entries) == 1:
        entry = family_entries[0]
        return DraftSuggestion(row.account_code, entry.statement, entry.line_key, 0.70, "suggested", (f"sat_family:{family}", f"nature:{row.nature}", f"parent_present:{bool(row.parent_code)}"))
    if len(family_entries) > 1:
        return DraftSuggestion(row.account_code, None, None, 0.0, "ambiguous", tuple(sorted(f"sat_family:{entry.statement}:{entry.line_key}" for entry in family_entries)))
    name_candidates = _name_suggestion_candidates(row, previous_profile)
    if len(name_candidates) == 1:
        statement, line_key, rule_id = name_candidates[0]
        return DraftSuggestion(
            row.account_code, statement, line_key, 0.20, "name_suggestion_only",
            (f"name_suggestion:{rule_id}", f"sat_unmapped:{row.sat_group_code}"),
        )
    if len(name_candidates) > 1:
        return DraftSuggestion(
            row.account_code, None, None, 0.0, "ambiguous",
            tuple(sorted(f"name_suggestion:{rule_id}" for _, _, rule_id in name_candidates)),
        )
    return DraftSuggestion(row.account_code, None, None, 0.0, "unmapped", (f"sat_unmapped:{row.sat_group_code}",))


def _previous_exact_candidates(
    row: AccountCatalogRow, previous_profile: AccountingProfile | None
) -> list[tuple[str, str, str]]:
    if previous_profile is None:
        return []
    return [
        (rule.statement, rule.line_key, rule.rule_id)
        for rule in previous_profile.rules
        if rule.approved and rule.kind in {"exact_code", "subtree"}
        and rule.account_code == row.account_code
    ]


def _name_suggestion_candidates(
    row: AccountCatalogRow, previous_profile: AccountingProfile | None
) -> list[tuple[str, str, str]]:
    if previous_profile is None:
        return []
    normalized = _normalize_name(row.account_name)
    return [
        (rule.statement, rule.line_key, rule.rule_id)
        for rule in previous_profile.rules
        if rule.kind == "name_suggestion" and rule.normalized_name == normalized
    ]


def _normalize_name(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    return " ".join("".join(character for character in normalized if not unicodedata.combining(character)).upper().split())


def _mapping_collisions(suggestions: Sequence[DraftSuggestion]) -> list[dict[str, object]]:
    by_code: dict[str, list[DraftSuggestion]] = defaultdict(list)
    for suggestion in suggestions:
        by_code[suggestion.account_code].append(suggestion)
    return [
        {
            "account_code": account_code,
            "suggestions": [item.to_dict() for item in items],
            "reason": "multiple_catalog_records_same_account_code",
        }
        for account_code, items in sorted(by_code.items())
        if len(items) > 1
    ]


def _previous_profile_diff(
    catalog: ParsedAccountCatalog,
    previous_profile: AccountingProfile | None,
    *,
    applicable: bool,
) -> dict[str, object]:
    if previous_profile is None:
        return {"available": False, "reason": "no_previous_profile_supplied"}
    identity = previous_profile.catalog_identity
    return {
        "available": True,
        "applicable": applicable,
        "profile_id": previous_profile.profile_id,
        "profile_version": previous_profile.profile_version,
        "source_sha256_changed": identity.source_sha256 != catalog.source_sha256,
        "semantic_sha256_changed": identity.semantic_sha256 != catalog.semantic_sha256,
        "previous_source_sha256": identity.source_sha256,
        "previous_semantic_sha256": identity.semantic_sha256,
        "reason": None if applicable else "previous_profile_belongs_to_a_different_catalog_hash",
    }


def _runtime_gates(rfc_confirmed: str | None, representative_balance_path: str | Path | None) -> list[str]:
    gates = ["approved_profile_required"]
    if not rfc_confirmed:
        gates.append("confirmed_rfc_required")
    if not representative_balance_path:
        gates.append("representative_balance_required")
    return gates


def _redacted_catalog_ref(source_name: str) -> str:
    return "catalog-" + hashlib.sha256(source_name.encode("utf-8")).hexdigest()[:16]


def _redact_issue(issue: Mapping[str, Any]) -> dict[str, object]:
    """Keep row/code evidence but never emit raw account names or filename."""

    return {
        "severity": issue["severity"],
        "code": issue["code"],
        "source_rows": issue.get("source_rows", []),
        "account_codes": issue.get("account_codes", []),
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Offline redacted catalog-to-draft-profile analysis.")
    parser.add_argument("catalogs", nargs="+", type=Path, help="Private CSV paths; never place them in Git.")
    parser.add_argument("--output", required=True, type=Path, help="Redacted JSON report path.")
    parser.add_argument("--previous-profile", type=Path, help="Optional local profile used only when its catalog hash matches.")
    parser.add_argument("--rfc-confirmed", help="Recorded draft evidence only; never approves a profile.")
    parser.add_argument("--representative-balance", type=Path, help="Recorded draft evidence only; balance is not evaluated here.")
    args = parser.parse_args(argv)
    previous = load_accounting_profile(args.previous_profile) if args.previous_profile else None
    write_redacted_report(
        analyze_catalog_paths(
            args.catalogs,
            previous_profile=previous,
            rfc_confirmed=args.rfc_confirmed,
            representative_balance_path=args.representative_balance,
        ),
        args.output,
    )
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
