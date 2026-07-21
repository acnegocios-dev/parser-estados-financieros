"""Versioned accounting-profile schema and offline SAT taxonomy resolution.

This module deliberately has no dependency on the financial-statement engine,
workbook builder, HTTP API, or a network service.  It validates an accounting
overlay before a later integration stage is allowed to use it.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Mapping


ACCOUNTING_PROFILE_SCHEMA_VERSION = "accounting-profile-v1"
SAT_TAXONOMY_SCHEMA_VERSION = "sat-taxonomy-v1"
SAT_RMF_2026_V1 = "sat-rmf-2026-v1"
PROFILE_STATUSES = frozenset({"draft", "approved", "retired"})
STATEMENTS = frozenset({"BG", "ER"})
RULE_KINDS = frozenset(
    {"exact_code", "subtree", "sat_exact", "sat_family", "context", "name_suggestion"}
)
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_RFC_RE = re.compile(r"^[A-Z&Ñ]{3,4}\d{6}[A-Z0-9]{3}$")
_SAT_EXACT_RE = re.compile(r"^\d{3}(?:\.\d{1,2})?$")
_SAT_FAMILY_RE = re.compile(r"^\d{3}$")


@dataclass(frozen=True)
class CatalogIdentity:
    source_sha256: str
    semantic_sha256: str


@dataclass(frozen=True)
class Approval:
    approved_by: str
    approved_at: date
    approval_reference: str


@dataclass(frozen=True)
class ProfileReference:
    profile_id: str
    profile_version: str


@dataclass(frozen=True)
class GeneratorProfileReference:
    profile_id: str
    profile_version: str


@dataclass(frozen=True)
class MappingRule:
    """A company overlay rule; it is not a universal account-code dictionary."""

    rule_id: str
    statement: str
    line_key: str
    kind: str
    account_code: str | None = None
    sat_group_code: str | None = None
    parent_code: str | None = None
    nature: str | None = None
    section_guard: str | None = None
    normalized_name: str | None = None
    exclude_account_codes: tuple[str, ...] = ()
    approved: bool = False
    composite: bool = False
    presentation_sign: int = 1
    resolution_policy: str = "exact_accumulator_then_leaf_fallback"


@dataclass(frozen=True)
class AccountingProfile:
    profile_id: str
    profile_version: str
    status: str
    company_name: str
    rfc: str | None
    valid_from: date | None
    valid_to: date | None
    catalog_identity: CatalogIdentity
    taxonomy_version: str
    enabled_lines: Mapping[str, tuple[str, ...]]
    rules: tuple[MappingRule, ...] = ()
    supersedes: ProfileReference | None = None
    approval: Approval | None = None
    generator_profile: GeneratorProfileReference | None = None
    regression_controls: Mapping[str, str] = field(default_factory=dict)
    schema_version: str = ACCOUNTING_PROFILE_SCHEMA_VERSION

    def to_dict(self) -> dict[str, object]:
        return {
            "schema_version": self.schema_version,
            "profile_id": self.profile_id,
            "profile_version": self.profile_version,
            "status": self.status,
            "company_name": self.company_name,
            "rfc": self.rfc,
            "valid_from": self.valid_from.isoformat() if self.valid_from else None,
            "valid_to": self.valid_to.isoformat() if self.valid_to else None,
            "catalog_identity": asdict(self.catalog_identity),
            "taxonomy_version": self.taxonomy_version,
            "enabled_lines": {statement: list(lines) for statement, lines in self.enabled_lines.items()},
            "rules": [asdict(rule) for rule in self.rules],
            "supersedes": asdict(self.supersedes) if self.supersedes else None,
            "generator_profile": asdict(self.generator_profile) if self.generator_profile else None,
            "regression_controls": dict(self.regression_controls),
            "approval": (
                {
                    **asdict(self.approval),
                    "approved_at": self.approval.approved_at.isoformat(),
                }
                if self.approval
                else None
            ),
        }


@dataclass(frozen=True)
class TaxonomyEntry:
    sat_group_code: str
    statement: str
    line_key: str
    label: str


@dataclass(frozen=True)
class OfflineSatTaxonomy:
    schema_version: str
    taxonomy_id: str
    source: str
    entries: tuple[TaxonomyEntry, ...]


@dataclass(frozen=True)
class AccountClassificationInput:
    account_code: str
    account_name: str
    parent_code: str
    nature: str
    sat_group_code: str


@dataclass(frozen=True)
class ProfileIssue:
    code: str
    message: str
    rule_ids: tuple[str, ...] = ()


class ProfileValidationError(ValueError):
    def __init__(self, issues: Iterable[ProfileIssue]):
        self.issues = tuple(issues)
        super().__init__("; ".join(issue.code for issue in self.issues))


@dataclass(frozen=True)
class Resolution:
    status: str
    line_keys: tuple[str, ...]
    precedence: str | None
    evidence: tuple[str, ...]
    suggestions: tuple[str, ...] = ()


def load_sat_rmf_2026_v1(path: str | Path | None = None) -> OfflineSatTaxonomy:
    """Load the versioned local taxonomy; this function performs no I/O over a network."""

    taxonomy_path = Path(path) if path else Path(__file__).with_name("sat_rmf_2026_v1.json")
    payload = json.loads(taxonomy_path.read_text(encoding="utf-8"))
    entries = tuple(TaxonomyEntry(**entry) for entry in payload["entries"])
    taxonomy = OfflineSatTaxonomy(
        schema_version=payload["schema_version"],
        taxonomy_id=payload["taxonomy_id"],
        source=payload["source"],
        entries=entries,
    )
    issues = validate_taxonomy(taxonomy)
    if issues:
        raise ProfileValidationError(issues)
    return taxonomy


def profile_from_dict(payload: Mapping[str, Any]) -> AccountingProfile:
    """Read a profile document without silently normalizing business values."""

    identity = CatalogIdentity(**payload["catalog_identity"])
    rules = tuple(
        MappingRule(
            **{
                **rule,
                "exclude_account_codes": tuple(rule.get("exclude_account_codes", ())),
            }
        )
        for rule in payload.get("rules", ())
    )
    supersedes = payload.get("supersedes")
    approval = payload.get("approval")
    generator_profile = payload.get("generator_profile")
    return AccountingProfile(
        schema_version=payload.get("schema_version", ACCOUNTING_PROFILE_SCHEMA_VERSION),
        profile_id=payload["profile_id"],
        profile_version=payload["profile_version"],
        status=payload["status"],
        company_name=payload["company_name"],
        rfc=payload.get("rfc"),
        valid_from=_parse_date(payload.get("valid_from")),
        valid_to=_parse_date(payload.get("valid_to")),
        catalog_identity=identity,
        taxonomy_version=payload["taxonomy_version"],
        enabled_lines={key: tuple(value) for key, value in payload["enabled_lines"].items()},
        rules=rules,
        supersedes=ProfileReference(**supersedes) if supersedes else None,
        generator_profile=GeneratorProfileReference(**generator_profile) if generator_profile else None,
        regression_controls=dict(payload.get("regression_controls", {})),
        approval=(
            Approval(
                approved_by=approval["approved_by"],
                approved_at=_parse_date(approval["approved_at"]),
                approval_reference=approval["approval_reference"],
            )
            if approval
            else None
        ),
    )


def load_accounting_profile(path: str | Path) -> AccountingProfile:
    """Load a versioned local profile; runtime code never fetches one online."""

    profile_path = Path(path)
    return profile_from_dict(json.loads(profile_path.read_text(encoding="utf-8")))


def load_accounting_profiles(directory: str | Path) -> tuple[AccountingProfile, ...]:
    """Load local profile documents only; profile filenames are not selectors."""

    profile_dir = Path(directory)
    return tuple(
        load_accounting_profile(path)
        for path in sorted(profile_dir.glob("*.json"))
    )


def select_profile_for_runtime(
    profiles: Iterable[AccountingProfile],
    *,
    rfc: str,
    as_of: date,
    catalog_identity: CatalogIdentity,
) -> AccountingProfile:
    """Select exactly one approved, effective profile by RFC and catalog hashes."""

    candidates = tuple(profile for profile in profiles if profile.rfc == rfc and _is_effective(profile, as_of))
    if not candidates:
        raise ProfileValidationError((
            ProfileIssue("profile_not_found", "No profile matches RFC and effective date."),
        ))
    approved = tuple(profile for profile in candidates if profile.status == "approved")
    if not approved:
        raise ProfileValidationError((
            ProfileIssue("profile_not_approved", "Matching profile is not approved for runtime."),
        ))
    hash_matches = tuple(
        profile for profile in approved
        if profile.catalog_identity == catalog_identity
    )
    if not hash_matches:
        raise ProfileValidationError((
            ProfileIssue("catalog_hash_mismatch", "Catalog source or semantic hash differs from approved profile."),
        ))
    if len(hash_matches) != 1:
        raise ProfileValidationError((
            ProfileIssue("ambiguous_mapping", "More than one approved profile matches runtime identity."),
        ))
    profile = hash_matches[0]
    require_profile_for_runtime(profile, rfc=rfc, as_of=as_of, catalog_identity=catalog_identity)
    return profile


def _is_effective(profile: AccountingProfile, as_of: date) -> bool:
    return bool(
        profile.valid_from
        and as_of >= profile.valid_from
        and (profile.valid_to is None or as_of <= profile.valid_to)
    )


def validate_taxonomy(taxonomy: OfflineSatTaxonomy) -> tuple[ProfileIssue, ...]:
    issues: list[ProfileIssue] = []
    if taxonomy.schema_version != SAT_TAXONOMY_SCHEMA_VERSION:
        issues.append(ProfileIssue("taxonomy_schema_version_invalid", "Unexpected taxonomy schema version."))
    if taxonomy.taxonomy_id != SAT_RMF_2026_V1:
        issues.append(ProfileIssue("taxonomy_version_invalid", "Unexpected offline SAT taxonomy version."))
    seen: set[tuple[str, str]] = set()
    for entry in taxonomy.entries:
        key = (entry.statement, entry.sat_group_code)
        if entry.statement not in STATEMENTS:
            issues.append(ProfileIssue("taxonomy_statement_invalid", f"Invalid statement {entry.statement}."))
        if not _SAT_EXACT_RE.fullmatch(entry.sat_group_code):
            issues.append(ProfileIssue("taxonomy_sat_code_invalid", f"Invalid SAT code {entry.sat_group_code}."))
        if key in seen:
            issues.append(ProfileIssue("taxonomy_duplicate_entry", f"Duplicate taxonomy entry {key}."))
        seen.add(key)
    return tuple(issues)


def validate_accounting_profile(
    profile: AccountingProfile, *, taxonomy: OfflineSatTaxonomy | None = None
) -> tuple[ProfileIssue, ...]:
    """Validate a profile document before it can be considered for runtime use."""

    taxonomy = taxonomy or load_sat_rmf_2026_v1()
    issues: list[ProfileIssue] = list(validate_taxonomy(taxonomy))
    if profile.schema_version != ACCOUNTING_PROFILE_SCHEMA_VERSION:
        issues.append(ProfileIssue("profile_schema_version_invalid", "Unexpected accounting profile schema version."))
    if profile.status not in PROFILE_STATUSES:
        issues.append(ProfileIssue("profile_status_invalid", f"Unsupported profile status {profile.status}."))
    if not profile.profile_id or not profile.profile_version:
        issues.append(ProfileIssue("profile_identity_missing", "profile_id and profile_version are required."))
    if not profile.company_name:
        issues.append(ProfileIssue("company_name_missing", "company_name is required."))
    if profile.taxonomy_version != taxonomy.taxonomy_id:
        issues.append(ProfileIssue("taxonomy_version_mismatch", "Profile must reference the loaded offline taxonomy."))
    if not _SHA256_RE.fullmatch(profile.catalog_identity.source_sha256):
        issues.append(ProfileIssue("source_hash_invalid", "source_sha256 must be lowercase SHA-256 hex."))
    if not _SHA256_RE.fullmatch(profile.catalog_identity.semantic_sha256):
        issues.append(ProfileIssue("semantic_hash_invalid", "semantic_sha256 must be lowercase SHA-256 hex."))
    if profile.valid_from is None:
        issues.append(ProfileIssue("valid_from_missing", "Profile validity start is required."))
    if profile.valid_from and profile.valid_to and profile.valid_from > profile.valid_to:
        issues.append(ProfileIssue("validity_range_invalid", "valid_from must not be after valid_to."))
    if profile.supersedes and profile.supersedes.profile_id == profile.profile_id and profile.supersedes.profile_version == profile.profile_version:
        issues.append(ProfileIssue("supersedes_self", "A profile cannot supersede itself."))
    if profile.generator_profile and (
        not profile.generator_profile.profile_id or not profile.generator_profile.profile_version
    ):
        issues.append(ProfileIssue("generator_profile_invalid", "Generator profile reference is incomplete."))
    if profile.status == "approved":
        if not profile.rfc or not _RFC_RE.fullmatch(profile.rfc):
            issues.append(ProfileIssue("approved_rfc_required", "Approved profiles require a valid uppercase RFC."))
        if profile.approval is None:
            issues.append(ProfileIssue("approved_approval_required", "Approved profiles require approval evidence."))
        elif not (profile.approval.approved_by and profile.approval.approval_reference):
            issues.append(ProfileIssue("approval_evidence_incomplete", "Approval evidence is incomplete."))
    elif profile.approval is not None:
        issues.append(ProfileIssue("approval_for_nonapproved_profile", "Only approved profiles may carry approval evidence."))

    issues.extend(_validate_enabled_lines(profile))
    issues.extend(_validate_rules(profile))
    return tuple(issues)


def require_profile_for_runtime(
    profile: AccountingProfile,
    *,
    rfc: str,
    as_of: date,
    catalog_identity: CatalogIdentity,
    taxonomy: OfflineSatTaxonomy | None = None,
) -> None:
    """Raise structured validation errors instead of selecting an unsafe profile."""

    issues = list(validate_accounting_profile(profile, taxonomy=taxonomy))
    if profile.status != "approved":
        issues.append(ProfileIssue("profile_not_approved", "Runtime selection requires an approved profile."))
    if profile.rfc != rfc:
        issues.append(ProfileIssue("profile_rfc_mismatch", "Profile RFC does not match the requested RFC."))
    if profile.valid_from and as_of < profile.valid_from or profile.valid_to and as_of > profile.valid_to:
        issues.append(ProfileIssue("profile_not_effective", "Profile is outside its effective validity range."))
    if profile.catalog_identity.source_sha256 != catalog_identity.source_sha256:
        issues.append(ProfileIssue("catalog_source_hash_mismatch", "Catalog source hash differs from the profile."))
    if profile.catalog_identity.semantic_sha256 != catalog_identity.semantic_sha256:
        issues.append(ProfileIssue("catalog_semantic_hash_mismatch", "Catalog semantic hash differs from the profile."))
    if issues:
        raise ProfileValidationError(issues)


def resolve_account(
    profile: AccountingProfile,
    account: AccountClassificationInput,
    *,
    taxonomy: OfflineSatTaxonomy | None = None,
) -> Resolution:
    """Resolve one account using explicit evidence, never its leading digit alone."""

    taxonomy = taxonomy or load_sat_rmf_2026_v1()
    profile_issues = validate_accounting_profile(profile, taxonomy=taxonomy)
    if profile_issues:
        raise ProfileValidationError(profile_issues)

    stages = (
        ("approved_override", ("exact_code", "subtree")),
        ("sat_exact", ("sat_exact",)),
        ("sat_family", ("sat_family",)),
        ("context", ("context",)),
    )
    for stage, kinds in stages:
        candidates = _profile_candidates(profile, account, kinds)
        if stage == "sat_exact":
            candidates.extend(_taxonomy_candidates(profile, taxonomy, account, exact=True))
        elif stage == "sat_family":
            candidates.extend(_taxonomy_candidates(profile, taxonomy, account, exact=False))
        result = _resolution_from_candidates(stage, candidates)
        if result is not None:
            return result

    suggestions = tuple(
        sorted(
            {
                rule.line_key
                for rule in profile.rules
                if rule.kind == "name_suggestion" and _rule_matches(rule, account)
            }
        )
    )
    return Resolution(
        status="pending_classification",
        line_keys=(),
        precedence=None,
        evidence=(),
        suggestions=suggestions,
    )


def _validate_enabled_lines(profile: AccountingProfile) -> list[ProfileIssue]:
    issues: list[ProfileIssue] = []
    for statement, lines in profile.enabled_lines.items():
        if statement not in STATEMENTS:
            issues.append(ProfileIssue("enabled_lines_statement_invalid", f"Unknown statement {statement}."))
        if not lines:
            issues.append(ProfileIssue("enabled_lines_empty", f"No enabled lines for {statement}."))
        if len(set(lines)) != len(lines):
            issues.append(ProfileIssue("enabled_line_duplicate", f"Duplicate enabled line for {statement}."))
    for statement in STATEMENTS:
        if statement not in profile.enabled_lines:
            issues.append(ProfileIssue("enabled_lines_missing", f"enabled_lines must include {statement}."))
    return issues


def _validate_rules(profile: AccountingProfile) -> list[ProfileIssue]:
    issues: list[ProfileIssue] = []
    seen_rule_ids: set[str] = set()
    selectors: dict[tuple[str, str, str], list[MappingRule]] = {}
    for rule in profile.rules:
        if not rule.rule_id or rule.rule_id in seen_rule_ids:
            issues.append(ProfileIssue("rule_id_duplicate", "Rule IDs must be unique and non-empty.", (rule.rule_id,)))
        seen_rule_ids.add(rule.rule_id)
        if rule.statement not in STATEMENTS:
            issues.append(ProfileIssue("rule_statement_invalid", f"Invalid rule statement {rule.statement}.", (rule.rule_id,)))
        if rule.kind not in RULE_KINDS:
            issues.append(ProfileIssue("rule_kind_invalid", f"Invalid rule kind {rule.kind}.", (rule.rule_id,)))
            continue
        if rule.line_key not in profile.enabled_lines.get(rule.statement, ()):
            issues.append(ProfileIssue("rule_line_not_enabled", "Rule line must be enabled by the profile.", (rule.rule_id,)))
        if rule.kind in {"exact_code", "subtree"}:
            if not rule.account_code:
                issues.append(ProfileIssue("override_account_code_missing", "Code/subtree override requires account_code.", (rule.rule_id,)))
            if not rule.approved:
                issues.append(ProfileIssue("override_not_approved", "Code/subtree override requires explicit approval.", (rule.rule_id,)))
        if rule.kind == "sat_exact" and not rule.sat_group_code or rule.kind == "sat_exact" and not _SAT_EXACT_RE.fullmatch(rule.sat_group_code or ""):
            issues.append(ProfileIssue("sat_exact_invalid", "sat_exact requires a valid SAT group code.", (rule.rule_id,)))
        if rule.kind == "sat_family" and not _SAT_FAMILY_RE.fullmatch(rule.sat_group_code or ""):
            issues.append(ProfileIssue("sat_family_invalid", "sat_family requires a three-digit SAT family.", (rule.rule_id,)))
        if rule.kind == "context" and not (rule.parent_code or rule.nature):
            issues.append(ProfileIssue("context_without_evidence", "Context needs parent_code or nature; section_guard alone cannot resolve.", (rule.rule_id,)))
        if rule.kind == "name_suggestion" and not rule.normalized_name:
            issues.append(ProfileIssue("name_suggestion_missing", "name_suggestion requires normalized_name.", (rule.rule_id,)))
        if rule.nature is not None and rule.nature not in {"A", "D"}:
            issues.append(ProfileIssue("rule_nature_invalid", "Rule nature must be A or D.", (rule.rule_id,)))
        if rule.section_guard is not None and rule.section_guard not in {"1", "2", "3"}:
            issues.append(ProfileIssue("section_guard_invalid", "section_guard must be 1, 2 or 3.", (rule.rule_id,)))
        if rule.presentation_sign not in {-1, 1}:
            issues.append(ProfileIssue("presentation_sign_invalid", "presentation_sign must be 1 or -1.", (rule.rule_id,)))
        if rule.resolution_policy != "exact_accumulator_then_leaf_fallback":
            issues.append(
                ProfileIssue(
                    "resolution_policy_invalid",
                    "Only the approved exact-accumulator then leaf-fallback policy is supported.",
                    (rule.rule_id,),
                )
            )
        selector = _rule_selector(rule)
        if selector:
            selectors.setdefault(selector, []).append(rule)

    for selector, rules in selectors.items():
        line_keys = {rule.line_key for rule in rules}
        if len(line_keys) > 1 and not all(rule.composite for rule in rules):
            issues.append(
                ProfileIssue(
                    "double_assignment_rule",
                    f"Selector {selector} assigns more than one line without an explicit composite rule.",
                    tuple(rule.rule_id for rule in rules),
                )
            )
    return issues


def _profile_candidates(
    profile: AccountingProfile, account: AccountClassificationInput, kinds: tuple[str, ...]
) -> list[tuple[str, str, bool, str]]:
    candidates: list[tuple[str, str, bool, str]] = []
    for rule in profile.rules:
        if rule.kind in kinds and _rule_matches(rule, account):
            candidates.append((rule.statement, rule.line_key, rule.composite, f"profile:{rule.rule_id}"))
    return candidates


def _taxonomy_candidates(
    profile: AccountingProfile,
    taxonomy: OfflineSatTaxonomy,
    account: AccountClassificationInput,
    *,
    exact: bool,
) -> list[tuple[str, str, bool, str]]:
    sat_code = account.sat_group_code if exact else _sat_family(account.sat_group_code)
    if not sat_code:
        return []
    candidates = []
    for entry in taxonomy.entries:
        if entry.sat_group_code == sat_code and entry.line_key in profile.enabled_lines.get(entry.statement, ()):
            candidates.append((entry.statement, entry.line_key, False, f"taxonomy:{entry.sat_group_code}"))
    return candidates


def _resolution_from_candidates(
    stage: str, candidates: list[tuple[str, str, bool, str]]
) -> Resolution | None:
    if not candidates:
        return None
    assignments = {(statement, line_key) for statement, line_key, _, _ in candidates}
    evidence = tuple(sorted({evidence for _, _, _, evidence in candidates}))
    if len(assignments) == 1:
        return Resolution("assigned", tuple(sorted(line_key for _, line_key in assignments)), stage, evidence)
    if all(composite for _, _, composite, _ in candidates):
        return Resolution("assigned", tuple(sorted(line_key for _, line_key in assignments)), stage, evidence)
    return Resolution("ambiguous_mapping", (), stage, evidence)


def _rule_matches(rule: MappingRule, account: AccountClassificationInput) -> bool:
    if not _matches_section_guard(rule.section_guard, account.account_code):
        return False
    if rule.kind == "exact_code":
        return rule.account_code == account.account_code
    if rule.kind == "subtree":
        return bool(rule.account_code) and (
            account.account_code == rule.account_code or account.account_code.startswith(f"{rule.account_code}-")
        )
    if rule.kind == "sat_exact":
        return rule.sat_group_code == account.sat_group_code
    if rule.kind == "sat_family":
        return rule.sat_group_code == _sat_family(account.sat_group_code)
    if rule.kind == "context":
        return (not rule.parent_code or rule.parent_code == account.parent_code) and (
            not rule.nature or rule.nature == account.nature
        )
    if rule.kind == "name_suggestion":
        return rule.normalized_name == normalize_account_name(account.account_name)
    return False


def _matches_section_guard(section_guard: str | None, account_code: str) -> bool:
    return section_guard is None or account_code.startswith(section_guard)


def _rule_selector(rule: MappingRule) -> tuple[str, str, str] | None:
    if rule.kind in {"exact_code", "subtree"} and rule.account_code:
        return (rule.kind, rule.account_code, rule.section_guard or "")
    if rule.kind in {"sat_exact", "sat_family"} and rule.sat_group_code:
        return (rule.kind, rule.sat_group_code, rule.section_guard or "")
    if rule.kind == "context" and (rule.parent_code or rule.nature):
        return ("context", f"{rule.parent_code or ''}|{rule.nature or ''}", rule.section_guard or "")
    return None


def _sat_family(sat_group_code: str) -> str | None:
    if not _SAT_EXACT_RE.fullmatch(sat_group_code):
        return None
    return sat_group_code.split(".", 1)[0]


def normalize_account_name(value: str) -> str:
    """Normalize only for a non-final suggestion; source names remain unchanged."""

    decomposed = unicodedata.normalize("NFKD", value)
    without_marks = "".join(char for char in decomposed if not unicodedata.combining(char))
    return " ".join(without_marks.upper().split())


def _parse_date(value: str | None) -> date | None:
    return date.fromisoformat(value) if value else None
