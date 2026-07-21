from __future__ import annotations

import unittest
from datetime import date

from src.accounting_profiles import (
    ACCOUNTING_PROFILE_SCHEMA_VERSION,
    SAT_RMF_2026_V1,
    AccountingProfile,
    AccountClassificationInput,
    Approval,
    CatalogIdentity,
    MappingRule,
    ProfileValidationError,
    load_sat_rmf_2026_v1,
    profile_from_dict,
    require_profile_for_runtime,
    resolve_account,
    validate_accounting_profile,
)


SOURCE_HASH = "a" * 64
SEMANTIC_HASH = "b" * 64


def make_profile(*, status: str = "draft", rules: tuple[MappingRule, ...] = (), **changes) -> AccountingProfile:
    values = {
        "profile_id": "example-company",
        "profile_version": "2026.1",
        "status": status,
        "company_name": "Example Company",
        "rfc": None,
        "valid_from": date(2026, 1, 1),
        "valid_to": None,
        "catalog_identity": CatalogIdentity(SOURCE_HASH, SEMANTIC_HASH),
        "taxonomy_version": SAT_RMF_2026_V1,
        "enabled_lines": {
            "BG": ("bg.cash", "bg.banks", "bg.accounts_receivable", "bg.fixed_assets"),
            "ER": ("er.revenue", "er.cost_of_sales", "er.operating_expenses", "er.financial_expenses"),
        },
        "rules": rules,
        "approval": None,
        "schema_version": ACCOUNTING_PROFILE_SCHEMA_VERSION,
    }
    values.update(changes)
    return AccountingProfile(**values)


class AccountingProfileTest(unittest.TestCase):
    def test_offline_taxonomy_is_versioned_and_does_not_need_network(self) -> None:
        taxonomy = load_sat_rmf_2026_v1()

        self.assertEqual(taxonomy.taxonomy_id, SAT_RMF_2026_V1)
        self.assertTrue(taxonomy.entries)
        self.assertEqual({entry.statement for entry in taxonomy.entries}, {"BG", "ER"})

    def test_approved_profile_requires_rfc_and_approval_evidence(self) -> None:
        profile = make_profile(status="approved")

        codes = {issue.code for issue in validate_accounting_profile(profile)}

        self.assertIn("approved_rfc_required", codes)
        self.assertIn("approved_approval_required", codes)

    def test_approved_profile_serializes_and_validates(self) -> None:
        profile = make_profile(
            status="approved",
            rfc="ABC010101AAA",
            approval=Approval("contador.revisor", date(2026, 1, 3), "acta-2026-01"),
        )

        rebuilt = profile_from_dict(profile.to_dict())

        self.assertEqual(validate_accounting_profile(rebuilt), ())
        self.assertEqual(rebuilt.to_dict()["catalog_identity"]["source_sha256"], SOURCE_HASH)

    def test_runtime_rejects_draft_mismatched_hash_and_outside_validity(self) -> None:
        profile = make_profile()

        with self.assertRaises(ProfileValidationError) as context:
            require_profile_for_runtime(
                profile,
                rfc="ABC010101AAA",
                as_of=date(2025, 12, 31),
                catalog_identity=CatalogIdentity("c" * 64, SEMANTIC_HASH),
            )

        codes = {issue.code for issue in context.exception.issues}
        self.assertIn("profile_not_approved", codes)
        self.assertIn("profile_rfc_mismatch", codes)
        self.assertIn("profile_not_effective", codes)
        self.assertIn("catalog_source_hash_mismatch", codes)

    def test_approved_code_override_wins_over_sat_taxonomy(self) -> None:
        profile = make_profile(
            rules=(
                MappingRule(
                    "override-bank-special",
                    "BG",
                    "bg.cash",
                    "exact_code",
                    account_code="1120-SPECIAL",
                    approved=True,
                ),
            )
        )
        account = AccountClassificationInput("1120-SPECIAL", "BANCO", "1120", "D", "102.01")

        resolved = resolve_account(profile, account)

        self.assertEqual(resolved.status, "assigned")
        self.assertEqual(resolved.line_keys, ("bg.cash",))
        self.assertEqual(resolved.precedence, "approved_override")

    def test_sat_exact_and_family_precede_context_without_using_account_first_digit(self) -> None:
        profile = make_profile(
            rules=(
                MappingRule("sat-exact", "BG", "bg.banks", "sat_exact", sat_group_code="102.01"),
                MappingRule(
                    "context-only",
                    "BG",
                    "bg.cash",
                    "context",
                    parent_code="1120",
                    nature="D",
                    section_guard="1",
                ),
            )
        )
        account = AccountClassificationInput("9999-ALT", "CUENTA", "1120", "D", "102.01")

        resolved = resolve_account(profile, account)

        self.assertEqual(resolved.line_keys, ("bg.banks",))
        self.assertEqual(resolved.precedence, "sat_exact")

    def test_section_guard_alone_cannot_be_a_final_rule(self) -> None:
        profile = make_profile(
            rules=(MappingRule("bad-guard", "BG", "bg.cash", "context", section_guard="1"),)
        )

        codes = {issue.code for issue in validate_accounting_profile(profile)}

        self.assertIn("context_without_evidence", codes)

    def test_name_is_suggestion_only_then_pending_classification(self) -> None:
        profile = make_profile(
            rules=(
                MappingRule(
                    "name-only",
                    "BG",
                    "bg.fixed_assets",
                    "name_suggestion",
                    normalized_name="EQUIPO DE COMPUTO",
                ),
            )
        )

        resolved = resolve_account(
            profile,
            AccountClassificationInput("1220-X", "Equipo de Cómputo", "1220", "D", "999.99"),
        )

        self.assertEqual(resolved.status, "pending_classification")
        self.assertEqual(resolved.suggestions, ("bg.fixed_assets",))

    def test_ambiguous_and_double_assignment_rules_are_blocked(self) -> None:
        rules = (
            MappingRule("dup-one", "BG", "bg.cash", "exact_code", account_code="1110-X", approved=True),
            MappingRule("dup-two", "BG", "bg.banks", "exact_code", account_code="1110-X", approved=True),
        )
        profile = make_profile(rules=rules)

        codes = {issue.code for issue in validate_accounting_profile(profile)}

        self.assertIn("double_assignment_rule", codes)
        with self.assertRaises(ProfileValidationError):
            resolve_account(profile, AccountClassificationInput("1110-X", "CAJA", "1110", "D", "101"))

    def test_ambiguous_taxonomy_entries_do_not_silently_assign(self) -> None:
        taxonomy = load_sat_rmf_2026_v1()
        profile = make_profile()
        duplicate_entry = taxonomy.entries[0]
        ambiguous_taxonomy = type(taxonomy)(
            taxonomy.schema_version,
            taxonomy.taxonomy_id,
            taxonomy.source,
            taxonomy.entries + (type(duplicate_entry)("101", "BG", "bg.banks", "duplicate"),),
        )

        with self.assertRaises(ProfileValidationError) as context:
            resolve_account(profile, AccountClassificationInput("1110", "CAJA", "", "D", "101"), taxonomy=ambiguous_taxonomy)

        self.assertIn("taxonomy_duplicate_entry", {issue.code for issue in context.exception.issues})


if __name__ == "__main__":
    unittest.main()
