"""Synthetic-only catalog guardrails; these fixtures do not validate real balances.

They exercise parser, hierarchy, profile selection, and blocking behavior with
invented account values.  No assertion in this module is evidence of financial
validation for a company catalog without a representative balanza.
"""

from __future__ import annotations

import unittest
from dataclasses import replace
from datetime import date
import json
from pathlib import Path

from src.account_catalog import (
    BLOCKING,
    OBSERVATION,
    WARNING,
    compare_catalog_semantics,
    parse_account_catalog_bytes,
)
from src.accounting_profiles import (
    AccountClassificationInput,
    Approval,
    CatalogIdentity,
    MappingRule,
    ProfileValidationError,
    require_profile_for_runtime,
    resolve_account,
    profile_from_dict,
    load_sat_rmf_2026_v1,
)
from src.catalog_profile_draft import ParallelCompanyEvidence, controlled_parallel_readiness
from src.engine import build_bg_dataset, build_er_dataset, build_profile_coverage, require_profile_coverage
from tests.test_accounting_profiles import make_profile


ROOT = Path(__file__).resolve().parents[1]
AL_PROFILE = ROOT / "src" / "profiles" / "SME170717GA0-2026-07-v1.json"


class SyntheticCatalogGuardrailsTest(unittest.TestCase):
    def test_new_renamed_and_semantically_reused_code_do_not_inherit_a_mapping(self) -> None:
        previous = parse_account_catalog_bytes(
            b"A-100,OLD LABEL,ROOT,D,101\n",
            source_name="synthetic-previous.csv",
        )
        current = parse_account_catalog_bytes(
            b"A-100,RENAMED LABEL,ROOT,D,101\n"
            b"B-200,NEW LABEL,ROOT,D,102\n"
            b"C-300,OLD MEANING,ROOT,D,101\n",
            source_name="synthetic-current.csv",
        )
        semantic_reuse = parse_account_catalog_bytes(
            b"C-300,NEW MEANING,OTHER,A,201\n",
            source_name="synthetic-reused.csv",
        )
        original_c300 = parse_account_catalog_bytes(b"C-300,OLD MEANING,ROOT,D,101\n")

        delta_codes = {issue.code: issue.severity for issue in compare_catalog_semantics(previous, current)}
        reuse_codes = {issue.code: issue.severity for issue in compare_catalog_semantics(original_c300, semantic_reuse)}

        self.assertEqual(delta_codes["renamed_account_code"], WARNING)
        self.assertEqual(delta_codes["new_account_code"], OBSERVATION)
        self.assertEqual(reuse_codes["account_code_semantics_changed"], BLOCKING)

    def test_orphan_sat_unknown_and_incompatible_nature_remain_unresolved_or_blocked(self) -> None:
        orphan = parse_account_catalog_bytes(b"A-100,ORPHAN,MISSING,D,101\n")
        self.assertIn("orphan_parent_code", {issue.code for issue in orphan.issues})

        context_profile = make_profile(
            rules=(
                MappingRule("context-cash", "BG", "bg.cash", "context", parent_code="ROOT", nature="D"),
            )
        )
        sat_unknown = resolve_account(make_profile(), AccountClassificationInput("A-200", "SYNTH", "ROOT", "D", "999"))
        base_taxonomy = load_sat_rmf_2026_v1()
        no_sat_fallback = type(base_taxonomy)(
            base_taxonomy.schema_version, base_taxonomy.taxonomy_id, base_taxonomy.source, ()
        )
        nature_incompatible = resolve_account(
            context_profile,
            AccountClassificationInput("A-201", "SYNTH", "ROOT", "A", "101"),
            taxonomy=no_sat_fallback,
        )

        self.assertEqual(sat_unknown.status, "pending_classification")
        self.assertEqual(nature_incompatible.status, "pending_classification")

    def test_contra_asset_result_and_duplicate_material_mapping_keep_explicit_controls(self) -> None:
        profile = profile_from_dict(json.loads(AL_PROFILE.read_text(encoding="utf-8")))
        bg = build_bg_dataset(({"account_code": "1215", "saldo_final": "100"},), profile=profile)
        er = build_er_dataset(({"account_code": "4110", "saldo_final": "250"},), profile=profile)
        duplicate = replace(
            profile,
            rules=profile.rules + (
                MappingRule("duplicate-asset", "BG", "bancos", "exact_code", account_code="1215", approved=True),
            ),
        )
        coverage = build_profile_coverage(({"account_code": "1215", "saldo_final": "100"},), duplicate)

        self.assertEqual({line["key"]: line["amount"] for line in bg["lines"]}["depreciacion_acumulada"], -100.0)
        self.assertEqual({line["key"]: line["amount"] for line in er["lines"]}["ingresos_por_servicios"], 250.0)
        self.assertTrue(any(issue["code"] == "profile_mapping_duplicate" for issue in coverage["blockers"]))
        with self.assertRaises(Exception):
            require_profile_coverage(coverage)

    def test_expired_profile_and_less_than_three_reviewed_companies_block_parallel_pilot(self) -> None:
        approved = make_profile(
            status="approved",
            rfc="ABC010101AAA",
            approval=Approval("reviewer", date(2026, 1, 1), "synthetic-review"),
            valid_to=date(2026, 1, 31),
        )
        with self.assertRaises(ProfileValidationError) as context:
            require_profile_for_runtime(
                approved,
                rfc=approved.rfc,
                as_of=date(2026, 2, 1),
                catalog_identity=approved.catalog_identity,
            )
        self.assertIn("profile_not_effective", {issue.code for issue in context.exception.issues})

        readiness = controlled_parallel_readiness((
            ParallelCompanyEvidence(approved, representative_balance_verified=True, accounting_review_completed=True),
            ParallelCompanyEvidence(replace(approved, rfc="DEF010101AAA", profile_id="synthetic-2"), True, True),
        ))
        self.assertFalse(readiness["allowed"])
        self.assertEqual(readiness["qualified_companies"], 2)
        self.assertIn("minimum_three_independently_qualified_companies_required", readiness["pending"])


if __name__ == "__main__":
    unittest.main()
