from __future__ import annotations

import json
import unittest
from dataclasses import replace
from decimal import Decimal
from pathlib import Path

from src.account_catalog import AccountCatalogRow
from src.accounting_profiles import AccountClassificationInput, MappingRule, profile_from_dict, resolve_account
from src.engine import (
    ProfileCoverageError,
    build_bg_dataset,
    build_er_dataset,
    build_profile_coverage,
    require_profile_coverage,
    resolve_account_code,
)
from src.parser import BalanzaRow, enrich_balanza_rows


ROOT = Path(__file__).resolve().parents[1]
PROFILE_PATH = ROOT / "src" / "profiles" / "SME170717GA0-2026-07-v1.json"


def profile():
    return profile_from_dict(json.loads(PROFILE_PATH.read_text(encoding="utf-8")))


class ParserAndHierarchyRefactorTest(unittest.TestCase):
    def test_catalog_enrichment_preserves_parent_nature_sat_and_match_evidence(self) -> None:
        balance = BalanzaRow(7, "AB-01 Cuenta", "AB-01", "Cuenta", "AB", Decimal(0), Decimal(0), Decimal(0), Decimal("12"))
        catalog = AccountCatalogRow(1, "AB-01", "Cuenta", "ROOT-A", "D", "156")

        enriched = enrich_balanza_rows((balance,), (catalog,))[0]

        self.assertEqual(enriched.parent_code, "ROOT-A")
        self.assertEqual(enriched.nature, "D")
        self.assertEqual(enriched.sat_group_code, "156")
        self.assertEqual(enriched.catalog_match, "matched")

    def test_parent_code_is_primary_even_when_codes_are_not_prefix_related(self) -> None:
        rows = [
            {"source_row": 1, "account_code": "ACTIVO", "saldo_final": "100"},
            {"source_row": 2, "account_code": "A-01", "parent_code": "ACTIVO", "saldo_final": "40"},
            {"source_row": 3, "account_code": "A-02", "parent_code": "ACTIVO", "saldo_final": "60"},
        ]

        resolved = resolve_account_code(rows, "ACTIVO")

        self.assertEqual(resolved["policy"], "exact_accumulator")
        self.assertEqual(resolved["hierarchy_method"], "parent_code")
        self.assertEqual(resolved["leaf_amount"], Decimal("100"))


class ProfileMappingRefactorTest(unittest.TestCase):
    def test_1220_meanings_are_resolved_by_profile_evidence_not_first_digit(self) -> None:
        base = profile()
        custom = replace(
            base,
            rules=(
                MappingRule("cash-1220", "BG", "caja_chica", "context", parent_code="EFECTIVO", nature="D", section_guard="1"),
                MappingRule("computer-1220", "BG", "equipo_de_computo", "sat_exact", sat_group_code="156"),
                MappingRule("tax-1220", "BG", "contribuciones", "sat_exact", sat_group_code="118"),
            ),
        )
        cases = (
            (AccountClassificationInput("1220", "Caja", "EFECTIVO", "D", "999"), "caja_chica"),
            (AccountClassificationInput("1220", "Equipo", "ACTIVO", "D", "156"), "equipo_de_computo"),
            (AccountClassificationInput("1220", "Impuesto", "ACTIVO", "D", "118"), "contribuciones"),
        )

        self.assertEqual([resolve_account(custom, account).line_keys[0] for account, _ in cases], [line for _, line in cases])

    def test_1180_and_1197_and_tax_213_216_do_not_share_prefix_or_section_assignment(self) -> None:
        active = profile()
        rows = [
            {"account_code": "1180-01", "saldo_final": "10"},
            {"account_code": "1197-01", "saldo_final": "20"},
            {"account_code": "2130-01", "saldo_final": "30"},
            {"account_code": "2160-01", "saldo_final": "40"},
        ]
        bg = build_bg_dataset(rows, profile=active)
        amounts = {line["key"]: line["amount"] for line in bg["lines"]}

        self.assertEqual(amounts["anticipo_a_proveedores"], 20.0)
        self.assertEqual(amounts["anticipos_de_clientes"], 30.0)
        self.assertEqual(amounts["otros_pasivos_ptu"], 40.0)
        self.assertNotIn("1180", {code for line in bg["lines"] for code in line["account_codes"]})

    def test_3110_6410_sat701_and_internal_isr_ptu_use_explicit_profile_rules(self) -> None:
        active = profile()
        rows = [
            {"account_code": "3110", "saldo_final": "250"},
            {"account_code": "6410-01", "saldo_final": "-20", "sat_group_code": "701"},
            {"account_code": "6510-0001", "saldo_final": "7"},
            {"account_code": "6510-0002", "saldo_final": "8"},
        ]
        er = build_er_dataset(rows, profile=active)
        bg = build_bg_dataset(rows, profile=active)
        er_amounts = {line["key"]: line["amount"] for line in er["lines"]}
        bg_amounts = {line["key"]: line["amount"] for line in bg["lines"]}

        self.assertEqual(bg_amounts["aportaciones_para_aumentos_de_capital"], 250.0)
        self.assertEqual(er_amounts["gastos_financieros"], 20.0)
        self.assertEqual(er_amounts["isr_del_ejercicio"], 7.0)
        self.assertEqual(er_amounts["ptu_del_ejercicio"], 8.0)

    def test_material_gap_duplicate_and_one_peso_section_difference_block_but_zero_gap_warns(self) -> None:
        active = profile()
        missing = build_profile_coverage(({"account_code": "1999", "saldo_final": "2"},), active)
        zero_missing = build_profile_coverage(({"account_code": "1999", "saldo_final": "0"},), active)
        duplicate_profile = replace(
            active,
            rules=active.rules + (
                MappingRule("duplicate-1110", "BG", "bancos", "exact_code", account_code="1110", approved=True),
            ),
        )
        duplicate = build_profile_coverage(({"account_code": "1110", "saldo_final": "2"},), duplicate_profile)

        with self.assertRaises(ProfileCoverageError):
            require_profile_coverage(missing)
        with self.assertRaises(ProfileCoverageError):
            require_profile_coverage(duplicate)
        self.assertEqual(zero_missing["warnings"][0]["code"], "profile_mapping_missing_zero")
        self.assertTrue(any(item["code"] == "profile_section_coverage_difference" for item in missing["blockers"]))


if __name__ == "__main__":
    unittest.main()
