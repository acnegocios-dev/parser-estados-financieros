from __future__ import annotations

import json
import unittest
from pathlib import Path

from src.accounting_profiles import (
    AccountClassificationInput,
    profile_from_dict,
    resolve_account,
    validate_accounting_profile,
)
from src.engine import build_bg_dataset, build_er_dataset
from src.parser import parse_balanza


ROOT = Path(__file__).resolve().parents[1]
PROFILE_PATH = ROOT / "src" / "profiles" / "SME170717GA0-2026-07-v1.json"
SAMPLE = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"


def load_profile():
    return profile_from_dict(json.loads(PROFILE_PATH.read_text(encoding="utf-8")))


def rules_by_code(profile, statement: str) -> dict[str, tuple[str, int]]:
    return {
        rule.account_code: (rule.line_key, rule.presentation_sign)
        for rule in profile.rules
        if rule.statement == statement and rule.account_code
    }


class SmeAccountingProfileTest(unittest.TestCase):
    def test_profile_is_approved_bound_to_the_expected_catalog_and_generator(self) -> None:
        profile = load_profile()

        self.assertEqual(validate_accounting_profile(profile), ())
        self.assertEqual(profile.status, "approved")
        self.assertEqual(profile.rfc, "SME170717GA0")
        self.assertEqual(
            profile.catalog_identity.source_sha256,
            "bc6ea8f397689550d6d19cfef01eeb1cbb9d74405c74684d2a90e5ac11a1d419",
        )
        self.assertEqual(profile.generator_profile.profile_id, "manual-eeff-three-sheet")
        self.assertEqual(profile.generator_profile.profile_version, "2026-07-20.exact-style-print-v2")

    def test_bg_rules_are_company_overlay_rules_with_depreciation_negative_and_2120_once(self) -> None:
        profile = load_profile()
        rules = rules_by_code(profile, "BG")

        self.assertEqual(rules["2120"], ("acreedores_diversos", 1))
        self.assertEqual(sum(code == "2120" for code in rules), 1)
        self.assertEqual(
            {code for code, (line, sign) in rules.items() if line == "depreciacion_acumulada" and sign == -1},
            {"1215", "1225", "1235", "1245"},
        )
        self.assertEqual(rules["3110"], ("aportaciones_para_aumentos_de_capital", 1))
        self.assertEqual(rules["1230"], ("equipo_de_transporte", 1))
        self.assertEqual(profile.regression_controls["legacy_2120_division"], "prohibited")
        self.assertEqual(profile.regression_controls["legacy_2120_cent_adjustment"], "prohibited")

    def test_er_rules_retain_varios_and_financial_presentation_sign(self) -> None:
        profile = load_profile()
        rules = rules_by_code(profile, "ER")
        varios = {code for code, (line, _) in rules.items() if line == "varios"}

        self.assertEqual(varios, {"6148", "6176", "6195"})
        self.assertEqual(rules["6410"], ("gastos_financieros", -1))
        self.assertEqual(rules["4110"], ("ingresos_por_servicios", 1))
        self.assertEqual(rules["4110-9999"], ("otros_productos", 1))
        ingresos = next(rule for rule in profile.rules if rule.rule_id == "er-4110")
        self.assertEqual(ingresos.exclude_account_codes, ("4110-9999",))

    def test_profile_resolves_only_its_explicit_al_services_overrides(self) -> None:
        profile = load_profile()

        resolved = resolve_account(
            profile,
            AccountClassificationInput("2120", "ACREEDORES", "", "A", "205"),
        )
        pending = resolve_account(
            profile,
            AccountClassificationInput("2121", "NUEVA CUENTA", "", "A", "205"),
        )

        self.assertEqual(resolved.line_keys, ("acreedores_diversos",))
        self.assertEqual(resolved.precedence, "approved_override")
        self.assertEqual(pending.status, "pending_classification")

    def test_current_engine_regression_matches_profile_controls_to_the_cent(self) -> None:
        profile = load_profile()
        parsed = parse_balanza(SAMPLE)
        er = build_er_dataset(parsed.rows)
        bg = build_bg_dataset(parsed.rows, result_ejercicio=er["raw_amounts"]["resultado_ejercicio"])

        self.assertAlmostEqual(float(er["raw_amounts"]["varios"]), float(profile.regression_controls["er_h46"]), places=2)
        self.assertAlmostEqual(
            float(er["raw_amounts"]["resultado_ejercicio"]),
            float(profile.regression_controls["er_h70"]),
            places=2,
        )
        self.assertAlmostEqual(bg["total_activo"], 2654692.85, places=2)
        self.assertAlmostEqual(bg["total_pasivo"], 11900467.83, places=2)
        self.assertAlmostEqual(bg["diferencia_cuadre"], -0.059663, places=6)
        self.assertTrue(bg["cuadra"])


if __name__ == "__main__":
    unittest.main()
