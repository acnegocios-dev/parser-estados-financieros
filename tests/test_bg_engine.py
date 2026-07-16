from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import unittest

from src.engine import (
    build_bal_dataset,
    build_bg_dataset,
    build_er_dataset,
    build_input_views,
    resolve_account_code,
)
from src.parser import parse_balanza
from src.validation import validate_balance_sheet


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"


def bg_line(dataset: dict, key: str) -> dict:
    return next(line for line in dataset["lines"] if line["key"] == key)


class AccountResolverTest(unittest.TestCase):
    def test_preserves_all_rows_for_bal_and_leaf_rows_for_calculation(self) -> None:
        rows = [
            {"source_row": 10, "account_code": "6110", "account_raw": "6110 SUELDOS", "saldo_final": "100"},
            {"source_row": 11, "account_code": "6110-0001", "account_raw": "6110-0001 A", "saldo_final": "40"},
            {"source_row": 12, "account_code": "6110-0002", "account_raw": "6110-0002 B", "saldo_final": "60"},
        ]

        views = build_input_views(rows)
        bal = build_bal_dataset(rows)

        self.assertEqual([row["source_row"] for row in views["all_rows"]], [10, 11, 12])
        self.assertEqual([row["source_row"] for row in views["calculation_rows"]], [11, 12])
        self.assertEqual([row["source_row"] for row in bal["rows"]], [10, 11, 12])
        self.assertEqual(bal["sumas_iguales"]["accumulator_source_rows"], [10])

    def test_exact_accumulator_is_used_without_double_counting_descendants(self) -> None:
        rows = [
            {"account_code": "2120", "saldo_final": "100"},
            {"account_code": "2120-0001", "saldo_final": "40"},
            {"account_code": "2120-0002", "saldo_final": "60"},
        ]

        resolution = resolve_account_code(rows, "2120")

        self.assertEqual(resolution["policy"], "exact_accumulator")
        self.assertEqual(resolution["amount"], Decimal("100"))
        self.assertEqual(resolution["accumulator_amount"], Decimal("100"))
        self.assertEqual(resolution["leaf_amount"], Decimal("100"))
        self.assertFalse(resolution["aggregate_detail_mismatch"])

    def test_missing_accumulator_falls_back_to_leaf_descendants(self) -> None:
        rows = [
            {"account_code": "1130-0001", "saldo_final": "40"},
            {"account_code": "1130-0002", "saldo_final": "60"},
        ]

        resolution = resolve_account_code(rows, "1130")

        self.assertEqual(resolution["policy"], "leaf_fallback")
        self.assertEqual(resolution["amount"], Decimal("100"))
        self.assertIsNone(resolution["accumulator_amount"])
        self.assertEqual(resolution["leaf_amount"], Decimal("100"))

    def test_mismatch_is_structured_and_keeps_exact_accumulator_policy(self) -> None:
        rows = [
            {"account_code": "1110", "saldo_final": "100"},
            {"account_code": "1110-0001", "saldo_final": "40"},
            {"account_code": "1110-0002", "saldo_final": "50"},
        ]

        dataset = build_bg_dataset(rows, tolerance=1)
        warning = next(item for item in dataset["warnings"] if item["code"] == "aggregate_detail_mismatch")

        self.assertEqual(bg_line(dataset, "caja_chica")["amount"], 100.0)
        self.assertEqual(warning["account_code"], "1110")
        self.assertEqual(warning["accumulator"], 100.0)
        self.assertEqual(warning["leaf_sum"], 90.0)
        self.assertEqual(warning["difference"], 10.0)
        self.assertEqual(warning["policy"], "exact_accumulator")


class BgAndBalDatasetTest(unittest.TestCase):
    def test_acreedores_diversos_is_included_once(self) -> None:
        rows = [
            {"account_code": "2120", "saldo_final": "100"},
            {"account_code": "2120-0001", "saldo_final": "40"},
            {"account_code": "2120-0002", "saldo_final": "60"},
        ]

        dataset = build_bg_dataset(rows)

        self.assertEqual(bg_line(dataset, "acreedores_diversos")["amount"], 100.0)
        self.assertEqual(dataset["total_pasivo"], 100.0)

    def test_sample_bg_and_validation_share_the_bg_l47_equation(self) -> None:
        parsed = parse_balanza(SAMPLE)
        er = build_er_dataset(parsed.rows)
        bg = build_bg_dataset(parsed.rows, result_ejercicio=er["raw_amounts"]["resultado_ejercicio"])
        validation = validate_balance_sheet(
            parsed.rows,
            result_ejercicio=er["raw_amounts"]["resultado_ejercicio"],
        )

        self.assertAlmostEqual(bg["total_activo"], 2654692.85)
        self.assertAlmostEqual(bg["total_pasivo"], 11900467.83)
        self.assertAlmostEqual(bg["capital_contable"], -9245774.92)
        self.assertAlmostEqual(bg["diferencia_cuadre"], -0.059663, places=6)
        self.assertTrue(bg["cuadra"])
        self.assertEqual(bg["reference"], "BG!L47")
        self.assertEqual(bg["formula"], "F45-L45")
        self.assertAlmostEqual(validation.diferencia_cuadre, bg["diferencia_cuadre"], places=8)
        self.assertEqual(validation.total_activo, bg["total_activo"])

    def test_bal_uses_all_sample_rows_and_accumulator_totals(self) -> None:
        parsed = parse_balanza(SAMPLE)

        dataset = build_bal_dataset(parsed.rows)

        self.assertEqual(len(dataset["rows"]), 157)
        self.assertEqual(dataset["rows"][0]["source_row"], parsed.rows[0].source_row)
        self.assertEqual(dataset["sumas_iguales"]["totals"]["debe"], 584.64)
        self.assertEqual(dataset["sumas_iguales"]["totals"]["haber"], 584.64)


if __name__ == "__main__":
    unittest.main()
