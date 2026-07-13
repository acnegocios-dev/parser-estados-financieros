from __future__ import annotations

import unittest
from pathlib import Path

from src.parser import parse_balanza
from src.engine import build_er_dataset
from src.validation import validate_balance_difference, validate_balance_sheet


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"


class BalanceValidationTest(unittest.TestCase):
    def test_validate_balance_sheet_uses_programmatic_formula(self) -> None:
        rows = [
            {"account_code": "1110-0001", "account_name": "Caja", "saldo_final": "120.00"},
            {"account_code": "1120-0001", "account_name": "Bancos", "saldo_final": "380.00"},
            {"account_code": "2110-0001", "account_name": "Proveedores", "saldo_final": "300.00"},
            {"account_code": "3100-0001", "account_name": "Capital", "saldo_final": "200.00"},
        ]

        result = validate_balance_sheet(rows, tolerance=1.0)

        self.assertAlmostEqual(result.total_activo, 500.0)
        self.assertAlmostEqual(result.total_pasivo, 300.0)
        self.assertAlmostEqual(result.capital_contable, 200.0)
        self.assertAlmostEqual(result.diferencia_cuadre, 0.0)
        self.assertTrue(result.cuadra)
        self.assertFalse(result.balanza_no_cuadra)
        self.assertEqual(len(result.componentes), 3)

    def test_validate_balance_sheet_flags_nonzero_residual(self) -> None:
        rows = [
            {"account_code": "1110-0001", "account_name": "Caja", "saldo_final": "120.00"},
            {"account_code": "2110-0001", "account_name": "Proveedores", "saldo_final": "300.00"},
            {"account_code": "3100-0001", "account_name": "Capital", "saldo_final": "180.05"},
        ]

        result = validate_balance_sheet(rows, tolerance=1.0)

        self.assertAlmostEqual(result.diferencia_cuadre, -360.05)
        self.assertFalse(result.cuadra)
        self.assertTrue(result.balanza_no_cuadra)

    def test_validate_balance_sheet_with_sample_returns_detail(self) -> None:
        parsed = parse_balanza(SAMPLE)
        dataset = build_er_dataset([row.to_dict() for row in parsed.rows])
        result_ejercicio = dataset["raw_amounts"]["resultado_ejercicio"]
        result = validate_balance_sheet(parsed.rows, tolerance=1.0, result_ejercicio=result_ejercicio)

        self.assertEqual(result.tolerance, 1.0)
        self.assertEqual(len(result.componentes), 3)
        self.assertEqual(result.componentes[0].rubro, "activo")
        self.assertEqual(result.componentes[1].rubro, "pasivo")
        self.assertEqual(result.componentes[2].rubro, "capital_contable")
        self.assertAlmostEqual(result.diferencia_cuadre, -0.059663, places=6)
        self.assertTrue(result.cuadra)

    def test_balance_tolerance_is_strictly_less_than_one(self) -> None:
        self.assertTrue(validate_balance_difference("0.999999", tolerance=1.0))
        self.assertTrue(validate_balance_difference("-0.999999", tolerance=1.0))
        self.assertFalse(validate_balance_difference("1.0", tolerance=1.0))
        self.assertFalse(validate_balance_difference("-1.0", tolerance=1.0))


if __name__ == "__main__":
    unittest.main()
