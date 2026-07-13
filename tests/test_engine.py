from __future__ import annotations

import unittest
from pathlib import Path

from src.engine import build_er_dataset
from src.parser import parse_balanza


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"


def line(dataset: dict, key: str) -> dict:
    return next(item for item in dataset["lines"] if item["key"] == key)


class AuditaloErEngineTest(unittest.TestCase):
    def test_builds_er_dataset_from_sample_by_account_code(self) -> None:
        parsed = parse_balanza(SAMPLE)
        dataset = build_er_dataset(
            [row.to_dict() for row in parsed.rows],
            company=parsed.company_name,
            period=parsed.period.period_ym,
            source_path=parsed.source_path,
        )

        self.assertEqual(dataset["period"], "2026-07")
        self.assertEqual(dataset["source_rows_used"], 126)
        self.assertEqual(line(dataset, "ingresos_por_servicios")["amount_cell"], "H18")
        self.assertAlmostEqual(line(dataset, "ingresos_por_servicios")["accumulated_amount"], 1977920.26)
        self.assertAlmostEqual(line(dataset, "otros_productos")["accumulated_amount"], 98.92)
        self.assertAlmostEqual(line(dataset, "ingresos_netos")["accumulated_amount"], 1977920.26)
        self.assertAlmostEqual(line(dataset, "costo_de_ventas")["accumulated_amount"], 1654681.76)
        self.assertAlmostEqual(line(dataset, "utilidad_bruta")["accumulated_amount"], 323238.50)
        self.assertAlmostEqual(line(dataset, "gastos_de_operacion")["accumulated_amount"], 2890224.43)
        self.assertAlmostEqual(line(dataset, "utilidad_perdida_operacion")["accumulated_amount"], -2566985.93)
        self.assertAlmostEqual(line(dataset, "productos_financieros")["accumulated_amount"], 4574.35)
        self.assertAlmostEqual(line(dataset, "gastos_financieros")["accumulated_amount"], -229978.65)
        self.assertAlmostEqual(line(dataset, "resultado_integral_financiamiento")["accumulated_amount"], -225404.30)
        self.assertAlmostEqual(line(dataset, "resultado_antes_impuestos")["accumulated_amount"], -2792291.31)
        self.assertAlmostEqual(line(dataset, "resultado_ejercicio")["accumulated_amount"], -2792291.31)

    def test_generates_column_j_percentages_against_h18(self) -> None:
        parsed = parse_balanza(SAMPLE)
        dataset = build_er_dataset([row.to_dict() for row in parsed.rows])

        self.assertEqual(line(dataset, "ingresos_por_servicios")["percentage_of"], "H18")
        self.assertAlmostEqual(line(dataset, "ingresos_por_servicios")["percentage"], 1.0)
        self.assertAlmostEqual(line(dataset, "utilidad_bruta")["percentage"], 0.1634)
        self.assertAlmostEqual(line(dataset, "gastos_financieros")["percentage"], -0.1163)
        self.assertAlmostEqual(line(dataset, "resultado_ejercicio")["percentage"], -1.4117)

    def test_missing_expected_accounts_return_zero_and_warning(self) -> None:
        parsed = parse_balanza(SAMPLE)
        dataset = build_er_dataset([row.to_dict() for row in parsed.rows])

        self.assertEqual(line(dataset, "seguridad_e_higiene")["accumulated_amount"], 0.0)
        self.assertEqual(line(dataset, "isr_del_ejercicio")["accumulated_amount"], 0.0)
        missing_codes = {
            warning["account_code"]
            for warning in dataset["warnings"]
            if warning["code"] == "cuenta_no_encontrada"
        }
        self.assertIn("6147", missing_codes)
        self.assertIn("6510-0001", missing_codes)

    def test_mapping_does_not_depend_on_excel_source_row(self) -> None:
        rows = [
            {
                "source_row": 9999,
                "account_code": "4110-0001",
                "account_name": "SERVICIOS",
                "top_account": "4110",
                "saldo_final": "1000.00",
            },
            {
                "source_row": 2,
                "account_code": "5110-0001",
                "account_name": "COSTO",
                "top_account": "5110",
                "saldo_final": "300.00",
            },
            {
                "source_row": 1,
                "account_code": "6410-0001",
                "account_name": "GASTO FINANCIERO",
                "top_account": "6410",
                "saldo_final": "50.00",
            },
        ]

        dataset = build_er_dataset(rows)

        self.assertAlmostEqual(line(dataset, "ingresos_por_servicios")["accumulated_amount"], 1000.0)
        self.assertAlmostEqual(line(dataset, "costo_de_ventas")["accumulated_amount"], 300.0)
        self.assertAlmostEqual(line(dataset, "gastos_financieros")["accumulated_amount"], -50.0)

    def test_varios_includes_all_manual_composite_accounts(self) -> None:
        rows = [
            {"account_code": "6148-0001", "account_name": "BOTIQUIN Y ESTUDIOS MEDICOS", "saldo_final": "1173.30"},
            {"account_code": "6176-0001", "account_name": "GASTOS EN EL EXTRANJERO", "saldo_final": "38441.61"},
            {"account_code": "6195-0001", "account_name": "GASTOS DIVERSOS", "saldo_final": "0.00"},
        ]

        dataset = build_er_dataset(rows)

        varios = line(dataset, "varios")
        self.assertAlmostEqual(varios["accumulated_amount"], 39614.91)
        self.assertEqual(
            {account["account_code"] for account in varios["accounts"]},
            {"6148-0001", "6176-0001", "6195-0001"},
        )

    def test_leaf_accounts_prevent_accumulator_double_counting(self) -> None:
        rows = [
            {"account_code": "6110", "account_name": "SUELDOS ACUMULADOR", "saldo_final": "1000.00"},
            {"account_code": "6110-0001", "account_name": "SUELDOS A", "saldo_final": "300.00"},
            {"account_code": "6110-0002", "account_name": "SUELDOS B", "saldo_final": "200.00"},
        ]

        dataset = build_er_dataset(rows)

        self.assertAlmostEqual(line(dataset, "sueldos_y_salarios")["accumulated_amount"], 500.0)
        self.assertEqual(dataset["source_rows_used"], 2)

    def test_zero_income_produces_zero_percentages_without_division_error(self) -> None:
        dataset = build_er_dataset([])

        self.assertEqual(line(dataset, "ingresos_por_servicios")["accumulated_amount"], 0.0)
        self.assertTrue(all(item["percentage"] == 0.0 for item in dataset["lines"]))

    def test_negative_account_and_financial_sign_policy_are_preserved(self) -> None:
        rows = [
            {"account_code": "4110-0001", "account_name": "INGRESO NEGATIVO", "saldo_final": "-100.00"},
            {"account_code": "6410-0001", "account_name": "GASTO FINANCIERO", "saldo_final": "100.00"},
        ]

        dataset = build_er_dataset(rows)

        self.assertAlmostEqual(line(dataset, "ingresos_por_servicios")["accumulated_amount"], -100.0)
        self.assertAlmostEqual(line(dataset, "gastos_financieros")["accumulated_amount"], -100.0)
        self.assertLess(line(dataset, "resultado_ejercicio")["accumulated_amount"], 0.0)


if __name__ == "__main__":
    unittest.main()
