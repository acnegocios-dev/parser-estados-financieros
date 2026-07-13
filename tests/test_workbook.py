from __future__ import annotations

import unittest
from pathlib import Path

from src.engine import build_er_dataset
from src.parser import parse_balanza
from src.validation import validate_generated_workbook
from src.workbook import build_er_workbook


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"


class ErWorkbookGenerationTest(unittest.TestCase):
    def _build_result(self):
        parsed = parse_balanza(SAMPLE)
        dataset = build_er_dataset(
            [row.to_dict() for row in parsed.rows],
            company=parsed.company_name,
            period=parsed.period.period_ym,
            source_path=parsed.source_path,
        )
        return build_er_workbook(dataset, source_path=parsed.source_path)

    def test_builds_visible_er_sheet_with_manual_titles(self) -> None:
        result = self._build_result()
        workbook = result.workbook
        worksheet = workbook["ER"]

        self.assertEqual(workbook.sheetnames, ["ER"])
        self.assertEqual(worksheet.sheet_state, "visible")
        self.assertEqual(worksheet["B9"].value, "AL SERVICIOS MULTIPLES EMPRESARIALES SA DE CV")
        self.assertEqual(worksheet["B10"].value, "Estado de Resultados")
        self.assertEqual(worksheet["B11"].value, "Del 1ro de Enero al 31 de Julio 2026")
        self.assertEqual(worksheet["B12"].value, "(Importes expresados en pesos)")

    def test_writes_accumulated_amounts_in_h_and_percentages_in_j_only(self) -> None:
        result = self._build_result()
        worksheet = result.workbook["ER"]

        self.assertAlmostEqual(worksheet["H18"].value, 1977920.26)
        self.assertAlmostEqual(worksheet["H23"].value, 1654681.76)
        self.assertAlmostEqual(worksheet["H62"].value, -229978.65)
        self.assertEqual(worksheet["J18"].value, "=IF($H$18=0,0,H18/$H$18)")
        self.assertEqual(worksheet["J70"].value, "=IF($H$18=0,0,H70/$H$18)")
        self.assertIsNone(worksheet["D18"].value)
        self.assertIsNone(worksheet["F18"].value)
        self.assertFalse(any(str(cell).startswith("ER!D") or str(cell).startswith("ER!F") for cell in result.formula_cells))

    def test_uses_required_subtotal_formulas(self) -> None:
        result = self._build_result()
        worksheet = result.workbook["ER"]

        self.assertEqual(worksheet["H20"].value, "=SUM(H18:H19)")
        self.assertEqual(worksheet["H25"].value, "=H20-H23")
        self.assertEqual(worksheet["H51"].value, "=SUM(H28:H50)")
        self.assertEqual(worksheet["H53"].value, "=H25-H51")
        self.assertEqual(worksheet["H58"].value, "=SUM(H56:H57)")
        self.assertEqual(worksheet["H63"].value, "=SUM(H61:H62)")
        self.assertEqual(worksheet["H65"].value, "=H53+H58+H63")
        self.assertEqual(worksheet["H70"].value, "=H65-H67-H68")

    def test_uses_manual_money_format_and_has_no_formula_errors(self) -> None:
        result = self._build_result()
        worksheet = result.workbook["ER"]
        validation = validate_generated_workbook(result.workbook, balance_difference=0.0)

        self.assertEqual(worksheet["H18"].number_format, "#,##0.00_ ;[Red]\\-#,##0.00\\ ")
        self.assertEqual(worksheet["H62"].number_format, "#,##0.00")
        self.assertTrue(validation.formula_static_validation)
        self.assertFalse(validation.formula_recalculation_performed)
        self.assertIsNone(validation.formula_evaluated_error_count)
        self.assertFalse(validation.formula_cached_values_available)
        self.assertTrue(validation.ok)

    def test_requests_full_formula_recalculation_on_open(self) -> None:
        calculation = self._build_result().workbook.calculation
        self.assertEqual(calculation.calcMode, "auto")
        self.assertTrue(calculation.fullCalcOnLoad)
        self.assertTrue(calculation.forceFullCalc)


if __name__ == "__main__":
    unittest.main()
