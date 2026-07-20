from __future__ import annotations

import unittest
from pathlib import Path

from src.engine import build_er_dataset
from src.parser import parse_balanza
from src.validation import find_formula_issues, validate_balance_sheet, validate_generated_workbook
from src.workbook import _default_output_path, build_er_workbook


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

    def test_uses_approved_manual_number_format_and_has_no_formula_errors(self) -> None:
        result = self._build_result()
        worksheet = result.workbook["ER"]
        validation = validate_generated_workbook(result.workbook, balance_difference=0.0)

        self.assertEqual(worksheet["H18"].number_format, "General")
        self.assertEqual(worksheet["H62"].number_format, "General")
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


class ThreeSheetWorkbookContractTest(unittest.TestCase):
    """Characterize the approved BG + ER + BAL delivery contract in memory."""

    FORMULA_ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")

    def _build_result(self):
        parsed = parse_balanza(SAMPLE)
        dataset = build_er_dataset(
            [row.to_dict() for row in parsed.rows],
            company=parsed.company_name,
            period=parsed.period.period_ym,
            source_path=parsed.source_path,
        )
        return build_er_workbook(dataset, source_path=parsed.source_path)

    def _require_final_sheet_contract(self, workbook) -> None:
        self.assertEqual(workbook.sheetnames, ["BG", "ER", "BAL"])
        self.assertEqual([sheet.sheet_state for sheet in workbook.worksheets], ["visible"] * 3)

    def test_has_exactly_three_visible_sheets_in_delivery_order(self) -> None:
        self._require_final_sheet_contract(self._build_result().workbook)

    def test_bg_and_er_use_white_effective_areas_without_gridlines(self) -> None:
        workbook = self._build_result().workbook
        self._require_final_sheet_contract(workbook)

        for sheet_name, last_column, last_row in (("BG", 12, 47), ("ER", 10, 70)):
            worksheet = workbook[sheet_name]
            self.assertFalse(worksheet.sheet_view.showGridLines)
            for row in worksheet.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=last_column):
                for cell in row:
                    self.assertEqual(cell.fill.fill_type, "solid", f"{sheet_name}!{cell.coordinate}")
                    self.assertEqual(cell.fill.fgColor.type, "rgb", f"{sheet_name}!{cell.coordinate}")
                    self.assertEqual(cell.fill.fgColor.rgb, "FFFFFFFF", f"{sheet_name}!{cell.coordinate}")

    def test_bal_is_a_clean_c_to_g_table_with_dynamic_sumas_iguales(self) -> None:
        parsed = parse_balanza(SAMPLE)
        self.assertEqual(len(parsed.rows), 157)
        workbook = self._build_result().workbook
        self._require_final_sheet_contract(workbook)
        worksheet = workbook["BAL"]

        self.assertEqual(worksheet["C1"].value, parsed.company_name)
        self.assertEqual(worksheet["C2"].value, "Balanza de Comprobaci\u00f3n")
        self.assertEqual(worksheet["C3"].value, parsed.period.period_label_bal)
        self.assertEqual(worksheet["C4"].value, f"RFC: {parsed.period.rfc}")
        self.assertEqual(
            [worksheet.cell(6, column).value for column in range(3, 8)],
            ["CUENTA", "SALDO INICIAL", "DEBE", "HABER", "SALDO FINAL"],
        )
        self.assertIsNotNone(worksheet["C7"].value)
        self.assertIsNotNone(worksheet["C163"].value)
        self.assertIsNone(worksheet["C164"].value)
        self.assertEqual(worksheet["C165"].value, "SUMAS IGUALES")
        accumulator_rows = [row for row in parsed.rows if "-" not in row.account_code]
        self.assertAlmostEqual(sum(float(row.debe) for row in accumulator_rows), 584.64)
        self.assertAlmostEqual(sum(float(row.haber) for row in accumulator_rows), 584.64)
        self.assertEqual(worksheet["E165"].value, "=SUM(" + ",".join(
            f"E{row}" for row in range(7, 164) if "-" not in parsed.rows[row - 7].account_code
        ) + ")")
        self.assertEqual(worksheet["F165"].value, "=SUM(" + ",".join(
            f"F{row}" for row in range(7, 164) if "-" not in parsed.rows[row - 7].account_code
        ) + ")")
        self.assertEqual(worksheet["G165"].value, "=D165+E165-F165")
        for row in range(1, 166):
            for column in (1, 2, 8, 9, 10, 11, 12, 13):
                self.assertIsNone(worksheet.cell(row, column).value)

    def test_period_rfc_er_amount_and_bg_balance_share_the_loaded_input(self) -> None:
        parsed = parse_balanza(SAMPLE)
        result = self._build_result()
        self._require_final_sheet_contract(result.workbook)

        self.assertEqual(result.period, parsed.period.period_ym)
        self.assertEqual(result.workbook["ER"]["H46"].value, 39614.91)
        self.assertEqual(result.workbook["BG"]["B9"].value, parsed.period.period_label_bg)
        self.assertEqual(result.workbook["BAL"]["C4"].value, f"RFC: {parsed.period.rfc}")
        dataset = build_er_dataset(
            [row.to_dict() for row in parsed.rows],
            company=parsed.company_name,
            period=parsed.period.period_ym,
            source_path=parsed.source_path,
        )
        balance = validate_balance_sheet(
            parsed.rows,
            result_ejercicio=dataset["raw_amounts"]["resultado_ejercicio"],
        )
        self.assertTrue(balance.cuadra)
        self.assertEqual(result.workbook["BG"]["L47"].value, "=F45-L45")
        self.assertAlmostEqual(balance.diferencia_cuadre, -0.059663, places=6)

    def test_bg_and_bal_geometry_use_the_versioned_contract(self) -> None:
        workbook = self._build_result().workbook
        bg = workbook["BG"]
        bal = workbook["BAL"]

        self.assertEqual(
            {str(item) for item in bg.merged_cells.ranges},
            {"B7:L7", "B8:L8", "B9:L9", "B10:L10"},
        )
        self.assertEqual(bg["F26"].value, "=SUM(E16:E24)")
        self.assertEqual(bg["F36"].value, "=SUM(E29:E34)")
        self.assertEqual(bg["F42"].value, "=SUM(E39:E40)")
        self.assertEqual(bg["F45"].value, "=F26+F36+F42")
        self.assertEqual(bg["L26"].value, "=SUM(K16:K20)")
        self.assertEqual(bg["K34"].value, "=ER!H70")
        self.assertEqual(bg["L36"].value, "=SUM(K31:K34)")
        self.assertEqual(bg["L45"].value, "=L26+L36")
        self.assertEqual(bg["L47"].value, "=F45-L45")
        self.assertEqual(
            {str(item) for item in bal.merged_cells.ranges},
            {"C1:G1", "C2:G2", "C3:G3", "C4:G4"},
        )
        self.assertEqual(bal.column_dimensions["C"].width, 28.46)
        self.assertEqual(bal.column_dimensions["D"].width, 15.7)
        self.assertEqual(bal.print_area, "'BAL'!$C$1:$G$165")
        self.assertIsNone(bal.freeze_panes)

    def test_print_titles_repeat_and_each_sheet_fits_one_page_wide(self) -> None:
        workbook = self._build_result().workbook
        expected = {
            "BG": ("'BG'!$B$7:$L$47", "$7:$10"),
            "ER": ("'ER'!$B$9:$J$70", "$9:$15"),
            "BAL": ("'BAL'!$C$1:$G$165", "$1:$6"),
        }
        for sheet_name, (area, titles) in expected.items():
            worksheet = workbook[sheet_name]
            self.assertEqual(worksheet.print_area, area)
            self.assertEqual(worksheet.print_title_rows, titles)
            self.assertEqual(worksheet.page_setup.fitToWidth, 1)
            self.assertEqual(worksheet.page_setup.fitToHeight, 0)
            self.assertTrue(worksheet.sheet_properties.pageSetUpPr.fitToPage)

    def test_default_output_name_is_the_financial_statements_name(self) -> None:
        self.assertEqual(
            _default_output_path("AL SERVICIOS MULTIPLES EMPRESARIALES SA DE CV", "2026-07").name,
            "estados_financieros_al_servicios_multiples_empresariales_sa_de_cv_2026_07.xlsx",
        )

    def test_three_sheet_book_has_no_external_links_or_formula_error_tokens(self) -> None:
        workbook = self._build_result().workbook
        self._require_final_sheet_contract(workbook)

        self.assertEqual(getattr(workbook, "_external_links", []), [])
        self.assertEqual(find_formula_issues(workbook), [])
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        self.assertFalse(
                            any(token in cell.value.upper() for token in self.FORMULA_ERROR_TOKENS),
                            f"{worksheet.title}!{cell.coordinate}",
                        )


if __name__ == "__main__":
    unittest.main()
