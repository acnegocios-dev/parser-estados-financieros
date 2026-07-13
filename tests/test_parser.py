from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from src.parser import OOXML_ZIP, detect_mime_kind, parse_balanza
from src.period import extract_period_variables


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "sample-inputs" / "balanza_SME170717GA0_2026_07.xls"


class AuditaloParserTest(unittest.TestCase):
    def _write_fixture(
        self,
        path: Path,
        *,
        company: str = "AL SERVICIOS MULTIPLES EMPRESARIALES SA DE CV SME170717GA0",
        period: str = "Periodo: 2026-07",
        rows: list[tuple[object, ...]] | None = None,
    ) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Balanza"
        worksheet.append(("Cuenta", "Saldo Inicial", "Debe", "Haber", "SaldoFinal"))
        worksheet.append((None, None, None, None, None))
        worksheet.append((company, None, None, None, None))
        worksheet.append((period, None, None, None, None))
        for row in rows or [("4110-0001 SERVICIOS", "0", "0", "0", "100")]:
            worksheet.append(row)
        workbook.save(path)

    def test_detects_ooxml_content_even_with_xls_extension(self) -> None:
        self.assertEqual(detect_mime_kind(SAMPLE), OOXML_ZIP)

    def test_extracts_period_variables_from_filename(self) -> None:
        period = extract_period_variables(SAMPLE)

        self.assertEqual(period.rfc, "SME170717GA0")
        self.assertEqual(period.period_year, 2026)
        self.assertEqual(period.period_month, 7)
        self.assertEqual(period.period_ym, "2026-07")
        self.assertEqual(period.period_compact, "202607")
        self.assertEqual(period.period_label_bg, "Al 31 de Julio de 2026")
        self.assertEqual(period.period_label_er, "Del 1ro de Enero al 31 de Julio de 2026")
        self.assertEqual(period.period_label_bal, "JULIO.2026")

    def test_normalizes_auditalo_balanza_rows(self) -> None:
        parsed = parse_balanza(SAMPLE)

        self.assertEqual(parsed.detected_mime_kind, OOXML_ZIP)
        self.assertEqual(parsed.sheet_name, "Balanza")
        self.assertEqual(parsed.period.period_ym, "2026-07")
        self.assertEqual(parsed.content_period_ym, "2026-07")
        self.assertEqual(parsed.content_rfc, "SME170717GA0")
        self.assertEqual(parsed.company_name, "AL SERVICIOS MULTIPLES EMPRESARIALES SA DE CV")
        self.assertEqual(parsed.header_row, 1)
        self.assertEqual(len(parsed.rows), 157)

        first = parsed.rows[0]
        self.assertEqual(first.source_row, 7)
        self.assertEqual(first.account_raw, "1110   CAJA Y EFECTIVO")
        self.assertEqual(first.account_code, "1110")
        self.assertEqual(first.account_name, "CAJA Y EFECTIVO")
        self.assertEqual(first.top_account, "1110")
        self.assertEqual(str(first.saldo_inicial), "1324.44")
        self.assertEqual(str(first.debe), "0")
        self.assertEqual(str(first.haber), "0")
        self.assertEqual(str(first.saldo_final), "1324.44")

    def test_rejects_filename_period_that_differs_from_sheet_period(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            target = Path(tmp_dir) / "balanza_SME170717GA0_2026_06.xls"
            shutil.copyfile(SAMPLE, target)

            with self.assertRaisesRegex(ValueError, "Filename period 2026-06 differs"):
                parse_balanza(target)

    def test_rejects_filename_rfc_that_differs_from_sheet_rfc(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            target = Path(tmp_dir) / "balanza_ABC010101ABC_2026_07.xls"
            shutil.copyfile(SAMPLE, target)

            with self.assertRaisesRegex(ValueError, "Filename RFC ABC010101ABC differs"):
                parse_balanza(target)

    def test_warns_when_content_rfc_is_missing_and_keeps_filename_rfc(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            target = Path(tmp_dir) / "balanza_SME170717GA0_2026_07.xls"
            self._write_fixture(target, company="AL SERVICIOS MULTIPLES EMPRESARIALES SA DE CV")

            parsed = parse_balanza(target)

            self.assertIsNone(parsed.content_rfc)
            self.assertEqual(parsed.period.rfc, "SME170717GA0")
            self.assertEqual(parsed.company_name, "AL SERVICIOS MULTIPLES EMPRESARIALES SA DE CV")
            self.assertIn(
                "rfc_no_encontrado_en_contenido",
                {warning.code for warning in parsed.warnings},
            )

    def test_ignores_repeated_headers_records_blank_rows_and_warns_on_duplicate_accounts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            target = Path(tmp_dir) / "balanza_SME170717GA0_2026_07.xls"
            self._write_fixture(
                target,
                rows=[
                    ("4110-0001 SERVICIOS", "0", "0", "0", "100"),
                    (None, None, None, None, None),
                    ("Cuenta", "Saldo Inicial", "Debe", "Haber", "SaldoFinal"),
                    ("4110-0001 SERVICIOS", "0", "0", "0", "25"),
                ],
            )

            parsed = parse_balanza(target)

            self.assertEqual(len(parsed.rows), 2)
            self.assertEqual(parsed.empty_rows, (2, 6))
            warning_codes = {warning.code for warning in parsed.warnings}
            self.assertIn("encabezado_repetido_ignorado", warning_codes)
            self.assertIn("cuenta_repetida_agregada", warning_codes)
            self.assertEqual(parsed.structure_issues, ())


if __name__ == "__main__":
    unittest.main()
