from __future__ import annotations

import hashlib
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from src.accounting_profiles import CatalogIdentity, MappingRule
from src.catalog_profile_draft import BLOCKED, DRAFT, REJECTED, analyze_catalog_paths
from tests.test_accounting_profiles import make_profile


class CatalogProfileDraftTest(unittest.TestCase):
    def _catalog(self, directory: Path, name: str, text: str, *, encoding: str = "utf-8") -> Path:
        path = directory / name
        path.write_bytes(text.encode(encoding))
        return path

    def test_empty_delta_is_rejected_and_never_runtime_selectable(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            report = analyze_catalog_paths((self._catalog(Path(temporary), "DELTA cuentas.csv", ""),))

        catalog = report["catalogs"][0]
        self.assertEqual(catalog["status"], REJECTED)
        self.assertFalse(catalog["draft_profile"]["runtime_selectable"])
        self.assertIn("confirmed_rfc_required", catalog["draft_profile"]["runtime_gates"])
        self.assertIn("representative_balance_required", catalog["draft_profile"]["runtime_gates"])

    def test_duplicate_folvaz_6150_0005_is_blocked_without_exposing_names(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = self._catalog(
                Path(temporary),
                "FOLVAZ cuentas.csv",
                "6150-0005,NOMBRE PRIVADO UNO,6150,D,601\n6150-0005,NOMBRE PRIVADO DOS,6150,D,601\n",
            )
            report = analyze_catalog_paths((path,))

        catalog = report["catalogs"][0]
        self.assertEqual(catalog["status"], BLOCKED)
        self.assertTrue(any(item["code"] == "duplicate_account_code" for item in catalog["validations"]))
        self.assertNotIn("NOMBRE PRIVADO", str(catalog))
        self.assertTrue(catalog["collisions"])

    def test_code_1220_is_not_global_and_sat_drives_draft_suggestion(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = self._catalog(Path(temporary), "empresa.csv", "1220,EQUIPO,ROOT,D,156\n")
            report = analyze_catalog_paths((path,))

        suggestion = report["catalogs"][0]["suggestions"][0]
        self.assertEqual(suggestion["line_key"], "bg.computer_equipment")
        self.assertIn("sat_exact:156", suggestion["evidence"])

    def test_folvaz_1230_and_rgcv_6185_anomalies_remain_draft_unmapped(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            folvaz = self._catalog(directory, "FOLVAZ.csv", "1230,PRIVADO,,D,156\n1230-0003,PRIVADO,1230,D,601.16\n")
            rgcv = self._catalog(directory, "RGCV.csv", "6185,PRIVADO,,D,601\n6185-0005,PRIVADO,6185,D,183.04\n")
            report = analyze_catalog_paths((folvaz, rgcv))

        self.assertEqual([item["status"] for item in report["catalogs"]], [DRAFT, DRAFT])
        self.assertTrue(any(item["code"] == "section_sat_mismatch" for item in report["catalogs"][0]["validations"]))
        self.assertEqual(report["catalogs"][1]["unmapped_accounts"][0]["account_code"], "6185-0005")

    def test_name_rule_is_only_a_low_confidence_suggestion(self) -> None:
        raw = "X-1,Banco Privado,ROOT,D,999\n"
        profile = make_profile(
            rules=(MappingRule("name", "BG", "bg.banks", "name_suggestion", normalized_name="BANCO PRIVADO"),)
        )
        profile = replace(profile, catalog_identity=CatalogIdentity(hashlib.sha256(raw.encode()).hexdigest(), profile.catalog_identity.semantic_sha256))
        with tempfile.TemporaryDirectory() as temporary:
            path = self._catalog(Path(temporary), "empresa.csv", raw)
            report = analyze_catalog_paths((path,), previous_profile=profile)

        suggestion = report["catalogs"][0]["suggestions"][0]
        self.assertEqual(suggestion["status"], "name_suggestion_only")
        self.assertEqual(suggestion["confidence"], 0.20)
        self.assertNotIn("Banco Privado", str(suggestion))


if __name__ == "__main__":
    unittest.main()
