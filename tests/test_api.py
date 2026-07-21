from __future__ import annotations

import base64
import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from src.api import build_health_payload, build_process_payload


class ApiRuntimeMetadataTest(unittest.TestCase):
    def test_health_exposes_identity_without_server_paths(self) -> None:
        payload = build_health_payload()

        self.assertEqual(payload["status"], "ok")
        self.assertTrue(payload["service_version"])
        self.assertIn("git_commit", payload)
        self.assertIsInstance(payload["worktree_dirty"], bool)
        self.assertTrue(payload["process_started_at"])
        self.assertEqual(payload["generator_profile"], "manual-eeff-three-sheet")
        self.assertTrue(payload["generator_profile_version"])
        self.assertEqual(payload["formula_validation_mode"], "static_only")
        self.assertFalse(payload["formula_recalculation_performed"])

        serialized = json.dumps(payload)
        self.assertNotIn("/opt/n8n", serialized)
        self.assertNotIn("output_xlsx", serialized)
        self.assertNotIn("report_path", serialized)

    def test_process_exposes_output_identity_but_not_server_paths(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "estados_financieros_empresa_de_prueba_2026_07.xlsx"
            workbook = Workbook()
            workbook.active["A1"] = "evidence"
            workbook.save(output_path)

            report = {
                "period": {"period_ym": "2026-07"},
                "company_name": "Empresa de prueba",
                "content_period_ym": "2026-07",
                "parser": {"normalized_rows": 1, "leaf_rows_used_for_calculation": 1},
                "balance_check": {
                    "difference_cuadre": 0.0,
                    "tolerance": 1.0,
                    "cuadra": True,
                    "balanza_no_cuadra": False,
                    "total_activo": 1.0,
                    "total_pasivo": 1.0,
                    "capital_contable": 0.0,
                    "componentes": [],
                },
                "validation": {
                    "formula_static_validation": True,
                    "formula_recalculation_performed": False,
                    "formula_recalculation_engine": "none",
                    "formula_validation_mode": "static_only",
                    "formula_evaluated_error_count": None,
                    "formula_cached_values_available": False,
                    "ok": True,
                    "warnings": [],
                },
                "output_xlsx": str(output_path),
                "workbook": {"sheet_names": ["BG", "ER", "BAL"]},
                "runtime": {
                    "service_version": "estados-financieros-api-test",
                    "git_commit": "abcdef1234567890",
                    "worktree_dirty": True,
                    "process_started_at": "2026-07-13T18:00:00+00:00",
                    "generator_profile": "manual-eeff-three-sheet",
                    "generator_profile_version": "2026-07-16.three-sheet-v1",
                    "generated_at": "2026-07-16T18:01:00+00:00",
                    "output_sha256": "a" * 64,
                    "formula_validation_mode": "static_only",
                },
                "engine": {
                    "profile": {
                        "accounting_profile_id": "profile-test",
                        "accounting_profile_version": "1.0.0",
                        "accounting_profile_status": "approved",
                        "accounting_profile_company": "Empresa de prueba",
                        "accounting_profile_rfc": "SME170717GA0",
                        "accounting_profile_valid_from": "2026-07-01",
                        "base_taxonomy_version": "sat-rmf-2026-v1",
                        "catalog_semantic_sha256": "b" * 64,
                    },
                    "coverage": {
                        "assigned": 2,
                        "unassigned": 1,
                        "ambiguous": 0,
                        "duplicates": 0,
                        "entries": [{"account_code": "6150-0005", "source_row": 8}],
                        "section_controls": {"1": {"difference": 0.0, "status": "ok"}},
                    },
                    "warnings": [{"code": "profile_mapping_missing_zero", "account_code": "6150-0005"}],
                },
            }

            payload = build_process_payload(report)

        self.assertTrue(payload["output_filename"].startswith("estados_financieros_"))
        self.assertEqual(payload["sheet_names"], ["BG", "ER", "BAL"])
        self.assertEqual(payload["generator_profile"], "manual-eeff-three-sheet")
        self.assertEqual(payload["output_sha256"], "a" * 64)
        self.assertEqual(payload["formula_validation_mode"], "static_only")
        self.assertEqual(payload["coverage"]["total_accounts"], 3)
        self.assertFalse(payload["coverage"]["complete"])
        self.assertEqual(payload["warnings"], [{"code": "profile_mapping_missing_zero"}])
        self.assertEqual(payload["accounting_profile_rfc_hint"], "SME••••••GA0")
        self.assertEqual(payload["catalog_semantic_sha256_short"], "bbbbbbbbbbbb")
        self.assertEqual(
            base64.b64decode(payload["output_xlsx_base64"]).decode("utf-8", errors="ignore")[:2],
            "PK",
        )
        serialized = json.dumps(payload)
        self.assertNotIn(str(output_path.parent), serialized)
        self.assertNotIn("report_path", serialized)
        self.assertNotIn("componentes_balance", serialized)
        self.assertNotIn("6150-0005", serialized)
        self.assertNotIn("SME170717GA0", serialized)


if __name__ == "__main__":
    unittest.main()
