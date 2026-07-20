from __future__ import annotations

import hashlib
import importlib.util
import json
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REFERENCE_DIR = ROOT / "sample-inputs" / "reference-2026-07-20"
MANUAL = REFERENCE_DIR / "EEFF_202602_AL_Serv_Prueba.xlsx"
COLORED_MANUAL = REFERENCE_DIR / "EEFF_202602_AL_Serv_Prueba_referencia_colore_hoja_bal.xlsx"
EXPECTED_SOURCE_SHA256 = "991daeaa5b9f957e490e231164825640127cb850f01db865c04cfbb25e72b12c"
EXPECTED_MASK_SHA256 = "c27d3b4f40737e00a01dc83a2bc8745f6d62fd4a17cd8d78600a5ec09764dda2"

SPEC_CONTRACTS = {
    "BG": ("bg_style_spec.json", "A1:L47", "2026-07-20.bg-v2", {"B7:L7", "B8:L8", "B9:L9", "B10:L10"}),
    "ER": ("er_style_spec.json", "A1:J70", "2026-07-20.er-v2", None),
    "BAL": ("bal_style_spec.json", "C1:G185", "2026-07-20.bal-v2", {"C1:G1", "C2:G2", "C3:G3", "C4:G4"}),
}


def _load_extractor():
    path = ROOT / "docs" / "extract_er_style_spec.py"
    module_spec = importlib.util.spec_from_file_location("style_spec_extractor", path)
    assert module_spec and module_spec.loader
    module = importlib.util.module_from_spec(module_spec)
    module_spec.loader.exec_module(module)
    return module


class StyleSpecContractTest(unittest.TestCase):
    def test_reference_hashes_are_the_approved_ones(self) -> None:
        self.assertEqual(hashlib.sha256(MANUAL.read_bytes()).hexdigest(), EXPECTED_SOURCE_SHA256)
        self.assertEqual(
            hashlib.sha256(COLORED_MANUAL.read_bytes()).hexdigest(), EXPECTED_MASK_SHA256
        )

    def test_specs_are_versioned_and_self_contained(self) -> None:
        required_style_keys = {"font", "alignment", "fill", "border", "numberFormat", "protection"}
        required_geometry_keys = {
            "default_row_height", "default_column_width", "base_column_width",
            "column_widths", "hidden_columns", "row_heights", "hidden_rows",
            "merged_ranges", "page_margins", "page_setup", "page_setup_properties",
            "print_options", "print_area", "print_title_rows", "print_title_cols",
            "sheet_view", "freeze_panes",
        }
        for sheet, (filename, visible_range, version, merges) in SPEC_CONTRACTS.items():
            with self.subTest(sheet=sheet):
                spec = json.loads((ROOT / "src" / filename).read_text(encoding="utf-8"))
                self.assertEqual(spec["version"], version)
                self.assertEqual(spec["sheet"], sheet)
                self.assertEqual(spec["source"], MANUAL.name)
                self.assertEqual(spec["source_sha256"], EXPECTED_SOURCE_SHA256)
                self.assertEqual(spec["visible_range"], visible_range)
                self.assertTrue(spec["styles"])
                self.assertTrue(spec["cells"])
                self.assertTrue(required_geometry_keys.issubset(spec["geometry"]))
                self.assertTrue(all(required_style_keys.issubset(item) for item in spec["styles"]))
                if merges is not None:
                    self.assertEqual(set(spec["geometry"]["merged_ranges"]), merges)

    def test_bal_mask_is_documented_and_cleanly_scoped(self) -> None:
        spec = json.loads((ROOT / "src" / "bal_style_spec.json").read_text(encoding="utf-8"))
        self.assertEqual(spec["mask"]["source"], COLORED_MANUAL.name)
        self.assertEqual(spec["mask"]["source_sha256"], EXPECTED_MASK_SHA256)
        self.assertEqual(spec["mask"]["yellow_ranges"], ["C1:C4", "C5:G185"])
        self.assertEqual(spec["mask"]["green_ranges"], ["H7:M183"])
        self.assertIn("only the clean C:G mask", spec["mask"]["output_decision"])
        self.assertEqual(set(spec["dynamic_profiles"]), {
            "light_regular", "dark_regular", "light_bold", "dark_bold", "separator", "total",
        })

    def test_extractor_reads_and_writes_only_the_approved_reference_contract(self) -> None:
        extractor = _load_extractor()
        self.assertEqual(extractor.REFERENCE_DIR, REFERENCE_DIR)
        self.assertEqual(extractor.SOURCE, MANUAL)
        self.assertEqual(extractor.COLORED_SOURCE, COLORED_MANUAL)
        self.assertEqual(set(extractor.TARGETS), {"BG", "ER", "BAL"})
        self.assertEqual(extractor.BOUNDS["BG"], (1, 47, 1, 12))
        self.assertEqual(extractor.BOUNDS["ER"], (1, 70, 1, 10))
        self.assertEqual(extractor.BOUNDS["BAL"], (1, 185, 3, 7))

    def test_runtime_loads_only_versioned_json_specs(self) -> None:
        runtime_source = (ROOT / "src" / "workbook.py").read_text(encoding="utf-8")
        self.assertNotIn("reference-2026-07-20", runtime_source)
        self.assertNotIn("EEFF_202602", runtime_source)
        self.assertIn('with path.open(encoding="utf-8")', runtime_source)


if __name__ == "__main__":
    unittest.main()
