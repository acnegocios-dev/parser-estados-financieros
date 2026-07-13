from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from src.validation import recalculate_workbook, validate_generated_workbook


def _write_workbook(path: Path, values: dict[str, object]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    for coordinate, value in values.items():
        worksheet[coordinate] = value
    workbook.save(path)


class FormulaValidationModesTest(unittest.TestCase):
    def test_recalculation_without_engine_is_explicitly_unperformed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            formula_path = Path(directory) / "formula.xlsx"
            _write_workbook(formula_path, {"A1": "=1+1"})
            with patch("src.validation.shutil.which", return_value=None):
                result = recalculate_workbook(formula_path)

        self.assertFalse(result.performed)
        self.assertEqual(result.engine, "none")
        self.assertIsNone(result.evaluated_error_count)
        self.assertFalse(result.cached_values_available)
        self.assertFalse(result.blocked)

    def test_static_only_does_not_claim_evaluated_zero_errors(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            formula_path = Path(directory) / "formula.xlsx"
            _write_workbook(formula_path, {"A1": "=1+1", "A2": "=1/0"})

            result = validate_generated_workbook(formula_path, formula_mode="static_only")

        self.assertTrue(result.formula_static_validation)
        self.assertFalse(result.formula_recalculation_performed)
        self.assertEqual(result.formula_recalculation_engine, "none")
        self.assertIsNone(result.formula_evaluated_error_count)
        self.assertFalse(result.formula_cached_values_available)
        self.assertTrue(result.ok)

    def test_recalculated_ok_reports_cached_values_and_zero_evaluated_errors(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            formula_path = Path(directory) / "formula.xlsx"
            evaluated_path = Path(directory) / "evaluated.xlsx"
            _write_workbook(formula_path, {"A1": "=1+1", "A2": "=1/0"})
            _write_workbook(evaluated_path, {"A1": 2, "A2": 0})

            result = validate_generated_workbook(
                formula_path,
                formula_mode="recalculated_ok",
                evaluated_workbook_or_path=evaluated_path,
                formula_recalculation_engine="test-engine",
            )

        self.assertTrue(result.formula_static_validation)
        self.assertTrue(result.formula_recalculation_performed)
        self.assertEqual(result.formula_recalculation_engine, "test-engine")
        self.assertEqual(result.formula_evaluated_error_count, 0)
        self.assertTrue(result.formula_cached_values_available)
        self.assertTrue(result.ok)

    def test_recalculated_error_is_not_presented_as_success(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            formula_path = Path(directory) / "formula.xlsx"
            evaluated_path = Path(directory) / "evaluated.xlsx"
            _write_workbook(formula_path, {"A1": "=1+1", "A2": "=1/0"})
            _write_workbook(evaluated_path, {"A1": 2, "A2": "#DIV/0!"})

            result = validate_generated_workbook(
                formula_path,
                formula_mode="recalculated_error",
                evaluated_workbook_or_path=evaluated_path,
                formula_recalculation_engine="test-engine",
            )

        self.assertTrue(result.formula_static_validation)
        self.assertTrue(result.formula_recalculation_performed)
        self.assertEqual(result.formula_evaluated_error_count, 1)
        self.assertTrue(result.formula_cached_values_available)
        self.assertFalse(result.ok)
        self.assertEqual(result.formula_evaluated_issues[0].reason, "Evaluated formula contains #DIV/0!")
