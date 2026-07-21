from __future__ import annotations

import io
import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from src.accounting_profiles import (
    CatalogIdentity,
    ProfileValidationError,
    load_accounting_profile,
    select_profile_for_runtime,
)
from src.api import APIRequestError, FinancialStatementsAPI, normalize_upload_filename


ROOT = Path(__file__).resolve().parents[1]
PROFILE_PATH = ROOT / "src" / "profiles" / "SME170717GA0-2026-07-v1.json"


class _Upload:
    def __init__(self, filename: str, data: bytes):
        self.filename = filename
        self.file = io.BytesIO(data)


class RuntimeProfileSelectionTest(unittest.TestCase):
    def setUp(self) -> None:
        self.profile = load_accounting_profile(PROFILE_PATH)
        self.identity = self.profile.catalog_identity

    def test_selects_only_by_rfc_validity_and_exact_catalog_hashes(self) -> None:
        selected = select_profile_for_runtime(
            (self.profile,),
            rfc="SME170717GA0",
            as_of=date(2026, 7, 1),
            catalog_identity=self.identity,
        )

        self.assertEqual(selected.profile_id, self.profile.profile_id)

    def test_rejects_missing_rfc_unapproved_hash_mismatch_and_duplicate_profile(self) -> None:
        cases = (
            ("profile_not_found", "OTHER010101AAA", self.identity, (self.profile,)),
            ("catalog_hash_mismatch", "SME170717GA0", CatalogIdentity("0" * 64, self.identity.semantic_sha256), (self.profile,)),
            ("ambiguous_mapping", "SME170717GA0", self.identity, (self.profile, self.profile)),
        )
        for expected, rfc, identity, profiles in cases:
            with self.subTest(expected=expected), self.assertRaises(ProfileValidationError) as context:
                select_profile_for_runtime(profiles, rfc=rfc, as_of=date(2026, 7, 1), catalog_identity=identity)
            self.assertIn(expected, {issue.code for issue in context.exception.issues})

        draft = type(self.profile)(**{**self.profile.__dict__, "status": "draft", "approval": None})
        with self.assertRaises(ProfileValidationError) as context:
            select_profile_for_runtime((draft,), rfc="SME170717GA0", as_of=date(2026, 7, 1), catalog_identity=self.identity)
        self.assertIn("profile_not_approved", {issue.code for issue in context.exception.issues})


class MultipartBoundaryTest(unittest.TestCase):
    def _handler(self, *, content_length: int = 1):
        handler = FinancialStatementsAPI.__new__(FinancialStatementsAPI)
        handler.headers = {
            "Content-Type": "multipart/form-data; boundary=test",
            "Content-Length": str(content_length),
        }
        handler.rfile = io.BytesIO(b"ignored")
        return handler

    def test_requires_multipart_and_enforces_content_length(self) -> None:
        handler = self._handler()
        handler.headers["Content-Type"] = "application/json"
        with self.assertRaises(APIRequestError) as context:
            handler._resolve_multipart_input()
        self.assertEqual(context.exception.code, "multipart_required")

        handler = self._handler(content_length=FinancialStatementsAPI.MAX_CONTENT_LENGTH + 1)
        with self.assertRaises(APIRequestError) as context:
            handler._resolve_multipart_input()
        self.assertEqual(context.exception.code, "content_length_exceeded")

    def test_normalizes_traversal_and_rejects_invalid_extension(self) -> None:
        self.assertEqual(normalize_upload_filename("../../SME170717GA0_2026_07.xlsx", allowed_suffixes={".xlsx"}), "SME170717GA0_2026_07.xlsx")
        with self.assertRaises(APIRequestError) as context:
            normalize_upload_filename("../../payload.exe", allowed_suffixes={".xlsx"})
        self.assertEqual(context.exception.code, "invalid_file_extension")

    def test_invalid_balanza_signature_is_rejected_and_temp_is_cleaned(self) -> None:
        handler = self._handler()
        fields = {
            "file": _Upload("../../SME170717GA0_2026_07.xlsx", b"not-a-workbook"),
            "catalog": _Upload("catalog.csv", b"1110,CAJA,,D,101\n"),
        }
        with patch("src.api.cgi.FieldStorage", return_value=fields):
            with self.assertRaises(APIRequestError) as context:
                handler._resolve_multipart_input()
        self.assertEqual(context.exception.code, "invalid_balanza_signature")


if __name__ == "__main__":
    unittest.main()
