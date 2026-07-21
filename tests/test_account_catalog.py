from __future__ import annotations

import hashlib
import unittest

from src.account_catalog import (
    BLOCKING,
    OBSERVATION,
    WARNING,
    build_catalog_matrix,
    parse_account_catalog_bytes,
)


class AccountCatalogParserTest(unittest.TestCase):
    def test_empty_delta_catalog_is_blocking(self) -> None:
        parsed = parse_account_catalog_bytes(b"   ", source_name="DELTA cuentas.csv")

        self.assertFalse(parsed.is_valid)
        self.assertEqual(parsed.rows, ())
        self.assertEqual(parsed.source_sha256, hashlib.sha256(b"   ").hexdigest())
        self.assertIn("empty_catalog", {issue.code for issue in parsed.issues})
        self.assertEqual(parsed.issue_count(BLOCKING), 1)

    def test_folvaz_duplicate_6150_0005_is_blocking_without_deduplication(self) -> None:
        source = (
            "6155,\"PUBLICIDAD\",,D,604.61\n"
            "6150-0005,\"PUBLICIDAD DIGITAL\",6155,D,604.61\n"
            "6150-0005,\"PUBLICIDAD IMPRESA\",6155,D,604.61\n"
        ).encode("utf-8")

        parsed = parse_account_catalog_bytes(source, source_name="FOLVAZ cuentas.csv")

        duplicate = next(issue for issue in parsed.issues if issue.code == "duplicate_account_code")
        self.assertFalse(parsed.is_valid)
        self.assertEqual([row.account_code for row in parsed.rows].count("6150-0005"), 2)
        self.assertEqual(duplicate.severity, BLOCKING)
        self.assertEqual(duplicate.source_rows, (2, 3))

    def test_windows_1252_catalog_decodes_and_preserves_text_codes(self) -> None:
        source = "1130-AB12,\"CLIENTE PEÑA\",1130,D,105.01\n1130,CLIENTES,,D,105\n".encode(
            "cp1252"
        )

        parsed = parse_account_catalog_bytes(source, source_name="windows-1252.csv")

        self.assertTrue(parsed.is_valid)
        self.assertEqual(parsed.encoding, "windows-1252")
        self.assertEqual(parsed.rows[0].account_code, "1130-AB12")
        self.assertEqual(parsed.rows[0].account_name, "CLIENTE PEÑA")
        self.assertEqual(parsed.rows[0].parent_code, "1130")
        self.assertEqual(parsed.rows[0].sat_group_code, "105.01")

    def test_alphanumeric_parent_code_is_resolved_as_text(self) -> None:
        source = (
            "1130-AB12,\"CLIENTE ALFA\",1130,D,105.01\n"
            "1130-AB12-0001,\"SUBCUENTA\",1130-AB12,D,105.01\n"
            "1130,CLIENTES,,D,105\n"
        ).encode("utf-8")

        parsed = parse_account_catalog_bytes(source)

        self.assertTrue(parsed.is_valid)
        self.assertEqual(parsed.rows[1].parent_code, "1130-AB12")
        self.assertNotIn("orphan_parent_code", {issue.code for issue in parsed.issues})

    def test_hashes_distinguish_raw_source_but_not_equivalent_semantics(self) -> None:
        utf8_bom = b"\xef\xbb\xbf" + "1130,CLIENTES,,D,105\n".encode("utf-8")
        plain_utf8 = "1130,CLIENTES,,D,105\n".encode("utf-8")

        bom_catalog = parse_account_catalog_bytes(utf8_bom)
        plain_catalog = parse_account_catalog_bytes(plain_utf8)

        self.assertEqual(bom_catalog.encoding, "utf-8-sig")
        self.assertNotEqual(bom_catalog.source_sha256, plain_catalog.source_sha256)
        self.assertEqual(bom_catalog.semantic_sha256, plain_catalog.semantic_sha256)

    def test_invalid_fields_orphans_cycles_and_section_mismatch_keep_severity(self) -> None:
        source = (
            "1110,CAJA,9999,D,601.16\n"
            "2110,PROVEEDORES,2120,A,201\n"
            "2120,ACREEDORES,2110,X,201\n"
            "3110,CAPITAL,,D,30.12\n"
        ).encode("utf-8")

        parsed = parse_account_catalog_bytes(source)

        by_code = {issue.code: issue.severity for issue in parsed.issues}
        self.assertEqual(by_code["orphan_parent_code"], BLOCKING)
        self.assertEqual(by_code["parent_cycle"], BLOCKING)
        self.assertEqual(by_code["invalid_nature"], BLOCKING)
        self.assertEqual(by_code["invalid_sat_group_code"], BLOCKING)
        self.assertEqual(by_code["section_sat_mismatch"], WARNING)
        self.assertEqual(by_code["encoding_detected"], OBSERVATION)

    def test_matrix_excludes_raw_rows_and_customer_account_names(self) -> None:
        parsed = parse_account_catalog_bytes(b"1130,CLIENTES SECRETOS,,D,105\n", source_name="example.csv")

        matrix = build_catalog_matrix([parsed])

        self.assertEqual(matrix[0]["source_name"], "example.csv")
        self.assertEqual(matrix[0]["row_count"], 1)
        self.assertNotIn("rows", matrix[0])
        self.assertNotIn("CLIENTES SECRETOS", repr(matrix[0]))


if __name__ == "__main__":
    unittest.main()
