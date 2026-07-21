"""Minimal local HTTP API for the financial-statements prototype."""

from __future__ import annotations

import argparse
import base64
import cgi
import json
import re
import shutil
import tempfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

try:
    from .parser import OOXML_ZIP, detect_mime_kind
    from .prototype import RuntimeGenerationError, run_prototype
    from .runtime_metadata import build_runtime_metadata
except ImportError:  # pragma: no cover - supports direct module execution.
    from parser import OOXML_ZIP, detect_mime_kind
    from prototype import RuntimeGenerationError, run_prototype
    from runtime_metadata import build_runtime_metadata


def build_health_payload() -> dict[str, Any]:
    return {
        "status": "ok",
        "ready": True,
        "service": "estados-financieros",
        **build_runtime_metadata(),
    }


def _public_coverage_summary(coverage: dict[str, Any]) -> dict[str, Any]:
    """Expose only aggregate mapping controls to an untrusted browser."""

    assigned = int(coverage.get("assigned") or 0)
    unassigned = int(coverage.get("unassigned") or 0)
    ambiguous = int(coverage.get("ambiguous") or 0)
    duplicates = int(coverage.get("duplicates") or 0)
    total_accounts = assigned + unassigned + ambiguous
    percent = 0 if total_accounts == 0 else round((assigned / total_accounts) * 100, 2)
    controls = coverage.get("section_controls") or {}
    return {
        "assigned": assigned,
        "unassigned": unassigned,
        "ambiguous": ambiguous,
        "duplicates": duplicates,
        "total_accounts": total_accounts,
        "coverage_percent": percent,
        "complete": percent == 100 and unassigned == 0 and ambiguous == 0 and duplicates == 0,
        "section_controls": {
            str(section): {
                "difference": (control or {}).get("difference"),
                "status": (control or {}).get("status"),
            }
            for section, control in controls.items()
        },
    }


def _public_warning_codes(warnings: list[Any]) -> list[dict[str, str]]:
    """Warnings may contain account evidence; return only safe category codes."""

    safe: list[dict[str, str]] = []
    for warning in warnings:
        if isinstance(warning, dict):
            code = str(warning.get("code") or "review_required")
        else:
            code = "review_required"
        safe.append({"code": re.sub(r"[^A-Za-z0-9_.-]", "_", code)[:80] or "review_required"})
    return safe


def _safe_rfc_hint(value: Any) -> str:
    rfc = str(value or "").strip().upper()
    if len(rfc) < 7:
        return ""
    return f"{rfc[:3]}••••••{rfc[-3:]}"


def build_process_payload(report: dict[str, Any]) -> dict[str, Any]:
    runtime = report.get("runtime") or build_runtime_metadata(
        generated_at=report.get("generated_at"),
        output_sha256=report.get("output_sha256"),
    )
    profile = (report.get("engine") or {}).get("profile") or {}
    warnings = [
        *(report.get("engine") or {}).get("warnings", []),
        *(report.get("validation") or {}).get("warnings", []),
    ]
    return {
        "status": "ok",
        "period_ym": report["period"]["period_ym"],
        "period": {"period_ym": report["period"]["period_ym"]},
        "content_period_ym": report["content_period_ym"],
        "normalized_rows": report["parser"]["normalized_rows"],
        "leaf_rows_used_for_calculation": report["parser"]["leaf_rows_used_for_calculation"],
        "difference_cuadre": report["balance_check"]["difference_cuadre"],
        "tolerance": report["balance_check"]["tolerance"],
        "cuadra": report["balance_check"]["cuadra"],
        "balanza_no_cuadra": report["balance_check"]["balanza_no_cuadra"],
        "total_activo": report["balance_check"]["total_activo"],
        "total_pasivo": report["balance_check"]["total_pasivo"],
        "capital_contable": report["balance_check"]["capital_contable"],
        "formula_static_validation": report["validation"]["formula_static_validation"],
        "formula_recalculation_performed": report["validation"]["formula_recalculation_performed"],
        "formula_recalculation_engine": report["validation"]["formula_recalculation_engine"],
        "formula_validation_mode": report["validation"].get("formula_validation_mode")
        or runtime.get("formula_validation_mode"),
        "formula_evaluated_error_count": report["validation"]["formula_evaluated_error_count"],
        "formula_cached_values_available": report["validation"]["formula_cached_values_available"],
        "validation_ok": report["validation"]["ok"],
        "output_filename": Path(report["output_xlsx"]).name,
        "sheet_names": report["workbook"]["sheet_names"],
        "output_xlsx_base64": FinancialStatementsAPI._read_file_base64_static(report["output_xlsx"]),
        "accounting_profile_id": profile.get("accounting_profile_id"),
        "accounting_profile_version": profile.get("accounting_profile_version"),
        "accounting_profile_status": profile.get("accounting_profile_status"),
        "accounting_profile_company": profile.get("accounting_profile_company"),
        "accounting_profile_rfc_hint": _safe_rfc_hint(profile.get("accounting_profile_rfc")),
        "accounting_profile_valid_from": profile.get("accounting_profile_valid_from"),
        "accounting_profile_valid_to": profile.get("accounting_profile_valid_to"),
        "base_taxonomy_version": profile.get("base_taxonomy_version"),
        "catalog_source_sha256": profile.get("catalog_source_sha256"),
        "catalog_semantic_sha256_short": str(profile.get("catalog_semantic_sha256") or "")[:12],
        "generator_profile": profile.get("generator_profile_id") or runtime.get("generator_profile"),
        "generator_profile_version": profile.get("generator_profile_version") or runtime.get("generator_profile_version"),
        "coverage": _public_coverage_summary((report.get("engine") or {}).get("coverage", {})),
        "warnings": _public_warning_codes(warnings),
        **{key: value for key, value in runtime.items() if key not in {"generator_profile", "generator_profile_version"}},
    }


class APIRequestError(ValueError):
    def __init__(self, status: HTTPStatus, code: str):
        self.status = status
        self.code = code
        super().__init__(code)


def normalize_upload_filename(filename: str | None, *, allowed_suffixes: set[str]) -> str:
    """Keep only a safe basename; the original path is never persisted or returned."""

    candidate = Path(filename or "upload").name
    suffix = Path(candidate).suffix.lower()
    if suffix not in allowed_suffixes:
        raise APIRequestError(HTTPStatus.UNPROCESSABLE_ENTITY, "invalid_file_extension")
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", Path(candidate).stem).strip("._")[:96]
    if not stem:
        stem = "upload"
    return f"{stem}{suffix}"


def _single_upload_field(form: cgi.FieldStorage, name: str):
    if name not in form:
        return None
    field = form[name]
    if isinstance(field, list):
        raise APIRequestError(HTTPStatus.BAD_REQUEST, "duplicate_multipart_field")
    if not getattr(field, "file", None):
        raise APIRequestError(HTTPStatus.BAD_REQUEST, "invalid_multipart_field")
    return field


def _write_limited_upload(upload: Any, destination: Path, maximum: int) -> None:
    data = upload.file.read(maximum + 1)
    if len(data) > maximum:
        raise APIRequestError(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "uploaded_file_too_large")
    if not data:
        raise APIRequestError(HTTPStatus.BAD_REQUEST, "empty_uploaded_file")
    destination.write_bytes(data)


class FinancialStatementsAPI(BaseHTTPRequestHandler):
    server_version = "FinancialStatementsAPI/1.0"
    MAX_CONTENT_LENGTH = 10 * 1024 * 1024
    MAX_BALANZA_BYTES = 8 * 1024 * 1024
    MAX_CATALOG_BYTES = 2 * 1024 * 1024

    def do_GET(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler contract.
        if self.path == "/health":
            self._send_json(HTTPStatus.OK, build_health_payload())
            return
        self._send_json(
            HTTPStatus.NOT_FOUND,
            {"status": "error", "error": "not_found"},
        )

    def do_POST(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler contract.
        parsed_path = urlparse(self.path)
        if parsed_path.path != "/process":
            self._send_json(
                HTTPStatus.NOT_FOUND,
                {"status": "error", "error": "not_found"},
            )
            return

        cleanup_path: Path | None = None
        output_dir: Path | None = None
        try:
            input_path, catalog_path, cleanup_path = self._resolve_multipart_input()
            output_dir = Path(tempfile.mkdtemp(prefix="estados_financieros_output_"))
            try:
                report = run_prototype(
                    input_path,
                    catalog_path=catalog_path,
                    output_dir=output_dir,
                )
                payload = build_process_payload(report)
            finally:
                pass
        except RuntimeGenerationError as exc:
            self._send_json(
                HTTPStatus.UNPROCESSABLE_ENTITY,
                {"status": "error", "error": exc.code},
            )
            return
        except APIRequestError as exc:
            self._send_json(exc.status, {"status": "error", "error": exc.code})
            return
        except Exception as exc:  # pragma: no cover - surfaced to caller.
            self._send_json(
                HTTPStatus.UNPROCESSABLE_ENTITY,
                {
                    "status": "error",
                    "error": "invalid_upload",
                },
            )
            return
        finally:
            if cleanup_path is not None:
                shutil.rmtree(cleanup_path, ignore_errors=True)
            if output_dir is not None:
                shutil.rmtree(output_dir, ignore_errors=True)

        self._send_json(HTTPStatus.OK, payload)

    def do_HEAD(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler contract.
        if self.path == "/health":
            self.send_response(HTTPStatus.OK)
            self.end_headers()
            return
        self.send_response(HTTPStatus.NOT_FOUND)
        self.end_headers()

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A003
        # Keep the local API quiet by default.
        return

    def _resolve_multipart_input(self) -> tuple[Path, Path, Path]:
        content_type = self.headers.get("Content-Type", "")
        if not content_type.startswith("multipart/form-data"):
            raise APIRequestError(HTTPStatus.UNSUPPORTED_MEDIA_TYPE, "multipart_required")
        try:
            length = int(self.headers.get("Content-Length", "0") or 0)
        except ValueError as exc:
            raise APIRequestError(HTTPStatus.BAD_REQUEST, "invalid_content_length") from exc
        if length <= 0:
            raise APIRequestError(HTTPStatus.BAD_REQUEST, "empty_upload")
        if length > self.MAX_CONTENT_LENGTH:
            raise APIRequestError(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "content_length_exceeded")

        environ = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            "CONTENT_LENGTH": str(length),
        }
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=environ)
        upload = _single_upload_field(form, "file")
        catalog = _single_upload_field(form, "catalog")
        if upload is None or catalog is None:
            raise APIRequestError(HTTPStatus.BAD_REQUEST, "multipart_file_and_catalog_required")
        temp_dir = Path(tempfile.mkdtemp(prefix="estados_financieros_"))
        try:
            balance_name = normalize_upload_filename(getattr(upload, "filename", None), allowed_suffixes={".xls", ".xlsx"})
            catalog_name = normalize_upload_filename(getattr(catalog, "filename", None), allowed_suffixes={".csv"})
            temp_path = temp_dir / balance_name
            catalog_path = temp_dir / catalog_name
            _write_limited_upload(upload, temp_path, self.MAX_BALANZA_BYTES)
            _write_limited_upload(catalog, catalog_path, self.MAX_CATALOG_BYTES)
            if detect_mime_kind(temp_path) != OOXML_ZIP:
                raise APIRequestError(HTTPStatus.UNPROCESSABLE_ENTITY, "invalid_balanza_signature")
            return temp_path, catalog_path, temp_dir
        except Exception:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise

    def _send_json(self, status: HTTPStatus, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_file_base64(self, value: str | Path) -> str:
        return self._read_file_base64_static(value)

    @staticmethod
    def _read_file_base64_static(value: str | Path) -> str:
        path = Path(value)
        if not path.exists() or not path.is_file():
            return ""
        return base64.b64encode(path.read_bytes()).decode("ascii")


def serve(host: str = "127.0.0.1", port: int = 8080) -> None:
    server = ThreadingHTTPServer((host, port), FinancialStatementsAPI)
    print(f"Serving http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Local HTTP API for the states-financial prototype.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8080)
    args = parser.parse_args()
    serve(host=args.host, port=args.port)


if __name__ == "__main__":
    main()
