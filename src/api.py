"""Minimal local HTTP API for the financial-statements prototype."""

from __future__ import annotations

import argparse
import base64
import cgi
import json
import shutil
import tempfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

try:
    from .prototype import run_prototype
    from .runtime_metadata import build_runtime_metadata
except ImportError:  # pragma: no cover - supports direct module execution.
    from prototype import run_prototype
    from runtime_metadata import build_runtime_metadata


def build_health_payload() -> dict[str, Any]:
    return {
        "status": "ok",
        "ready": True,
        "service": "estados-financieros",
        **build_runtime_metadata(),
    }


def build_process_payload(report: dict[str, Any]) -> dict[str, Any]:
    runtime = report.get("runtime") or build_runtime_metadata(
        generated_at=report.get("generated_at"),
        output_sha256=report.get("output_sha256"),
    )
    return {
        "status": "ok",
        "period_ym": report["period"]["period_ym"],
        "period": report["period"],
        "company_name": report["company_name"],
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
        "componentes_balance": report["balance_check"]["componentes"],
        "formula_static_validation": report["validation"]["formula_static_validation"],
        "formula_recalculation_performed": report["validation"]["formula_recalculation_performed"],
        "formula_recalculation_engine": report["validation"]["formula_recalculation_engine"],
        "formula_validation_mode": report["validation"].get("formula_validation_mode")
        or runtime.get("formula_validation_mode"),
        "formula_evaluated_error_count": report["validation"]["formula_evaluated_error_count"],
        "formula_cached_values_available": report["validation"]["formula_cached_values_available"],
        "validation_ok": report["validation"]["ok"],
        "output_filename": Path(report["output_xlsx"]).name,
        "output_xlsx_base64": FinancialStatementsAPI._read_file_base64_static(report["output_xlsx"]),
        "warnings": report["validation"]["warnings"],
        **runtime,
    }


class FinancialStatementsAPI(BaseHTTPRequestHandler):
    server_version = "FinancialStatementsAPI/1.0"

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

        try:
            input_path, cleanup_path = self._resolve_input_path()
            try:
                report = run_prototype(input_path)
            finally:
                if cleanup_path is not None:
                    if cleanup_path.is_dir():
                        shutil.rmtree(cleanup_path, ignore_errors=True)
                    else:
                        cleanup_path.unlink(missing_ok=True)
        except Exception as exc:  # pragma: no cover - surfaced to caller.
            self._send_json(
                HTTPStatus.BAD_REQUEST,
                {
                    "status": "error",
                    "error": type(exc).__name__,
                    "message": str(exc),
                },
            )
            return

        payload = build_process_payload(report)
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

    def _resolve_input_path(self) -> tuple[Path, Path | None]:
        content_type = self.headers.get("Content-Type", "")

        if content_type.startswith("multipart/form-data"):
            return self._resolve_multipart_input()

        if content_type.startswith("application/json"):
            raw_length = int(self.headers.get("Content-Length", "0") or 0)
            body = self.rfile.read(raw_length) if raw_length else b"{}"
            data = json.loads(body.decode("utf-8") or "{}")
            value = data.get("input_path") or data.get("path")
            if not value:
                raise ValueError("Missing 'input_path' in JSON body.")
            return Path(value), None

        if content_type.startswith("application/x-www-form-urlencoded"):
            raw_length = int(self.headers.get("Content-Length", "0") or 0)
            body = self.rfile.read(raw_length).decode("utf-8")
            params = parse_qs(body)
            value = (params.get("input_path") or params.get("path") or [None])[0]
            if not value:
                raise ValueError("Missing 'input_path' in form body.")
            return Path(value), None

        raise ValueError(
            "Unsupported content type. Use multipart/form-data, application/json or application/x-www-form-urlencoded."
        )

    def _resolve_multipart_input(self) -> tuple[Path, Path | None]:
        length = int(self.headers.get("Content-Length", "0") or 0)
        if length <= 0:
            raise ValueError("Empty multipart body.")

        environ = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            "CONTENT_LENGTH": str(length),
        }
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=environ)
        upload = form["file"] if "file" in form else None
        if upload is None or not getattr(upload, "file", None):
            raise ValueError("Missing multipart field 'file'.")

        filename = getattr(upload, "filename", None) or "upload.xls"
        temp_dir = Path(tempfile.mkdtemp(prefix="estados_financieros_"))
        temp_path = temp_dir / Path(filename).name
        temp_path.write_bytes(upload.file.read())
        return temp_path, temp_dir

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
