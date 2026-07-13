"""Non-sensitive identity metadata for the running financial-statements API."""

from __future__ import annotations

from datetime import datetime, timezone
from functools import lru_cache
import json
from pathlib import Path
import shutil
import subprocess
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
SERVICE_VERSION = "estados-financieros-api-2026-07-13"
PROCESS_STARTED_AT = datetime.now(timezone.utc).isoformat()


def _generator_profile() -> tuple[str, str]:
    try:
        spec = json.loads((ROOT / "src" / "er_style_spec.json").read_text(encoding="utf-8"))
        return "manual-er-parity", str(spec.get("version") or "unknown")
    except (OSError, ValueError):
        return "manual-er-parity", "unknown"


def _git_value(*args: str) -> str:
    try:
        result = subprocess.run(
            ["git", *args],
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=2,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return "unknown"
    value = result.stdout.strip()
    return value if result.returncode == 0 and value else "unknown"


def _worktree_dirty() -> bool:
    try:
        result = subprocess.run(
            ["git", "status", "--porcelain", "--untracked-files=normal"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=2,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return bool(result.returncode == 0 and result.stdout.strip())


def _recalculation_engine() -> str:
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    return Path(executable).name if executable else "none"


@lru_cache(maxsize=1)
def runtime_identity() -> dict[str, Any]:
    profile, profile_version = _generator_profile()
    return {
        "service_version": SERVICE_VERSION,
        "git_commit": _git_value("rev-parse", "HEAD"),
        "worktree_dirty": _worktree_dirty(),
        "process_started_at": PROCESS_STARTED_AT,
        "generator_profile": profile,
        "generator_profile_version": profile_version,
    }


def build_runtime_metadata(
    *,
    generated_at: str | None = None,
    output_sha256: str | None = None,
    formula_static_validation: bool | None = None,
    formula_recalculation_performed: bool = False,
    formula_evaluated_error_count: int | None = None,
    formula_cached_values_available: bool = False,
) -> dict[str, Any]:
    metadata = dict(runtime_identity())
    metadata.update(
        {
            "generated_at": generated_at,
            "output_sha256": output_sha256,
            "formula_static_validation": formula_static_validation,
            "formula_recalculation_performed": formula_recalculation_performed,
            "formula_recalculation_engine": (
                _recalculation_engine() if formula_recalculation_performed else "none"
            ),
            "formula_validation_mode": (
                "recalculated" if formula_recalculation_performed else "static_only"
            ),
            "formula_evaluated_error_count": formula_evaluated_error_count,
            "formula_cached_values_available": formula_cached_values_available,
        }
    )
    return metadata


def sha256_file(path: str | Path) -> str:
    import hashlib

    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
