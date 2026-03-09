#!/usr/bin/env python3
"""Run Python and optional reference validators on a pinned ODF corpus."""

from __future__ import annotations

import argparse
import json
import os
import shlex
import subprocess
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import FileFormat, OdfValidator  # noqa: E402
from openxml_audit.parity_normalization import (  # noqa: E402
    normalize_description,
    normalize_error_tuple,
)

DEFAULT_CORPUS_MANIFEST = Path("data/odf/reference_corpus/manifest.json")
DEFAULT_FIXTURES_ROOT = Path("tests/fixtures/odf")
DEFAULT_OUTPUT = Path("reports/odf/reference_runs.json")
DEFAULT_STAGING_ROOT = Path("/tmp/openxml_audit_odf_reference")


@dataclass(frozen=True)
class RunnerConfig:
    """External reference runner configuration."""

    name: str
    command_template: list[str] | None


def _resolve_repo_path(path: Path) -> Path:
    if path.is_absolute():
        return path
    return (ROOT / path).resolve()


def _display_path(path: Path) -> str:
    resolved = path.resolve()
    try:
        return str(resolved.relative_to(ROOT))
    except ValueError:
        return str(resolved)


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _iter_files(source_dir: Path) -> list[Path]:
    return sorted(path for path in source_dir.rglob("*") if path.is_file())


def _build_odf_from_fixture_dir(source_dir: Path, output_path: Path) -> None:
    """Build a deterministic ODF ZIP package from a fixture directory."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    files = _iter_files(source_dir)
    mimetype_file = source_dir / "mimetype"

    with zipfile.ZipFile(output_path, "w") as zf:
        if mimetype_file.exists():
            zf.writestr(
                "mimetype",
                mimetype_file.read_bytes(),
                compress_type=zipfile.ZIP_STORED,
            )
        for file_path in files:
            rel = file_path.relative_to(source_dir).as_posix()
            if rel == "mimetype" and mimetype_file.exists():
                continue
            zf.writestr(rel, file_path.read_bytes(), compress_type=zipfile.ZIP_DEFLATED)


def _parse_command_template(raw: str | None) -> list[str] | None:
    if raw is None or not raw.strip():
        return None
    tokens = shlex.split(raw)
    return tokens or None


def _format_command(template: list[str], file_path: Path) -> list[str]:
    placeholder = "{file}"
    file_text = str(file_path)
    if any(placeholder in token for token in template):
        return [token.replace(placeholder, file_text) for token in template]
    return [*template, file_text]


def _looks_like_issue_line(line: str) -> bool:
    lower = line.lower()
    if lower.startswith("picked up _java_options"):
        return False
    return any(
        token in lower
        for token in (
            "error",
            "warning",
            "fatal",
            "invalid",
            "violation",
            "failed",
        )
    )


def _dedupe_stable(values: list[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        output.append(value)
    return output


def _extract_messages_from_json(payload: Any) -> list[str]:
    message_keys = ("message", "description", "detail", "text", "error", "warning")
    container_keys = ("errors", "warnings", "issues", "messages", "results", "violations")

    if isinstance(payload, str):
        value = payload.strip()
        return [value] if value else []
    if isinstance(payload, list):
        rows: list[str] = []
        for item in payload:
            rows.extend(_extract_messages_from_json(item))
        return rows
    if isinstance(payload, dict):
        rows: list[str] = []
        for key in message_keys:
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                rows.append(value.strip())
        for key in container_keys:
            rows.extend(_extract_messages_from_json(payload.get(key)))
        return rows
    return []


def _try_extract_json_messages(text: str) -> list[str] | None:
    stripped = text.strip()
    if not stripped:
        return []
    if stripped[0] not in "[{":
        return None
    try:
        payload = json.loads(stripped)
    except json.JSONDecodeError:
        return None
    return _extract_messages_from_json(payload)


def _extract_messages_from_text(text: str) -> list[str]:
    rows = [line.strip() for line in text.splitlines() if line.strip()]
    if not rows:
        return []
    issue_like = [line for line in rows if _looks_like_issue_line(line)]
    if issue_like:
        rows = issue_like
    return rows


def _parse_reference_messages(stdout: str, stderr: str) -> list[str]:
    messages: list[str] = []
    for chunk in (stdout, stderr):
        json_rows = _try_extract_json_messages(chunk)
        if json_rows is None:
            messages.extend(_extract_messages_from_text(chunk))
        else:
            messages.extend(json_rows)
    return _dedupe_stable([row for row in messages if row.strip()])


def _severity_from_message(message: str) -> str:
    lower = message.lower()
    if "warn" in lower:
        return "warning"
    if "info" in lower:
        return "info"
    return "error"


def _reference_issue_rows(tool: str, messages: list[str]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for message in messages:
        normalized = normalize_description(message)
        severity = _severity_from_message(message)
        rows.append(
            {
                "id": f"Ref_{tool}",
                "error_type": "reference",
                "part": "/",
                "path": "/",
                "severity": severity,
                "description": normalized,
                "family_key": f"{tool}|{severity}|{normalized}",
                "comparison_key": f"{severity}|{normalized}",
            }
        )
    return rows


def _preview(text: str, max_lines: int = 20, max_chars: int = 4000) -> str:
    if not text:
        return ""
    rows = text.splitlines()
    if len(rows) > max_lines:
        rows = rows[:max_lines]
        rows.append(f"... ({len(text.splitlines()) - max_lines} additional lines)")
    out = "\n".join(rows)
    if len(out) > max_chars:
        return out[:max_chars] + "... (truncated)"
    return out


def _run_python_validator(path: Path, strict: bool) -> dict[str, Any]:
    started = perf_counter()
    validator = OdfValidator(file_format=FileFormat.ODF_1_3, strict=strict)
    result = validator.validate(path)
    duration = perf_counter() - started

    issues: list[dict[str, str]] = []
    for error in result.errors:
        normalized = normalize_error_tuple(error)
        normalized_description = normalized["description"]
        severity = error.severity.value
        issues.append(
            {
                **normalized,
                "severity": severity,
                "comparison_key": f"{severity}|{normalized_description}",
            }
        )

    return {
        "status": "ok",
        "duration_seconds": round(duration, 6),
        "exit_code": 0,
        "error_count": result.error_count,
        "warning_count": result.warning_count,
        "issues": issues,
    }


def _run_reference_runner(
    config: RunnerConfig,
    file_path: Path,
    timeout_seconds: int,
) -> dict[str, Any]:
    if config.command_template is None:
        return {
            "status": "unavailable",
            "reason": "Command template not configured.",
            "issues": [],
        }

    command = _format_command(config.command_template, file_path)
    started = perf_counter()
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            timeout=timeout_seconds,
        )
    except FileNotFoundError as exc:
        duration = perf_counter() - started
        return {
            "status": "unavailable",
            "reason": f"Executable not found: {exc.filename}",
            "duration_seconds": round(duration, 6),
            "issues": [],
            "command": command,
        }
    except subprocess.TimeoutExpired:
        duration = perf_counter() - started
        return {
            "status": "timeout",
            "reason": f"Timed out after {timeout_seconds}s",
            "duration_seconds": round(duration, 6),
            "issues": [],
            "command": command,
        }

    duration = perf_counter() - started
    messages = _parse_reference_messages(completed.stdout, completed.stderr)
    issues = _reference_issue_rows(config.name, messages)
    status = "ok"
    if completed.returncode != 0 and not issues:
        status = "error"

    return {
        "status": status,
        "duration_seconds": round(duration, 6),
        "exit_code": completed.returncode,
        "issue_count": len(issues),
        "issues": issues,
        "command": command,
        "stdout_preview": _preview(completed.stdout),
        "stderr_preview": _preview(completed.stderr),
    }


def _load_samples(corpus_manifest: dict[str, Any]) -> list[dict[str, Any]]:
    samples = corpus_manifest.get("samples")
    if not isinstance(samples, list):
        return []
    valid_samples: list[dict[str, Any]] = []
    for sample in samples:
        if not isinstance(sample, dict):
            continue
        sample_id = sample.get("id")
        fixture_dir = sample.get("fixture_dir")
        filename = sample.get("filename")
        if not isinstance(sample_id, str):
            continue
        if not isinstance(fixture_dir, str):
            continue
        if not isinstance(filename, str):
            continue
        valid_samples.append(sample)
    return valid_samples


def _materialize_sample(
    sample: dict[str, Any],
    fixtures_root: Path,
    staging_root: Path,
) -> Path:
    source_dir = (fixtures_root / str(sample["fixture_dir"])).resolve()
    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"Fixture directory not found: {source_dir}")

    sample_id = str(sample["id"])
    filename = str(sample["filename"])
    out_path = staging_root / sample_id / filename
    _build_odf_from_fixture_dir(source_dir, out_path)
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run Python and reference validators for ODF corpus."
    )
    parser.add_argument(
        "--corpus-manifest",
        type=Path,
        default=DEFAULT_CORPUS_MANIFEST,
        help=f"Pinned corpus manifest path (default: {DEFAULT_CORPUS_MANIFEST})",
    )
    parser.add_argument(
        "--fixtures-root",
        type=Path,
        default=DEFAULT_FIXTURES_ROOT,
        help=f"Fixture root used by corpus samples (default: {DEFAULT_FIXTURES_ROOT})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output JSON report path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--staging-root",
        type=Path,
        default=DEFAULT_STAGING_ROOT,
        help=f"Staging directory for materialized ODF files (default: {DEFAULT_STAGING_ROOT})",
    )
    parser.add_argument(
        "--keep-staging",
        action="store_true",
        help="Keep staged files on disk (default is temporary directory cleanup).",
    )
    parser.add_argument(
        "--strict",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Use strict mode for Python ODF validator.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=30,
        help="Timeout per external validator invocation.",
    )
    parser.add_argument(
        "--odf-toolkit-cmd",
        type=str,
        default=os.getenv("ODF_TOOLKIT_CMD"),
        help=(
            "ODF Toolkit command template (optional). "
            "Use '{file}' placeholder, or file path is appended."
        ),
    )
    parser.add_argument(
        "--opf-cmd",
        type=str,
        default=os.getenv("OPF_ODF_VALIDATOR_CMD"),
        help=(
            "OPF validator command template (optional). "
            "Use '{file}' placeholder, or file path is appended."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute report and print summary without writing output file.",
    )
    args = parser.parse_args()

    manifest_path = _resolve_repo_path(args.corpus_manifest)
    fixtures_root = _resolve_repo_path(args.fixtures_root)
    output_path = _resolve_repo_path(args.output)
    staging_root = _resolve_repo_path(args.staging_root)

    if not manifest_path.exists():
        print(f"Corpus manifest not found: {manifest_path}")
        return 2
    if not fixtures_root.exists():
        print(f"Fixtures root not found: {fixtures_root}")
        return 2

    manifest = _load_json(manifest_path)
    samples = _load_samples(manifest)
    if not samples:
        print("No valid corpus samples found in manifest.")
        return 2

    runners = [
        RunnerConfig(
            name="odf_toolkit",
            command_template=_parse_command_template(args.odf_toolkit_cmd),
        ),
        RunnerConfig(
            name="opf",
            command_template=_parse_command_template(args.opf_cmd),
        ),
    ]

    temp_dir: TemporaryDirectory[str] | None = None
    if args.keep_staging:
        staging_root.mkdir(parents=True, exist_ok=True)
        active_staging_root = staging_root
    else:
        temp_dir = TemporaryDirectory(prefix="openxml_audit_odf_reference_")
        active_staging_root = Path(temp_dir.name)

    started = perf_counter()
    sample_reports: list[dict[str, Any]] = []
    runner_status: dict[str, Counter[str]] = {
        "python": Counter(),
        "odf_toolkit": Counter(),
        "opf": Counter(),
    }

    try:
        for sample in samples:
            sample_id = str(sample["id"])
            staged_file = _materialize_sample(sample, fixtures_root, active_staging_root)

            runs: dict[str, Any] = {}
            python_run = _run_python_validator(staged_file, strict=args.strict)
            runs["python"] = python_run
            runner_status["python"][str(python_run.get("status", "unknown"))] += 1

            for runner in runners:
                ref_run = _run_reference_runner(
                    config=runner,
                    file_path=staged_file,
                    timeout_seconds=args.timeout_seconds,
                )
                runs[runner.name] = ref_run
                runner_status[runner.name][str(ref_run.get("status", "unknown"))] += 1

            sample_reports.append(
                {
                    "id": sample_id,
                    "profile": sample.get("profile", ""),
                    "fixture_dir": sample["fixture_dir"],
                    "filename": sample["filename"],
                    "staged_relpath": f"{sample_id}/{sample['filename']}",
                    "runs": runs,
                }
            )
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()

    duration = perf_counter() - started
    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "contract_version": "odf-reference-v1",
        "corpus_manifest": _display_path(manifest_path),
        "fixtures_root": _display_path(fixtures_root),
        "strict": args.strict,
        "sample_count": len(sample_reports),
        "duration_seconds": round(duration, 6),
        "runners": {
            "python": {
                "status_counts": dict(runner_status["python"]),
                "command_template": None,
            },
            "odf_toolkit": {
                "status_counts": dict(runner_status["odf_toolkit"]),
                "command_template": runners[0].command_template,
            },
            "opf": {
                "status_counts": dict(runner_status["opf"]),
                "command_template": runners[1].command_template,
            },
        },
        "samples": sample_reports,
    }

    print(
        f"Processed {len(sample_reports)} samples in {report['duration_seconds']}s "
        f"(strict={args.strict})."
    )
    for runner_name in ("python", "odf_toolkit", "opf"):
        counts = report["runners"][runner_name]["status_counts"]
        print(f"- {runner_name}: {counts}")

    if not args.dry_run:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
        print(f"Report written to {output_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
