"""Euro-Office Document Server conversion oracle.

This engine exercises the released ``/converter`` contract.  It establishes
that Document Server can fetch and convert a file, then validates the returned
OOXML package and (for same-format conversions) records a canonical package
diff.  It does *not* drive the browser editor, a save callback, or coauthoring,
so its observations must not be presented as end-to-end editing evidence.

Secrets are intentionally environment-only::

    export EUROOFFICE_ORACLE_URL=https://office.example.test/
    export EUROOFFICE_ORACLE_JWT_SECRET=...
    openxml-audit-oracle eurooffice FILES... --source-base-url https://files.example.test/

HTTP(S) input URLs need no ``--source-base-url``.  Local inputs do: Document
Server must be able to fetch them from a URL reachable from its own container.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
import urllib.parse
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Literal, Protocol, cast

from openxml_audit.eurooffice import (
    CONNECTOR_VERSION,
    DOCUMENT_FORMATS_COMMIT,
    DOCUMENT_FORMATS_VERSION,
    DOCUMENT_SERVER_RELEASE,
    ConversionResult,
    EuroOfficeClient,
    EuroOfficeError,
    EuroOfficeFormatMode,
    format_support,
)
from openxml_audit.package_diff import compare_packages
from openxml_audit.validator import OpenXmlValidator

Outcome = Literal[
    "preserved",
    "rewritten",
    "converted",
    "unsupported",
    "request_failed",
    "download_failed",
    "invalid_output",
    "source_unavailable",
]

_OFFICE_EXTENSIONS = {
    "docx",
    "docm",
    "dotx",
    "dotm",
    "xlsx",
    "xlsm",
    "xltx",
    "xltm",
    "pptx",
    "pptm",
    "potx",
    "potm",
    "ppsx",
    "ppsm",
    "odt",
    "ott",
    "ods",
    "ots",
    "odp",
    "otp",
    "odg",
    # Validator-recognized families that the pinned connector does not edit.
    "odm",
    "oth",
    "odc",
    "odi",
    "odf",
    "odb",
    "otm",
    "otg",
    "xlam",
    "thmx",
    "ppam",
}


class ConversionClient(Protocol):
    """Surface needed by the orchestrator; enables network-free tests."""

    def healthcheck(self) -> bool: ...

    def version(self) -> str: ...

    def convert(
        self,
        *,
        source_url: str,
        source_format: str,
        target_format: str,
        key: str,
        title: str | None = None,
    ) -> ConversionResult: ...

    def download(self, file_url: str) -> bytes: ...


@dataclass
class EuroOfficeConversionObservation:
    """One conversion-path observation."""

    source_relpath: str
    source_format: str
    target_format: str | None
    format_mode: str
    outcome: Outcome
    duration_seconds: float
    server_version: str | None = None
    source_valid: bool | None = None
    source_error_count: int | None = None
    target_valid: bool | None = None
    target_error_count: int | None = None
    changed_parts: list[str] = field(default_factory=list)
    added_parts: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    size_in: int | None = None
    size_out: int | None = None
    sha256_in: str | None = None
    sha256_out: str | None = None
    artifact_path: str | None = None
    diff_dir: str | None = None
    notes: list[str] = field(default_factory=list)


def _is_url(value: str) -> bool:
    return urllib.parse.urlsplit(value).scheme in {"http", "https"}


def _source_name(value: str) -> str:
    if _is_url(value):
        return urllib.parse.unquote(Path(urllib.parse.urlsplit(value).path).name)
    return Path(value).name


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _conversion_key(source_url: str, target_format: str) -> str:
    return hashlib.sha256(f"{source_url}\0{target_format}".encode()).hexdigest()[:20]


def _local_source_url(path: Path, source_base_url: str | None) -> str:
    if not source_base_url:
        raise ValueError(
            "local inputs require --source-base-url or EUROOFFICE_ORACLE_SOURCE_BASE_URL"
        )
    parsed = urllib.parse.urlsplit(source_base_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("source base URL must use HTTP(S)")
    return urllib.parse.urljoin(
        source_base_url.rstrip("/") + "/",
        urllib.parse.quote(path.name),
    )


def _prepare_source(
    source: str,
    *,
    client: ConversionClient,
    work_dir: Path,
    source_base_url: str | None,
) -> tuple[str, Path | None, bytes | None, list[str]]:
    notes: list[str] = []
    if not _is_url(source):
        path = Path(source).expanduser().resolve()
        data = path.read_bytes()
        staged = work_dir / f"source{path.suffix.lower()}"
        shutil.copy2(path, staged)
        return _local_source_url(path, source_base_url), staged, data, notes

    suffix = Path(urllib.parse.urlsplit(source).path).suffix.lower()
    staged = work_dir / f"source{suffix}"
    try:
        data = client.download(source)
    except EuroOfficeError:
        notes.append("source was not downloadable locally; package diff is unavailable")
        return source, None, None, notes
    staged.write_bytes(data)
    return source, staged, data, notes


def _validate_ooxml(path: Path) -> tuple[bool, int, frozenset[tuple[str, str, str, str]]]:
    result = OpenXmlValidator().validate(path)
    signatures = frozenset(
        (
            error.severity.value,
            error.error_type.value,
            error.part_uri,
            error.description,
        )
        for error in result.errors
    )
    return result.is_valid, result.error_count, signatures


def observe(
    source: str | Path,
    *,
    client: ConversionClient,
    server_version: str | None = None,
    source_base_url: str | None = None,
    work_root: Path | None = None,
    keep_artifacts: bool = False,
) -> EuroOfficeConversionObservation:
    """Exercise one file through the synchronous conversion endpoint."""

    started = time.perf_counter()
    source_text = str(source)
    name = _source_name(source_text)
    capability = format_support(name)
    observation = EuroOfficeConversionObservation(
        source_relpath=name or source_text,
        source_format=capability.extension,
        target_format=capability.conversion_target,
        format_mode=capability.mode.value,
        outcome="unsupported",
        duration_seconds=0.0,
        server_version=server_version,
    )
    if capability.mode is EuroOfficeFormatMode.UNSUPPORTED:
        observation.notes.append(
            f"not supported by connector {CONNECTOR_VERSION}'s pinned format matrix"
        )
        observation.duration_seconds = time.perf_counter() - started
        return observation

    root = work_root.expanduser().resolve() if work_root else None
    if root:
        root.mkdir(parents=True, exist_ok=True)
    work_dir = Path(tempfile.mkdtemp(prefix="eurooffice-oracle-", dir=root))
    try:
        try:
            source_url, source_path, source_bytes, source_notes = _prepare_source(
                source_text,
                client=client,
                work_dir=work_dir,
                source_base_url=source_base_url,
            )
        except (OSError, ValueError) as exc:
            observation.outcome = "source_unavailable"
            observation.notes.append(str(exc))
            return observation

        observation.notes.extend(source_notes)
        if source_bytes is not None:
            observation.size_in = len(source_bytes)
            observation.sha256_in = _sha256(source_bytes)

        target_format = capability.conversion_target
        if target_format is None:  # guarded above; keeps the invariant explicit
            observation.outcome = "unsupported"
            return observation
        source_findings: frozenset[tuple[str, str, str, str]] | None = None
        if source_path is not None and capability.extension == target_format:
            source_valid, source_error_count, source_findings = _validate_ooxml(source_path)
            observation.source_valid = source_valid
            observation.source_error_count = source_error_count
        key = _conversion_key(source_url, target_format)
        try:
            result = client.convert(
                source_url=source_url,
                source_format=capability.extension,
                target_format=target_format,
                key=key,
                title=name,
            )
        except EuroOfficeError as exc:
            observation.outcome = "request_failed"
            observation.notes.append(str(exc))
            return observation

        try:
            output_bytes = client.download(result.file_url)
        except EuroOfficeError as exc:
            observation.outcome = "download_failed"
            observation.notes.append(str(exc))
            return observation

        output_path = work_dir / f"converted.{target_format}"
        output_path.write_bytes(output_bytes)
        observation.size_out = len(output_bytes)
        observation.sha256_out = _sha256(output_bytes)

        same_format_diff: dict[str, object] | None = None
        if source_path is not None and capability.extension == target_format:
            diff_dir = work_dir / "package-diff"
            try:
                same_format_diff = compare_packages(
                    base_path=source_path,
                    head_path=output_path,
                    output_dir=diff_dir,
                )
            except (OSError, ValueError, zipfile.BadZipFile):
                observation.notes.append("converted package could not be diffed")
            else:
                observation.changed_parts = list(same_format_diff["changed_files"])
                observation.added_parts = list(same_format_diff["added_files"])
                observation.removed_parts = list(same_format_diff["removed_files"])
                if keep_artifacts:
                    observation.diff_dir = str(diff_dir)

        valid, error_count, target_findings = _validate_ooxml(output_path)
        observation.target_valid = valid
        observation.target_error_count = error_count
        if not valid:
            observation.outcome = "invalid_output"
            observation.notes.append("converted package failed openxml-audit validation")
            if source_findings is not None and source_findings == target_findings:
                observation.notes.append(
                    "target validator findings match the source; no new finding was introduced"
                )
            return observation

        if same_format_diff is not None:
            observation.outcome = (
                "rewritten"
                if observation.changed_parts or observation.added_parts or observation.removed_parts
                else "preserved"
            )
        elif source_path is None and capability.extension == target_format:
            observation.outcome = "source_unavailable"
        else:
            observation.outcome = "converted"

        if capability.mode is EuroOfficeFormatMode.LOSSY_EDIT:
            observation.notes.append(
                f"connector editing path converts {capability.extension} to {target_format}"
            )
        elif capability.mode is EuroOfficeFormatMode.VIEW_ONLY:
            observation.notes.append(
                f"{capability.extension} is view/convert-only, not an editable connector format"
            )
        return observation
    finally:
        observation.duration_seconds = time.perf_counter() - started
        if keep_artifacts:
            converted = work_dir / f"converted.{capability.conversion_target}"
            observation.artifact_path = str(converted) if converted.exists() else None
        else:
            shutil.rmtree(work_dir, ignore_errors=True)


def build_report(
    observations: list[EuroOfficeConversionObservation],
    *,
    server_version: str | None,
) -> dict[str, object]:
    """Build the stable JSON report envelope."""

    counts: dict[str, int] = {}
    for observation in observations:
        counts[observation.outcome] = counts.get(observation.outcome, 0) + 1
    return {
        "schema_version": 1,
        "engine": "eurooffice",
        "evidence_scope": "Document Server conversion endpoint; not browser editing",
        "upstream": {
            "document_server_version": server_version,
            "document_server_release": DOCUMENT_SERVER_RELEASE,
            "nextcloud_connector_version": CONNECTOR_VERSION,
            "document_formats_version": DOCUMENT_FORMATS_VERSION,
            "document_formats_commit": DOCUMENT_FORMATS_COMMIT,
        },
        "observations": [asdict(observation) for observation in observations],
        "summary": {"total": len(observations), "outcomes": counts},
    }


def _expand_inputs(values: list[str]) -> list[str]:
    inputs: list[str] = []
    for value in values:
        if _is_url(value):
            inputs.append(value)
            continue
        path = Path(value).expanduser()
        if path.is_dir():
            inputs.extend(
                str(candidate)
                for candidate in sorted(path.rglob("*"))
                if candidate.is_file()
                and candidate.suffix.lower().lstrip(".") in _OFFICE_EXTENSIONS
            )
        elif path.is_file():
            inputs.append(str(path))
    return inputs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="+", help="office files, directories, or HTTP(S) URLs")
    parser.add_argument(
        "--server-url",
        default=None,
        help="Document Server base URL; defaults to EUROOFFICE_ORACLE_URL",
    )
    parser.add_argument(
        "--source-base-url",
        default=None,
        help="public base URL serving local inputs; defaults to EUROOFFICE_ORACLE_SOURCE_BASE_URL",
    )
    parser.add_argument("--output", type=Path, default=None, help="write JSON report here")
    parser.add_argument("--work-root", type=Path, default=None, help="parent for temporary runs")
    parser.add_argument("--timeout", type=float, default=120.0, help="HTTP timeout in seconds")
    parser.add_argument(
        "--keep-artifacts", action="store_true", help="retain converted files and XML diffs"
    )
    args = parser.parse_args()

    inputs = _expand_inputs(args.input)
    if not inputs:
        print("no office inputs found", file=sys.stderr)
        return 2

    try:
        client = EuroOfficeClient.from_env(args.server_url, timeout=args.timeout)
        if not client.healthcheck():
            print("Euro-Office healthcheck did not return true", file=sys.stderr)
            return 1
        server_version = client.version()
    except EuroOfficeError as exc:
        print(f"Euro-Office preflight failed: {exc}", file=sys.stderr)
        return 1

    source_base_url = args.source_base_url
    if source_base_url is None:
        source_base_url = os.environ.get("EUROOFFICE_ORACLE_SOURCE_BASE_URL")

    observations = [
        observe(
            item,
            client=client,
            server_version=server_version,
            source_base_url=source_base_url,
            work_root=args.work_root,
            keep_artifacts=args.keep_artifacts,
        )
        for item in inputs
    ]
    report = build_report(observations, server_version=server_version)
    rendered = json.dumps(report, indent=2, sort_keys=True)
    if args.output:
        args.output.write_text(rendered + "\n", encoding="utf-8")
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(rendered)

    summary = cast(dict[str, object], report["summary"])
    outcomes = cast(dict[str, int], summary["outcomes"])
    print(
        "eurooffice-oracle: "
        + " ".join(f"{name}={count}" for name, count in sorted(outcomes.items())),
        file=sys.stderr,
    )
    hard_failures = sum(
        outcomes.get(name, 0)
        for name in ("request_failed", "download_failed", "invalid_output", "source_unavailable")
    )
    return 1 if hard_failures else 0


if __name__ == "__main__":
    sys.exit(main())
