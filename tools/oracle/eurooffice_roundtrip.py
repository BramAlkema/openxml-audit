"""EuroOffice DOCX editor round-trip oracle.

This is an editor oracle, not a converter wrapper.  The production transport
uploads a DOCX to EuroOffice's Apache-licensed example app, opens it in a
pinned Playwright browser, inserts and saves a unique marker, reopens it,
removes and saves the marker, then downloads the final package.  Both passes
must make EuroOffice's own document-state event report a dirty session and
must produce a new stored package hash.

Package diff, DOCX semantic comparison, and strict schema/security validation
remain independent evidence.  A successful callback save is never reported
as semantic preservation by itself.
"""

from __future__ import annotations

import argparse
import contextlib
import hashlib
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
import uuid
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal, Protocol, cast

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from openxml_audit.docx.semantic_snapshot import (  # noqa: E402
    compare_docx_semantics,
    snapshot_docx,
)
from openxml_audit.package_diff import compare_packages  # noqa: E402
from openxml_audit.validator import OpenXmlValidator  # noqa: E402

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DEFAULT_PLAYWRIGHT_IMAGE = "mcr.microsoft.com/playwright:v1.62.0-noble"
DEFAULT_STAGE_PARENT = Path.home() / "Documents" / ".eurooffice_oracle_runs"
SAFE_IDENTIFIER = re.compile(r"^[A-Za-z0-9_.-]+$")


class EuroOfficeTransportError(RuntimeError):
    """One externally observable editor-transport stage failed."""

    def __init__(self, stage: str, message: str) -> None:
        super().__init__(message)
        self.stage = stage


@dataclass(frozen=True)
class EditorPassEvidence:
    action: str
    opened: bool
    dirty: bool
    disconnect: str
    page_url: str
    frame_url: str
    diagnostics: tuple[str, ...] = ()


@dataclass(frozen=True)
class EditorTransportEvidence:
    remote_filename: str
    browser_image: str
    uploaded_sha256: str
    inserted_sha256: str
    final_sha256: str
    passes: tuple[EditorPassEvidence, ...]


class EuroOfficeClient(Protocol):
    def roundtrip(self, input_path: Path, output_path: Path) -> EditorTransportEvidence:
        """Produce an editor-saved DOCX at ``output_path``."""


CommandRunner = Callable[..., subprocess.CompletedProcess[bytes]]


class SSHExampleEuroOfficeClient:
    """Drive the bundled EuroOffice example app on a Docker host over SSH."""

    def __init__(
        self,
        *,
        host: str = "dtda-server",
        container: str = "eurooffice-documentserver",
        browser_image: str = DEFAULT_PLAYWRIGHT_IMAGE,
        browser_script: Path | None = None,
        save_timeout_seconds: float = 45.0,
        poll_interval_seconds: float = 2.0,
        runner: CommandRunner = subprocess.run,
    ) -> None:
        for label, value in {
            "host": host,
            "container": container,
            "browser image": browser_image,
        }.items():
            if not SAFE_IDENTIFIER.fullmatch(value.replace("/", "_", 1).replace(":", "_", 1)):
                raise ValueError(f"unsafe {label}: {value!r}")
        self.host = host
        self.container = container
        self.browser_image = browser_image
        self.browser_script = browser_script or Path(__file__).with_name(
            "eurooffice_editor_roundtrip.mjs"
        )
        self.save_timeout_seconds = save_timeout_seconds
        self.poll_interval_seconds = poll_interval_seconds
        self._runner = runner
        self._remote_script = "/tmp/openxml-audit-eurooffice-editor-roundtrip.mjs"

    def _run(
        self,
        argv: list[str],
        *,
        stage: str,
        input_bytes: bytes | None = None,
    ) -> bytes:
        try:
            result = self._runner(
                argv,
                input=input_bytes,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )
        except OSError as exc:
            raise EuroOfficeTransportError(stage, str(exc)) from exc
        if result.returncode != 0:
            stderr = result.stderr.decode("utf-8", errors="replace").strip()
            raise EuroOfficeTransportError(
                stage,
                f"command failed ({result.returncode}): {stderr[:1000]}",
            )
        return result.stdout

    def _ssh(self, command: str, *, stage: str, input_bytes: bytes | None = None) -> bytes:
        return self._run(
            ["ssh", self.host, command],
            stage=stage,
            input_bytes=input_bytes,
        )

    def _stage_browser_script(self) -> None:
        if not self.browser_script.is_file():
            raise EuroOfficeTransportError(
                "browser_stage",
                f"browser script not found: {self.browser_script}",
            )
        self._run(
            [
                "rsync",
                "-a",
                str(self.browser_script),
                f"{self.host}:{self._remote_script}",
            ],
            stage="browser_stage",
        )

    def _upload(self, payload: bytes, remote_filename: str) -> str:
        form_value = f"uploadedFile=@-;filename={remote_filename};type={DOCX_MIME}"
        command = shlex.join(
            [
                "docker",
                "exec",
                "-i",
                self.container,
                "curl",
                "-fsS",
                "-F",
                form_value,
                "http://127.0.0.1/example/upload",
            ]
        )
        response = self._ssh(command, stage="upload", input_bytes=payload)
        try:
            data = json.loads(response)
        except json.JSONDecodeError as exc:
            raise EuroOfficeTransportError("upload", "upload returned invalid JSON") from exc
        actual_filename = data.get("filename")
        if (
            data.get("documentType") != "word"
            or not isinstance(actual_filename, str)
            or not SAFE_IDENTIFIER.fullmatch(actual_filename)
        ):
            raise EuroOfficeTransportError("upload", f"unexpected upload response: {data!r}")
        return actual_filename

    def _browser_pass(
        self,
        remote_filename: str,
        marker: str,
        action: Literal["insert", "remove"],
    ) -> EditorPassEvidence:
        node_command = (
            "mkdir -p /tmp/pw && cd /tmp/pw && "
            "npm install --silent playwright@1.62.0 && "
            "cp /tmp/probe.mjs ./probe.mjs && node probe.mjs"
        )
        command = shlex.join(
            [
                "docker",
                "run",
                "--rm",
                "--init",
                "--ipc=host",
                "--network",
                f"container:{self.container}",
                "-v",
                f"{self._remote_script}:/tmp/probe.mjs:ro",
                "-e",
                "HOME=/tmp",
                "-e",
                "EUROOFFICE_EXAMPLE_URL=http://127.0.0.1/example/",
                "-e",
                f"EUROOFFICE_FILE_NAME={remote_filename}",
                "-e",
                f"EUROOFFICE_PROBE_ACTION={action}",
                "-e",
                f"EUROOFFICE_PROBE_MARKER={marker}",
                self.browser_image,
                "sh",
                "-lc",
                node_command,
            ]
        )
        raw = self._ssh(command, stage=f"editor_{action}")
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise EuroOfficeTransportError(
                f"editor_{action}",
                f"browser returned invalid JSON: {raw[-1000:]!r}",
            ) from exc
        if data.get("opened") is not True or data.get("dirty") is not True:
            raise EuroOfficeTransportError(
                f"editor_{action}",
                f"browser did not prove an opened dirty session: {data!r}",
            )
        return EditorPassEvidence(
            action=action,
            opened=True,
            dirty=True,
            disconnect=str(data.get("disconnect", "")),
            page_url=str(data.get("pageUrl", "")),
            frame_url=str(data.get("frameUrl", "")),
            diagnostics=tuple(str(value) for value in data.get("diagnostics", [])),
        )

    def _download(self, remote_filename: str) -> bytes:
        url = f"http://127.0.0.1/example/download?fileName={remote_filename}"
        command = shlex.join(
            ["docker", "exec", self.container, "curl", "-fsS", url]
        )
        return self._ssh(command, stage="download")

    def _wait_for_new_hash(self, remote_filename: str, previous_hash: str) -> tuple[bytes, str]:
        deadline = time.monotonic() + self.save_timeout_seconds
        last_hash = previous_hash
        while time.monotonic() < deadline:
            payload = self._download(remote_filename)
            last_hash = hashlib.sha256(payload).hexdigest()
            if last_hash != previous_hash:
                return payload, last_hash
            time.sleep(self.poll_interval_seconds)
        raise EuroOfficeTransportError(
            "save_callback",
            f"stored DOCX hash did not change within {self.save_timeout_seconds:g}s "
            f"(last={last_hash})",
        )

    def _delete(self, remote_filename: str) -> None:
        url = f"http://127.0.0.1/example/file?filename={remote_filename}"
        command = shlex.join(
            ["docker", "exec", self.container, "curl", "-fsS", "-X", "DELETE", url]
        )
        self._ssh(command, stage="cleanup")

    def roundtrip(self, input_path: Path, output_path: Path) -> EditorTransportEvidence:
        payload = input_path.read_bytes()
        uploaded_hash = hashlib.sha256(payload).hexdigest()
        remote_filename = f"oa-{uuid.uuid4().hex[:12]}.docx"
        marker = f"ooxmlprobe{uuid.uuid4().hex}"
        passes: list[EditorPassEvidence] = []
        self._stage_browser_script()
        remote_filename = self._upload(payload, remote_filename)
        try:
            stored = self._download(remote_filename)
            if hashlib.sha256(stored).hexdigest() != uploaded_hash:
                raise EuroOfficeTransportError(
                    "upload",
                    "example storage bytes differ from the uploaded DOCX before editing",
                )
            passes.append(self._browser_pass(remote_filename, marker, "insert"))
            _inserted, inserted_hash = self._wait_for_new_hash(
                remote_filename,
                uploaded_hash,
            )
            passes.append(self._browser_pass(remote_filename, marker, "remove"))
            final, final_hash = self._wait_for_new_hash(remote_filename, inserted_hash)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_bytes(final)
            return EditorTransportEvidence(
                remote_filename=remote_filename,
                browser_image=self.browser_image,
                uploaded_sha256=uploaded_hash,
                inserted_sha256=inserted_hash,
                final_sha256=final_hash,
                passes=tuple(passes),
            )
        finally:
            with contextlib.suppress(EuroOfficeTransportError):
                self._delete(remote_filename)


Outcome = Literal[
    "preserved",
    "schema_regression",
    "semantic_drift",
    "editor_failed",
    "missing_output",
]
OUTCOMES: tuple[Outcome, ...] = (
    "preserved",
    "schema_regression",
    "semantic_drift",
    "editor_failed",
    "missing_output",
)


@dataclass
class EuroOfficeRoundtripObservation:
    source_relpath: str
    outcome: Outcome
    duration_seconds: float
    target_format: str = "docx"
    changed_parts: list[str] = field(default_factory=list)
    added_parts: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    semantic_comparison: dict[str, Any] | None = None
    validation_before: dict[str, Any] | None = None
    validation_after: dict[str, Any] | None = None
    editor_evidence: dict[str, Any] | None = None
    diff_dir: str | None = None
    roundtripped_path: str | None = None
    notes: list[str] = field(default_factory=list)


def _stage_root() -> Path:
    override = os.environ.get("EUROOFFICE_ORACLE_STAGE")
    root = Path(override).expanduser() if override else DEFAULT_STAGE_PARENT
    root.mkdir(parents=True, exist_ok=True)
    return root


def _validation_summary(path: Path) -> dict[str, Any]:
    result = OpenXmlValidator(
        max_errors=250,
        security_validation=True,
        strict=True,
    ).validate(path)
    return {
        "valid": result.is_valid,
        "error_count": result.error_count,
        "warning_count": result.warning_count,
        "findings": [
            {
                "type": finding.error_type.value,
                "severity": finding.severity.value,
                "description": finding.description,
                "part_uri": finding.part_uri,
                "path": finding.path,
            }
            for finding in result.errors[:100]
        ],
        "findings_truncated": len(result.errors) > 100,
    }


def observe(
    input_path: Path,
    *,
    client: EuroOfficeClient | None = None,
    keep_artifacts: bool = False,
) -> EuroOfficeRoundtripObservation:
    """Roundtrip one DOCX and keep transport, semantic, and validity evidence separate."""

    source = input_path.resolve()
    if not source.is_file() or source.suffix.lower() != ".docx":
        return EuroOfficeRoundtripObservation(
            source_relpath=source.name,
            outcome="missing_output",
            duration_seconds=0.0,
            notes=[f"DOCX input does not exist or is unsupported: {source}"],
        )

    started = time.monotonic()
    work_dir = _stage_root() / f"run-{uuid.uuid4().hex[:8]}"
    work_dir.mkdir(parents=True, exist_ok=False)
    original = work_dir / f"original_{source.name}"
    output = work_dir / f"roundtripped_{source.name}"
    shutil.copy2(source, original)
    active_client = client or SSHExampleEuroOfficeClient()

    try:
        try:
            editor = active_client.roundtrip(original, output)
        except EuroOfficeTransportError as exc:
            return EuroOfficeRoundtripObservation(
                source_relpath=source.name,
                outcome="editor_failed",
                duration_seconds=time.monotonic() - started,
                notes=[f"{exc.stage}: {exc}"],
            )
        if not output.is_file():
            return EuroOfficeRoundtripObservation(
                source_relpath=source.name,
                outcome="missing_output",
                duration_seconds=time.monotonic() - started,
                editor_evidence=asdict(editor),
                notes=["editor transport returned without a DOCX output"],
            )

        compare_dir = work_dir / "compare"
        package_report = compare_packages(
            base_path=original,
            head_path=output,
            output_dir=compare_dir,
        )
        semantics = compare_docx_semantics(snapshot_docx(original), snapshot_docx(output))
        validation_before = _validation_summary(original)
        validation_after = _validation_summary(output)
        schema_regression = (
            validation_after["error_count"] > validation_before["error_count"]
            or (validation_before["valid"] and not validation_after["valid"])
        )
        if not semantics.preserved:
            outcome: Outcome = "semantic_drift"
        elif schema_regression:
            outcome = "schema_regression"
        else:
            outcome = "preserved"

        return EuroOfficeRoundtripObservation(
            source_relpath=source.name,
            outcome=outcome,
            duration_seconds=time.monotonic() - started,
            changed_parts=list(cast(list[str], package_report.get("changed_files", []))),
            added_parts=list(cast(list[str], package_report.get("added_files", []))),
            removed_parts=list(cast(list[str], package_report.get("removed_files", []))),
            semantic_comparison=semantics.to_dict(),
            validation_before=validation_before,
            validation_after=validation_after,
            editor_evidence=asdict(editor),
            diff_dir=str(compare_dir) if keep_artifacts else None,
            roundtripped_path=str(output) if keep_artifacts else None,
            notes=["strict schema/security regression detected"] if schema_regression else [],
        )
    finally:
        if not keep_artifacts:
            shutil.rmtree(work_dir, ignore_errors=True)


def _to_jsonable(observations: list[EuroOfficeRoundtripObservation]) -> dict[str, Any]:
    outcomes = dict.fromkeys(OUTCOMES, 0)
    changed_features: dict[str, int] = {}
    for observation in observations:
        outcomes[observation.outcome] += 1
        comparison = observation.semantic_comparison or {}
        for feature in comparison.get("changed_features", []):
            changed_features[feature] = changed_features.get(feature, 0) + 1
    return {
        "schema_version": 1,
        "engine": "eurooffice-editor",
        "observations": [asdict(observation) for observation in observations],
        "summary": {
            "total": len(observations),
            **outcomes,
            "changed_feature_counts": changed_features,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="+", type=Path, help="DOCX files or directories")
    parser.add_argument("--output", type=Path, help="write JSON report to this path")
    parser.add_argument("--host", default="dtda-server", help="SSH Docker host")
    parser.add_argument(
        "--container",
        default="eurooffice-documentserver",
        help="EuroOffice Document Server container",
    )
    parser.add_argument(
        "--browser-image",
        default=DEFAULT_PLAYWRIGHT_IMAGE,
        help="pinned Playwright browser image",
    )
    parser.add_argument("--keep-artifacts", action="store_true")
    args = parser.parse_args()

    inputs: list[Path] = []
    for entry in args.input:
        if entry.is_dir():
            inputs.extend(sorted(entry.rglob("*.docx")))
        elif entry.is_file() and entry.suffix.lower() == ".docx":
            inputs.append(entry)
    if not inputs:
        print("no .docx inputs found", file=sys.stderr)
        return 2

    client = SSHExampleEuroOfficeClient(
        host=args.host,
        container=args.container,
        browser_image=args.browser_image,
    )
    observations = [
        observe(path, client=client, keep_artifacts=args.keep_artifacts) for path in inputs
    ]
    report = _to_jsonable(observations)
    rendered = json.dumps(report, indent=2, sort_keys=True)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered + "\n", encoding="utf-8")
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(rendered)
    hard_outcomes = {"schema_regression", "semantic_drift", "editor_failed", "missing_output"}
    return 1 if any(observation.outcome in hard_outcomes for observation in observations) else 0


if __name__ == "__main__":
    raise SystemExit(main())
