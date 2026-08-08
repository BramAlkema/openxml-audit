"""Run one DOCX through EuroOffice and Google Docs, then align the evidence."""

from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import asdict
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC_DIR = REPO_ROOT / "src"
TOOLS_DIR = REPO_ROOT / "tools"
for path in (SRC_DIR, TOOLS_DIR):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from openxml_audit.docx.differential_oracle import (  # noqa: E402
    build_docx_differential_report,
)
from openxml_audit.gsuite import DEFAULT_CREDS_PATH  # noqa: E402
from oracle.eurooffice_roundtrip import (  # noqa: E402
    DEFAULT_PLAYWRIGHT_IMAGE,
    SSHExampleEuroOfficeClient,
)
from oracle.eurooffice_roundtrip import (  # noqa: E402
    observe as observe_eurooffice,
)
from oracle.gsuite_roundtrip import observe as observe_google  # noqa: E402


def _resolve_google_config(
    *,
    folder_id: str | None,
    subject: str | None,
    creds: Path | None,
) -> tuple[str, str, Path]:
    resolved_folder = folder_id or os.environ.get("GSUITE_ORACLE_FOLDER_ID")
    resolved_subject = subject or os.environ.get("GSUITE_ORACLE_SUBJECT")
    resolved_creds = Path(
        creds or os.environ.get("GSUITE_ORACLE_CREDS") or DEFAULT_CREDS_PATH
    ).expanduser()
    missing = []
    if not resolved_folder:
        missing.append("GSUITE_ORACLE_FOLDER_ID/--google-folder-id")
    if not resolved_subject:
        missing.append("GSUITE_ORACLE_SUBJECT/--google-subject")
    if not resolved_creds.is_file():
        missing.append(f"Google service-account credentials at {resolved_creds}")
    if missing:
        raise ValueError("missing Google oracle configuration: " + ", ".join(missing))
    assert resolved_folder is not None
    assert resolved_subject is not None
    return resolved_folder, resolved_subject, resolved_creds


def run_paired(
    source: Path,
    *,
    eurooffice_client: SSHExampleEuroOfficeClient,
    google_folder_id: str,
    google_subject: str,
    google_creds: Path,
) -> dict[str, Any]:
    """Run both live targets and produce a matrix when both outputs exist."""

    eurooffice = observe_eurooffice(
        source,
        client=eurooffice_client,
        keep_artifacts=True,
    )
    google = observe_google(
        source,
        folder_id=google_folder_id,
        subject=google_subject,
        creds_path=google_creds,
        keep_artifacts=True,
    )
    differential = None
    if eurooffice.roundtripped_path and google.roundtripped_path:
        differential = build_docx_differential_report(
            source,
            {
                "eurooffice": Path(eurooffice.roundtripped_path),
                "google_docs": Path(google.roundtripped_path),
            },
        )
    return {
        "schema_version": 1,
        "engine": "paired-docx-roundtrip",
        "source": str(source.resolve()),
        "eurooffice": asdict(eurooffice),
        "google_docs": asdict(google),
        "differential": differential,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="source DOCX")
    parser.add_argument("--output", type=Path, help="write JSON report")
    parser.add_argument("--eurooffice-host", default="dtda-server")
    parser.add_argument("--eurooffice-container", default="eurooffice-documentserver")
    parser.add_argument("--browser-image", default=DEFAULT_PLAYWRIGHT_IMAGE)
    parser.add_argument("--google-folder-id")
    parser.add_argument("--google-subject")
    parser.add_argument("--google-creds", type=Path)
    args = parser.parse_args()

    if not args.source.is_file() or args.source.suffix.lower() != ".docx":
        parser.error("source must be an existing .docx file")
    try:
        folder_id, subject, creds = _resolve_google_config(
            folder_id=args.google_folder_id,
            subject=args.google_subject,
            creds=args.google_creds,
        )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    report = run_paired(
        args.source,
        eurooffice_client=SSHExampleEuroOfficeClient(
            host=args.eurooffice_host,
            container=args.eurooffice_container,
            browser_image=args.browser_image,
        ),
        google_folder_id=folder_id,
        google_subject=subject,
        google_creds=creds,
    )
    rendered = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(rendered, end="")

    eurooffice_ok = report["eurooffice"]["outcome"] == "preserved"
    google_ok = report["google_docs"]["outcome"] in {"preserved", "lossy_conversion"}
    differential = report["differential"]
    semantics_ok = bool(
        differential and differential["all_targets_semantically_preserved"]
    )
    return 0 if eurooffice_ok and google_ok and semantics_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
