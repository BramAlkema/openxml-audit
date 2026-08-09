"""Single entrypoint for the roundtrip oracles.

Usage:

  python -m openxml_audit.oracle word     FILES... [--output X.json]
  python -m openxml_audit.oracle excel    FILES... [--output X.json]
  python -m openxml_audit.oracle pptx     FILES... [--output X.json]
  python -m openxml_audit.oracle odf      FILES... [--output X.json]
  python -m openxml_audit.oracle gsuite   FILES... [--output X.json]
  python -m openxml_audit.oracle eurooffice FILES... [--output X.json]

Each subcommand defers to the format's existing CLI in
`tools/oracle/`. This module is a thin dispatcher introduced in 0.6.8
so callers don't need to remember which file lives where (and so
shell history accumulates one verb across formats). The `gsuite`
engine added in spec 031 (alias `google`) needs the optional
`gsuite` extra installed and a service-account JSON key plus
domain-wide delegation configured — see
`specs/031-gsuite-roundtrip-oracle.md`.

The `eurooffice` engine (aliases `euro-office` and `euro`) exercises
Euro-Office Document Server's conversion endpoint. It reads JWT
credentials from environment variables only; see
`specs/038-eurooffice-conversion-oracle.md`.

Use `python -m openxml_audit.oracle preflight` to run the macOS
permission / install check across the desktop-app engines before a
corpus walk.
"""

from __future__ import annotations

import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[3]
_TOOLS_DIR = _REPO_ROOT / "tools"


def _ensure_tools_on_path() -> None:
    if str(_TOOLS_DIR) not in sys.path:
        sys.path.insert(0, str(_TOOLS_DIR))
    if str(_REPO_ROOT) not in sys.path:
        sys.path.insert(0, str(_REPO_ROOT))


def _run_word(args: list[str]) -> int:
    _ensure_tools_on_path()
    from tools.oracle.word_repair_corpus import main as word_main

    sys.argv = ["word_repair_corpus.py", *args]
    return word_main()


def _run_excel(args: list[str]) -> int:
    _ensure_tools_on_path()
    from tools.oracle.xlsx_repair_oracle import main as xlsx_main

    sys.argv = ["xlsx_repair_oracle.py", *args]
    return xlsx_main()


def _run_pptx(args: list[str]) -> int:
    _ensure_tools_on_path()
    from tools.oracle.pptx_repair_oracle import main as pptx_main

    sys.argv = ["pptx_repair_oracle.py", *args]
    return pptx_main()


def _run_odf(args: list[str]) -> int:
    _ensure_tools_on_path()
    from tools.oracle.odf_repair_oracle import main as odf_main

    sys.argv = ["odf_repair_oracle.py", *args]
    return odf_main()


def _run_gsuite(args: list[str]) -> int:
    _ensure_tools_on_path()
    from tools.oracle.gsuite_roundtrip import main as gsuite_main

    sys.argv = ["gsuite_roundtrip.py", *args]
    return gsuite_main()


def _run_eurooffice(args: list[str]) -> int:
    _ensure_tools_on_path()
    from tools.oracle.eurooffice_conversion_oracle import main as eurooffice_main

    sys.argv = ["eurooffice_conversion_oracle.py", *args]
    return eurooffice_main()


def _run_preflight(args: list[str]) -> int:
    _ensure_tools_on_path()
    from tools.oracle.preflight import main as preflight_main

    sys.argv = ["preflight.py", *args]
    return preflight_main()


_DISPATCH = {
    "word": _run_word,
    "excel": _run_excel,
    "xlsx": _run_excel,  # alias
    "pptx": _run_pptx,
    "powerpoint": _run_pptx,  # alias
    "odf": _run_odf,
    "gsuite": _run_gsuite,
    "google": _run_gsuite,  # alias
    "eurooffice": _run_eurooffice,
    "euro-office": _run_eurooffice,  # alias
    "euro": _run_eurooffice,  # alias
    "preflight": _run_preflight,
}


def main() -> int:
    if len(sys.argv) < 2 or sys.argv[1] in {"-h", "--help"}:
        print(__doc__, file=sys.stderr)
        print(
            "\nAvailable engines: " + ", ".join(sorted(_DISPATCH.keys())),
            file=sys.stderr,
        )
        return 0 if len(sys.argv) >= 2 else 2

    engine = sys.argv[1].lower()
    if engine not in _DISPATCH:
        print(f"unknown engine: {engine}", file=sys.stderr)
        print("Available: " + ", ".join(sorted(_DISPATCH.keys())), file=sys.stderr)
        return 2

    return _DISPATCH[engine](sys.argv[2:])


if __name__ == "__main__":
    sys.exit(main())
