"""Per-app survival verdicts derived from validation findings (Spec 035).

Maps a `ValidationResult` to the question the mission actually asks:
"will this file open — and survive — in its target app?". The mapping
is rule-based over the finding taxonomy the validator already emits:

1. package-integrity findings → the package layer is broken; apps
   cannot be expected to open the file (`reject-likely`)
2. app-compat findings for the target app → rules that encode observed
   app behavior: repair dialogs and silent canonicalization rewrites
   (`repair-or-rewrite-likely`)
3. other schema/semantic errors → the app may open, repair, or reject
   (`at-risk`)
4. no errors → `opens-clean`

These are rule-derived predictions, not statistically calibrated
claims. Scoring them against oracle ground truth on a real corpus is
future work; until then no verdict carries a probability.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from openxml_audit.errors import (
    SourceClass,
    ValidationResult,
    ValidationSeverity,
)

__all__ = ["AppPrediction", "AppVerdict", "predict", "target_app"]

_BASIS_LIMIT = 3

_WORD = "Microsoft Word"
_EXCEL = "Microsoft Excel"
_POWERPOINT = "Microsoft PowerPoint"
_GENERIC_APP = "the target application"

_APP_BY_EXTENSION = {
    ".docx": _WORD,
    ".docm": _WORD,
    ".dotx": _WORD,
    ".dotm": _WORD,
    ".xlsx": _EXCEL,
    ".xlsm": _EXCEL,
    ".xltx": _EXCEL,
    ".xltm": _EXCEL,
    ".xlam": _EXCEL,
    ".pptx": _POWERPOINT,
    ".pptm": _POWERPOINT,
    ".potx": _POWERPOINT,
    ".potm": _POWERPOINT,
    ".ppsx": _POWERPOINT,
    ".ppsm": _POWERPOINT,
    ".ppam": _POWERPOINT,
    ".odt": "LibreOffice Writer",
    ".ott": "LibreOffice Writer",
    ".odm": "LibreOffice Writer",
    ".oth": "LibreOffice Writer",
    ".ods": "LibreOffice Calc",
    ".ots": "LibreOffice Calc",
    ".odp": "LibreOffice Impress",
    ".otp": "LibreOffice Impress",
    ".odg": "LibreOffice Draw",
    ".otg": "LibreOffice Draw",
}

_APP_COMPAT_CLASS_BY_APP = {
    _WORD: SourceClass.WORD_APP_COMPAT,
    _EXCEL: SourceClass.EXCEL_APP_COMPAT,
    _POWERPOINT: SourceClass.POWERPOINT_APP_COMPAT,
}


class AppPrediction(str, Enum):
    """Predicted target-app behavior, most to least survivable."""

    OPENS_CLEAN = "opens-clean"
    AT_RISK = "at-risk"
    REPAIR_OR_REWRITE_LIKELY = "repair-or-rewrite-likely"
    REJECT_LIKELY = "reject-likely"


@dataclass(frozen=True, slots=True)
class AppVerdict:
    """One per-app survival prediction with its evidence basis."""

    app: str
    prediction: AppPrediction
    headline: str
    evidence: str
    basis: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "app": self.app,
            "prediction": self.prediction.value,
            "headline": self.headline,
            "evidence": self.evidence,
            "basis": list(self.basis),
        }


def target_app(file_path: str | Path) -> str:
    """Resolve the app a file is destined for from its extension."""
    return _APP_BY_EXTENSION.get(Path(file_path).suffix.lower(), _GENERIC_APP)


def _basis(errors: list[Any]) -> tuple[str, ...]:
    descriptions = []
    for error in errors[:_BASIS_LIMIT]:
        descriptions.append(error.description)
    if len(errors) > _BASIS_LIMIT:
        descriptions.append(f"... and {len(errors) - _BASIS_LIMIT} more")
    return tuple(descriptions)


def predict(result: ValidationResult) -> AppVerdict:
    """Return the survival verdict for the file's target app."""
    app = target_app(result.file_path)
    errors = [error for error in result.errors if error.severity == ValidationSeverity.ERROR]
    warning_count = sum(
        1 for error in result.errors if error.severity == ValidationSeverity.WARNING
    )

    integrity = [error for error in errors if error.source_class is SourceClass.PACKAGE_INTEGRITY]
    if integrity:
        noun = "finding" if len(integrity) == 1 else "findings"
        return AppVerdict(
            app=app,
            prediction=AppPrediction.REJECT_LIKELY,
            headline=(
                f"{app}: unlikely to open this file ({len(integrity)} package-integrity {noun})"
            ),
            evidence="package-integrity findings (OPC/package layer)",
            basis=_basis(integrity),
        )

    # App-compat rules encode observed repair/rewrite behavior, so they
    # drive the verdict at any severity — a warning-severity "Excel will
    # move every inline string to shared strings on save" is still a
    # rewrite prediction (confirmed by the v0.7.2 oracle baseline).
    compat_class = _APP_COMPAT_CLASS_BY_APP.get(app)
    compat = [finding for finding in result.errors if finding.source_class is compat_class]
    if compat:
        noun = "finding" if len(compat) == 1 else "findings"
        return AppVerdict(
            app=app,
            prediction=AppPrediction.REPAIR_OR_REWRITE_LIKELY,
            headline=(
                f"{app}: expected to repair or rewrite this file "
                f"({len(compat)} {app} app-compat {noun})"
            ),
            evidence=f"{app} app-compat rules (encode observed repair/rewrite behavior)",
            basis=_basis(compat),
        )

    if errors:
        noun = "violation" if len(errors) == 1 else "violations"
        return AppVerdict(
            app=app,
            prediction=AppPrediction.AT_RISK,
            headline=(
                f"{app}: at risk — {len(errors)} schema/semantic {noun}; "
                "the app may open, repair, or reject this file"
            ),
            evidence="schema/semantic findings (SDK-proxy / native rules)",
            basis=_basis(errors),
        )

    if warning_count:
        noun = "warning" if warning_count == 1 else "warnings"
        headline = f"{app}: expected to open cleanly ({warning_count} {noun} only)"
    else:
        headline = f"{app}: expected to open cleanly (no findings)"
    return AppVerdict(
        app=app,
        prediction=AppPrediction.OPENS_CLEAN,
        headline=headline,
        evidence="no error-severity findings",
        basis=(),
    )
