"""PPTX-specific validation for PowerPoint presentations."""

from __future__ import annotations

from typing import Any

from openxml_audit.pptx.masters import (
    MasterValidator,
    validate_slide_layout,
    validate_slide_master,
)
from openxml_audit.pptx.presentation import (
    PresentationValidator,
    validate_presentation,
)
from openxml_audit.pptx.slides import (
    SlideValidator,
    validate_slide,
)
from openxml_audit.pptx.themes import (
    ThemeValidator,
    validate_theme,
)


def __getattr__(name: str) -> Any:
    if name in {
        "check_capability",
        "get_capability_finding",
        "list_capability_findings",
    }:
        from openxml_audit.pptx import capabilities as _capabilities

        return getattr(_capabilities, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    # Capability helpers (PPTX-specific registry)
    "check_capability",
    "get_capability_finding",
    "list_capability_findings",
    # Presentation
    "PresentationValidator",
    "validate_presentation",
    # Slides
    "SlideValidator",
    "validate_slide",
    # Themes
    "ThemeValidator",
    "validate_theme",
    # Masters
    "MasterValidator",
    "validate_slide_master",
    "validate_slide_layout",
]
