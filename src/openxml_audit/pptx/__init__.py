"""PPTX-specific validation for PowerPoint presentations."""

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
from openxml_audit.pptx.masters import (
    MasterValidator,
    validate_slide_master,
    validate_slide_layout,
)

__all__ = [
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
