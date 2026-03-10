"""ODF constraint registry — reusable semantic validation constraints."""

from __future__ import annotations

from openxml_audit.odf.constraints.base import (
    ConstraintRegistry,
    EvaluationContext,
    OdfConstraint,
    OdfSemanticRule,
)
from openxml_audit.odf.constraints.chart import (
    ChartAxisConstraint,
    ChartPlotAreaConstraint,
    ChartSeriesDataRangeConstraint,
    ChartStyleRefConstraint,
)
from openxml_audit.odf.constraints.core import (
    BodyDocumentClassConstraint,
    CoreRootConstraint,
    MetaStructureConstraint,
    SettingsStructureConstraint,
)
from openxml_audit.odf.constraints.drawing import (
    Draw3dSceneConstraint,
    DrawConnectorResolveConstraint,
    DrawCustomGeometryConstraint,
    DrawFrameHrefConstraint,
    DrawGroupNestingConstraint,
    DrawShapePositionConstraint,
    DrawStyleRefConstraint,
)
from openxml_audit.odf.constraints.forms import (
    FormColumnRefConstraint,
    FormControlIdUniqueConstraint,
    FormControlNameUniqueConstraint,
    FormEventListenerConstraint,
)
from openxml_audit.odf.constraints.manifest import ManifestMediaTypeConstraint
from openxml_audit.odf.constraints.metadata import MetaStatisticsConstraint
from openxml_audit.odf.constraints.presentation import (
    AnimationTargetConstraint,
    CustomShowSlideRefConstraint,
    DrawLayerUniqueConstraint,
    HeaderFooterDeclRefConstraint,
    HeaderFooterDeclUniqueConstraint,
    NotesPageRefConstraint,
    PresentationClassConstraint,
    PresentationMinPagesConstraint,
    PresentationPageLayoutConstraint,
    PresentationPageNameConstraint,
    PresentationSettingsConstraint,
    SoundHrefConstraint,
    TransitionTypeConstraint,
)
from openxml_audit.odf.constraints.reference import (
    EmbeddedObjectRefConstraint,
    FontFaceCrossPartConstraint,
    ImageRefConstraint,
    MasterPageReferenceConstraint,
)
from openxml_audit.odf.constraints.spreadsheet import (
    CellStyleRefConstraint,
    CellValidationRefConstraint,
    CellValidationUniqueConstraint,
    ColumnStyleRefConstraint,
    ConditionalStyleRefConstraint,
    DatabaseRangeTableConstraint,
    DataPilotSourceConstraint,
    FilterFieldConstraint,
    RepeatCountConstraint,
    SpreadsheetColumnCountConstraint,
    SpreadsheetMinTableConstraint,
    SpreadsheetNamedRangeConstraint,
    SpreadsheetTableNameConstraint,
)
from openxml_audit.odf.constraints.style import (
    DataStyleRefConstraint,
    FontFaceDeclarationConstraint,
    ListStyleRefConstraint,
    MasterPageLayoutConstraint,
    StyleParentRefConstraint,
)
from openxml_audit.odf.constraints.style_chain import (
    DeepInheritanceConstraint,
    DefaultStyleFamilyConstraint,
    FontFamilyConsistencyConstraint,
    MasterPageHeaderFooterConstraint,
    MasterPageNextRefConstraint,
    NextStyleRefConstraint,
    OrphanedAutoStyleConstraint,
    StyleCycleConstraint,
    StyleDuplicateNameConstraint,
    StyleFamilyMismatchConstraint,
    StyleMapTargetConstraint,
    StyleNameEmptyConstraint,
)
from openxml_audit.odf.constraints.text import (
    HeadingLevelSkipConstraint,
    NoteRefConstraint,
    SectionNameUniqueConstraint,
    SequenceDeclUniqueConstraint,
    TextBookmarkRefConstraint,
    TextListLevelConstraint,
    TextStyleReferenceConstraint,
    TextTableStructureConstraint,
    TrackedChangeIdConstraint,
    UserFieldDeclUniqueConstraint,
    UserFieldGetRefConstraint,
    VariableDeclUniqueConstraint,
    VariableGetRefConstraint,
)
from openxml_audit.odf.constraints.version_specific import (
    ChangeTrackingVersionConstraint,
    DigitalSignatureVersionConstraint,
    DrawEnhancedGeometryVersionConstraint,
    NamedExpressionsVersionConstraint,
    PresentationAnimationVersionConstraint,
    RdfMetadataVersionConstraint,
    VersionAttributePresentConstraint,
    VersionConsistencyConstraint,
)

__all__ = [
    "ConstraintRegistry",
    "EvaluationContext",
    "OdfConstraint",
    "OdfSemanticRule",
    "build_default_registry",
    "get_odf_semantic_rules",
]

# All constraint classes in evaluation order
_DEFAULT_CONSTRAINTS: tuple[type[OdfConstraint], ...] = (
    # core
    CoreRootConstraint,
    BodyDocumentClassConstraint,
    MetaStructureConstraint,
    SettingsStructureConstraint,
    # manifest
    ManifestMediaTypeConstraint,
    # style
    FontFaceDeclarationConstraint,
    StyleParentRefConstraint,
    DataStyleRefConstraint,
    ListStyleRefConstraint,
    MasterPageLayoutConstraint,
    # text (existing)
    TextStyleReferenceConstraint,
    TextListLevelConstraint,
    TextBookmarkRefConstraint,
    # text (M2)
    HeadingLevelSkipConstraint,
    NoteRefConstraint,
    SectionNameUniqueConstraint,
    TrackedChangeIdConstraint,
    SequenceDeclUniqueConstraint,
    VariableDeclUniqueConstraint,
    VariableGetRefConstraint,
    UserFieldDeclUniqueConstraint,
    UserFieldGetRefConstraint,
    TextTableStructureConstraint,
    # spreadsheet (existing)
    SpreadsheetTableNameConstraint,
    SpreadsheetNamedRangeConstraint,
    SpreadsheetColumnCountConstraint,
    # spreadsheet (M2)
    SpreadsheetMinTableConstraint,
    DatabaseRangeTableConstraint,
    DataPilotSourceConstraint,
    CellValidationUniqueConstraint,
    CellValidationRefConstraint,
    RepeatCountConstraint,
    ColumnStyleRefConstraint,
    CellStyleRefConstraint,
    ConditionalStyleRefConstraint,
    FilterFieldConstraint,
    # presentation (existing)
    PresentationPageNameConstraint,
    PresentationMinPagesConstraint,
    PresentationPageLayoutConstraint,
    # presentation (M2)
    CustomShowSlideRefConstraint,
    DrawLayerUniqueConstraint,
    SoundHrefConstraint,
    HeaderFooterDeclUniqueConstraint,
    HeaderFooterDeclRefConstraint,
    PresentationSettingsConstraint,
    AnimationTargetConstraint,
    TransitionTypeConstraint,
    NotesPageRefConstraint,
    PresentationClassConstraint,
    # style chain (M3)
    StyleCycleConstraint,
    OrphanedAutoStyleConstraint,
    DefaultStyleFamilyConstraint,
    StyleMapTargetConstraint,
    MasterPageHeaderFooterConstraint,
    StyleFamilyMismatchConstraint,
    StyleNameEmptyConstraint,
    StyleDuplicateNameConstraint,
    DeepInheritanceConstraint,
    NextStyleRefConstraint,
    MasterPageNextRefConstraint,
    FontFamilyConsistencyConstraint,
    # version-specific (M4)
    VersionAttributePresentConstraint,
    VersionConsistencyConstraint,
    RdfMetadataVersionConstraint,
    NamedExpressionsVersionConstraint,
    DigitalSignatureVersionConstraint,
    ChangeTrackingVersionConstraint,
    DrawEnhancedGeometryVersionConstraint,
    PresentationAnimationVersionConstraint,
    # drawing (M6)
    DrawShapePositionConstraint,
    DrawGroupNestingConstraint,
    DrawConnectorResolveConstraint,
    DrawCustomGeometryConstraint,
    DrawFrameHrefConstraint,
    Draw3dSceneConstraint,
    DrawStyleRefConstraint,
    # forms (M6)
    FormControlNameUniqueConstraint,
    FormControlIdUniqueConstraint,
    FormColumnRefConstraint,
    FormEventListenerConstraint,
    # chart (M6)
    ChartPlotAreaConstraint,
    ChartAxisConstraint,
    ChartSeriesDataRangeConstraint,
    ChartStyleRefConstraint,
    # cross-part references
    MasterPageReferenceConstraint,
    FontFaceCrossPartConstraint,
    EmbeddedObjectRefConstraint,
    ImageRefConstraint,
    # metadata
    MetaStatisticsConstraint,
)


def build_default_registry() -> ConstraintRegistry:
    """Build the default constraint registry with all built-in constraints."""
    registry = ConstraintRegistry()
    for cls in _DEFAULT_CONSTRAINTS:
        registry.register(cls())
    return registry


def get_odf_semantic_rules() -> tuple[OdfSemanticRule, ...]:
    """Return semantic-core rule metadata with stable identifiers."""
    return build_default_registry().rules()
