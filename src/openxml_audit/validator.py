"""Main OOXML validator - entry point for validation."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING, Callable

from lxml import etree

from openxml_audit.context import ValidationContext
from openxml_audit.errors import (
    FileFormat,
    PackageValidationError,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.binary import parse_font_key, validate_binary_content
from openxml_audit.namespaces import (
    OFFICE_DOC_RELATIONSHIPS,
    REL_COMMENTS,
    REL_CUSTOM_XML,
    REL_CUSTOM_XML_PROPS,
    REL_ENDNOTES,
    REL_FONT_TABLE,
    REL_FOOTNOTES,
    REL_NUMBERING,
    REL_SETTINGS,
    REL_SLIDE,
    REL_SLIDE_MASTER,
    REL_STYLES,
    REL_STYLES_WITH_EFFECTS,
    REL_SHARED_STRINGS,
    REL_THEME,
    REL_WEB_SETTINGS,
    WORDPROCESSINGML,
)
from openxml_audit.package import OpenXmlPackage
from openxml_audit.parts import (
    DocumentPart,
    OpenXmlPart,
    PresentationPart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
    ThemePart,
    WorkbookPart,
)
from openxml_audit.pptx.masters import MasterValidator
from openxml_audit.pptx.presentation import PresentationValidator
from openxml_audit.pptx.slides import SlideValidator
from openxml_audit.pptx.themes import ThemeValidator
from openxml_audit.schema.validator import SchemaValidator
from openxml_audit.semantic.validator import (
    SemanticValidator,
    create_pptx_semantic_validator,
    create_spreadsheet_semantic_validator,
    create_word_semantic_validator,
)
from openxml_audit.relationships import get_rels_path
from openxml_audit.word.document import DocumentValidator
from openxml_audit.excel.workbook import WorkbookValidator

if TYPE_CHECKING:
    pass


class DocumentKind(Enum):
    """Supported Open XML document kinds."""

    PRESENTATION = "presentation"
    WORD = "word"
    SPREADSHEET = "spreadsheet"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class DocumentProfile:
    """Validation profile for a specific Open XML document kind."""

    kind: DocumentKind
    structure_validators: tuple[
        Callable[[OpenXmlPackage], list[ValidationError]], ...
    ] = ()
    semantic_validator: Callable[[OpenXmlPackage], list[ValidationError]] | None = None
    specific_validator: Callable[[OpenXmlPackage], list[ValidationError]] | None = None


class OpenXmlValidator:
    """Validator for Open XML documents.

    Validates PPTX files with PPTX-specific checks. For non-presentation
    OOXML files (DOCX/XLSX), runs package, schema, and core structural
    validation for the main document part.

    Example:
        validator = OpenXmlValidator()
        result = validator.validate("presentation.pptx")
        if result.is_valid:
            print("File is valid!")
        else:
            for error in result.errors:
                print(error)
    """

    def __init__(
        self,
        file_format: FileFormat = FileFormat.OFFICE_2019,
        max_errors: int = 1000,
        schema_validation: bool = True,
        semantic_validation: bool = True,
        strict: bool = True,
    ):
        """Initialize the validator.

        Args:
            file_format: The Office version to validate against.
            max_errors: Maximum number of errors to collect before stopping.
                       Set to 0 for unlimited.
            schema_validation: Enable schema validation of XML content.
            semantic_validation: Enable semantic validation (relationships, IDs, etc.).
        """
        self._file_format = file_format
        self._max_errors = max_errors
        self._schema_validation = schema_validation
        self._semantic_validation = semantic_validation
        self._schema_validator = SchemaValidator() if schema_validation else None
        self._semantic_validator_pptx = (
            create_pptx_semantic_validator() if semantic_validation else None
        )
        self._semantic_validator_word = (
            create_word_semantic_validator() if semantic_validation else None
        )
        self._semantic_validator_spreadsheet = (
            create_spreadsheet_semantic_validator() if semantic_validation else None
        )
        self._strict = strict

        # PPTX-specific validators
        self._presentation_validator = PresentationValidator()
        self._slide_validator = SlideValidator()
        self._theme_validator = ThemeValidator()
        self._master_validator = MasterValidator()
        self._document_validator = DocumentValidator()
        self._workbook_validator = WorkbookValidator()
        self._profiles: dict[DocumentKind, DocumentProfile] = {
            DocumentKind.PRESENTATION: DocumentProfile(
                kind=DocumentKind.PRESENTATION,
                structure_validators=(
                    self._validate_presentation_structure,
                    self._validate_slides,
                ),
                semantic_validator=self._validate_semantic,
                specific_validator=self._validate_pptx_specific,
            ),
            DocumentKind.WORD: DocumentProfile(
                kind=DocumentKind.WORD,
                structure_validators=(self._validate_word_structure,),
                semantic_validator=self._validate_semantic_word,
            ),
            DocumentKind.SPREADSHEET: DocumentProfile(
                kind=DocumentKind.SPREADSHEET,
                structure_validators=(self._validate_spreadsheet_structure,),
                semantic_validator=self._validate_semantic_spreadsheet,
            ),
            DocumentKind.UNKNOWN: DocumentProfile(kind=DocumentKind.UNKNOWN),
        }

    @property
    def file_format(self) -> FileFormat:
        """Get the target file format version."""
        return self._file_format

    @property
    def max_errors(self) -> int:
        """Get the maximum number of errors to collect."""
        return self._max_errors

    def validate(self, path: str | Path) -> ValidationResult:
        """Validate an Open XML file.

        Args:
            path: Path to the OOXML file.

        Returns:
            ValidationResult containing validation status and any errors.
        """
        path = Path(path)
        errors: list[ValidationError] = []

        try:
            with OpenXmlPackage(path) as package:
                # Phase 1: Package structure validation
                errors.extend(self._validate_package_structure(package))

                if self._should_stop(errors):
                    return self._create_result(path, errors)

                doc_kind = self._detect_document_type(package)
                profile = self._profiles.get(
                    doc_kind,
                    self._profiles[DocumentKind.UNKNOWN],
                )

                for validator in profile.structure_validators:
                    errors.extend(validator(package))

                    if self._should_stop(errors):
                        return self._create_result(path, errors)

                # Phase 4: Relationship integrity
                errors.extend(self._validate_relationships(package))

                if self._should_stop(errors):
                    return self._create_result(path, errors)

                # Phase 4.5: Binary payload validation
                errors.extend(self._validate_binary_parts(package))

                if self._should_stop(errors):
                    return self._create_result(path, errors)

                # Phase 5: Schema validation
                if self._schema_validator is not None:
                    errors.extend(self._validate_schema(package))

                if self._should_stop(errors):
                    return self._create_result(path, errors)

                # Phase 6: Semantic validation
                if profile.semantic_validator is not None:
                    errors.extend(profile.semantic_validator(package))

                if self._should_stop(errors):
                    return self._create_result(path, errors)

                # Phase 7: Document-specific validation
                if profile.specific_validator is not None:
                    errors.extend(profile.specific_validator(package))

        except PackageValidationError as e:
            errors.extend(e.errors)

        return self._create_result(path, errors)

    def is_valid(self, path: str | Path) -> bool:
        """Quick check if a file is valid.

        Args:
            path: Path to the PPTX file.

        Returns:
            True if the file is valid, False otherwise.
        """
        result = self.validate(path)
        return result.is_valid

    def _should_stop(self, errors: list[ValidationError]) -> bool:
        """Check if we should stop collecting errors."""
        if self._max_errors == 0:
            return False
        error_count = sum(1 for e in errors if e.severity == ValidationSeverity.ERROR)
        return error_count >= self._max_errors

    def _create_result(self, path: Path, errors: list[ValidationError]) -> ValidationResult:
        """Create a ValidationResult from collected errors."""
        is_valid = not any(e.severity == ValidationSeverity.ERROR for e in errors)
        return ValidationResult(
            is_valid=is_valid,
            errors=errors,
            file_path=str(path),
            file_format=self._file_format,
        )

    def _validate_package_structure(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate the OPC package structure."""
        return package.validate_structure()

    def _validate_presentation_structure(
        self, package: OpenXmlPackage
    ) -> list[ValidationError]:
        """Validate the presentation.xml structure."""
        errors: list[ValidationError] = []

        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri is None:
            return errors  # Already reported in package validation

        presentation = PresentationPart(package, main_doc_uri)

        if presentation.xml is None:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description="Cannot parse presentation.xml",
                    part_uri=main_doc_uri,
                    severity=ValidationSeverity.ERROR,
                )
            )
            return errors

        # Check for slide masters
        master_ids = presentation.slide_master_ids
        if not master_ids:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description="Presentation has no slide masters",
                    part_uri=main_doc_uri,
                    severity=ValidationSeverity.ERROR,
                )
            )

        # Validate each slide master exists
        for _id_val, rel_id in master_ids:
            target = presentation.relationships.resolve_target(rel_id)
            if target is None:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.RELATIONSHIP,
                        description=f"Slide master relationship {rel_id} not found",
                        part_uri=main_doc_uri,
                        severity=ValidationSeverity.ERROR,
                    )
                )
            elif not package.has_part(target):
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description=f"Slide master part not found: {target}",
                        part_uri=target,
                        severity=ValidationSeverity.ERROR,
                    )
                )

        return errors

    def _validate_slides(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate all slides in the presentation."""
        errors: list[ValidationError] = []

        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri is None:
            return errors

        presentation = PresentationPart(package, main_doc_uri)
        slide_ids = presentation.slide_ids

        for slide_id, rel_id in slide_ids:
            target = presentation.relationships.resolve_target(rel_id)
            if target is None:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.RELATIONSHIP,
                        description=f"Slide relationship {rel_id} not found",
                        part_uri=main_doc_uri,
                        severity=ValidationSeverity.ERROR,
                    )
                )
                continue

            if not package.has_part(target):
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description=f"Slide part not found: {target}",
                        part_uri=target,
                        severity=ValidationSeverity.ERROR,
                    )
                )
                continue

            # Validate slide XML
            slide = SlidePart(package, target)
            if slide.xml is None:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.SCHEMA,
                        description="Cannot parse slide XML",
                        part_uri=target,
                        severity=ValidationSeverity.ERROR,
                    )
                )
                continue

            # Check slide has a layout relationship
            layout_rels = list(slide.relationships.get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
            ))
            if not layout_rels:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.RELATIONSHIP,
                        description="Slide has no slideLayout relationship",
                        part_uri=target,
                        severity=ValidationSeverity.ERROR,
                    )
                )

        return errors

    def _validate_word_structure(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate core Word document structure."""
        errors: list[ValidationError] = []

        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri is None:
            return errors

        document = DocumentPart(package, main_doc_uri)
        if document.xml is None:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description="Cannot parse document.xml",
                    part_uri=main_doc_uri,
                    severity=ValidationSeverity.ERROR,
                )
            )
            return errors

        context = ValidationContext(
            package=package,
            file_format=self._file_format,
            max_errors=self._max_errors,
            strict=self._strict,
        )

        self._document_validator.validate(document, context)
        errors.extend(context.errors)
        return errors

    def _validate_spreadsheet_structure(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate core Excel workbook structure."""
        errors: list[ValidationError] = []

        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri is None:
            return errors

        workbook = WorkbookPart(package, main_doc_uri)
        if workbook.xml is None:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description="Cannot parse workbook.xml",
                    part_uri=main_doc_uri,
                    severity=ValidationSeverity.ERROR,
                )
            )
            return errors

        context = ValidationContext(
            package=package,
            file_format=self._file_format,
            max_errors=self._max_errors,
            strict=self._strict,
        )

        self._workbook_validator.validate(workbook, context)
        errors.extend(context.errors)
        return errors

    def _validate_relationships(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate that all internal relationships point to existing parts."""
        errors: list[ValidationError] = []

        # Check package relationships
        for rel in package.relationships:
            if not rel.is_external:
                target = rel.resolve_target("/")
                if not package.has_part(target):
                    errors.append(
                        ValidationError(
                            error_type=ValidationErrorType.RELATIONSHIP,
                            description=f"Relationship target not found: {target}",
                            part_uri="/_rels/.rels",
                            node=rel.id,
                            severity=ValidationSeverity.ERROR,
                        )
                    )

        # Check main document relationships
        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri:
            main_part = OpenXmlPart(package, main_doc_uri)
            for rel in main_part.relationships:
                if not rel.is_external:
                    target = rel.resolve_target(main_doc_uri)
                    if not package.has_part(target):
                        errors.append(
                            ValidationError(
                                error_type=ValidationErrorType.RELATIONSHIP,
                                description=f"Relationship target not found: {target}",
                                part_uri=main_doc_uri,
                                node=rel.id,
                                severity=ValidationSeverity.ERROR,
                            )
                        )

        return errors

    def _validate_binary_parts(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate binary payloads (images, embeddings) by magic bytes."""
        errors: list[ValidationError] = []
        font_keys = self._collect_font_keys(package)

        for part_uri in package.list_parts():
            content_type = package.content_types.get_content_type(part_uri)
            if content_type and "xml" in content_type:
                continue
            data = package.get_part_content(part_uri)
            if data is None:
                continue
            font_key = font_keys.get(part_uri)
            result = validate_binary_content(content_type, part_uri, data, font_key=font_key)
            if result:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.BINARY,
                        description=result.message,
                        part_uri=part_uri,
                        severity=result.severity,
                    )
                )

        return errors

    def _collect_font_keys(self, package: OpenXmlPackage) -> dict[str, bytes]:
        """Collect font keys for obfuscated font parts."""
        font_keys: dict[str, bytes] = {}
        font_keys.update(self._collect_word_font_keys(package))
        return font_keys

    def _collect_word_font_keys(self, package: OpenXmlPackage) -> dict[str, bytes]:
        """Collect font keys from Word fontTable.xml relationships."""
        font_keys: dict[str, bytes] = {}
        font_table_uri = "/word/fontTable.xml"
        if not package.has_part(font_table_uri):
            return font_keys

        part = OpenXmlPart(package, font_table_uri)
        xml = part.xml
        if xml is None:
            return font_keys

        ns = {"w": WORDPROCESSINGML, "r": OFFICE_DOC_RELATIONSHIPS}
        embed_tags = ("embedRegular", "embedBold", "embedItalic", "embedBoldItalic")
        for tag in embed_tags:
            for embed in xml.findall(f".//w:{tag}", ns):
                rel_id = embed.get(f"{{{OFFICE_DOC_RELATIONSHIPS}}}id", "")
                font_key = embed.get(f"{{{WORDPROCESSINGML}}}fontKey", "")
                if not rel_id or not font_key:
                    continue
                rel = part.relationships.get_by_id(rel_id)
                if rel is None or rel.is_external:
                    continue
                target = part.relationships.resolve_target(rel_id)
                if target is None:
                    continue
                key_bytes = parse_font_key(font_key)
                if key_bytes:
                    font_keys[target] = key_bytes

        return font_keys

    def _detect_document_type(self, package: OpenXmlPackage) -> DocumentKind:
        """Detect the OOXML document type by main part content type and path."""
        main_doc_uri = package.get_main_document_uri()
        if not main_doc_uri:
            return DocumentKind.UNKNOWN
        content_type = package.content_types.get_content_type(main_doc_uri) or ""
        uri_lower = main_doc_uri.lower()
        if "presentationml" in content_type or "/ppt/" in uri_lower:
            return DocumentKind.PRESENTATION
        if "wordprocessingml" in content_type or "/word/" in uri_lower:
            return DocumentKind.WORD
        if "spreadsheetml" in content_type or "/xl/" in uri_lower:
            return DocumentKind.SPREADSHEET
        return DocumentKind.UNKNOWN

    def _validate_schema(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate XML content against schema constraints."""
        errors: list[ValidationError] = []

        if self._schema_validator is None:
            return errors

        context = ValidationContext(
            package=package,
            file_format=self._file_format,
            max_errors=self._max_errors,
            strict=self._strict,
        )

        def is_xml_content_type(content_type: str | None) -> bool:
            return bool(content_type and "xml" in content_type)

        for part_uri in package.list_parts():
            content_type = package.content_types.get_content_type(part_uri)
            if not is_xml_content_type(content_type):
                continue
            part = OpenXmlPart(package, part_uri)
            self._schema_validator.validate_part(part, context)

            if context.should_stop:
                break

        errors.extend(context.errors)
        return errors

    def _validate_semantic(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate semantic constraints (relationships, IDs, etc.)."""
        errors: list[ValidationError] = []

        if self._semantic_validator_pptx is None:
            return errors

        context = ValidationContext(
            package=package,
            file_format=self._file_format,
            max_errors=self._max_errors,
            strict=self._strict,
        )

        # Validate presentation.xml
        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri:
            presentation = PresentationPart(package, main_doc_uri)
            self._semantic_validator_pptx.validate_part(presentation, context)

        # Validate slides
        if main_doc_uri:
            presentation = PresentationPart(package, main_doc_uri)
            for _slide_id, rel_id in presentation.slide_ids:
                target = presentation.relationships.resolve_target(rel_id)
                if target and package.has_part(target):
                    slide = SlidePart(package, target)
                    self._semantic_validator_pptx.validate_part(slide, context)

                    if context.should_stop:
                        break

        errors.extend(context.errors)
        return errors

    def _validate_semantic_word(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate semantic constraints for Word documents."""
        errors: list[ValidationError] = []

        if self._semantic_validator_word is None:
            return errors

        context = ValidationContext(
            package=package,
            file_format=self._file_format,
            max_errors=self._max_errors,
            strict=self._strict,
        )

        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri:
            main_doc = OpenXmlPart(package, main_doc_uri)
            self._validate_required_relationships(
                main_doc,
                self._word_required_relationships(),
                context,
            )
            self._validate_relationship_content_types(
                main_doc,
                self._word_relationship_content_types(),
                context,
            )
            self._validate_word_cross_part(package, context)

        errors.extend(self._validate_semantic_all_parts(package, self._semantic_validator_word, context))
        return errors

    def _validate_semantic_spreadsheet(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate semantic constraints for Excel workbooks."""
        errors: list[ValidationError] = []

        if self._semantic_validator_spreadsheet is None:
            return errors

        context = ValidationContext(
            package=package,
            file_format=self._file_format,
            max_errors=self._max_errors,
            strict=self._strict,
        )

        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri:
            main_doc = OpenXmlPart(package, main_doc_uri)
            self._validate_required_relationships(
                main_doc,
                self._spreadsheet_required_relationships(),
                context,
            )
            self._validate_relationship_content_types(
                main_doc,
                self._spreadsheet_relationship_content_types(),
                context,
            )
            self._validate_spreadsheet_cross_part(package, context)

        errors.extend(
            self._validate_semantic_all_parts(
                package,
                self._semantic_validator_spreadsheet,
                context,
            )
        )
        return errors

    def _validate_semantic_all_parts(
        self,
        package: OpenXmlPackage,
        validator: SemanticValidator,
        context: ValidationContext,
    ) -> list[ValidationError]:
        def is_xml_content_type(content_type: str | None) -> bool:
            return bool(content_type and "xml" in content_type)

        for part_uri in package.list_parts():
            content_type = package.content_types.get_content_type(part_uri)
            if not is_xml_content_type(content_type):
                continue
            part = OpenXmlPart(package, part_uri)
            validator.validate_part(part, context)

            if context.should_stop:
                break

        return context.errors

    def _validate_required_relationships(
        self,
        part: OpenXmlPart,
        required_types: tuple[str, ...],
        context: ValidationContext,
    ) -> None:
        if not required_types:
            return
        context.set_part(part)
        present = {rel.type for rel in part.relationships}
        for rel_type in required_types:
            if rel_type not in present:
                rel_name = rel_type.rsplit("/", 1)[-1]
                context.add_semantic_error(
                    f"Missing required relationship type '{rel_name}' ({rel_type})",
                    node="Relationship",
                )

    def _word_required_relationships(self) -> tuple[str, ...]:
        if not self._strict:
            return ()
        return (
            REL_STYLES,
            REL_STYLES_WITH_EFFECTS,
            REL_SETTINGS,
            REL_WEB_SETTINGS,
            REL_FONT_TABLE,
            REL_NUMBERING,
            REL_THEME,
        )

    def _spreadsheet_required_relationships(self) -> tuple[str, ...]:
        if not self._strict:
            return ()
        return (
            REL_STYLES,
            REL_THEME,
        )

    def _word_relationship_content_types(self) -> dict[str, tuple[str, ...]]:
        return {
            REL_STYLES: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            ),
            REL_STYLES_WITH_EFFECTS: ("application/vnd.ms-word.stylesWithEffects+xml",),
            REL_SETTINGS: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            ),
            REL_WEB_SETTINGS: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
            ),
            REL_FONT_TABLE: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
            ),
            REL_NUMBERING: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            ),
            REL_THEME: ("application/vnd.openxmlformats-officedocument.theme+xml",),
            REL_COMMENTS: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            ),
            REL_FOOTNOTES: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            ),
            REL_ENDNOTES: (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
            ),
            REL_CUSTOM_XML: ("application/xml",),
            REL_CUSTOM_XML_PROPS: (
                "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            ),
        }

    def _spreadsheet_relationship_content_types(self) -> dict[str, tuple[str, ...]]:
        return {
            REL_STYLES: (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            ),
            REL_THEME: ("application/vnd.openxmlformats-officedocument.theme+xml",),
            REL_SHARED_STRINGS: (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            ),
        }

    def _validate_relationship_content_types(
        self,
        part: OpenXmlPart,
        expected: dict[str, tuple[str, ...]],
        context: ValidationContext,
    ) -> None:
        if not expected or context.package is None:
            return
        context.set_part(part)
        package = context.package
        for rel in part.relationships:
            if rel.is_external:
                continue
            expected_types = expected.get(rel.type)
            if not expected_types:
                continue
            target = rel.resolve_target(part.uri)
            if not target:
                continue
            content_type = package.content_types.get_content_type(target)
            if content_type not in expected_types:
                expected_desc = ", ".join(expected_types)
                context.add_semantic_error(
                    f"Relationship target '{target}' has content type '{content_type}', "
                    f"expected {expected_desc}",
                    node="ContentType",
                )

    def _validate_word_cross_part(
        self,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> None:
        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri is None:
            return

        main_doc = OpenXmlPart(package, main_doc_uri)
        self._validate_word_main_content_type(main_doc, context)
        self._validate_custom_xml_parts(package, main_doc, context)

        styles_xml = package.get_part_xml("/word/styles.xml")
        style_ids = self._collect_word_style_ids(styles_xml)
        numbering_xml = package.get_part_xml("/word/numbering.xml")
        numbering_ids, abstract_ids = self._collect_word_numbering_ids(numbering_xml)
        comments_xml = package.get_part_xml("/word/comments.xml")
        comment_ids = self._collect_word_part_ids(comments_xml, "comment", "id")
        footnotes_xml = package.get_part_xml("/word/footnotes.xml")
        footnote_ids = self._collect_word_part_ids(footnotes_xml, "footnote", "id")
        endnotes_xml = package.get_part_xml("/word/endnotes.xml")
        endnote_ids = self._collect_word_part_ids(endnotes_xml, "endnote", "id")
        self._validate_word_settings_notes(package, footnote_ids, endnote_ids, context)

        needs_styles = False
        needs_numbering = False
        needs_comments = False
        needs_footnotes = False
        needs_endnotes = False
        missing_styles_reported = False
        missing_numbering_reported = False
        missing_comments_reported = False
        missing_footnotes_reported = False
        missing_endnotes_reported = False

        for part_uri, xml in self._iter_word_reference_parts(package):
            part = OpenXmlPart(package, part_uri)
            context.set_part(part)

            for style_ref in self._iter_word_style_refs(xml):
                needs_styles = True
                if styles_xml is None:
                    if not missing_styles_reported:
                        context.add_semantic_error(
                            "Document references styles but styles.xml is missing",
                            node="pStyle",
                        )
                        missing_styles_reported = True
                    break
                elif style_ref not in style_ids:
                    context.add_semantic_error(
                        f"Style '{style_ref}' referenced but not defined in styles.xml",
                        node="pStyle",
                    )

            for num_ref in self._iter_word_num_refs(xml):
                needs_numbering = True
                if numbering_xml is None:
                    if not missing_numbering_reported:
                        context.add_semantic_error(
                            "Document references numbering but numbering.xml is missing",
                            node="numId",
                        )
                        missing_numbering_reported = True
                    break
                elif num_ref not in numbering_ids:
                    context.add_semantic_error(
                        f"Numbering definition '{num_ref}' referenced but not found",
                        node="numId",
                    )

            for comment_ref in self._iter_word_id_refs(
                xml, ("commentReference", "commentRangeStart", "commentRangeEnd")
            ):
                needs_comments = True
                if comments_xml is None:
                    if not missing_comments_reported:
                        context.add_semantic_error(
                            "Document references comments but comments.xml is missing",
                            node="commentReference",
                        )
                        missing_comments_reported = True
                    break
                elif comment_ref not in comment_ids:
                    context.add_semantic_error(
                        f"Comment '{comment_ref}' referenced but not found",
                        node="commentReference",
                    )

            for footnote_ref in self._iter_word_id_refs(xml, ("footnoteReference",)):
                needs_footnotes = True
                if footnotes_xml is None:
                    if not missing_footnotes_reported:
                        context.add_semantic_error(
                            "Document references footnotes but footnotes.xml is missing",
                            node="footnoteReference",
                        )
                        missing_footnotes_reported = True
                    break
                elif footnote_ref not in footnote_ids:
                    context.add_semantic_error(
                        f"Footnote '{footnote_ref}' referenced but not found",
                        node="footnoteReference",
                    )

            for endnote_ref in self._iter_word_id_refs(xml, ("endnoteReference",)):
                needs_endnotes = True
                if endnotes_xml is None:
                    if not missing_endnotes_reported:
                        context.add_semantic_error(
                            "Document references endnotes but endnotes.xml is missing",
                            node="endnoteReference",
                        )
                        missing_endnotes_reported = True
                    break
                elif endnote_ref not in endnote_ids:
                    context.add_semantic_error(
                        f"Endnote '{endnote_ref}' referenced but not found",
                        node="endnoteReference",
                    )

        if styles_xml is not None:
            styles_part = OpenXmlPart(package, "/word/styles.xml")
            self._validate_word_style_links(styles_xml, style_ids, context, styles_part)
        if numbering_xml is not None:
            self._validate_word_numbering_links(numbering_xml, abstract_ids, context)

        dynamic_rels: list[str] = []
        if needs_comments:
            dynamic_rels.append(REL_COMMENTS)
        if needs_footnotes:
            dynamic_rels.append(REL_FOOTNOTES)
        if needs_endnotes:
            dynamic_rels.append(REL_ENDNOTES)
        if needs_styles:
            dynamic_rels.append(REL_STYLES)
        if needs_numbering:
            dynamic_rels.append(REL_NUMBERING)
        if dynamic_rels:
            self._validate_required_relationships(main_doc, tuple(dynamic_rels), context)

    def _validate_spreadsheet_cross_part(
        self,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> None:
        main_doc_uri = package.get_main_document_uri()
        if main_doc_uri is None:
            return

        main_doc = OpenXmlPart(package, main_doc_uri)
        self._validate_spreadsheet_main_content_type(main_doc, context)

        shared_strings_xml = package.get_part_xml("/xl/sharedStrings.xml")
        shared_strings_count = None
        if shared_strings_xml is not None:
            shared_strings_count = len(
                shared_strings_xml.findall(
                    ".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"
                )
            )

        needs_shared_strings = False
        for part_uri, xml in self._iter_spreadsheet_worksheet_parts(package):
            part = OpenXmlPart(package, part_uri)
            context.set_part(part)
            for ref in self._iter_shared_string_refs(xml):
                needs_shared_strings = True
                if shared_strings_xml is None:
                    context.add_semantic_error(
                        "Worksheet references shared strings but sharedStrings.xml is missing",
                        node="sharedStrings",
                    )
                    break
                if ref < 0:
                    context.add_semantic_error(
                        f"Shared string index {ref} is invalid",
                        node="sharedStrings",
                    )
                if shared_strings_count is not None and ref >= shared_strings_count:
                    context.add_semantic_error(
                        f"Shared string index {ref} out of range (count {shared_strings_count})",
                        node="sharedStrings",
                    )

        if needs_shared_strings:
            self._validate_required_relationships(
                main_doc,
                (REL_SHARED_STRINGS,),
                context,
            )

    def _validate_word_main_content_type(
        self,
        main_doc: OpenXmlPart,
        context: ValidationContext,
    ) -> None:
        if context.package is None:
            return
        ext = context.package.path.suffix.lower()
        expected = {
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            ".docm": "application/vnd.ms-word.document.macroEnabled.main+xml",
            ".dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
            ".dotm": "application/vnd.ms-word.template.macroEnabled.main+xml",
        }.get(ext)
        if not expected:
            return
        content_type = context.package.content_types.get_content_type(main_doc.uri)
        if content_type != expected:
            context.set_part(main_doc)
            context.add_semantic_error(
                f"Main document content type '{content_type}' does not match extension '{ext}'",
                node="ContentType",
            )

    def _validate_presentation_main_content_type(
        self,
        main_doc: OpenXmlPart,
        context: ValidationContext,
    ) -> None:
        if context.package is None:
            return
        ext = context.package.path.suffix.lower()
        expected = {
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            ".pptm": "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
            ".potx": "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
            ".potm": "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
            ".ppsx": "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
            ".ppsm": "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
            ".ppam": "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml",
        }.get(ext)
        if not expected:
            return
        content_type = context.package.content_types.get_content_type(main_doc.uri)
        if content_type != expected:
            context.set_part(main_doc)
            context.add_semantic_error(
                f"Main presentation content type '{content_type}' does not match extension '{ext}'",
                node="ContentType",
            )

    def _validate_spreadsheet_main_content_type(
        self,
        main_doc: OpenXmlPart,
        context: ValidationContext,
    ) -> None:
        if context.package is None:
            return
        ext = context.package.path.suffix.lower()
        expected = {
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            ".xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
        }.get(ext)
        if not expected:
            return
        content_type = context.package.content_types.get_content_type(main_doc.uri)
        if content_type != expected:
            context.set_part(main_doc)
            context.add_semantic_error(
                f"Main workbook content type '{content_type}' does not match extension '{ext}'",
                node="ContentType",
            )

    def _iter_word_reference_parts(
        self,
        package: OpenXmlPackage,
    ) -> list[tuple[str, etree._Element]]:
        parts: list[tuple[str, etree._Element]] = []
        skip = {
            "/word/styles.xml",
            "/word/numbering.xml",
            "/word/fontTable.xml",
            "/word/comments.xml",
            "/word/footnotes.xml",
            "/word/endnotes.xml",
        }
        for part_uri in package.list_parts():
            if part_uri in skip:
                continue
            if not part_uri.startswith("/word/"):
                continue
            content_type = package.content_types.get_content_type(part_uri) or ""
            if "xml" not in content_type:
                continue
            xml = package.get_part_xml(part_uri)
            if xml is None:
                continue
            parts.append((part_uri, xml))
        return parts

    def _iter_spreadsheet_worksheet_parts(
        self,
        package: OpenXmlPackage,
    ) -> list[tuple[str, etree._Element]]:
        parts: list[tuple[str, etree._Element]] = []
        for part_uri in package.list_parts():
            if not part_uri.startswith("/xl/worksheets/"):
                continue
            content_type = package.content_types.get_content_type(part_uri) or ""
            if "xml" not in content_type:
                continue
            xml = package.get_part_xml(part_uri)
            if xml is None:
                continue
            parts.append((part_uri, xml))
        return parts

    def _iter_word_style_refs(self, xml: etree._Element) -> list[str]:
        ns = {"w": WORDPROCESSINGML}
        refs: list[str] = []
        for tag in ("pStyle", "rStyle", "tblStyle"):
            for elem in xml.findall(f".//w:{tag}", ns):
                val = elem.get(f"{{{WORDPROCESSINGML}}}val")
                if val:
                    refs.append(val)
        return refs

    def _iter_word_num_refs(self, xml: etree._Element) -> list[str]:
        ns = {"w": WORDPROCESSINGML}
        refs: list[str] = []
        for elem in xml.findall(".//w:numId", ns):
            val = elem.get(f"{{{WORDPROCESSINGML}}}val")
            if val:
                refs.append(val)
        return refs

    def _iter_word_id_refs(self, xml: etree._Element, tags: tuple[str, ...]) -> list[str]:
        ns = {"w": WORDPROCESSINGML}
        refs: list[str] = []
        for tag in tags:
            for elem in xml.findall(f".//w:{tag}", ns):
                val = elem.get(f"{{{WORDPROCESSINGML}}}id")
                if val:
                    refs.append(val)
        return refs

    def _iter_shared_string_refs(self, xml: etree._Element) -> list[int]:
        refs: list[int] = []
        for cell in xml.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
            if cell.get("t") != "s":
                continue
            value = cell.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
            if value is None or value.text is None:
                continue
            try:
                refs.append(int(value.text))
            except ValueError:
                continue
        return refs

    def _collect_word_style_ids(self, styles_xml: etree._Element | None) -> set[str]:
        if styles_xml is None:
            return set()
        ns = {"w": WORDPROCESSINGML}
        style_ids = set()
        for style in styles_xml.findall(".//w:style", ns):
            style_id = style.get(f"{{{WORDPROCESSINGML}}}styleId")
            if style_id:
                style_ids.add(style_id)
        return style_ids

    def _collect_word_numbering_ids(
        self, numbering_xml: etree._Element | None
    ) -> tuple[set[str], set[str]]:
        if numbering_xml is None:
            return set(), set()
        ns = {"w": WORDPROCESSINGML}
        num_ids = set()
        abstract_ids = set()
        for num in numbering_xml.findall(".//w:num", ns):
            num_id = num.get(f"{{{WORDPROCESSINGML}}}numId")
            if num_id:
                num_ids.add(num_id)
        for abstract_num in numbering_xml.findall(".//w:abstractNum", ns):
            abstract_id = abstract_num.get(f"{{{WORDPROCESSINGML}}}abstractNumId")
            if abstract_id:
                abstract_ids.add(abstract_id)
        return num_ids, abstract_ids

    def _collect_word_part_ids(
        self,
        xml: etree._Element | None,
        element_name: str,
        id_attribute: str,
    ) -> set[str]:
        if xml is None:
            return set()
        ns = {"w": WORDPROCESSINGML}
        ids = set()
        for elem in xml.findall(f".//w:{element_name}", ns):
            elem_id = elem.get(f"{{{WORDPROCESSINGML}}}{id_attribute}")
            if elem_id:
                ids.add(elem_id)
        return ids

    def _validate_word_style_links(
        self,
        styles_xml: etree._Element,
        style_ids: set[str],
        context: ValidationContext,
        styles_part: OpenXmlPart,
    ) -> None:
        ns = {"w": WORDPROCESSINGML}
        context.set_part(styles_part)
        for style in styles_xml.findall(".//w:style", ns):
            style_id = style.get(f"{{{WORDPROCESSINGML}}}styleId") or ""
            for tag in ("basedOn", "next", "link"):
                ref = style.find(f"w:{tag}", ns)
                if ref is None:
                    continue
                val = ref.get(f"{{{WORDPROCESSINGML}}}val")
                if val and val not in style_ids:
                    context.add_semantic_error(
                        f"Style '{style_id}' references missing style '{val}' via {tag}",
                        node=tag,
                    )

    def _validate_word_numbering_links(
        self,
        numbering_xml: etree._Element,
        abstract_ids: set[str],
        context: ValidationContext,
    ) -> None:
        ns = {"w": WORDPROCESSINGML}
        if context.package is None:
            return
        context.set_part(OpenXmlPart(context.package, "/word/numbering.xml"))
        for num in numbering_xml.findall(".//w:num", ns):
            num_id = num.get(f"{{{WORDPROCESSINGML}}}numId") or ""
            abstract_ref = num.find("w:abstractNumId", ns)
            if abstract_ref is None:
                context.add_semantic_error(
                    f"Numbering definition '{num_id}' missing abstractNumId",
                    node="abstractNumId",
                )
                continue
            val = abstract_ref.get(f"{{{WORDPROCESSINGML}}}val")
            if val and val not in abstract_ids:
                context.add_semantic_error(
                    f"Numbering definition '{num_id}' references missing abstractNumId '{val}'",
                    node="abstractNumId",
                )

    def _validate_word_settings_notes(
        self,
        package: OpenXmlPackage,
        footnote_ids: set[str],
        endnote_ids: set[str],
        context: ValidationContext,
    ) -> None:
        settings_xml = package.get_part_xml("/word/settings.xml")
        if settings_xml is None:
            return
        settings_part = OpenXmlPart(package, "/word/settings.xml")
        context.set_part(settings_part)
        ns = {"w": WORDPROCESSINGML}
        for elem in settings_xml.findall(".//w:footnotePr/w:footnote", ns):
            note_id = elem.get(f"{{{WORDPROCESSINGML}}}id")
            if note_id and note_id not in footnote_ids:
                context.add_semantic_error(
                    f"Footnote '{note_id}' referenced in settings.xml was not found in footnotes.xml",
                    node="footnote",
                )
        for elem in settings_xml.findall(".//w:endnotePr/w:endnote", ns):
            note_id = elem.get(f"{{{WORDPROCESSINGML}}}id")
            if note_id and note_id not in endnote_ids:
                context.add_semantic_error(
                    f"Endnote '{note_id}' referenced in settings.xml was not found in endnotes.xml",
                    node="endnote",
                )

    def _validate_custom_xml_parts(
        self,
        package: OpenXmlPackage,
        main_doc: OpenXmlPart,
        context: ValidationContext,
    ) -> None:
        items = [
            part_uri
            for part_uri in package.list_parts()
            if part_uri.startswith("/customXml/")
            and part_uri.endswith(".xml")
            and "/itemProps" not in part_uri
            and "/_rels/" not in part_uri
        ]
        if not items:
            return

        context.set_part(main_doc)
        has_rel = any(rel.type == REL_CUSTOM_XML for rel in main_doc.relationships)
        if not has_rel:
            context.add_semantic_error(
                f"Missing required relationship type 'customXml' ({REL_CUSTOM_XML})",
                node="Relationship",
            )

        for item_uri in items:
            rels_path = get_rels_path(item_uri)
            if not package.has_part(rels_path):
                context.set_part(OpenXmlPart(package, item_uri))
                context.add_semantic_error(
                    "customXml item missing relationships to customXmlProps",
                    node="Relationship",
                )
                continue
            item_part = OpenXmlPart(package, item_uri)
            rels = item_part.relationships
            rel = rels.get_first_by_type(REL_CUSTOM_XML_PROPS)
            if rel is None:
                context.set_part(item_part)
                context.add_semantic_error(
                    "customXml item missing customXmlProps relationship",
                    node="Relationship",
                )
                continue
            target = rel.resolve_target(item_uri)
            if not target or not package.has_part(target):
                context.set_part(item_part)
                context.add_semantic_error(
                    "customXmlProps target not found",
                    node="Relationship",
                )
                continue
            content_type = package.content_types.get_content_type(target)
            expected = self._word_relationship_content_types().get(REL_CUSTOM_XML_PROPS, ())
            if content_type not in expected:
                context.set_part(item_part)
                context.add_semantic_error(
                    f"customXmlProps target '{target}' has content type '{content_type}'",
                    node="ContentType",
                )

    def _validate_pptx_specific(self, package: OpenXmlPackage) -> list[ValidationError]:
        """Validate PPTX-specific structure (presentation, slides, themes, masters)."""
        errors: list[ValidationError] = []

        context = ValidationContext(
            package=package,
            file_format=self._file_format,
            max_errors=self._max_errors,
            strict=self._strict,
        )

        main_doc_uri = package.get_main_document_uri()
        if not main_doc_uri:
            return errors

        self._validate_presentation_main_content_type(
            OpenXmlPart(package, main_doc_uri), context
        )

        # Validate presentation.xml
        presentation = PresentationPart(package, main_doc_uri)
        self._presentation_validator.validate(presentation, context)

        if context.should_stop:
            errors.extend(context.errors)
            return errors

        # Validate slide masters
        for _id_val, rel_id in presentation.slide_master_ids:
            target = presentation.relationships.resolve_target(rel_id)
            if target and package.has_part(target):
                master = SlideMasterPart(package, target)
                self._master_validator.validate_master(master, context)

                if context.should_stop:
                    break

                # Validate theme for this master
                theme_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                for theme_rel in master.relationships.get_by_type(theme_rel_type):
                    theme_target = theme_rel.resolve_target(target)
                    if theme_target and package.has_part(theme_target):
                        theme = ThemePart(package, theme_target)
                        self._theme_validator.validate(theme, context)

                # Validate layouts for this master
                layout_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
                for layout_rel in master.relationships.get_by_type(layout_rel_type):
                    layout_target = layout_rel.resolve_target(target)
                    if layout_target and package.has_part(layout_target):
                        layout = SlideLayoutPart(package, layout_target)
                        self._master_validator.validate_layout(layout, context)

                        if context.should_stop:
                            break

        if context.should_stop:
            errors.extend(context.errors)
            return errors

        # Validate slides
        for _slide_id, rel_id in presentation.slide_ids:
            target = presentation.relationships.resolve_target(rel_id)
            if target and package.has_part(target):
                slide = SlidePart(package, target)
                self._slide_validator.validate(slide, context)

                if context.should_stop:
                    break

        errors.extend(context.errors)
        return errors


def validate_pptx(path: str | Path) -> ValidationResult:
    """Convenience function to validate a PPTX file.

    Args:
        path: Path to the PPTX file.

    Returns:
        ValidationResult containing validation status and any errors.
    """
    validator = OpenXmlValidator()
    return validator.validate(path)


def is_valid_pptx(path: str | Path) -> bool:
    """Convenience function to check if a PPTX file is valid.

    Args:
        path: Path to the PPTX file.

    Returns:
        True if the file is valid, False otherwise.
    """
    validator = OpenXmlValidator()
    return validator.is_valid(path)
