"""Command-line interface for openxml-audit."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table

from openxml_audit.errors import FileFormat, ValidationSeverity
from openxml_audit.odf import OdfValidator
from openxml_audit.validator import OpenXmlValidator

console = Console()
error_console = Console(stderr=True)

OOXML_FORMATS = {
    FileFormat.OFFICE_2007,
    FileFormat.OFFICE_2010,
    FileFormat.OFFICE_2013,
    FileFormat.OFFICE_2016,
    FileFormat.OFFICE_2019,
    FileFormat.OFFICE_2021,
    FileFormat.MICROSOFT_365,
}
ODF_FORMATS = {
    FileFormat.ODF_1_2,
    FileFormat.ODF_1_3,
}

DEFAULT_FORMAT_BY_VALIDATOR = {
    "ooxml": FileFormat.OFFICE_2019,
    "odf": FileFormat.ODF_1_3,
}

OOXML_EXTENSIONS = {
    ".docx",
    ".docm",
    ".dotx",
    ".dotm",
    ".pptx",
    ".pptm",
    ".potx",
    ".potm",
    ".ppsx",
    ".ppsm",
    ".thmx",
    ".ppam",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".xlam",
}
ODF_EXTENSIONS = {
    ".odt",
    ".ods",
    ".odp",
    ".odg",
    ".odc",
    ".odi",
    ".odf",
    ".odb",
    ".odm",
    ".ott",
    ".ots",
    ".otp",
    ".otg",
    ".otm",
    ".oth",
}


def _detect_validator_for_path(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in ODF_EXTENSIONS:
        return "odf"
    if ext in OOXML_EXTENSIONS:
        return "ooxml"
    return ""


def _resolve_format(requested: FileFormat | None, validator_kind: str) -> FileFormat:
    allowed = OOXML_FORMATS if validator_kind == "ooxml" else ODF_FORMATS
    if requested is None:
        return DEFAULT_FORMAT_BY_VALIDATOR[validator_kind]
    if requested not in allowed:
        raise ValueError(
            f"Format '{requested.value}' is not supported for validator '{validator_kind}'."
        )
    return requested


def _collect_files(path: Path, recursive: bool, validator: str) -> list[Path]:
    if path.is_dir():
        if not recursive:
            raise ValueError(
                f"{path} is a directory. Use --recursive to validate all files."
            )
        if validator == "ooxml":
            extensions = OOXML_EXTENSIONS
        elif validator == "odf":
            extensions = ODF_EXTENSIONS
        else:
            extensions = OOXML_EXTENSIONS | ODF_EXTENSIONS
        return sorted({p for ext in extensions for p in path.rglob(f"*{ext}")})
    return [path]


@click.command()
@click.argument("path", type=click.Path(exists=True, path_type=Path))
@click.option(
    "--format",
    "-f",
    "file_format",
    type=click.Choice([f.value for f in FileFormat], case_sensitive=False),
    default=None,
    show_default=False,
    help="Office/ODF version to validate against.",
)
@click.option(
    "--validator",
    type=click.Choice(["ooxml", "odf", "auto"], case_sensitive=False),
    default="auto",
    help="Validator to use (auto selects by file extension).",
)
@click.option(
    "--policy",
    type=click.Choice(["strict", "permissive"], case_sensitive=False),
    default="strict",
    help="Validation policy (strict treats errors as fatal).",
)
@click.option(
    "--output",
    "-o",
    type=click.Choice(["text", "json", "xml"], case_sensitive=False),
    default="text",
    help="Output format.",
)
@click.option(
    "--max-errors",
    "-m",
    type=int,
    default=100,
    help="Maximum errors to report (0 for unlimited).",
)
@click.option(
    "--recursive",
    "-r",
    is_flag=True,
    help="Validate all matching files in directory.",
)
@click.option(
    "--quiet",
    "-q",
    is_flag=True,
    help="Only output errors, no success messages.",
)
def main(
    path: Path,
    file_format: str | None,
    validator: str,
    policy: str,
    output: str,
    max_errors: int,
    recursive: bool,
    quiet: bool,
) -> None:
    """Validate Open XML or ODF files against their specifications.

    PATH can be a single file or a directory (with --recursive).
    """
    requested_format = FileFormat(file_format) if file_format else None
    strict = policy == "strict"

    # Collect files to validate
    try:
        files = _collect_files(path, recursive, validator)
    except ValueError as exc:
        error_console.print(f"[red]Error:[/red] {exc}")
        sys.exit(1)
    if path.is_dir() and not files:
        error_console.print(f"[yellow]Warning:[/yellow] No matching files found in {path}")
        sys.exit(0)

    # Validate each file
    results = []
    all_valid = True
    validators: dict[tuple[str, FileFormat], object] = {}

    for file_path in files:
        file_validator = validator
        if file_validator == "auto":
            file_validator = _detect_validator_for_path(file_path)
            if not file_validator:
                error_console.print(
                    f"[red]Error:[/red] Cannot determine validator for {file_path}."
                )
                sys.exit(1)
        try:
            format_enum = _resolve_format(requested_format, file_validator)
        except ValueError as exc:
            error_console.print(f"[red]Error:[/red] {exc}")
            sys.exit(1)
        cache_key = (file_validator, format_enum)
        if cache_key not in validators:
            if file_validator == "ooxml":
                validators[cache_key] = OpenXmlValidator(
                    file_format=format_enum,
                    max_errors=max_errors,
                    strict=strict,
                )
            else:
                validators[cache_key] = OdfValidator(
                    file_format=format_enum,
                    strict=strict,
                )
        result = validators[cache_key].validate(file_path)
        results.append(result)
        if not result.is_valid:
            all_valid = False

    # Output results
    if output == "json":
        _output_json(results)
    elif output == "xml":
        _output_xml(results)
    else:
        _output_text(results, quiet)

    sys.exit(0 if all_valid else 1)


def _output_text(results: list, quiet: bool) -> None:
    """Output results as formatted text."""
    for result in results:
        if result.is_valid:
            if not quiet:
                console.print(f"[green]✓[/green] {result.file_path} - Valid")
        else:
            console.print(f"[red]✗[/red] {result.file_path} - Invalid")

            # Create error table
            table = Table(show_header=True, header_style="bold")
            table.add_column("Type", style="dim", width=12)
            table.add_column("Severity", width=8)
            table.add_column("Location", width=30)
            table.add_column("Description")

            for error in result.errors:
                severity_style = {
                    ValidationSeverity.ERROR: "red",
                    ValidationSeverity.WARNING: "yellow",
                    ValidationSeverity.INFO: "blue",
                }.get(error.severity, "white")

                location = error.part_uri
                if error.path:
                    location = f"{location}:{error.path}"

                table.add_row(
                    error.error_type.value,
                    f"[{severity_style}]{error.severity.value}[/{severity_style}]",
                    location,
                    error.description,
                )

            console.print(table)
            console.print()

    # Summary
    total = len(results)
    valid = sum(1 for r in results if r.is_valid)
    invalid = total - valid

    if total > 1:
        console.print(f"\n[bold]Summary:[/bold] {valid}/{total} files valid", end="")
        if invalid > 0:
            console.print(f", [red]{invalid} invalid[/red]")
        else:
            console.print()


def _output_json(results: list) -> None:
    """Output results as JSON."""
    output = []
    for result in results:
        output.append({
            "file": result.file_path,
            "valid": result.is_valid,
            "format": result.file_format.value,
            "error_count": result.error_count,
            "warning_count": result.warning_count,
            "errors": [
                {
                    "type": e.error_type.value,
                    "severity": e.severity.value,
                    "description": e.description,
                    "part_uri": e.part_uri,
                    "path": e.path,
                    "node": e.node,
                }
                for e in result.errors
            ],
        })

    console.print_json(json.dumps(output, indent=2))


def _output_xml(results: list) -> None:
    """Output results as XML (matching .NET SDK format)."""
    from xml.etree.ElementTree import Element, SubElement, tostring
    from xml.dom.minidom import parseString

    root = Element("ValidationResults")
    root.set("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    for result in results:
        file_elem = SubElement(root, "File")
        file_elem.set("path", str(result.file_path))
        file_elem.set("valid", str(result.is_valid).lower())
        file_elem.set("format", result.file_format.value)

        if result.errors:
            errors_elem = SubElement(file_elem, "Errors")
            errors_elem.set("count", str(result.error_count))

            for error in result.errors:
                error_elem = SubElement(errors_elem, "ValidationError")
                error_elem.set("type", error.error_type.value)
                error_elem.set("severity", error.severity.value)

                desc_elem = SubElement(error_elem, "Description")
                desc_elem.text = error.description

                if error.part_uri:
                    part_elem = SubElement(error_elem, "PartUri")
                    part_elem.text = error.part_uri

                if error.path:
                    path_elem = SubElement(error_elem, "Path")
                    path_elem.text = error.path

                if error.node:
                    node_elem = SubElement(error_elem, "Node")
                    node_elem.text = error.node

        if result.warning_count > 0:
            warnings_elem = SubElement(file_elem, "Warnings")
            warnings_elem.set("count", str(result.warning_count))

    # Pretty print XML
    xml_str = tostring(root, encoding="unicode")
    dom = parseString(xml_str)
    console.print(dom.toprettyxml(indent="  "))


if __name__ == "__main__":
    main()
