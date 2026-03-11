#!/usr/bin/env python3
"""Generate a spreadsheet with openpyxl and validate it.

Requirements:
    pip install openxml-audit openpyxl
"""

from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from openxml_audit import OpenXmlValidator, FileFormat


def generate_spreadsheet(path: Path) -> None:
    """Create a sample spreadsheet with openpyxl."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Data"

    # Headers
    ws.append(["Month", "Revenue", "Expenses"])

    # Data
    data = [
        ("Jan", 12000, 8000),
        ("Feb", 15000, 9000),
        ("Mar", 13000, 8500),
        ("Apr", 16000, 9200),
    ]
    for row in data:
        ws.append(row)

    # Add a chart
    chart = BarChart()
    chart.title = "Monthly Performance"
    chart.y_axis.title = "Amount ($)"
    chart_data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=5)
    categories = Reference(ws, min_col=1, min_row=2, max_row=5)
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "E2")

    wb.save(str(path))


def main() -> None:
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "sales.xlsx"

        print("Generating spreadsheet...")
        generate_spreadsheet(path)

        print("Validating...")
        validator = OpenXmlValidator(file_format=FileFormat.OFFICE_2019)
        result = validator.validate(path)

        if result.is_valid:
            print(f"Valid! ({result.warning_count} warnings)")
        else:
            print(f"Invalid: {result.error_count} errors, {result.warning_count} warnings")
            for error in result.errors:
                print(f"  [{error.severity.value}] {error.description}")
                if error.part_uri:
                    print(f"    Part: {error.part_uri}")


if __name__ == "__main__":
    main()
