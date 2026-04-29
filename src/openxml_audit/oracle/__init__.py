"""Roundtrip-oracle meta-CLI dispatcher.

`python -m openxml_audit.oracle <format> FILES...` routes to the
right format-specific oracle (Word, ODF, PowerPoint, Excel) without
the caller having to know the path to each `tools/oracle/*.py` script.
"""
