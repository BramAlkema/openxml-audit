# Contributing to openxml-audit

Contributions are welcome! This guide covers the basics.

## Setup

```bash
git clone https://github.com/BramAlkema/openxml-audit.git
cd openxml-audit
pip install -e ".[dev]"
```

## Development

```bash
# Run tests
pytest

# Lint and format
ruff check src/ tests/
ruff format src/ tests/

# Type check
mypy src/openxml_audit
```

## Pull Requests

1. Fork the repo and create a branch from `main`
2. Add tests for new functionality
3. Make sure `pytest`, `ruff check`, and `mypy` pass
4. Keep PRs focused — one feature or fix per PR
5. Update documentation if behavior changes

## Validation Rules

- OOXML rules live in `src/openxml_audit/schema/` and `src/openxml_audit/semantic/`
- ODF rules live in `src/openxml_audit/odf/semantic.py` and `src/openxml_audit/odf/security.py`
- Each rule needs a stable ID (e.g. `ODFSEM001`) and a test fixture

## Reporting Issues

Use [GitHub Issues](https://github.com/BramAlkema/openxml-audit/issues). Include:
- The file that triggered the issue (or a minimal reproduction)
- Expected vs actual validation output
- Python version and OS
