# Integration Fixtures

This folder defines the golden fixtures for integration checks. The
fixtures are referenced by environment variables so large files are not
committed to the repo.

## Required environment variables

- WT4_MC601_PDF
- WT4_MC601_HUMAN_DOCX
- WT4_DYS830_PDF
- WT4_DYS830_HUMAN_DOCX
- WT4_E135_1B_PDF
- WT4_E135_1B_HUMAN_DOCX

## Example

export WT4_MC601_PDF="/path/to/CB MC-601.pdf"
export WT4_MC601_HUMAN_DOCX="/path/to/AST-B-MC-601.docx"

python -m tests.integration.runner --fixture MC-601

pytest tests/integration/test_golden_fixtures.py
