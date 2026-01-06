"""Tests for the CB PDF parser module."""
import pytest
from pathlib import Path
from io import BytesIO
from fastapi.testclient import TestClient

from src.main import app
from src.cb_parser import CBParser, ParseResult, ClauseItem


@pytest.fixture
def client():
    """Create test client."""
    return TestClient(app)


class TestCBParser:
    """Tests for CBParser class."""

    def test_parser_file_not_found(self):
        """Test parser raises error for non-existent file."""
        with pytest.raises(FileNotFoundError):
            CBParser("/nonexistent/path.pdf")

    def test_parse_result_to_dict(self):
        """Test ParseResult serialization."""
        result = ParseResult(
            filename="test.pdf",
            trf_no="IEC62368_1E",
            clauses=[
                ClauseItem(
                    clause_id="4.1.1",
                    requirement_test="Test requirement",
                    result_remark="Test result",
                    verdict="P",
                    page_number=1
                )
            ]
        )

        data = result.to_dict()

        assert data["filename"] == "test.pdf"
        assert data["trf_no"] == "IEC62368_1E"
        assert data["total_clauses"] == 1
        assert data["clauses"][0]["clause_id"] == "4.1.1"
        assert data["clauses"][0]["verdict"] == "P"

    def test_verdict_normalization(self):
        """Test verdict value normalization."""
        parser = CBParser.__new__(CBParser)

        assert parser._normalize_verdict("P") == "P"
        assert parser._normalize_verdict("p") == "P"
        assert parser._normalize_verdict("N/A") == "N/A"
        assert parser._normalize_verdict("NA") == "N/A"
        assert parser._normalize_verdict("n/a") == "N/A"
        assert parser._normalize_verdict("Fail") == "Fail"
        assert parser._normalize_verdict("F") == "Fail"
        assert parser._normalize_verdict("--") == "--"

    def test_clause_id_pattern(self):
        """Test clause ID pattern matching."""
        pattern = CBParser.CLAUSE_ID_PATTERN

        # Valid clause IDs
        assert pattern.match("4")
        assert pattern.match("4.1")
        assert pattern.match("4.1.1")
        assert pattern.match("4.1.1.2.3")
        assert pattern.match("G.7.3.2.1")

        # Invalid clause IDs
        assert not pattern.match("abc")
        assert not pattern.match("4.1.a")
        assert not pattern.match("")


class TestParseAPI:
    """Tests for /parse API endpoint."""

    def test_parse_non_pdf_rejected(self, client):
        """Test that non-PDF files are rejected."""
        response = client.post(
            "/parse",
            files={"file": ("test.txt", BytesIO(b"test"), "text/plain")}
        )
        assert response.status_code == 400

    def test_parse_invalid_pdf(self, client):
        """Test handling of invalid PDF content."""
        response = client.post(
            "/parse",
            files={"file": ("test.pdf", BytesIO(b"not a pdf"), "application/pdf")}
        )
        # Should return error but not crash
        assert response.status_code in [200, 500]
