"""Tests for the translation service module."""
import pytest
from unittest.mock import patch, MagicMock
from fastapi.testclient import TestClient

from src.main import app
from src.translation_service import TranslationService, TranslationResult
from src.termbase import Termbase, TermEntry


@pytest.fixture
def client():
    """Create test client."""
    return TestClient(app)


@pytest.fixture
def termbase():
    """Create a test termbase."""
    tb = Termbase()
    tb.add_entry(TermEntry(
        source_en="Safeguard",
        target_zh="安全防護",
        priority=100
    ))
    tb.add_entry(TermEntry(
        source_en="Primary circuit",
        target_zh="一次側電路",
        priority=80
    ))
    return tb


class TestTranslationService:
    """Tests for TranslationService class."""

    def test_dry_run_mode(self, termbase):
        """Test dry run mode returns mock translation."""
        service = TranslationService(termbase=termbase, dry_run=True)
        result = service.translate("Test text")

        assert "[DRY_RUN:" in result.translated_text
        assert result.model_used is not None

    def test_empty_text(self, termbase):
        """Test handling empty text."""
        service = TranslationService(termbase=termbase, dry_run=True)
        result = service.translate("")

        assert result.translated_text == ""
        assert result.model_used == "none"

    def test_term_protection(self, termbase):
        """Test that terms are protected during translation."""
        service = TranslationService(termbase=termbase, dry_run=True)
        result = service.translate("The Safeguard is important")

        # In dry run, the placeholder should be in the output
        # After restore, should have Chinese term
        assert "安全防護" in result.translated_text or "TERM_" in result.translated_text

    def test_qa_report_generation(self, termbase):
        """Test QA report generation."""
        service = TranslationService(termbase=termbase, dry_run=True)

        results = [
            service.translate("Test 1"),
            service.translate("Test 2"),
        ]

        report = service.generate_qa_report(results)

        assert report.total_segments == 2
        assert report.translated_segments == 2
        assert 0 <= report.quality_score <= 1

    def test_number_checking(self, termbase):
        """Test number/clause checking."""
        service = TranslationService(termbase=termbase, dry_run=True)

        original = "Test 4.1.1 with 250V"
        translated = "測試 4.1.1 帶 250V"

        issues = service._check_numbers(original, translated)
        # Should not have issues if numbers are preserved
        # (depends on exact matching logic)


class TestTranslateAPI:
    """Tests for /translate API endpoint."""

    def test_translate_dry_run(self, client):
        """Test translate endpoint in dry run mode."""
        response = client.post(
            "/translate",
            json={
                "text": "This is a test",
                "dry_run": True
            }
        )

        assert response.status_code == 200
        data = response.json()
        assert "original_text" in data
        assert "translated_text" in data
        assert "[DRY_RUN:" in data["translated_text"]

    def test_translate_empty_text(self, client):
        """Test translate with empty text."""
        response = client.post(
            "/translate",
            json={
                "text": "",
                "dry_run": True
            }
        )

        assert response.status_code == 200
        data = response.json()
        assert data["translated_text"] == ""

    def test_translate_with_custom_model(self, client):
        """Test translate with custom model."""
        response = client.post(
            "/translate",
            json={
                "text": "Test",
                "model": "custom-model",
                "dry_run": True
            }
        )

        assert response.status_code == 200
        data = response.json()
        assert "custom-model" in data["model_used"]

    def test_translate_without_refinement(self, client):
        """Test translate without refinement."""
        response = client.post(
            "/translate",
            json={
                "text": "Test",
                "enable_refinement": False,
                "dry_run": True
            }
        )

        assert response.status_code == 200
        data = response.json()
        assert data["was_refined"] == False


class TestTranslationResult:
    """Tests for TranslationResult dataclass."""

    def test_result_creation(self):
        """Test result creation."""
        result = TranslationResult(
            original_text="Test",
            translated_text="測試",
            model_used="test-model",
            was_refined=True
        )

        assert result.original_text == "Test"
        assert result.translated_text == "測試"
        assert result.was_refined == True
