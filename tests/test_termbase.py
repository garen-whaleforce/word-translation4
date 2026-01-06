"""Tests for the termbase module."""
import pytest
from pathlib import Path

from src.termbase import (
    Termbase,
    TermEntry,
    TermProtection,
    load_termbase_from_json,
    load_termbase_from_csv
)


class TestTermEntry:
    """Tests for TermEntry dataclass."""

    def test_term_entry_creation(self):
        """Test basic entry creation."""
        entry = TermEntry(
            source_en="Reinforced Safeguard",
            target_zh="強化安全防護",
            priority=10
        )
        assert entry.source_en == "Reinforced Safeguard"
        assert entry.target_zh == "強化安全防護"
        assert entry.priority == 10

    def test_term_entry_length(self):
        """Test length property."""
        entry = TermEntry(source_en="test", target_zh="測試")
        assert entry.length == 4


class TestTermbase:
    """Tests for Termbase class."""

    @pytest.fixture
    def termbase(self):
        """Create a termbase with test entries."""
        tb = Termbase()
        tb.add_entry(TermEntry(
            source_en="Reinforced Safeguard",
            target_zh="強化安全防護",
            priority=100
        ))
        tb.add_entry(TermEntry(
            source_en="Safeguard",
            target_zh="安全防護",
            priority=50
        ))
        tb.add_entry(TermEntry(
            source_en="Primary circuit",
            target_zh="一次側電路",
            priority=80
        ))
        tb.add_entry(TermEntry(
            source_en="N/A",
            target_zh="不適用",
            priority=10
        ))
        return tb

    def test_add_entry(self, termbase):
        """Test adding entries."""
        assert len(termbase) == 4
        assert "reinforced safeguard" in termbase
        assert "safeguard" in termbase

    def test_longest_match_priority(self, termbase):
        """Test that longer terms are matched first."""
        text = "Reinforced Safeguard is required"
        result = termbase.protect_terms(text)

        # Should match "Reinforced Safeguard", not just "Safeguard"
        assert "⟦TERM_" in result.protected_text
        # 只應有一個映射 (Reinforced Safeguard)
        assert len(result.mapping) == 1
        term = list(result.mapping.values())[0]
        assert term.source_en == "Reinforced Safeguard"

    def test_protect_and_restore(self, termbase):
        """Test full protect -> restore cycle."""
        original = "The Primary circuit requires a Safeguard"
        protected = termbase.protect_terms(original)

        # Check protection worked
        assert "⟦TERM_" in protected.protected_text
        assert len(protected.mapping) == 2

        # Simulate translation (keeping placeholders)
        translated = protected.protected_text  # In real case, this would be translated

        # Restore terms
        restored = termbase.restore_terms(translated, protected.mapping)

        # Check restoration
        assert "一次側電路" in restored
        assert "安全防護" in restored

    def test_case_insensitive_matching(self, termbase):
        """Test case insensitive matching."""
        text = "SAFEGUARD and safeguard and Safeguard"
        result = termbase.protect_terms(text)

        # All should be matched (case insensitive)
        assert result.protected_text.count("⟦TERM_") == 3

    def test_validate_unprotected_tokens(self, termbase):
        """Test validation catches unprotected tokens."""
        # Simulate a bad translation with leftover placeholder
        text = "這是⟦TERM_0001⟧的測試"
        violations = termbase.validate_terms(text)

        assert len(violations) >= 1
        assert violations[0].violation_type == "unprotected_token"

    def test_empty_text(self, termbase):
        """Test handling empty text."""
        result = termbase.protect_terms("")
        assert result.protected_text == ""
        assert result.mapping == {}

    def test_get_term(self, termbase):
        """Test getting specific term."""
        term = termbase.get_term("Safeguard")
        assert term is not None
        assert term.target_zh == "安全防護"

        term = termbase.get_term("nonexistent")
        assert term is None


class TestTermbaseLoading:
    """Tests for termbase loading functions."""

    def test_load_from_json(self, tmp_path):
        """Test loading from JSON file."""
        json_content = '''[
            {"en_norm": "Test", "zh_pref": "測試", "count": 10},
            {"en_norm": "Example", "zh_pref": "範例", "count": 5}
        ]'''

        json_file = tmp_path / "test_glossary.json"
        json_file.write_text(json_content, encoding='utf-8')

        termbase = load_termbase_from_json(json_file)

        assert len(termbase) == 2
        assert "test" in termbase
        assert "example" in termbase

    def test_load_from_csv(self, tmp_path):
        """Test loading from CSV file."""
        csv_content = '''source,context,en_raw,zh_raw,en_norm,zh_norm
E135-1B,4.1.1,Test,測試,Test,測試
E135-1B,4.1.2,Example,範例,Example,範例'''

        csv_file = tmp_path / "test_tm.csv"
        csv_file.write_text(csv_content, encoding='utf-8')

        termbase = load_termbase_from_csv(csv_file)

        assert len(termbase) == 2
        assert "test" in termbase


class TestRealTermbaseFiles:
    """Tests using real termbase files if available."""

    @pytest.fixture
    def rules_dir(self):
        """Get rules directory."""
        return Path(__file__).parent.parent / "rules"

    def test_load_real_glossary(self, rules_dir):
        """Test loading real glossary file."""
        glossary_path = rules_dir / "en_zh_glossary_preferred.json"
        if not glossary_path.exists():
            pytest.skip("Glossary file not found")

        termbase = load_termbase_from_json(glossary_path)
        assert len(termbase) > 0

        # Check some expected terms
        assert "general" in termbase or "General" in termbase.entries

    def test_reinforced_safeguard_protection(self, rules_dir):
        """Test that Reinforced Safeguard is properly protected."""
        glossary_path = rules_dir / "en_zh_glossary_preferred.json"
        if not glossary_path.exists():
            pytest.skip("Glossary file not found")

        termbase = load_termbase_from_json(glossary_path)

        # Add the critical term if not present
        if "reinforced safeguard" not in termbase:
            termbase.add_entry(TermEntry(
                source_en="Reinforced Safeguard",
                target_zh="強化安全防護",
                priority=1000
            ))

        text = "Reinforced Safeguard is required for this component"
        result = termbase.protect_terms(text)

        assert "⟦TERM_" in result.protected_text

        # Restore and verify
        restored = termbase.restore_terms(result.protected_text, result.mapping)
        assert "強化安全防護" in restored
