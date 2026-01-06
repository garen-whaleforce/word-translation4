#!/usr/bin/env python3
"""
Self-test script for CB PDF to Word Translation Service.
Verifies all components work correctly.
"""
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))


def test_imports():
    """Test all modules can be imported."""
    print("Testing imports...")

    try:
        from src.config import settings
        print("  ✓ config")
    except Exception as e:
        print(f"  ✗ config: {e}")
        return False

    try:
        from src.cb_parser import CBParser, ParseResult, ClauseItem, OverviewItem
        print("  ✓ cb_parser")
    except Exception as e:
        print(f"  ✗ cb_parser: {e}")
        return False

    try:
        from src.termbase import (
            Termbase, TermEntry, TermProtection,
            load_termbase_from_json, load_termbase_from_csv,
            create_combined_termbase
        )
        print("  ✓ termbase")
    except Exception as e:
        print(f"  ✗ termbase: {e}")
        return False

    try:
        from src.translation_service import TranslationService, TranslationResult
        print("  ✓ translation_service")
    except Exception as e:
        print(f"  ✗ translation_service: {e}")
        return False

    try:
        from src.word_filler import WordFiller, FillResult
        print("  ✓ word_filler")
    except Exception as e:
        print(f"  ✗ word_filler: {e}")
        return False

    try:
        from src.template_registry import TemplateRegistry, TemplateInfo
        print("  ✓ template_registry")
    except Exception as e:
        print(f"  ✗ template_registry: {e}")
        return False

    try:
        from src.validator import Validator, ValidationReport
        print("  ✓ validator")
    except Exception as e:
        print(f"  ✗ validator: {e}")
        return False

    try:
        from src.pipeline import Pipeline, PipelineResult
        print("  ✓ pipeline")
    except Exception as e:
        print(f"  ✗ pipeline: {e}")
        return False

    return True


def test_config():
    """Test configuration loading."""
    print("\nTesting configuration...")
    from src.config import settings

    print(f"  LiteLLM API Base: {settings.litellm_api_base}")
    print(f"  Bulk Model: {settings.bulk_model}")
    print(f"  Refine Model: {settings.refine_model}")
    print("  ✓ Configuration loaded")
    return True


def test_termbase():
    """Test termbase functionality."""
    print("\nTesting termbase...")
    from src.termbase import Termbase, TermEntry

    tb = Termbase()
    tb.add_entry(TermEntry(
        source_en="Reinforced Safeguard",
        target_zh="強化安全防護",
        priority=100
    ))
    tb.add_entry(TermEntry(
        source_en="Primary circuit",
        target_zh="一次側電路",
        priority=80
    ))

    # Test protection
    text = "The Reinforced Safeguard protects the Primary circuit"
    result = tb.protect_terms(text)
    print(f"  Original: {text}")
    print(f"  Protected: {result.protected_text}")
    print(f"  Mapping: {len(result.mapping)} terms")

    # Test restoration
    restored = tb.restore_terms(result.protected_text, result.mapping)
    print(f"  Restored: {restored}")

    assert "強化安全防護" in restored
    assert "一次側電路" in restored
    print("  ✓ Termbase protect/restore working")
    return True


def test_termbase_files():
    """Test loading real termbase files."""
    print("\nTesting termbase files...")
    from src.termbase import load_termbase_from_json, load_termbase_from_csv

    rules_dir = project_root / "rules"

    # Test glossary
    glossary_path = rules_dir / "en_zh_glossary_preferred.json"
    if glossary_path.exists():
        tb = load_termbase_from_json(glossary_path)
        print(f"  Glossary loaded: {len(tb)} terms")
    else:
        print("  ⚠ Glossary file not found")

    # Test translation memory
    tm_path = rules_dir / "en_zh_translation_memory.csv"
    if tm_path.exists():
        tb = load_termbase_from_csv(tm_path)
        print(f"  Translation memory loaded: {len(tb)} terms")
    else:
        print("  ⚠ Translation memory file not found")

    print("  ✓ Termbase files loaded")
    return True


def test_translation_dry_run():
    """Test translation service in dry run mode."""
    print("\nTesting translation service (dry run)...")
    from src.translation_service import TranslationService

    service = TranslationService(dry_run=True)
    result = service.translate("Primary circuit protection")

    print(f"  Input: Primary circuit protection")
    print(f"  Output: {result.translated_text}")
    print(f"  Model: {result.model_used}")
    print(f"  Was refined: {result.was_refined}")
    print("  ✓ Translation service dry run working")
    return True


def test_template_registry():
    """Test template registry."""
    print("\nTesting template registry...")
    from src.template_registry import TemplateRegistry

    templates_dir = project_root / "templates"
    registry = TemplateRegistry(templates_dir)

    templates = registry.list_templates()
    print(f"  Available templates: {len(templates)}")
    for t in templates:
        print(f"    - {t.name} ({t.id})")

    print("  ✓ Template registry working")
    return True


def test_word_filler():
    """Test Word filler with empty data."""
    print("\nTesting Word filler...")
    from src.word_filler import WordFiller
    from src.cb_parser import ParseResult, OverviewItem, ClauseItem
    import tempfile

    templates_dir = project_root / "templates"
    template_files = list(templates_dir.glob("*.docx"))

    if not template_files:
        print("  ⚠ No templates found, skipping")
        return True

    template_path = template_files[0]
    filler = WordFiller(template_path)

    # Create minimal test data
    parse_result = ParseResult(
        filename="test.pdf",
        overview_of_energy_sources=[
            OverviewItem(
                hazard_clause="5",
                description="Electrical shock",
                safeguards="Basic Insulation",
                remarks=""
            )
        ],
        energy_source_diagram_text="Test diagram",
        clauses=[
            ClauseItem(
                clause_id="4.1.1",
                requirement_test="General requirements",
                result_remark="Compliant",
                verdict="P",
                page_number=3
            )
        ]
    )

    # Fill template
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        output_path = Path(f.name)

    result = filler.fill(parse_result, output_path)
    print(f"  Template: {template_path.name}")
    print(f"  Output: {output_path}")
    print(f"  Overview rows filled: {result.overview_rows_filled}")
    print(f"  Clause rows filled: {result.clause_rows_filled}")

    # Cleanup
    if output_path.exists():
        output_path.unlink()

    print("  ✓ Word filler working")
    return True


def test_api_startup():
    """Test FastAPI app can start."""
    print("\nTesting API startup...")

    try:
        from src.main import app
        print(f"  App title: {app.title}")
        print(f"  Routes: {len(app.routes)}")

        # List key routes
        route_names = [r.path for r in app.routes if hasattr(r, 'path')]
        print(f"  Key endpoints: {', '.join(route_names[:10])}")

        print("  ✓ FastAPI app initialized")
        return True
    except Exception as e:
        print(f"  ✗ API startup failed: {e}")
        return False


def test_pipeline_dry_run():
    """Test pipeline in dry run mode."""
    print("\nTesting pipeline (dry run)...")
    from src.pipeline import Pipeline, PipelineConfig

    templates_dir = project_root / "templates"
    template_files = list(templates_dir.glob("*.docx"))

    if not template_files:
        print("  ⚠ No templates found, skipping")
        return True

    # Test pipeline initialization
    config = PipelineConfig(
        template_path=template_files[0],
        glossary_path=project_root / "rules" / "en_zh_glossary_preferred.json",
        dry_run=True
    )
    pipeline = Pipeline(config)

    print(f"  Pipeline initialized with dry_run=True")
    print(f"  Template: {template_files[0].name}")
    print("  ✓ Pipeline ready")
    return True


def main():
    """Run all self-tests."""
    print("=" * 60)
    print("CB PDF to Word Translation Service - Self Test")
    print("=" * 60)

    tests = [
        ("Module Imports", test_imports),
        ("Configuration", test_config),
        ("Termbase", test_termbase),
        ("Termbase Files", test_termbase_files),
        ("Translation Service", test_translation_dry_run),
        ("Template Registry", test_template_registry),
        ("Word Filler", test_word_filler),
        ("API Startup", test_api_startup),
        ("Pipeline", test_pipeline_dry_run),
    ]

    results = []
    for name, test_func in tests:
        try:
            success = test_func()
            results.append((name, success))
        except Exception as e:
            print(f"\n✗ {name} failed with exception: {e}")
            import traceback
            traceback.print_exc()
            results.append((name, False))

    # Summary
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)

    passed = sum(1 for _, s in results if s)
    total = len(results)

    for name, success in results:
        status = "✓" if success else "✗"
        print(f"  {status} {name}")

    print(f"\nResult: {passed}/{total} tests passed")

    if passed == total:
        print("\n✓ All self-tests passed!")
        return 0
    else:
        print("\n✗ Some tests failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
