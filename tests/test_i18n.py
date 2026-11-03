"""
Unit tests for the Internationalization (i18n) module.
Verifies language switching, dictionary retrieval, parameter interpolation,
and table toolbar translation snippet generation.
"""
import pytest
from src.i18n import (
    t,
    get_current_language,
    set_language,
    DEFAULT_LANGUAGE,
    LANGUAGES,
    get_table_toolbar_js
)

def test_default_language():
    """Default language should be Spanish ('es')."""
    assert DEFAULT_LANGUAGE == "es"
    assert "es" in LANGUAGES
    assert "en" in LANGUAGES

def test_translation_fallback():
    """Key translation should fallback gracefully."""
    # Known key
    es_val = t("app_title")
    assert "ReconcileAI" in es_val
    
    # Unknown key should return the key itself
    unknown = t("non_existent_key_12345")
    assert unknown == "non_existent_key_12345"
    
    # Default override
    custom_default = t("unknown_key", default="Custom Fallback")
    assert custom_default == "Custom Fallback"

def test_translation_interpolation():
    """Tests variable interpolation in localized strings."""
    res = t("files_found_in", folder="data/IMP-182", count=4)
    assert "data/IMP-182" in res
    assert "4" in res

def test_toolbar_js_generation():
    """Verifies that the generated JS script contains correct translated tooltips."""
    js_es = get_table_toolbar_js("es")
    assert "Pantalla completa" in js_es
    assert "Descargar como CSV" in js_es
    assert "Buscar" in js_es

    js_en = get_table_toolbar_js("en")
    assert "Fullscreen" in js_en
    assert "Download as CSV" in js_en
    assert "Search" in js_en
