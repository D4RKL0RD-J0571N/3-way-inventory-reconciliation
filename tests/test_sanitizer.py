"""
Comprehensive unit tests for data sanitization, edge cases, and business rule validators.
"""
import pytest
from src.sanitizer import (
    clean_string,
    clean_material_code,
    clean_number,
    clean_int,
    parse_multi_po
)
from src.models import PrepackItem, ChildSKU
from src.validator import validate_item_integrity, validate_all_items

def test_clean_string():
    assert clean_string("  Hello   World  ") == "Hello World"
    assert clean_string("BLUSA\xa0SM\tNegro\n") == "BLUSA SM Negro"
    assert clean_string(None) == ""
    assert clean_string("nan") == ""
    assert clean_string("#N/A") == ""
    assert clean_string("#VALUE!") == ""

def test_clean_material_code():
    # Scientific notation from Excel
    assert clean_material_code("1.001002001E+09") == "1001002001"
    # Float from Excel reading integer
    assert clean_material_code("1001003.0") == "1001003"
    # Leading zeroes
    assert clean_material_code("0000000001001003") == "1001003"
    # Clean whitespace and non-breaking space
    assert clean_material_code(" \xa01001003 \n") == "1001003"
    # None or NaN
    assert clean_material_code(None) == ""
    assert clean_material_code("NaN") == ""

def test_clean_number():
    # Standard integers and floats
    assert clean_number(100) == 100.0
    assert clean_number("1044") == 1044.0
    # US format with comma thousands
    assert clean_number("1,044.50") == 1044.5
    # European format with dot thousands and comma decimal
    assert clean_number("1.044,50") == 1044.5
    # Currency symbols
    assert clean_number("$ 1,250.75") == 1250.75
    assert clean_number("₡ 5,000.00") == 5000.0
    # Negative in parentheses (accountant format)
    assert clean_number("(148.00)") == -148.0
    # Negative with minus sign
    assert clean_number("-250.50") == -250.50
    # Error strings and empty values
    assert clean_number("#VALUE!") == 0.0
    assert clean_number("#N/A") == 0.0
    assert clean_number("-") == 0.0
    assert clean_number(None) == 0.0
    assert clean_number("N/A", default=99.0) == 99.0

def test_parse_multi_po():
    assert parse_multi_po("4300007063, 4300007064") == ["4300007063", "4300007064"]
    assert parse_multi_po("4300006764; 4300006804") == ["4300006764", "4300006804"]
    assert parse_multi_po(None) == []

def test_validator_curve_mismatch():
    # Curve says 12, but children sum to 10
    item = PrepackItem(
        material_code="1001003",
        vendor_material="STYLE-TEST-01",
        description="Falda",
        category="DAMA",
        order_po_43="4300007063",
        quantity_po_43_prk=10,
        quantity_po_43_units=120,
        quantity_inv_41_prk=10,
        quantity_inv_41_units=120,
        received_real_units=120,
        curve_total=12,
        unit_cost=10.0,
        children=[
            ChildSKU("1001002001", "S", 2, 20, 20),
            ChildSKU("1001002002", "M", 2, 20, 20),
            ChildSKU("1001002003", "L", 3, 30, 30),
            ChildSKU("1001002004", "XL", 3, 30, 30), # Sum = 2+2+3+3 = 10 != 12!
        ]
    )
    issues = validate_item_integrity(item)
    error_types = [i.issue_type for i in issues]
    assert "CURVA_DESCUADRADA" in error_types

def test_validator_uom_confusion():
    # Invoiced 100 units, but received 500 units (someone entered pieces instead of boxes)
    item = PrepackItem(
        material_code="1001003",
        vendor_material="STYLE-1",
        description="Falda",
        category="DAMA",
        order_po_43="4300007063",
        quantity_po_43_prk=10,
        quantity_po_43_units=100,
        quantity_inv_41_prk=10,
        quantity_inv_41_units=100,
        received_real_units=500, # 5x what was invoiced!
        curve_total=10,
        unit_cost=10.0
    )
    issues = validate_item_integrity(item)
    error_types = [i.issue_type for i in issues]
    assert "POSIBLE_CONFUSION_UDM" in error_types

def test_validator_missing_price_on_claim():
    # Has a shortage of 20 units, but price is $0.00
    item = PrepackItem(
        material_code="1001003",
        vendor_material="STYLE-1",
        description="Falda",
        category="DAMA",
        order_po_43="4300007063",
        quantity_po_43_prk=10,
        quantity_po_43_units=100,
        quantity_inv_41_prk=10,
        quantity_inv_41_units=100,
        received_real_units=80,
        curve_total=10,
        unit_cost=0.0, # Zero price!
        faltante=20.0
    )
    issues = validate_item_integrity(item)
    warning_types = [i.issue_type for i in issues]
    assert "PRECIO_CERO_EN_RECLAMO" in warning_types
