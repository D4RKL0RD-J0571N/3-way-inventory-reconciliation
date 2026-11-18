"""
Unit tests for Service 3: Material Discrepancy Classifier (src/discrepancy_classifier.py).

Validates the Cardinal Discrepancy Rule (no offsetting across distinct materials),
individual status classifications, and automatic revuelto pair detection.
"""

import pytest
from src.discrepancy_classifier import (
    DeclaredItem,
    PhysicalItem,
    DiscrepancyType,
    compare_merchandise,
)


class TestDiscrepancyClassifier:
    """Test suite for material-level discrepancy comparison."""

    def test_exact_match(self):
        decl = [DeclaredItem(material_code="1001", declared_units=500, declared_cartons=10)]
        phys = [PhysicalItem(material_code="1001", physical_units=500, physical_cartons=10)]
        summary = compare_merchandise(decl, phys)

        assert summary.clean_materials_count == 1
        assert summary.discrepant_materials_count == 0
        assert summary.total_shortage_units == 0
        assert summary.total_overage_units == 0
        assert summary.items[0].discrepancy_type == DiscrepancyType.EXACT_MATCH
        assert summary.items[0].difference_units == 0

    def test_cardinal_rule_no_offsetting(self):
        """
        Cardinal Discrepancy Rule:
        Material 1001 has shortage of -20 units.
        Material 1002 has overage of +20 units.
        Net sum is 0, but total shortages must be 20 and total overages must be 20.
        They must NEVER offset into zero discrepancies.
        """
        decl = [
            DeclaredItem(material_code="1001", declared_units=100, declared_cartons=2),
            DeclaredItem(material_code="1002", declared_units=100, declared_cartons=2),
        ]
        phys = [
            PhysicalItem(material_code="1001", physical_units=80, physical_cartons=2),
            PhysicalItem(material_code="1002", physical_units=120, physical_cartons=2),
        ]
        summary = compare_merchandise(decl, phys)

        assert summary.net_unit_difference == 0  # Net sum is 0
        assert summary.total_shortage_units == 20  # Never offset!
        assert summary.total_overage_units == 20   # Never offset!
        assert summary.discrepant_materials_count == 2
        assert summary.revuelto_pairs_count == 1  # Detected as candidate revuelto pair!

        # Check that both materials are linked as revuelto partners
        item_1001 = next(it for it in summary.items if it.material_code == "1001")
        item_1002 = next(it for it in summary.items if it.material_code == "1002")
        assert item_1001.discrepancy_type == DiscrepancyType.REVUELTO_CANDIDATE
        assert item_1002.discrepancy_type == DiscrepancyType.REVUELTO_CANDIDATE
        assert item_1001.revuelto_partner == "1002"
        assert item_1002.revuelto_partner == "1001"

    def test_unexpected_item_received(self):
        """Merchandise received on dock that was never declared on the invoice."""
        decl = [DeclaredItem(material_code="1001", declared_units=100, declared_cartons=2)]
        phys = [
            PhysicalItem(material_code="1001", physical_units=100, physical_cartons=2),
            PhysicalItem(material_code="9999", physical_units=50, physical_cartons=1),
        ]
        summary = compare_merchandise(decl, phys)

        unexp = next(it for it in summary.items if it.material_code == "9999")
        assert unexp.discrepancy_type == DiscrepancyType.UNEXPECTED_ITEM
        assert unexp.difference_units == 50
        assert summary.total_overage_units == 50

    def test_missing_item_completely(self):
        """Merchandise declared on invoice that was completely absent from dock."""
        decl = [
            DeclaredItem(material_code="1001", declared_units=100, declared_cartons=2),
            DeclaredItem(material_code="1002", declared_units=200, declared_cartons=4),
        ]
        phys = [PhysicalItem(material_code="1001", physical_units=100, physical_cartons=2)]
        summary = compare_merchandise(decl, phys)

        missing = next(it for it in summary.items if it.material_code == "1002")
        assert missing.discrepancy_type == DiscrepancyType.MISSING_ITEM
        assert missing.difference_units == -200
        assert summary.total_shortage_units == 200

    def test_carton_mismatch_only(self):
        """Units match exactly, but carton count differs."""
        decl = [DeclaredItem(material_code="1001", declared_units=100, declared_cartons=5)]
        phys = [PhysicalItem(material_code="1001", physical_units=100, physical_cartons=4)]
        summary = compare_merchandise(decl, phys)

        item = summary.items[0]
        assert item.discrepancy_type == DiscrepancyType.CARTON_MISMATCH_ONLY
        assert item.difference_units == 0
        assert item.difference_cartons == -1

    def test_damage_only(self):
        """Physical units match declared units, but some units are damaged."""
        decl = [DeclaredItem(material_code="1001", declared_units=100, declared_cartons=2)]
        phys = [PhysicalItem(material_code="1001", physical_units=100, physical_cartons=2, damaged_units=5)]
        summary = compare_merchandise(decl, phys)

        item = summary.items[0]
        assert item.discrepancy_type == DiscrepancyType.DAMAGE_ONLY
        assert item.damaged_units == 5
        assert summary.total_damaged_units == 5
