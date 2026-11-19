"""
Unit tests for Service 1: Physical Quantity & Packaging Engine (src/quantity_engine.py).

Validates all 12 operational edge cases from DC Receiving Closure operations,
specifically preventing Excel 0-multiplier bugs, loss of loose units, and shared box double-counts.
"""

import pytest
from src.quantity_engine import (
    CartonCalculationInput,
    CartonCalculationResult,
    parse_packaging_expression,
    calculate_carton_quantities,
    FLAG_QC_COMPLEMENT_WITHOUT_CARTON_COUNT,
    FLAG_QC_FORMULA_MISMATCH,
    FLAG_SHARED_CARTON_PRIMARY,
    FLAG_SHARED_CARTON_SECONDARY,
    FLAG_VARIANTE_SUELTA_ROW,
    FLAG_UNEVEN_PREPACK_CURVE,
    FLAG_NEGATIVE_REPORTED_DIFF,
    FLAG_QC_OBSERVATION_ADJUSTMENT,
    FLAG_EMPTY_COMPLEMENT_CARTONS
)


class TestPackagingExpressionParser:
    """Tests for string expressions in packing lists and QC documents."""

    def test_single_term_expression(self):
        res = parse_packaging_expression("50*12")
        assert res.total_cartons == 50
        assert res.total_units == 600
        assert res.full_cartons == 50
        assert res.units_per_full == 12
        assert not res.is_compound
        assert not res.has_loose

    def test_scalar_expression(self):
        res = parse_packaging_expression("192")
        assert res.total_cartons == 1
        assert res.total_units == 192
        assert res.units_per_full == 192
        assert not res.is_compound

    def test_x_delimiter(self):
        res = parse_packaging_expression("10x24")
        assert res.total_cartons == 10
        assert res.total_units == 240

    def test_compound_two_terms(self):
        res = parse_packaging_expression("3*192 3*204")
        assert res.total_cartons == 6
        assert res.total_units == 1188
        assert res.full_cartons == 3
        assert res.units_per_full == 192
        assert res.complement_cartons == 3
        assert res.units_per_complement == 204
        assert res.is_compound
        assert not res.has_loose

    def test_compound_with_loose_addition_row_37(self):
        """The critical '6*192 +48' pattern from Row 37 of IMP-103."""
        res = parse_packaging_expression("6*192 +48")
        assert res.total_cartons == 6
        assert res.total_units == 1200
        assert res.full_cartons == 6
        assert res.units_per_full == 192
        assert res.loose_units == 48
        assert res.is_compound
        assert res.has_loose

    def test_compound_with_plus_between_terms(self):
        res = parse_packaging_expression("40*60 + 1*72")
        assert res.total_cartons == 41
        assert res.total_units == 2472
        assert res.full_cartons == 40
        assert res.units_per_full == 60
        assert res.complement_cartons == 1
        assert res.units_per_complement == 72

    def test_empty_and_null_expressions(self):
        assert parse_packaging_expression(None).total_units == 0
        assert parse_packaging_expression("").total_units == 0
        assert parse_packaging_expression("None").total_units == 0
        assert parse_packaging_expression("nan").total_units == 0


class TestCartonCalculationEngine:
    """Tests for physical quantity calculations from spreadsheet columns."""

    def test_standard_full_cartons(self):
        inp = CartonCalculationInput(
            full_cartons=10,
            units_per_full_carton=24,
            curve=12,
            uom="PRK"
        )
        res = calculate_carton_quantities(inp)
        assert res.calculated_units == 240
        assert res.calculated_cartons == 10
        assert res.calculated_packs == 20
        assert res.loose_units == 0
        assert len(res.flags) == 0

    def test_compound_cartons_row_1(self):
        """Row 1: 3 cartons of 192 + 3 cartons of 204 = 1,188 units, 6 cartons."""
        inp = CartonCalculationInput(
            full_cartons=3,
            units_per_full_carton=192,
            complement_cartons=3,
            units_in_complement=204,
            reported_total_cartons=6,
            reported_total_units=1188,
            curve=12,
            uom="PRK"
        )
        res = calculate_carton_quantities(inp)
        assert res.calculated_units == 1188
        assert res.calculated_cartons == 6
        assert res.calculated_packs == 99
        assert res.loose_units == 0
        assert res.formula_discrepancy == 0
        assert len(res.flags) == 0

    def test_row_37_excel_zero_multiplier_bug_rescuing(self):
        """
        Row 37 bug: Full cartons=6, units=192, complement cartons=None, complement units=48.
        In Excel: 6*192 + None*48 = 1,152 (dropped 48 units and flashed -48).
        In ReconcileAI: Correctly retains 1,200 units, 6 cartons, 100 PRK packs.
        """
        inp = CartonCalculationInput(
            full_cartons=6,
            units_per_full_carton=192,
            complement_cartons=None,
            units_in_complement=48,
            reported_total_cartons=6,
            reported_total_units=1200,
            reported_diff=-48,
            curve=12,
            uom="PRK"
        )
        res = calculate_carton_quantities(inp)
        assert res.calculated_units == 1200
        assert res.calculated_cartons == 6
        assert res.calculated_packs == 100
        assert res.loose_units == 0
        assert res.formula_discrepancy == 0
        assert FLAG_QC_COMPLEMENT_WITHOUT_CARTON_COUNT in res.flags
        assert FLAG_NEGATIVE_REPORTED_DIFF in res.flags
        assert "Alerta: 48 unidades complementarias rescatadas" in res.explanation

    def test_variante_suelta_row(self):
        """Loose variant breakout from broken complement cartons."""
        inp = CartonCalculationInput(
            full_cartons=None,
            reported_total_cartons=0,
            reported_total_units=10,
            observation="VARIANTE SUELTA",
            uom="UN"
        )
        res = calculate_carton_quantities(inp)
        assert res.calculated_units == 10
        assert res.calculated_cartons == 0
        assert res.calculated_packs == 10
        assert FLAG_VARIANTE_SUELTA_ROW in res.flags

    def test_shared_carton_pair(self):
        """Two materials packed into one shared box (DOS MATERIALES POR BULTO)."""
        # Primary item
        inp1 = CartonCalculationInput(
            full_cartons=1,
            units_per_full_carton=300,
            reported_total_cartons=1,
            reported_total_units=300,
            observation="DOS MATERIALES POR BULTO",
            uom="UN"
        )
        res1 = calculate_carton_quantities(inp1)
        assert res1.calculated_units == 300
        assert res1.calculated_cartons == 1
        assert FLAG_SHARED_CARTON_PRIMARY in res1.flags

        # Secondary item
        inp2 = CartonCalculationInput(
            full_cartons=None,
            units_per_full_carton=300,
            reported_total_cartons=0,
            reported_total_units=300,
            observation="DOS MATERIALES POR BULTO",
            uom="UN"
        )
        res2 = calculate_carton_quantities(inp2)
        assert res2.calculated_units == 300
        assert res2.calculated_cartons == 0  # Crucial: 0 physical cartons to avoid double counting!
        assert FLAG_SHARED_CARTON_SECONDARY in res2.flags

        # Combined pair totals:
        total_boxes = res1.calculated_cartons + res2.calculated_cartons
        total_units = res1.calculated_units + res2.calculated_units
        assert total_boxes == 1
        assert total_units == 600

    def test_uneven_prepack_curve(self):
        """Units that do not divide evenly by curve trigger uneven curve flag."""
        inp = CartonCalculationInput(
            full_cartons=15,
            units_per_full_carton=48,
            complement_cartons=1,
            units_in_complement=22,
            curve=12,
            uom="PRK"
        )
        res = calculate_carton_quantities(inp)
        # 15*48 + 22 = 720 + 22 = 742 units
        assert res.calculated_units == 742
        assert res.calculated_cartons == 16
        assert res.calculated_packs == 61  # 61 * 12 = 732
        assert res.loose_units == 10       # 742 % 12 = 10
        assert FLAG_UNEVEN_PREPACK_CURVE in res.flags

    def test_formula_mismatch_detection(self):
        """Detect when reported Excel value contradicts calculated physical units."""
        inp = CartonCalculationInput(
            full_cartons=5,
            units_per_full_carton=100,
            reported_total_units=450,
            uom="UN"
        )
        res = calculate_carton_quantities(inp)
        assert res.calculated_units == 500
        assert res.formula_discrepancy == 50
        assert FLAG_QC_FORMULA_MISMATCH in res.flags

    def test_observation_adjustment_keywords(self):
        """Detect QC notes indicating manual additions/deductions."""
        inp_plus = CartonCalculationInput(
            full_cartons=5,
            units_per_full_carton=192,
            observation="MAS 1PK, 5685920"
        )
        res_plus = calculate_carton_quantities(inp_plus)
        assert FLAG_QC_OBSERVATION_ADJUSTMENT in res_plus.flags

        inp_minus = CartonCalculationInput(
            full_cartons=9,
            units_per_full_carton=192,
            observation="MENOS 1PK"
        )
        res_minus = calculate_carton_quantities(inp_minus)
        assert FLAG_QC_OBSERVATION_ADJUSTMENT in res_minus.flags

    def test_empty_complement_cartons_flag(self):
        """Detect complement cartons declared with 0 or missing units."""
        inp = CartonCalculationInput(
            full_cartons=10,
            units_per_full_carton=50,
            complement_cartons=2,
            units_in_complement=0
        )
        res = calculate_carton_quantities(inp)
        assert res.calculated_cartons == 12
        assert res.calculated_units == 500
        assert FLAG_EMPTY_COMPLEMENT_CARTONS in res.flags

    def test_fallback_to_packaging_string(self):
        """When spreadsheet columns are empty, engine falls back to packaging string."""
        inp = CartonCalculationInput(
            packaging_string="6*192 +48",
            curve=12,
            uom="PRK"
        )
        res = calculate_carton_quantities(inp)
        assert res.calculated_units == 1200
        assert res.calculated_cartons == 6
        assert res.calculated_packs == 100
        assert res.loose_units == 0
