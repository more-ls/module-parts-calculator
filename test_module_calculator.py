#!/usr/bin/env python3
"""Tests for module_calculator.py — run after each edit: python test_module_calculator.py"""

import tempfile
import unittest
from pathlib import Path

from module_calculator import (
    PART_DEFINITIONS_XLSX,
    INVENTORY_XLSX,
    REPORTS_DIR,
    _parts_sorted_by_category,
    _remaining_to_order,
    build_summary_report,
    calculate,
    load_bom_from_excel,
    load_inventory_from_excel,
    save_report_pdf,
)

PROJECT_DIR = Path(__file__).parent


class TestCalculations(unittest.TestCase):
    def test_remaining_to_order(self):
        self.assertEqual(_remaining_to_order(10, 3), 7)
        self.assertEqual(_remaining_to_order(5, 5), 0)
        self.assertEqual(_remaining_to_order(2, 10), 0)

    def test_calculate_totals(self):
        bom = {
            "MO:HEAT": {"bottom plate": 1, "Magnets": 4},
            "tip holder 1000": {"bottom plate": 1, "Magnets": 4},
        }
        counts = {"MO:HEAT": 2, "tip holder 1000": 1}
        total_modules, total_parts, parts_by_type = calculate(counts, bom)

        self.assertEqual(total_modules, 3)
        self.assertEqual(total_parts["bottom plate"], 3)
        self.assertEqual(total_parts["Magnets"], 12)
        self.assertEqual(parts_by_type["MO:HEAT"]["Magnets"], 8)

    def test_parts_sorted_by_category(self):
        parts = {"bottom plate": 1, "Magnets": 4, "heating mat": 1}
        categories = {
            "bottom plate": "cnc metal",
            "Magnets": "mechanical part",
            "heating mat": "electronics",
        }
        sorted_parts = _parts_sorted_by_category(parts, categories)
        self.assertEqual(sorted_parts[0][0], "mechanical part")
        self.assertEqual(sorted_parts[1][0], "electronics")
        self.assertEqual(sorted_parts[2][0], "cnc metal")


class TestReport(unittest.TestCase):
    def setUp(self):
        self.part_category = {
            "bottom plate": "cnc metal",
            "Magnets": "mechanical part",
            "heating mat": "electronics",
        }
        self.module_types_bom = {
            "MO:HEAT": {"bottom plate": 1, "Magnets": 4, "heating mat": 1},
        }
        self.module_counts = {"MO:HEAT": 2}
        self.total_parts = {"bottom plate": 2, "Magnets": 8, "heating mat": 2}
        self.parts_by_type = {
            "MO:HEAT": {"bottom plate": 2, "Magnets": 8, "heating mat": 2},
        }

    def test_report_has_inventory_columns(self):
        inventory = {"Magnets": 3, "bottom plate": 2, "heating mat": 0}
        _, sections = build_summary_report(
            "Test Project",
            ["MO:HEAT"],
            self.part_category,
            self.module_counts,
            2,
            self.total_parts,
            inventory,
        )

        summary = sections[0]
        self.assertEqual(
            summary.headers,
            ["Part", "Type", "Needed", "In Stock", "Remaining to Order"],
        )
        self.assertEqual(summary.group, "Parts Summary")

        magnets_row = next(row for row in summary.rows if row[0] == "Magnets")
        self.assertEqual(magnets_row, ["Magnets", "mechanical part", "8", "3", "5"])

    def test_report_creates_order_section(self):
        inventory = {"Magnets": 0}
        _, sections = build_summary_report(
            "Test Project",
            ["MO:HEAT"],
            self.part_category,
            self.module_counts,
            2,
            self.total_parts,
            inventory,
        )

        order_sections = [s for s in sections if s.group == "Parts to Order"]
        self.assertEqual(len(order_sections), 1)
        self.assertTrue(order_sections[0].page_break_before)

    def test_report_sections_have_groups(self):
        inventory = {}
        _, sections = build_summary_report(
            "Test Project",
            ["MO:HEAT"],
            self.part_category,
            self.module_counts,
            2,
            self.total_parts,
            inventory,
        )

        groups = [s.group for s in sections]
        self.assertIn("Parts Summary", groups)
        self.assertIn("Modules", groups)
        self.assertNotIn("Module Breakdown", groups)

    def test_pdf_is_created(self):
        inventory = {"Magnets": 1}
        title, sections = build_summary_report(
            "PDF Test",
            ["MO:HEAT"],
            self.part_category,
            self.module_counts,
            2,
            self.total_parts,
            inventory,
        )

        with tempfile.TemporaryDirectory() as tmp:
            original_reports = REPORTS_DIR
            try:
                import module_calculator as mc

                test_reports = Path(tmp) / "reports"
                mc.REPORTS_DIR = test_reports
                pdf_path = save_report_pdf(title, sections, "PDF Test")
                self.assertTrue(pdf_path.exists())
                self.assertGreater(pdf_path.stat().st_size, 500)
            finally:
                mc.REPORTS_DIR = original_reports


class TestExcelFiles(unittest.TestCase):
    def test_part_definitions_loads(self):
        if not PART_DEFINITIONS_XLSX.exists():
            self.skipTest("part definitions.xlsx not found")
        modules, categories, bom = load_bom_from_excel()
        self.assertGreater(len(modules), 0)
        self.assertGreater(len(categories), 0)
        self.assertGreater(sum(len(v) for v in bom.values()), 0)

    def test_inventory_loads(self):
        if not INVENTORY_XLSX.exists():
            self.skipTest("inventory.xlsx not found")
        stock = load_inventory_from_excel()
        self.assertGreater(len(stock), 0)
        self.assertIn("Magnets", stock)


class TestGeneratorScripts(unittest.TestCase):
    def test_create_bom_excel_runs(self):
        import create_bom_excel

        path = create_bom_excel.create_bom_excel()
        self.assertTrue(path.exists())

    def test_create_inventory_excel_runs(self):
        import create_inventory_excel

        path = create_inventory_excel.create_inventory_excel()
        self.assertTrue(path.exists())


if __name__ == "__main__":
    print("Running tests...\n")
    result = unittest.main(verbosity=2, exit=False)
    if result.result.wasSuccessful():
        print("\nAll tests passed.")
    else:
        print("\nSome tests failed.")
        raise SystemExit(1)
