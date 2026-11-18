import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from examples.create_sample import create_sample
from main import (
    ENTRY_COLUMNS,
    ValidationError,
    build_upload_from_records,
    build_upload_workbook,
    create_entry_template,
    parse_number,
)


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.source = Path(self.directory.name) / "source.xlsx"
        self.entries = self.source.with_name("entries.xlsx")
        self.output = self.source.with_name("output.xlsx")
        create_sample(self.source)
        self.original = self.source.read_bytes()
        self.record = dict.fromkeys(ENTRY_COLUMNS, "")
        self.record.update({
            "Ürün Adı*": "Yeni Defter", "Fiyat*": "1.250,50",
            "Şablon Ürün Modeli*": "K000001",
        })

    def test_excel_workflow_preserves_source_and_copies_template(self):
        create_entry_template(self.source, self.entries)
        workbook = load_workbook(self.entries)
        workbook["Yeni Ürünler"].append([self.record[key] for key in ENTRY_COLUMNS])
        workbook.save(self.entries)
        workbook.close()
        self.assertEqual(build_upload_workbook(self.source, self.entries, self.output), 1)
        self.assertEqual(self.source.read_bytes(), self.original)
        result = load_workbook(self.output)
        self.addCleanup(result.close)
        products = result["Products"]
        self.assertEqual(products.max_row, 3)
        self.assertEqual([products.cell(3, col).value for col in range(1, 7)],
                         [2, "Yeni Defter", "K000002", 1250.5, 50, "1"])
        self.assertEqual(products["I3"].value, "Yeni Defter")
        self.assertTrue(products["B3"].font.bold)
        self.assertEqual(products["D3"].number_format, "0.00")

    def test_duplicate_and_unknown_models_do_not_create_output(self):
        for changes in ({"Model (boşsa otomatik)": "K000001"},
                        {"Şablon Ürün Modeli*": "UNKNOWN"}):
            with self.subTest(changes=changes):
                with self.assertRaises(ValidationError):
                    build_upload_from_records(self.source, [(2, self.record | changes)], self.output)
                self.assertFalse(self.output.exists())
                self.assertEqual(self.source.read_bytes(), self.original)

    def test_input_files_cannot_be_overwritten(self):
        create_entry_template(self.source, self.entries)
        original_entries = self.entries.read_bytes()
        actions = [
            lambda: create_entry_template(self.source, self.source),
            lambda: build_upload_from_records(self.source, [(2, self.record)], self.source),
            lambda: build_upload_workbook(self.source, self.entries, self.entries),
        ]
        for action in actions:
            with self.assertRaises(ValidationError):
                action()
        self.assertEqual(self.source.read_bytes(), self.original)
        self.assertEqual(self.entries.read_bytes(), original_entries)

    def test_empty_input_does_not_create_output(self):
        with self.assertRaises(ValidationError):
            build_upload_from_records(self.source, [], self.output)
        self.assertFalse(self.output.exists())

    def test_invalid_numbers_are_rejected(self):
        for value, label in [("abc", "fiyat"), ("nan", "fiyat"),
                             (float("inf"), "fiyat"), (-1, "fiyat"),
                             ("-2", "stok"), ("1,5", "stok"), (True, "stok")]:
            with self.subTest(value=value, label=label):
                with self.assertRaises(ValidationError):
                    parse_number(value, 2, label)

    def test_multiple_products_get_unique_ids_and_allow_zero_stock(self):
        record = self.record | {"Stok (boşsa şablondan)": "0"}
        self.assertEqual(build_upload_from_records(
            self.source, [(2, record), (3, record)], self.output), 2)
        result = load_workbook(self.output)
        self.addCleanup(result.close)
        products = result["Products"]
        self.assertEqual([products["A3"].value, products["A4"].value], [2, 3])
        self.assertEqual([products["C3"].value, products["C4"].value], ["K000002", "K000003"])
        self.assertEqual(products["E3"].value, 0)


if __name__ == "__main__":
    unittest.main()
