"""
Tests for fatturazione_xml.excel_reader

Runs against the real .xlsm file at:
    /Users/gabrielebaudo/fatturazione/Database fatturazione 2026.xlsm
"""

import unittest
from fatturazione_xml.xlsm_parser import get_sheet_bindings, XmlBinding
from fatturazione_xml.excel_reader import read_cell_values, get_numinvio

XLSM_PATH = "/Users/gabrielebaudo/fatturazione/Database fatturazione 2026.xlsm"
SHEET_NAME = "XML con IVA"


class TestReadCellValues(unittest.TestCase):

    def setUp(self):
        sb = get_sheet_bindings(XLSM_PATH, SHEET_NAME)
        self.bindings = sb.bindings
        self.result = read_cell_values(XLSM_PATH, SHEET_NAME, self.bindings)

    def test_read_returns_list_of_tuples(self):
        """Result should be a list where each element is a (XmlBinding, ...) tuple."""
        self.assertIsInstance(self.result, list)
        first = self.result[0]
        self.assertIsInstance(first, tuple)
        self.assertIsInstance(first[0], XmlBinding)

    def test_read_length_matches_bindings(self):
        """One result per binding."""
        self.assertEqual(len(self.result), len(self.bindings))

    def test_versione_cell_has_value(self):
        """The cell bound to xpath containing @versione should have a non-None value."""
        versione_pairs = [
            (b, v) for b, v in self.result if "@versione" in b.xpath
        ]
        self.assertTrue(
            versione_pairs,
            msg="No binding with @versione found in result",
        )
        for binding, value in versione_pairs:
            self.assertIsNotNone(
                value,
                msg=f"Expected non-None value for {binding.cell_ref} ({binding.xpath})",
            )


class TestGetNuminvio(unittest.TestCase):

    def test_numinvio_is_positive_integer(self):
        """get_numinvio should return an int >= 0."""
        result = get_numinvio(XLSM_PATH)
        self.assertIsInstance(result, int)
        self.assertGreaterEqual(result, 0)


class TestErrorHandling(unittest.TestCase):

    def test_file_not_found_raises(self):
        """Calling with a nonexistent path should raise FileNotFoundError."""
        with self.assertRaises(FileNotFoundError):
            read_cell_values(
                "/nonexistent/path/file.xlsm",
                SHEET_NAME,
                [],
            )

    def test_get_numinvio_file_not_found_raises(self):
        """get_numinvio with a nonexistent path should raise FileNotFoundError."""
        with self.assertRaises(FileNotFoundError):
            get_numinvio("/nonexistent/path/file.xlsm")


if __name__ == "__main__":
    unittest.main()
