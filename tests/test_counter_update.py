"""
Tests for fatturazione_xml.counter_update

All tests that write use a temporary copy of the real workbook — the real
file is never modified.
"""

import shutil
import tempfile
import unittest

import openpyxl

from fatturazione_xml.counter_update import ConcurrentModificationError, increment_numinvio
from fatturazione_xml.excel_reader import get_numinvio

XLSM_PATH = "/Users/gabrielebaudo/fatturazione/Database fatturazione 2026.xlsm"
SHEET_NAME = "2026 XML"
CELL_REF = "K2"


class TestIncrementNuminvio(unittest.TestCase):

    def _make_temp_copy(self) -> str:
        """Create a temporary copy of the real xlsm and register cleanup."""
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False)
        tmp.close()
        shutil.copy2(XLSM_PATH, tmp.name)
        self.addCleanup(lambda: __import__("os").unlink(tmp.name))
        return tmp.name

    def test_increment_succeeds(self):
        """increment_numinvio returns current+1 and the file reflects the new value."""
        copy_path = self._make_temp_copy()

        current_value = get_numinvio(copy_path)
        result = increment_numinvio(copy_path, current_value)

        self.assertEqual(result, current_value + 1)
        # Verify the change was actually persisted.
        self.assertEqual(get_numinvio(copy_path), current_value + 1)

    def test_concurrent_modification_raises(self):
        """If K2 was changed between reading and writing, ConcurrentModificationError is raised."""
        copy_path = self._make_temp_copy()

        current_value = get_numinvio(copy_path)

        # Simulate a concurrent edit by bumping K2 by 99.
        wb = openpyxl.load_workbook(copy_path, keep_vba=True)
        wb[SHEET_NAME][CELL_REF].value = current_value + 99
        wb.save(copy_path)

        with self.assertRaises(ConcurrentModificationError):
            increment_numinvio(copy_path, current_value)

    def test_file_not_found_raises(self):
        """Calling with a nonexistent path raises FileNotFoundError."""
        with self.assertRaises(FileNotFoundError):
            increment_numinvio("/nonexistent/path/file.xlsm", 1)

    def test_vba_preserved(self):
        """After incrementing, the VBA archive must still be present in the saved file."""
        copy_path = self._make_temp_copy()

        current_value = get_numinvio(copy_path)
        increment_numinvio(copy_path, current_value)

        wb = openpyxl.load_workbook(copy_path, keep_vba=True)
        self.assertIsNotNone(
            wb.vba_archive,
            msg="VBA archive was lost after increment_numinvio saved the workbook",
        )


if __name__ == "__main__":
    unittest.main()
