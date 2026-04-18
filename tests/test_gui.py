"""
tests/test_gui.py

Pragmatic smoke tests for the GUI module.

We do NOT start a real tkinter event loop — instead we verify:
1. The gui.py module imports without errors and exposes a callable run().
2. Both gui.py and __main__.py are syntactically valid Python.
"""

import py_compile
import unittest
from pathlib import Path

# Absolute paths to the module files under test
_PKG = Path(__file__).parent.parent / "fatturazione_xml"
_GUI_FILE = _PKG / "gui.py"
_MAIN_FILE = _PKG / "__main__.py"


class TestGuiModuleSyntax(unittest.TestCase):
    def test_gui_py_compiles(self):
        """gui.py must be syntactically valid Python."""
        py_compile.compile(str(_GUI_FILE), doraise=True)

    def test_main_py_compiles(self):
        """__main__.py must be syntactically valid Python."""
        py_compile.compile(str(_MAIN_FILE), doraise=True)


class TestRunCallable(unittest.TestCase):
    def test_run_is_callable(self):
        """run() exported from fatturazione_xml.gui must be callable."""
        from fatturazione_xml.gui import run  # noqa: PLC0415
        self.assertTrue(callable(run))


if __name__ == "__main__":
    unittest.main()
