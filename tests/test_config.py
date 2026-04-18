"""Tests for fatturazione_xml.config module."""

import json
import unittest
from pathlib import Path
from unittest.mock import patch
import tempfile
import os


class TestLoadConfigDefaults(unittest.TestCase):
    def test_load_config_returns_defaults_when_no_file(self):
        """load_config() returns all 4 default keys when CONFIG_FILE does not exist."""
        import fatturazione_xml.config as cfg

        with tempfile.TemporaryDirectory() as tmpdir:
            fake_file = Path(tmpdir) / "nonexistent" / "config.json"
            with patch.object(cfg, "CONFIG_FILE", fake_file):
                result = cfg.load_config()

        self.assertIn("xlsm_path", result)
        self.assertIn("xml_output_dir", result)
        self.assertIn("filename_prefix", result)
        self.assertIn("year", result)
        self.assertEqual(result["xlsm_path"], cfg.DEFAULT_XLSM_PATH)
        self.assertEqual(result["xml_output_dir"], cfg.DEFAULT_XML_OUTPUT_DIR)
        self.assertEqual(result["filename_prefix"], cfg.DEFAULT_FILENAME_PREFIX)
        self.assertEqual(result["year"], cfg.DEFAULT_YEAR)


class TestSaveAndLoad(unittest.TestCase):
    def test_save_and_load_roundtrip(self):
        """Values saved with save_config() are returned unchanged by load_config()."""
        import fatturazione_xml.config as cfg

        custom = {
            "xlsm_path": "/tmp/my_file.xlsm",
            "xml_output_dir": "/tmp/xml_out",
            "filename_prefix": "IT99999_X",
            "year": 2025,
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            fake_dir = Path(tmpdir) / "FatturazioneXML"
            fake_file = fake_dir / "config.json"

            with patch.object(cfg, "CONFIG_DIR", fake_dir), \
                 patch.object(cfg, "CONFIG_FILE", fake_file):
                cfg.save_config(custom)
                loaded = cfg.load_config()

        self.assertEqual(loaded["xlsm_path"], "/tmp/my_file.xlsm")
        self.assertEqual(loaded["xml_output_dir"], "/tmp/xml_out")
        self.assertEqual(loaded["filename_prefix"], "IT99999_X")
        self.assertEqual(loaded["year"], 2025)

    def test_save_creates_directory(self):
        """save_config() creates the CONFIG_DIR directory tree if it doesn't exist."""
        import fatturazione_xml.config as cfg

        with tempfile.TemporaryDirectory() as tmpdir:
            fake_dir = Path(tmpdir) / "deep" / "nested" / "FatturazioneXML"
            fake_file = fake_dir / "config.json"

            self.assertFalse(fake_dir.exists())

            with patch.object(cfg, "CONFIG_DIR", fake_dir), \
                 patch.object(cfg, "CONFIG_FILE", fake_file):
                cfg.save_config(cfg._defaults())

            self.assertTrue(fake_dir.exists())
            self.assertTrue(fake_file.exists())


class TestIsConfigured(unittest.TestCase):
    def test_is_configured_false_when_no_file(self):
        """is_configured() returns False when CONFIG_FILE does not exist."""
        import fatturazione_xml.config as cfg

        with tempfile.TemporaryDirectory() as tmpdir:
            fake_file = Path(tmpdir) / "no_such_config.json"
            with patch.object(cfg, "CONFIG_FILE", fake_file):
                result = cfg.is_configured()

        self.assertFalse(result)

    def test_is_configured_false_when_xlsm_missing(self):
        """is_configured() returns False when xlsm_path points to a nonexistent file."""
        import fatturazione_xml.config as cfg

        with tempfile.TemporaryDirectory() as tmpdir:
            fake_dir = Path(tmpdir) / "FatturazioneXML"
            fake_file = fake_dir / "config.json"

            config = cfg._defaults()
            config["xlsm_path"] = "/nonexistent/path/file.xlsm"

            with patch.object(cfg, "CONFIG_DIR", fake_dir), \
                 patch.object(cfg, "CONFIG_FILE", fake_file):
                cfg.save_config(config)
                result = cfg.is_configured()

        self.assertFalse(result)

    def test_is_configured_true_when_valid(self):
        """is_configured() returns True when xlsm_path points to a real existing file."""
        import fatturazione_xml.config as cfg

        real_xlsm = Path.home() / "fatturazione" / "Database fatturazione 2026.xlsm"
        if not real_xlsm.exists():
            self.skipTest(f"Real xlsm not found at {real_xlsm}")

        with tempfile.TemporaryDirectory() as tmpdir:
            fake_dir = Path(tmpdir) / "FatturazioneXML"
            fake_file = fake_dir / "config.json"

            config = cfg._defaults()
            config["xlsm_path"] = str(real_xlsm)

            with patch.object(cfg, "CONFIG_DIR", fake_dir), \
                 patch.object(cfg, "CONFIG_FILE", fake_file):
                cfg.save_config(config)
                result = cfg.is_configured()

        self.assertTrue(result)


class TestGetOutputFilename(unittest.TestCase):
    def test_get_output_filename(self):
        """get_output_filename() returns prefix + numinvio_next + '.xml'."""
        import fatturazione_xml.config as cfg

        config = {"filename_prefix": "IT123_G", "xml_output_dir": "/tmp", "year": 2026}
        result = cfg.get_output_filename(config, 42)
        self.assertEqual(result, "IT123_G42.xml")


class TestGetOutputPath(unittest.TestCase):
    def test_get_output_path_creates_dir(self):
        """get_output_path() creates the output directory and returns the correct full path."""
        import fatturazione_xml.config as cfg

        with tempfile.TemporaryDirectory() as tmpdir:
            output_dir = Path(tmpdir) / "new_xml_output"
            self.assertFalse(output_dir.exists())

            config = {
                "filename_prefix": "IT01652160894_G",
                "xml_output_dir": str(output_dir),
                "year": 2026,
            }
            result = cfg.get_output_path(config, 7)

            self.assertTrue(output_dir.exists())
            self.assertEqual(result, output_dir / "IT01652160894_G7.xml")


if __name__ == "__main__":
    unittest.main()
