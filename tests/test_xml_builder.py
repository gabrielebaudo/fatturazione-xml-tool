"""
Tests for fatturazione_xml.xml_builder

Most tests use manually constructed bindings — no real xlsm required.
One integration test loads the real workbook.
"""

import datetime
import unittest
import xml.etree.ElementTree as ET

from fatturazione_xml.xlsm_parser import XmlBinding, XmlMapInfo
from fatturazione_xml.xml_builder import build_xml

XLSM_PATH = "/Users/gabrielebaudo/fatturazione/Database fatturazione 2026.xlsm"
FATTURA_NS = "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_map(namespace: str = FATTURA_NS) -> XmlMapInfo:
    return XmlMapInfo(
        map_id=1,
        name="TestMap",
        root_element="FatturaElettronica",
        namespace=namespace,
    )


def _make_binding(xpath: str, data_type: str = "string") -> XmlBinding:
    return XmlBinding(cell_ref="A1", xpath=xpath, data_type=data_type, map_id=1)


def _ns(tag: str, ns: str = FATTURA_NS) -> str:
    """Return Clark-notation tag: {ns}tag"""
    return f"{{{ns}}}{tag}"


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestMinimalInvoice(unittest.TestCase):
    """Build a minimal FatturaPA XML with ~10 key bindings."""

    def setUp(self):
        self.xml_map = _make_map()
        bindings_with_values = [
            (
                _make_binding("/ns1:FatturaElettronica/@versione"),
                "FPR12",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaHeader"
                    "/DatiTrasmissione/ProgressivoInvio"
                ),
                "00001",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/TipoDocumento"
                ),
                "TD01",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/Numero"
                ),
                "1",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/Data",
                    "date",
                ),
                "2026-04-01",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaHeader"
                    "/CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione"
                ),
                "Fornitore SRL",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaHeader"
                    "/CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione"
                ),
                "Cliente SPA",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaHeader"
                    "/DatiTrasmissione/FormatoTrasmissione"
                ),
                "FPR12",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaHeader"
                    "/CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdCodice"
                ),
                "12345678901",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaHeader"
                    "/DatiTrasmissione/CodiceDestinatario"
                ),
                "ABCDEFG",
            ),
        ]
        self.output = build_xml(bindings_with_values, self.xml_map)
        self.tree = ET.fromstring(self.output)  # parse for assertions

    def test_output_is_valid_xml(self):
        """Output must be parseable by ET.fromstring without errors."""
        self.assertIsNotNone(self.tree)

    def test_xml_declaration_present(self):
        """Output must start with the XML declaration."""
        self.assertTrue(
            self.output.startswith('<?xml version="1.0" encoding="UTF-8"?>'),
            msg=f"Missing XML declaration. Output starts with: {self.output[:60]!r}",
        )

    def test_fattura_namespace_in_output(self):
        """Root element must carry the FatturaPA namespace."""
        self.assertIn(FATTURA_NS, self.output)

    def test_root_tag_contains_fatturaelettronica(self):
        """Root element tag should include 'FatturaElettronica'."""
        self.assertIn("FatturaElettronica", self.tree.tag)

    def test_progressivo_invio_value(self):
        """ProgressivoInvio element must have the expected text."""
        el = self.tree.find(
            "./FatturaElettronicaHeader/DatiTrasmissione/ProgressivoInvio"
        )
        self.assertIsNotNone(el, msg="ProgressivoInvio element not found")
        self.assertEqual(el.text, "00001")

    def test_tipo_documento_value(self):
        """TipoDocumento element must be 'TD01'."""
        el = self.tree.find(
            "./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/TipoDocumento"
        )
        self.assertIsNotNone(el, msg="TipoDocumento element not found")
        self.assertEqual(el.text, "TD01")


class TestNoneValuesSkipped(unittest.TestCase):
    """Bindings with None value must not produce elements in the output."""

    def test_none_element_absent(self):
        xml_map = _make_map()
        bindings_with_values = [
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/TipoDocumento"
                ),
                "TD01",
            ),
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/Causale"
                ),
                None,  # should be skipped
            ),
        ]
        output = build_xml(bindings_with_values, xml_map)
        tree = ET.fromstring(output)

        # TipoDocumento should exist
        td = tree.find(
            "./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/TipoDocumento"
        )
        self.assertIsNotNone(td)

        # Causale should be absent
        causale = tree.find(
            "./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Causale"
        )
        self.assertIsNone(causale, msg="Causale should not appear when value is None")


class TestDateFormatting(unittest.TestCase):
    """datetime values must be formatted as YYYY-MM-DD."""

    def test_datetime_formatted(self):
        xml_map = _make_map()
        dt = datetime.datetime(2026, 4, 17, 10, 30, 0)
        bindings_with_values = [
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/Data",
                    "date",
                ),
                dt,
            ),
        ]
        output = build_xml(bindings_with_values, xml_map)
        tree = ET.fromstring(output)

        el = tree.find(
            "./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data"
        )
        self.assertIsNotNone(el, msg="Data element not found")
        self.assertEqual(el.text, "2026-04-17")

    def test_date_formatted(self):
        xml_map = _make_map()
        d = datetime.date(2026, 1, 5)
        bindings_with_values = [
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/Data",
                    "date",
                ),
                d,
            ),
        ]
        output = build_xml(bindings_with_values, xml_map)
        tree = ET.fromstring(output)

        el = tree.find(
            "./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data"
        )
        self.assertIsNotNone(el)
        self.assertEqual(el.text, "2026-01-05")


class TestFloatFormatting(unittest.TestCase):
    """float values must be formatted with 2 decimal places."""

    def test_float_22(self):
        xml_map = _make_map()
        bindings_with_values = [
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiBeniServizi/DettaglioLinee/AliquotaIVA",
                    "decimal",
                ),
                22.0,
            ),
        ]
        output = build_xml(bindings_with_values, xml_map)
        tree = ET.fromstring(output)

        el = tree.find(
            "./FatturaElettronicaBody/DatiBeniServizi/DettaglioLinee/AliquotaIVA"
        )
        self.assertIsNotNone(el, msg="AliquotaIVA element not found")
        self.assertEqual(el.text, "22.00")

    def test_float_1234_56(self):
        xml_map = _make_map()
        bindings_with_values = [
            (
                _make_binding(
                    "/ns1:FatturaElettronica/FatturaElettronicaBody"
                    "/DatiGenerali/DatiGeneraliDocumento/ImportoTotaleDocumento",
                    "decimal",
                ),
                1234.56,
            ),
        ]
        output = build_xml(bindings_with_values, xml_map)
        tree = ET.fromstring(output)

        el = tree.find(
            "./FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento"
            "/ImportoTotaleDocumento"
        )
        self.assertIsNotNone(el)
        self.assertEqual(el.text, "1234.56")


class TestAttributeBinding(unittest.TestCase):
    """The @versione xpath must set the attribute on the root element."""

    def test_versione_attribute(self):
        xml_map = _make_map()
        bindings_with_values = [
            (
                _make_binding("/ns1:FatturaElettronica/@versione"),
                "FPR12",
            ),
        ]
        output = build_xml(bindings_with_values, xml_map)
        tree = ET.fromstring(output)

        self.assertEqual(
            tree.get("versione"),
            "FPR12",
            msg=f"Expected versione='FPR12' on root, got: {tree.attrib!r}",
        )

    def test_versione_in_raw_output(self):
        """The raw output string must contain versione="FPR12"."""
        xml_map = _make_map()
        bindings_with_values = [
            (_make_binding("/ns1:FatturaElettronica/@versione"), "FPR12"),
        ]
        output = build_xml(bindings_with_values, xml_map)
        self.assertIn('versione="FPR12"', output)


class TestIntegrationWithRealWorkbook(unittest.TestCase):
    """Full round-trip: real xlsm → build_xml → parse."""

    def test_build_from_real_workbook(self):
        from fatturazione_xml.xlsm_parser import get_sheet_bindings
        from fatturazione_xml.excel_reader import read_cell_values

        sb = get_sheet_bindings(XLSM_PATH, "XML con IVA")
        bindings_with_values = read_cell_values(XLSM_PATH, "XML con IVA", sb.bindings)

        output = build_xml(bindings_with_values, sb.xml_map)

        # Must parse without error
        try:
            tree = ET.fromstring(output)
        except ET.ParseError as exc:
            self.fail(f"Output is not valid XML: {exc}\n\nOutput:\n{output[:500]}")

        # Root tag must contain FatturaElettronica
        self.assertIn("FatturaElettronica", tree.tag)

        # Must have the XML declaration
        self.assertTrue(
            output.startswith('<?xml version="1.0" encoding="UTF-8"?>'),
            msg=f"Missing XML declaration. Starts with: {output[:80]!r}",
        )

        # Namespace must be present
        self.assertIn(FATTURA_NS, output)


if __name__ == "__main__":
    unittest.main()
