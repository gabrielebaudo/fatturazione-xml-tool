"""
Tests for fatturazione_xml.xlsm_parser

Runs against the real .xlsm file at:
    /Users/gabrielebaudo/fatturazione/Database fatturazione 2026.xlsm
"""

import tempfile
import unittest
import zipfile
from fatturazione_xml.xlsm_parser import (
    _build_xml_maps,
    list_xml_sheets,
    get_sheet_bindings,
)

XLSM_PATH = "/Users/gabrielebaudo/fatturazione/Database fatturazione 2026.xlsm"

EXPECTED_SHEETS = ["XML con IVA", "XML intento", "XML no IVA", "XML Estere"]
FATTURA_NS = "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"


class TestListXmlSheets(unittest.TestCase):

    def test_returns_list(self):
        result = list_xml_sheets(XLSM_PATH)
        self.assertIsInstance(result, list)

    def test_contains_expected_sheets(self):
        result = list_xml_sheets(XLSM_PATH)
        for expected in EXPECTED_SHEETS:
            self.assertIn(
                expected,
                result,
                msg=f"Expected sheet {expected!r} not found in {result}",
            )

    def test_at_least_four_sheets(self):
        result = list_xml_sheets(XLSM_PATH)
        self.assertGreaterEqual(
            len(result),
            4,
            msg=f"Expected at least 4 XML sheets, got {len(result)}: {result}",
        )


class TestGetSheetBindingsXmlConIVA(unittest.TestCase):

    def setUp(self):
        self.sb = get_sheet_bindings(XLSM_PATH, "XML con IVA")

    def test_sheet_name(self):
        self.assertEqual(self.sb.sheet_name, "XML con IVA")

    def test_map_name(self):
        self.assertEqual(
            self.sb.xml_map.name,
            "FatturaOrdiva",
            msg=f"Expected map name 'FatturaOrdiva', got {self.sb.xml_map.name!r}",
        )

    def test_map_namespace(self):
        self.assertEqual(
            self.sb.xml_map.namespace,
            FATTURA_NS,
            msg=f"Unexpected namespace: {self.sb.xml_map.namespace!r}",
        )

    def test_map_root_element(self):
        self.assertEqual(self.sb.xml_map.root_element, "FatturaElettronica")

    def test_map_id(self):
        self.assertEqual(self.sb.xml_map.map_id, 28)

    def test_at_least_20_bindings(self):
        self.assertGreaterEqual(
            len(self.sb.bindings),
            20,
            msg=f"Expected at least 20 bindings, got {len(self.sb.bindings)}",
        )

    def test_cell_D3_progressivo_invio(self):
        cell_map = {b.cell_ref: b for b in self.sb.bindings}
        self.assertIn("D3", cell_map, msg="Cell D3 not found in bindings")
        self.assertIn(
            "ProgressivoInvio",
            cell_map["D3"].xpath,
            msg=f"Expected 'ProgressivoInvio' in D3 xpath, got: {cell_map['D3'].xpath!r}",
        )

    def test_cell_M7_numero(self):
        cell_map = {b.cell_ref: b for b in self.sb.bindings}
        self.assertIn("M7", cell_map, msg="Cell M7 not found in bindings")
        self.assertIn(
            "Numero",
            cell_map["M7"].xpath,
            msg=f"Expected 'Numero' in M7 xpath, got: {cell_map['M7'].xpath!r}",
        )

    def test_cell_J7_tipo_documento(self):
        cell_map = {b.cell_ref: b for b in self.sb.bindings}
        self.assertIn("J7", cell_map, msg="Cell J7 not found in bindings")
        self.assertIn(
            "TipoDocumento",
            cell_map["J7"].xpath,
            msg=f"Expected 'TipoDocumento' in J7 xpath, got: {cell_map['J7'].xpath!r}",
        )

    def test_bindings_have_correct_types(self):
        """All bindings should have non-empty cell_ref, xpath, and data_type."""
        for b in self.sb.bindings:
            self.assertTrue(b.cell_ref, f"Empty cell_ref in binding: {b}")
            self.assertTrue(b.xpath, f"Empty xpath in binding: {b}")
            self.assertTrue(b.data_type, f"Empty data_type in binding: {b}")
            self.assertIsInstance(b.map_id, int)


class TestGetSheetBindingsXmlEstere(unittest.TestCase):

    def setUp(self):
        self.sb = get_sheet_bindings(XLSM_PATH, "XML Estere")

    def test_sheet_name(self):
        self.assertEqual(self.sb.sheet_name, "XML Estere")

    def test_map_name(self):
        self.assertEqual(
            self.sb.xml_map.name,
            "FatturaStraniere",
            msg=f"Expected 'FatturaStraniere', got {self.sb.xml_map.name!r}",
        )

    def test_map_namespace(self):
        self.assertEqual(self.sb.xml_map.namespace, FATTURA_NS)

    def test_map_id(self):
        self.assertEqual(self.sb.xml_map.map_id, 29)

    def test_has_bindings(self):
        self.assertGreater(len(self.sb.bindings), 0)


class TestGetSheetBindingsOtherSheets(unittest.TestCase):

    def test_xml_intento(self):
        sb = get_sheet_bindings(XLSM_PATH, "XML intento")
        self.assertEqual(sb.xml_map.name, "FatturaLintento")
        self.assertEqual(sb.xml_map.namespace, FATTURA_NS)
        self.assertEqual(sb.xml_map.map_id, 32)
        self.assertGreater(len(sb.bindings), 0)

    def test_xml_no_iva(self):
        sb = get_sheet_bindings(XLSM_PATH, "XML no IVA")
        self.assertEqual(sb.xml_map.name, "FatturaOrdnoiva")
        self.assertEqual(sb.xml_map.namespace, FATTURA_NS)
        self.assertEqual(sb.xml_map.map_id, 30)
        self.assertGreater(len(sb.bindings), 0)


class TestErrorHandling(unittest.TestCase):

    def test_unknown_sheet_raises_key_error(self):
        with self.assertRaises(KeyError):
            get_sheet_bindings(XLSM_PATH, "NonExistentSheet")

    def test_sheet_without_bindings_raises_value_error(self):
        # "Fattura" is the first sheet and has no XML bindings
        with self.assertRaises((KeyError, ValueError)):
            get_sheet_bindings(XLSM_PATH, "Fattura")


class TestXmlMapSchemaOrder(unittest.TestCase):

    def test_schema_ref_order_is_extracted(self):
        xml_maps = """<?xml version="1.0" encoding="UTF-8"?>
<xmlMaps xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
         xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <Schema ID="SchemaA" SchemaRef="SchemaB"
          Namespace="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2">
    <xs:schema targetNamespace="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2">
      <xs:element name="FatturaElettronica">
        <xs:complexType>
          <xs:sequence>
            <xs:element ref="FatturaElettronicaHeader"/>
            <xs:element ref="FatturaElettronicaBody"/>
          </xs:sequence>
        </xs:complexType>
      </xs:element>
    </xs:schema>
  </Schema>
  <Schema ID="SchemaB">
    <xs:schema>
      <xs:element name="Sede">
        <xs:complexType>
          <xs:sequence>
            <xs:element name="Provincia"/>
            <xs:element name="Nazione"/>
          </xs:sequence>
        </xs:complexType>
      </xs:element>
    </xs:schema>
  </Schema>
  <Map ID="28" Name="FatturaOrdiva" RootElement="FatturaElettronica" SchemaID="SchemaA"/>
</xmlMaps>
"""
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False)
        tmp.close()
        self.addCleanup(lambda: __import__("os").unlink(tmp.name))

        with zipfile.ZipFile(tmp.name, "w") as zf:
            zf.writestr("xl/xmlMaps.xml", xml_maps)

        with zipfile.ZipFile(tmp.name, "r") as zf:
            maps = _build_xml_maps(zf)

        self.assertEqual(
            maps[28].element_order["FatturaElettronica"],
            ["FatturaElettronicaHeader", "FatturaElettronicaBody"],
        )
        self.assertEqual(maps[28].element_order["Sede"], ["Provincia", "Nazione"])


if __name__ == "__main__":
    unittest.main()
