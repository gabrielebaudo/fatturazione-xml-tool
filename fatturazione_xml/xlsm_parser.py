"""
xlsm_parser.py

Reads a .xlsm file as a zip archive and extracts XML Map bindings:
- Sheet name → filename mapping from xl/workbook.xml
- XML Maps info from xl/xmlMaps.xml
- Cell→XPath bindings from xl/tables/tableSingleCellsN.xml

Uses only stdlib: zipfile + xml.etree.ElementTree
"""

from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

# OpenXML namespaces
_NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_XSD = "http://www.w3.org/2001/XMLSchema"

# Relationship type for tableSingleCells
_REL_TYPE_SINGLE_CELLS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"
)


@dataclass
class XmlBinding:
    cell_ref: str    # e.g. "A3", "M7"
    xpath: str       # full xpath from the tableSingleCells XML
    data_type: str   # "string", "integer", "decimal", "date", etc.
    map_id: int


@dataclass
class XmlMapInfo:
    map_id: int
    name: str           # e.g. "FatturaOrdiva"
    root_element: str   # e.g. "FatturaElettronica"
    namespace: str      # e.g. "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"
    element_order: dict[str, list[str]] = field(default_factory=dict)


@dataclass
class SheetBindings:
    sheet_name: str     # e.g. "XML con IVA"
    xml_map: XmlMapInfo
    bindings: list[XmlBinding]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _read_zip_text(zf: zipfile.ZipFile, path: str) -> str:
    """Read a UTF-8 text file from an open ZipFile."""
    return zf.read(path).decode("utf-8")


def _parse_xml(zf: zipfile.ZipFile, path: str) -> ET.Element:
    """Parse an XML file from inside the zip and return the root element."""
    data = zf.read(path)
    return ET.fromstring(data)


def _build_sheet_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Return {sheet_name: sheet_filename} e.g. {'XML con IVA': 'sheet6.xml'}."""
    wb_root = _parse_xml(zf, "xl/workbook.xml")
    rels_root = _parse_xml(zf, "xl/_rels/workbook.xml.rels")

    # rId → target filename (e.g. "worksheets/sheet6.xml")
    rid_to_target: dict[str, str] = {}
    for rel in rels_root.findall(f"{{{_NS_RELS}}}Relationship"):
        rid_to_target[rel.get("Id")] = rel.get("Target", "")

    sheet_map: dict[str, str] = {}
    for sheet in wb_root.findall(
        f"{{{_NS_SPREADSHEET}}}sheets/{{{_NS_SPREADSHEET}}}sheet"
    ):
        name = sheet.get("name", "")
        rid = sheet.get(f"{{{_NS_RELATIONSHIPS}}}id", "")
        target = rid_to_target.get(rid, "")
        if target.startswith("worksheets/"):
            filename = target[len("worksheets/"):]
            sheet_map[name] = filename
    return sheet_map


def _build_xml_maps(zf: zipfile.ZipFile) -> dict[int, XmlMapInfo]:
    """Parse xl/xmlMaps.xml and return {map_id: XmlMapInfo}."""
    root = _parse_xml(zf, "xl/xmlMaps.xml")
    schema_elements = {
        schema_el.get("ID", ""): schema_el
        for schema_el in root.findall(f"{{{_NS_SPREADSHEET}}}Schema")
    }

    # Build schemaId → namespace from <Schema> / <xsd:schema targetNamespace="...">
    schema_ns: dict[str, str] = {}
    for schema_id, schema_el in schema_elements.items():
        for xsd_schema in schema_el.iter(f"{{{_NS_XSD}}}schema"):
            tns = xsd_schema.get("targetNamespace")
            if tns:
                schema_ns[schema_id] = tns
                break

    maps: dict[int, XmlMapInfo] = {}
    for map_el in root.findall(f"{{{_NS_SPREADSHEET}}}Map"):
        mid = int(map_el.get("ID", "0"))
        schema_id = map_el.get("SchemaID", "")
        info = XmlMapInfo(
            map_id=mid,
            name=map_el.get("Name", ""),
            root_element=map_el.get("RootElement", ""),
            namespace=schema_ns.get(schema_id, ""),
            element_order=_build_element_order(schema_elements, schema_id),
        )
        maps[mid] = info
    return maps


def _element_name(element: ET.Element) -> str | None:
    """Return the local xsd:element name, resolving simple ref attributes."""
    name = element.get("name") or element.get("ref")
    if not name:
        return None
    return name.split(":", 1)[-1]


def _merge_order(
    element_order: dict[str, list[str]],
    parent_name: str,
    child_names: list[str],
) -> None:
    """Merge child order while preserving the first schema order seen."""
    existing = element_order.setdefault(parent_name, [])
    for child_name in child_names:
        if child_name not in existing:
            existing.append(child_name)


def _parse_schema_element_order(schema_el: ET.Element) -> dict[str, list[str]]:
    """
    Extract {parent_element: [child_element, ...]} from xsd:sequence nodes.

    Excel stores most FatturaPA structures as inline complexType/sequence
    definitions, so walking every xsd:element is enough to recover ordering.
    """
    element_order: dict[str, list[str]] = {}
    xsd_element = f"{{{_NS_XSD}}}element"
    xsd_complex_type = f"{{{_NS_XSD}}}complexType"
    xsd_sequence = f"{{{_NS_XSD}}}sequence"

    for element in schema_el.iter(xsd_element):
        parent_name = _element_name(element)
        if not parent_name:
            continue

        complex_type = element.find(xsd_complex_type)
        if complex_type is None:
            continue

        sequence = complex_type.find(xsd_sequence)
        if sequence is None:
            continue

        child_names = [
            child_name
            for child in sequence.findall(xsd_element)
            if (child_name := _element_name(child))
        ]
        if child_names:
            _merge_order(element_order, parent_name, child_names)

    return element_order


def _schema_chain(
    schema_elements: dict[str, ET.Element],
    schema_id: str,
) -> list[ET.Element]:
    """Return the map schema plus any SchemaRef targets, avoiding cycles."""
    result: list[ET.Element] = []
    seen: set[str] = set()

    def visit(current_id: str) -> None:
        if not current_id or current_id in seen:
            return
        seen.add(current_id)

        schema_el = schema_elements.get(current_id)
        if schema_el is None:
            return

        result.append(schema_el)
        visit(schema_el.get("SchemaRef", ""))

    visit(schema_id)
    return result


def _build_element_order(
    schema_elements: dict[str, ET.Element],
    schema_id: str,
) -> dict[str, list[str]]:
    """Build element ordering from a map schema and its referenced schema."""
    element_order: dict[str, list[str]] = {}

    for schema_el in _schema_chain(schema_elements, schema_id):
        for parent_name, child_names in _parse_schema_element_order(schema_el).items():
            _merge_order(element_order, parent_name, child_names)

    return element_order


def _find_single_cells_file(zf: zipfile.ZipFile, sheet_filename: str) -> str | None:
    """
    Given a sheet filename (e.g. 'sheet6.xml'), look at the sheet's rels file
    (xl/worksheets/_rels/sheet6.xml.rels) and return the path to the
    tableSingleCells XML if one exists, or None.

    Returns a path relative to the xl/ directory, e.g. 'tables/tableSingleCells1.xml'.
    """
    rels_path = f"xl/worksheets/_rels/{sheet_filename}.rels"
    try:
        rels_root = _parse_xml(zf, rels_path)
    except (KeyError, ET.ParseError):
        return None

    for rel in rels_root.findall(f"{{{_NS_RELS}}}Relationship"):
        rel_type = rel.get("Type", "")
        if rel_type == _REL_TYPE_SINGLE_CELLS:
            # Target is like "../tables/tableSingleCells1.xml"
            target = rel.get("Target", "")
            # Resolve relative to xl/worksheets/ → strip "../"
            if target.startswith("../"):
                return f"xl/{target[3:]}"
            return f"xl/worksheets/{target}"
    return None


def _parse_bindings(zf: zipfile.ZipFile, tsc_path: str) -> list[XmlBinding]:
    """
    Parse a tableSingleCellsN.xml file and return a list of XmlBinding objects.
    """
    root = _parse_xml(zf, tsc_path)
    ns = _NS_SPREADSHEET  # tableSingleCells uses the spreadsheet namespace

    bindings: list[XmlBinding] = []
    for cell_el in root.findall(f"{{{ns}}}singleXmlCell"):
        cell_ref = cell_el.get("r", "")
        # <xmlCellPr> → <xmlPr mapId="..." xpath="..." xmlDataType="..."/>
        cell_pr = cell_el.find(f"{{{ns}}}xmlCellPr")
        if cell_pr is None:
            continue
        xml_pr = cell_pr.find(f"{{{ns}}}xmlPr")
        if xml_pr is None:
            continue

        map_id_str = xml_pr.get("mapId", "0")
        xpath = xml_pr.get("xpath", "")
        data_type = xml_pr.get("xmlDataType", "string")

        bindings.append(
            XmlBinding(
                cell_ref=cell_ref,
                xpath=xpath,
                data_type=data_type,
                map_id=int(map_id_str),
            )
        )
    return bindings


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def list_xml_sheets(xlsm_path: str) -> list[str]:
    """Return names of sheets that have XML single-cell bindings, in workbook order."""
    with zipfile.ZipFile(xlsm_path, "r") as zf:
        sheet_map = _build_sheet_map(zf)
        result: list[str] = []
        for sheet_name, sheet_file in sheet_map.items():
            tsc_path = _find_single_cells_file(zf, sheet_file)
            if tsc_path is not None:
                result.append(sheet_name)
    return result


def get_sheet_bindings(xlsm_path: str, sheet_name: str) -> SheetBindings:
    """Extract all cell→xpath bindings for the given sheet name.

    Raises:
        KeyError: if the sheet_name is not found in the workbook.
        ValueError: if the sheet has no XML single-cell bindings.
    """
    with zipfile.ZipFile(xlsm_path, "r") as zf:
        sheet_map = _build_sheet_map(zf)

        if sheet_name not in sheet_map:
            raise KeyError(f"Sheet not found: {sheet_name!r}")

        sheet_file = sheet_map[sheet_name]
        tsc_path = _find_single_cells_file(zf, sheet_file)
        if tsc_path is None:
            raise ValueError(
                f"Sheet {sheet_name!r} has no XML single-cell bindings"
            )

        bindings = _parse_bindings(zf, tsc_path)
        xml_maps = _build_xml_maps(zf)

        # Determine the map_id used by this sheet (take from the first binding)
        if not bindings:
            raise ValueError(f"No bindings found in {tsc_path}")

        map_id = bindings[0].map_id
        if map_id not in xml_maps:
            raise ValueError(f"XML Map ID {map_id} not found in xmlMaps.xml")

        return SheetBindings(
            sheet_name=sheet_name,
            xml_map=xml_maps[map_id],
            bindings=bindings,
        )
