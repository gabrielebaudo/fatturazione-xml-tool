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
from dataclasses import dataclass

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

    # Build schemaId → namespace from <Schema> / <xsd:schema targetNamespace="...">
    schema_ns: dict[str, str] = {}
    for schema_el in root.findall(f"{{{_NS_SPREADSHEET}}}Schema"):
        schema_id = schema_el.get("ID", "")
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
        )
        maps[mid] = info
    return maps


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
