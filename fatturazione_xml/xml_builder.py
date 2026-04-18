"""
xml_builder.py

Builds a FatturaPA v1.2 XML string from (XmlBinding, value) pairs.

Uses only stdlib: xml.etree.ElementTree
"""

from __future__ import annotations

import datetime
import re
import xml.etree.ElementTree as ET
from typing import Any

from .xlsm_parser import XmlBinding, XmlMapInfo

# FatturaPA namespace (also used in the output root element prefix)
_FATTURA_NS = "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"
_DS_NS = "http://www.w3.org/2000/09/xmldsig#"

# Regex to strip a position predicate like [1] from a segment
_PREDICATE_RE = re.compile(r"\[\d+\]$")


# ---------------------------------------------------------------------------
# XPath normalisation
# ---------------------------------------------------------------------------

def _normalise_xpath(xpath: str) -> tuple[list[str], str | None]:
    """
    Normalise a raw FatturaPA xpath into (path_segments, attr_name_or_None).

    Rules:
    1. Strip a leading '/'.
    2. Remove the first segment (the root element, e.g. 'ns1:FatturaElettronica').
    3. Strip any namespace prefix from every remaining segment.
    4. Strip position predicates like [1].
    5. If the last segment starts with '@', it is an attribute name on its
       *parent* element, not a child element; return it separately.

    Example:
        '/ns1:FatturaElettronica/@versione'
        → ([], 'versione')

        '/ns1:FatturaElettronica/FatturaElettronicaHeader/DatiTrasmissione/ProgressivoInvio'
        → (['FatturaElettronicaHeader', 'DatiTrasmissione', 'ProgressivoInvio'], None)
    """
    # Step 1 – strip leading slash
    if xpath.startswith("/"):
        xpath = xpath[1:]

    # Split on '/'
    parts = xpath.split("/")

    # Step 2 – drop the root element segment
    if parts:
        parts = parts[1:]

    segments: list[str] = []
    for part in parts:
        # Step 4 – strip position predicates
        part = _PREDICATE_RE.sub("", part)
        if not part:
            continue
        # Step 3 – strip namespace prefix (everything up to and including ':')
        if ":" in part and not part.startswith("@"):
            part = part.split(":", 1)[1]
        segments.append(part)

    # Step 5 – split off a trailing @attribute
    if segments and segments[-1].startswith("@"):
        attr_name = segments[-1][1:]  # drop the '@'
        path_segments = segments[:-1]
        return path_segments, attr_name

    return segments, None


# ---------------------------------------------------------------------------
# Value formatting
# ---------------------------------------------------------------------------

def _format_value(value: Any) -> str | None:
    """
    Convert a Python value to a string suitable for an XML element/attribute.
    Returns None if the value should be skipped entirely.
    """
    if value is None:
        return None

    if isinstance(value, bool):
        return "1" if value else "0"

    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, float):
        return f"{value:.2f}"

    if isinstance(value, int):
        return str(value)

    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None

    return str(value)


# ---------------------------------------------------------------------------
# Element-tree helpers
# ---------------------------------------------------------------------------

def _get_or_create(parent: ET.Element, tag: str) -> ET.Element:
    """Return the first child with *tag*, creating it if absent."""
    existing = parent.find(tag)
    if existing is None:
        existing = ET.SubElement(parent, tag)
    return existing


def _set_value(
    root: ET.Element,
    path_segments: list[str],
    attr_name: str | None,
    text: str,
) -> None:
    """
    Navigate / create path_segments under *root*, then either set the
    text of the final element or add an attribute to it.
    """
    node = root
    for seg in path_segments:
        node = _get_or_create(node, seg)

    if attr_name is not None:
        node.set(attr_name, text)
    else:
        node.text = text


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_xml(
    bindings_with_values: list[tuple[XmlBinding, Any]],
    xml_map: XmlMapInfo,
) -> str:
    """
    Build a FatturaPA XML string from cell-bound values.

    Returns a UTF-8 encoded XML string (decoded, not bytes) with:
    - XML declaration: <?xml version="1.0" encoding="UTF-8"?>
    - Root element: p:FatturaElettronica xmlns:p="<namespace>"
      (namespace from xml_map.namespace)
    - All bound cell values mapped to their xpath positions

    Raises:
        ValueError: if a required structural element cannot be built
    """
    namespace = xml_map.namespace or _FATTURA_NS

    # Register prefix 'p' for the FatturaPA namespace so ElementTree uses it.
    ET.register_namespace("p", namespace)
    ET.register_namespace("ds", _DS_NS)

    # Build root element with the prefixed tag
    root_tag = f"{{{namespace}}}FatturaElettronica"
    root = ET.Element(root_tag)

    # Collect attributes that need to be set on the root element
    # (e.g. versione="FPR12") separately so we can set them after
    # processing all bindings.
    root_attrs: dict[str, str] = {}

    for binding, value in bindings_with_values:
        text = _format_value(value)
        if text is None:
            continue

        path_segments, attr_name = _normalise_xpath(binding.xpath)

        if not path_segments and attr_name is not None:
            # Attribute on the root element itself (e.g. @versione)
            root_attrs[attr_name] = text
            continue

        _set_value(root, path_segments, attr_name, text)

    # Apply root-level attributes last so they appear on the root element
    for attr_name, attr_val in root_attrs.items():
        root.set(attr_name, attr_val)

    # Pretty-print (Python 3.9+)
    ET.indent(root, space="  ")

    xml_body = ET.tostring(root, encoding="unicode", xml_declaration=False)
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_body
