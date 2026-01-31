"""High-level Math class for OMML equation management."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.math import CT_OMath

if TYPE_CHECKING:
    from pptx.oxml.shapes import ShapeElement


class Math:
    """High-level interface for OMML math equations."""

    def __init__(self, shape_element: ShapeElement):
        """Initialize Math with shape element."""
        self._shape_element = shape_element
        self._omml_element = None

    @property
    def omml_element(self) -> CT_OMath:
        """Get or create the underlying OMML element."""
        if self._omml_element is None:
            # Create OMML element and add it to the shape
            self._omml_element = OxmlElement("m:oMath")
            # Add the OMML element to the shape's content
            # This will need to be implemented based on the shape structure
        return self._omml_element

    def get_omml(self) -> str:
        """Get OMML XML string."""
        from pptx.oxml.xmlchemy import serialize_for_reading
        return serialize_for_reading(self.omml_element)

    def set_omml(self, omml_xml: str):
        """Replace current OMML with new XML string."""
        from pptx.oxml import parse_xml
        new_element = parse_xml(omml_xml)
        if not isinstance(new_element, CT_OMath):
            raise ValueError("Root element must be m:oMath")

        # Replace the OMML element in the shape
        self._omml_element = new_element

    def add_omml(self, omml_xml: str):
        """Add OMML content to the equation."""
        from pptx.oxml import parse_xml
        new_element = parse_xml(omml_xml)
        if not isinstance(new_element, CT_OMath):
            raise ValueError("Root element must be m:oMath")

        # Add all children from the new element
        for child in new_element:
            self.omml_element.append(child)
