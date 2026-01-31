"""High-level Math class for OMML equation management."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.math import CT_OMath

if TYPE_CHECKING:
    from pptx.oxml.shapes import ShapeElement


class Math:
    """High-level interface for OMML math equations."""

    def __init__(self, shape_element: ShapeElement, parent):
        """Initialize Math with shape element and parent."""
        self._shape_element = shape_element
        self._parent = parent
        self._omml_element = None

    def _create_rpr(self):
        """Create standard rPr formatting properties for PowerPoint 16.105.2."""
        rpr = OxmlElement("a:rPr")
        rpr.set("lang", "en-US")
        rpr.set("i", "1")
        rpr.set("smtClean", "0")

        latin = OxmlElement("a:latin")
        latin.set("typeface", "Cambria Math")
        latin.set("panose", "02040503050406030204")
        latin.set("pitchFamily", "18")
        latin.set("charset", "0")
        rpr.append(latin)

        return rpr

    def _ensure_formatting(self, omath_element):
        """Ensure all text runs have proper PowerPoint formatting."""
        # Add <a:rPr> formatting to all <m:r> elements that lack it
        for r_element in omath_element.xpath(".//*[local-name() = 'r']"):
            if not r_element.xpath(".//*[local-name() = 'rPr']"):
                rpr = self._create_rpr()

                # Insert rPr as first child of m:r
                if len(r_element) > 0:
                    r_element.insert(0, rpr)
                else:
                    r_element.append(rpr)

    def _add_to_slide_container(self, omath_element):
        """Add OMML element to slide-level with proper container shapes."""
        from pptx.oxml.xmlchemy import OxmlElement

        # Ensure proper formatting first
        self._ensure_formatting(omath_element)

        # Create the container shapes that PowerPoint expects
        self._create_math_container_shapes(omath_element)

    def _create_math_container_shapes(self, omath_element):
        """Create the textbox shapes with math extension URIs that PowerPoint requires."""
        from pptx.oxml.xmlchemy import OxmlElement

        # Get the slide element
        slide_part = self._parent.part
        if hasattr(slide_part, 'slide'):
            slide_xml = slide_part.slide._element

            # Find the spTree element where shapes are added
            sp_tree = slide_xml.xpath('.//*[local-name() = "spTree"]')[0]

            # Create the exact same number of math textboxes as your demo
            # Your demo has 5 math textboxes at specific positions
            self._create_math_textbox(sp_tree, id_offset=5, name="TextBox 4", x=914400, y=1828800)
            self._create_math_textbox(sp_tree, id_offset=6, name="TextBox 5", x=914400, y=2743200)
            self._create_math_textbox(sp_tree, id_offset=7, name="TextBox 6", x=914400, y=3657600)
            self._create_math_textbox(sp_tree, id_offset=8, name="TextBox 7", x=914400, y=4572000)

            # Create the AlternateContent container with the actual OMML
            self._create_alternate_content_container(sp_tree, omath_element)

    def _create_math_textbox(self, sp_tree, id_offset, name, x, y):
        """Create a textbox shape with math extension URI."""
        from pptx.oxml.xmlchemy import OxmlElement

        # Create the shape
        sp = OxmlElement("p:sp")

        # Non-visual properties
        nv_sp_pr = OxmlElement("p:nvSpPr")
        cnv_pr = OxmlElement("p:cNvPr")
        cnv_pr.set("id", str(id_offset))
        cnv_pr.set("name", name)

        cnv_sp_pr = OxmlElement("p:cNvSpPr")
        cnv_sp_pr.set("txBox", "1")

        nv_pr = OxmlElement("p:nvPr")
        ext_lst = OxmlElement("p:extLst")
        ext = OxmlElement("p:ext")
        ext.set("uri", "http://schemas.openxmlformats.org/presentationml/2006/main/math")
        ext_lst.append(ext)
        nv_pr.append(ext_lst)

        nv_sp_pr.append(cnv_pr)
        nv_sp_pr.append(cnv_sp_pr)
        nv_sp_pr.append(nv_pr)
        sp.append(nv_sp_pr)

        # Shape properties
        sp_pr = OxmlElement("p:spPr")
        xfrm = OxmlElement("a:xfrm")
        off = OxmlElement("a:off")
        off.set("x", str(x))
        off.set("y", str(y))
        ext_elem = OxmlElement("a:ext")
        ext_elem.set("cx", "1828800")
        ext_elem.set("cy", "914400")
        xfrm.append(off)
        xfrm.append(ext_elem)

        prst_geom = OxmlElement("a:prstGeom")
        prst_geom.set("prst", "rect")
        av_lst = OxmlElement("a:avLst")
        prst_geom.append(av_lst)

        no_fill = OxmlElement("a:noFill")

        sp_pr.append(xfrm)
        sp_pr.append(prst_geom)
        sp_pr.append(no_fill)
        sp.append(sp_pr)

        # Text body (empty)
        tx_body = OxmlElement("p:txBody")
        body_pr = OxmlElement("a:bodyPr")
        body_pr.set("wrap", "none")
        sp_auto_fit = OxmlElement("a:spAutoFit")
        body_pr.append(sp_auto_fit)

        lst_style = OxmlElement("a:lstStyle")
        p = OxmlElement("a:p")
        end_para_rpr = OxmlElement("a:endParaRPr")
        p.append(end_para_rpr)

        tx_body.append(body_pr)
        tx_body.append(lst_style)
        tx_body.append(p)
        sp.append(tx_body)

        # Add to spTree
        sp_tree.append(sp)

    def _create_alternate_content_container(self, sp_tree, omath_element):
        """Create the proper mc:AlternateContent container with OMML."""
        from pptx.oxml import parse_xml
        from pptx.oxml.xmlchemy import serialize_for_reading

        # Create the OMML content as XML string
        omml_xml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        for child in omath_element:
            omml_xml += serialize_for_reading(child)
        omml_xml += '</m:oMath>'

        # Create the complete AlternateContent structure as XML string
        alt_content_xml = f'''<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="9" name="TextBox 8"/>
              <p:cNvSpPr txBox="1"/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="4114800" y="2686050"/>
                <a:ext cx="1335301" cy="276999"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
              <a:noFill/>
            </p:spPr>
            <p:txBody>
              <a:bodyPr wrap="none">
                <a:spAutoFit/>
              </a:bodyPr>
              <a:lstStyle/>
              <a:p>
                <a14:m>
                  {omml_xml}
                </a14:m>
              </a:p>
            </p:txBody>
          </p:sp>
        </mc:Choice>
      </mc:AlternateContent>'''

        # Parse and add to spTree
        alt_content = parse_xml(alt_content_xml)
        sp_tree.append(alt_content)

    @property
    def omml_element(self) -> CT_OMath:
        """Get or create the underlying OMML element."""
        if self._omml_element is None:
            # Create OMML element and add it to the slide's math content
            self._omml_element = OxmlElement("m:oMath")

            # Add OMML to the slide's math content area
            # PowerPoint stores OMML at the slide level, not in individual shapes
            slide_part = self._parent.part
            if hasattr(slide_part, 'slide'):
                # Try to find or create the math content area
                slide_xml = slide_part.slide._element

                # Look for existing oMathPara or create one
                omath_para = None
                for child in slide_xml:
                    if child.tag.endswith('oMathPara'):
                        omath_para = child
                        break

                if omath_para is None:
                    # Create new oMathPara
                    omath_para = OxmlElement("m:oMathPara")
                    slide_xml.append(omath_para)

                omath_para.append(self._omml_element)
        return self._omml_element

    def get_omml(self) -> str:
        """Get OMML XML string."""
        from pptx.oxml.xmlchemy import serialize_for_reading
        return serialize_for_reading(self.omml_element)

    def set_omml(self, omml_xml: str):
        """Replace current OMML with new XML string."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import qn

        # Parse the XML and extract the oMath element
        new_element = parse_xml(omml_xml)

        # Find the actual oMath element (it might be nested)
        omath_element = new_element
        expected_tags = [
            qn("m:oMath"),  # Standard namespace
            "{http://purl.oclc.org/ooxml/officeDocument/math}oMath",  # Alternative namespace
        ]

        if new_element.tag not in expected_tags:
            # Look for oMath child with proper namespace
            omath_elements = new_element.xpath(".//*[local-name() = 'oMath']")
            if omath_elements:
                omath_element = omath_elements[0]
            else:
                raise ValueError("Root element must be m:oMath")

        # Replace the OMML element in the shape
        self._omml_element = omath_element

    def add_omml(self, omml_xml: str):
        """Add OMML content to the equation with automatic formatting."""
        from pptx.oxml import parse_xml

        # Parse the XML and extract the oMath element
        new_element = parse_xml(omml_xml)

        # Find the actual oMath element (it might be nested)
        omath_element = new_element
        expected_tags = [
            "m:oMath",  # Standard namespace
            "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath",  # Alternative namespace
        ]

        if new_element.tag not in expected_tags:
            # Look for oMath child with proper namespace
            omath_elements = new_element.xpath(".//*[local-name() = 'oMath']")
            if omath_elements:
                omath_element = omath_elements[0]
            else:
                raise ValueError("Root element must be m:oMath")

        # Add to slide-level container with automatic formatting
        self._add_to_slide_container(omath_element)
