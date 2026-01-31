"""Custom element classes for OMML (Office Math Markup Language) elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    pass


class CT_OMath(BaseOxmlElement):
    """`m:oMath` custom element class - root element for a math equation."""

    add_r: Callable[[], "CT_R"]
    add_f: Callable[[], "CT_F"]
    add_sSup: Callable[[], "CT_SSup"]
    add_rad: Callable[[], "CT_Rad"]
    add_nary: Callable[[], "CT_Nary"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")
    rad: "ZeroOrMore" = ZeroOrMore("m:rad")
    nary: "ZeroOrMore" = ZeroOrMore("m:nary")


class CT_R(BaseOxmlElement):
    """`m:r` custom element class - math run (text container)."""

    add_t: Callable[[], "CT_T"]
    add_rPr: Callable[[], "CT_RPr"]

    t: "OneAndOnlyOne" = OneAndOnlyOne("m:t")
    rPr: "ZeroOrOne" = ZeroOrOne("m:rPr")


class CT_T(BaseOxmlElement):
    """`m:t` custom element class - text content."""

    @property
    def content(self) -> str:
        """Get text content."""
        return self.text or ""

    @content.setter
    def content(self, value: str):
        """Set text content."""
        # Use lxml's built-in text assignment
        object.__setattr__(self, 'text', value)


class CT_RPr(BaseOxmlElement):
    """`m:rPr` custom element class - run properties."""
    # Math run properties would be defined here
    pass


class CT_F(BaseOxmlElement):
    """`m:f` custom element class - fraction."""

    add_num: Callable[[], "CT_Num"]
    add_den: Callable[[], "CT_Den"]

    num: "OneAndOnlyOne" = OneAndOnlyOne("m:num")
    den: "OneAndOnlyOne" = OneAndOnlyOne("m:den")


class CT_Num(BaseOxmlElement):
    """`m:num` custom element class - fraction numerator."""

    add_r: Callable[[], "CT_R"]
    add_f: Callable[[], "CT_F"]
    add_sSup: Callable[[], "CT_SSup"]
    add_rad: Callable[[], "CT_Rad"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")
    rad: "ZeroOrMore" = ZeroOrMore("m:rad")


class CT_Den(BaseOxmlElement):
    """`m:den` custom element class - fraction denominator."""

    add_r: Callable[[], "CT_R"]
    add_f: Callable[[], "CT_F"]
    add_sSup: Callable[[], "CT_SSup"]
    add_rad: Callable[[], "CT_Rad"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")
    rad: "ZeroOrMore" = ZeroOrMore("m:rad")


class CT_SSup(BaseOxmlElement):
    """`m:sSup` custom element class - superscript."""

    add_e: Callable[[], "CT_E"]

    e: "OneAndOnlyOne" = OneAndOnlyOne("m:e")


class CT_E(BaseOxmlElement):
    """`m:e` custom element class - expression (base of superscript/subscript)."""

    add_r: Callable[[], "CT_R"]
    add_f: Callable[[], "CT_F"]
    add_sSup: Callable[[], "CT_SSup"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
    f: "ZeroOrMore" = ZeroOrMore("m:f")
    sSup: "ZeroOrMore" = ZeroOrMore("m:sSup")


class CT_Rad(BaseOxmlElement):
    """`m:rad` custom element class - radical (square root)."""

    add_radPr: Callable[[], "CT_RadPr"]
    add_deg: Callable[[], "CT_Deg"]
    add_e: Callable[[], "CT_E"]

    radPr: "ZeroOrOne" = ZeroOrOne("m:radPr")
    deg: "ZeroOrOne" = ZeroOrOne("m:deg")
    e: "OneAndOnlyOne" = OneAndOnlyOne("m:e")


class CT_RadPr(BaseOxmlElement):
    """`m:radPr` custom element class - radical properties."""
    # Radical properties would be defined here
    pass


class CT_Deg(BaseOxmlElement):
    """`m:deg` custom element class - radical degree (for nth roots)."""

    add_r: Callable[[], "CT_R"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")


class CT_Nary(BaseOxmlElement):
    """`m:nary` custom element class - n-ary operators (summation, integral, etc.)."""

    add_naryPr: Callable[[], "CT_NaryPr"]
    add_sub: Callable[[], "CT_Sub"]
    add_sup: Callable[[], "CT_Sup"]
    add_e: Callable[[], "CT_E"]

    naryPr: "ZeroOrOne" = ZeroOrOne("m:naryPr")
    sub: "ZeroOrOne" = ZeroOrOne("m:sub")
    sup: "ZeroOrOne" = ZeroOrOne("m:sup")
    e: "OneAndOnlyOne" = OneAndOnlyOne("m:e")


class CT_NaryPr(BaseOxmlElement):
    """`m:naryPr` custom element class - n-ary operator properties."""

    add_chr: Callable[[], "CT_Char"]
    add_limLoc: Callable[[], "CT_LimLoc"]

    chr: "ZeroOrOne" = ZeroOrOne("m:chr")
    limLoc: "ZeroOrOne" = ZeroOrOne("m:limLoc")


class CT_Char(BaseOxmlElement):
    """`m:chr` custom element class - character for n-ary operators."""
    pass


class CT_LimLoc(BaseOxmlElement):
    """`m:limLoc` custom element class - limit location for n-ary operators."""
    pass


class CT_Sub(BaseOxmlElement):
    """`m:sub` custom element class - subscript for n-ary operators."""

    add_r: Callable[[], "CT_R"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")


class CT_Sup(BaseOxmlElement):
    """`m:sup` custom element class - superscript for n-ary operators."""

    add_r: Callable[[], "CT_R"]

    r: "ZeroOrMore" = ZeroOrMore("m:r")
