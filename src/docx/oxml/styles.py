"""Custom element classes related to the styles part."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.simpletypes import ST_DecimalNumber, ST_OnOff, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.shared import CT_DecimalNumber, CT_OnOff, CT_String
    from docx.oxml.text.font import CT_RPr
    from docx.oxml.text.parfmt import CT_PPr


def styleId_from_name(name: str) -> str:
    """Return the style id corresponding to `name`, taking into account special-case
    names such as 'Heading 1'."""
    return {
        "caption": "Caption",
        "heading 1": "Heading1",
        "heading 2": "Heading2",
        "heading 3": "Heading3",
        "heading 4": "Heading4",
        "heading 5": "Heading5",
        "heading 6": "Heading6",
        "heading 7": "Heading7",
        "heading 8": "Heading8",
        "heading 9": "Heading9",
    }.get(name, name.replace(" ", ""))


class CT_LatentStyles(BaseOxmlElement):
    """`w:latentStyles` element, defining behavior defaults for latent styles and
    containing `w:lsdException` child elements that each override those defaults for a
    named latent style."""

    add_lsdException: Callable[[], CT_LsdException]
    lsdException_lst: List[CT_LsdException]

    lsdException = ZeroOrMore("w:lsdException", successors=())

    count: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:count", ST_DecimalNumber
    )
    defLockedState: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:defLockedState", ST_OnOff
    )
    defQFormat: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:defQFormat", ST_OnOff
    )
    defSemiHidden: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:defSemiHidden", ST_OnOff
    )
    defUIPriority: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:defUIPriority", ST_DecimalNumber
    )
    defUnhideWhenUsed: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:defUnhideWhenUsed", ST_OnOff
    )

    def bool_prop(self, attr_name: str) -> bool:
        """Return the boolean value of the attribute having `attr_name`, or |False| if
        not present."""
        value = getattr(self, attr_name)
        if value is None:
            return False
        return value

    def get_by_name(self, name: str) -> CT_LsdException | None:
        """Return the `w:lsdException` child having `name`, or |None| if not found."""
        found = self.xpath('w:lsdException[@w:name="%s"]' % name)
        if not found:
            return None
        return found[0]

    def set_bool_prop(self, attr_name: str, value: bool) -> None:
        """Set the on/off attribute having `attr_name` to `value`."""
        setattr(self, attr_name, bool(value))


class CT_LsdException(BaseOxmlElement):
    """``<w:lsdException>`` element, defining override visibility behaviors for a named
    latent style."""

    locked: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:locked", ST_OnOff
    )
    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]
    qFormat: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:qFormat", ST_OnOff
    )
    semiHidden: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:semiHidden", ST_OnOff
    )
    uiPriority: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:uiPriority", ST_DecimalNumber
    )
    unhideWhenUsed: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:unhideWhenUsed", ST_OnOff
    )

    def delete(self) -> None:
        """Remove this `w:lsdException` element from the XML document."""
        parent = self.getparent()
        assert parent is not None
        parent.remove(self)

    def on_off_prop(self, attr_name: str) -> bool | None:
        """Return the boolean value of the attribute having `attr_name`, or |None| if
        not present."""
        return getattr(self, attr_name)

    def set_on_off_prop(self, attr_name: str, value: bool | None) -> None:
        """Set the on/off attribute having `attr_name` to `value`."""
        setattr(self, attr_name, value)


class CT_Style(BaseOxmlElement):
    """A ``<w:style>`` element, representing a style definition."""

    _tag_seq = (
        "w:name",
        "w:aliases",
        "w:basedOn",
        "w:next",
        "w:link",
        "w:autoRedefine",
        "w:hidden",
        "w:uiPriority",
        "w:semiHidden",
        "w:unhideWhenUsed",
        "w:qFormat",
        "w:locked",
        "w:personal",
        "w:personalCompose",
        "w:personalReply",
        "w:rsid",
        "w:pPr",
        "w:rPr",
        "w:tblPr",
        "w:trPr",
        "w:tcPr",
        "w:tblStylePr",
    )
    get_or_add_basedOn: Callable[[], CT_String]
    get_or_add_next: Callable[[], CT_String]
    _add_locked: Callable[[], CT_OnOff]
    _add_name: Callable[[], CT_String]
    _add_qFormat: Callable[[], CT_OnOff]
    _add_semiHidden: Callable[[], CT_OnOff]
    _add_uiPriority: Callable[[], CT_DecimalNumber]
    _add_unhideWhenUsed: Callable[[], CT_OnOff]
    _remove_basedOn: Callable[[], None]
    _remove_next: Callable[[], None]
    _remove_locked: Callable[[], None]
    _remove_name: Callable[[], None]
    _remove_qFormat: Callable[[], None]
    _remove_semiHidden: Callable[[], None]
    _remove_uiPriority: Callable[[], None]
    _remove_unhideWhenUsed: Callable[[], None]

    name: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:name", successors=_tag_seq[1:]
    )
    basedOn: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:basedOn", successors=_tag_seq[3:]
    )
    next: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:next", successors=_tag_seq[4:]
    )
    uiPriority: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:uiPriority", successors=_tag_seq[8:]
    )
    semiHidden: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:semiHidden", successors=_tag_seq[9:]
    )
    unhideWhenUsed: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:unhideWhenUsed", successors=_tag_seq[10:]
    )
    qFormat: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:qFormat", successors=_tag_seq[11:]
    )
    locked: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:locked", successors=_tag_seq[12:]
    )
    pPr: CT_PPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:pPr", successors=_tag_seq[17:]
    )
    rPr: CT_RPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rPr", successors=_tag_seq[18:]
    )
    del _tag_seq

    type: WD_STYLE_TYPE | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:type", WD_STYLE_TYPE
    )
    styleId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:styleId", ST_String
    )
    default: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:default", ST_OnOff
    )
    customStyle: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:customStyle", ST_OnOff
    )

    @property
    def basedOn_val(self) -> str | None:
        """Value of `w:basedOn/@w:val` or |None| if not present."""
        basedOn = self.basedOn
        if basedOn is None:
            return None
        return basedOn.val

    @basedOn_val.setter
    def basedOn_val(self, value: str | None):
        if value is None:
            self._remove_basedOn()
        else:
            self.get_or_add_basedOn().val = value

    @property
    def base_style(self) -> CT_Style | None:
        """Sibling CT_Style element this style is based on or |None| if no base style or
        base style not found."""
        basedOn = self.basedOn
        if basedOn is None:
            return None
        styles = self.getparent()
        assert isinstance(styles, CT_Styles)
        base_style = styles.get_by_id(basedOn.val)
        if base_style is None:
            return None
        return base_style

    def delete(self) -> None:
        """Remove this `w:style` element from its parent `w:styles` element."""
        parent = self.getparent()
        assert parent is not None
        parent.remove(self)

    @property
    def locked_val(self) -> bool:
        """Value of `w:locked/@w:val` or |False| if not present."""
        locked = self.locked
        if locked is None:
            return False
        return locked.val

    @locked_val.setter
    def locked_val(self, value: bool):
        self._remove_locked()
        if bool(value) is True:
            locked = self._add_locked()
            locked.val = value

    @property
    def name_val(self) -> str | None:
        """Value of ``<w:name>`` child or |None| if not present."""
        name = self.name
        if name is None:
            return None
        return name.val

    @name_val.setter
    def name_val(self, value: str | None):
        self._remove_name()
        if value is not None:
            name = self._add_name()
            name.val = value

    @property
    def next_style(self) -> CT_Style | None:
        """Sibling CT_Style element identified by the value of `w:name/@w:val` or |None|
        if no value is present or no style with that style id is found."""
        next = self.next
        if next is None:
            return None
        styles = self.getparent()
        assert isinstance(styles, CT_Styles)
        return styles.get_by_id(next.val)  # None if not found

    @property
    def qFormat_val(self) -> bool:
        """Value of `w:qFormat/@w:val` or |False| if not present."""
        qFormat = self.qFormat
        if qFormat is None:
            return False
        return qFormat.val

    @qFormat_val.setter
    def qFormat_val(self, value: bool):
        self._remove_qFormat()
        if bool(value):
            self._add_qFormat()

    @property
    def semiHidden_val(self) -> bool:
        """Value of ``<w:semiHidden>`` child or |False| if not present."""
        semiHidden = self.semiHidden
        if semiHidden is None:
            return False
        return semiHidden.val

    @semiHidden_val.setter
    def semiHidden_val(self, value: bool):
        self._remove_semiHidden()
        if bool(value) is True:
            semiHidden = self._add_semiHidden()
            semiHidden.val = value

    @property
    def uiPriority_val(self) -> int | None:
        """Value of ``<w:uiPriority>`` child or |None| if not present."""
        uiPriority = self.uiPriority
        if uiPriority is None:
            return None
        return uiPriority.val

    @uiPriority_val.setter
    def uiPriority_val(self, value: int | None):
        self._remove_uiPriority()
        if value is not None:
            uiPriority = self._add_uiPriority()
            uiPriority.val = value

    @property
    def unhideWhenUsed_val(self) -> bool:
        """Value of `w:unhideWhenUsed/@w:val` or |False| if not present."""
        unhideWhenUsed = self.unhideWhenUsed
        if unhideWhenUsed is None:
            return False
        return unhideWhenUsed.val

    @unhideWhenUsed_val.setter
    def unhideWhenUsed_val(self, value: bool):
        self._remove_unhideWhenUsed()
        if bool(value) is True:
            unhideWhenUsed = self._add_unhideWhenUsed()
            unhideWhenUsed.val = value


class CT_Styles(BaseOxmlElement):
    """``<w:styles>`` element, the root element of a styles part, i.e. styles.xml."""

    add_style: Callable[[], CT_Style]
    get_or_add_latentStyles: Callable[[], CT_LatentStyles]
    latentStyles_lst: List[CT_LatentStyles]
    style_lst: List[CT_Style]

    _tag_seq = ("w:docDefaults", "w:latentStyles", "w:style")
    latentStyles: CT_LatentStyles | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:latentStyles", successors=_tag_seq[2:]
    )
    style = ZeroOrMore("w:style", successors=())
    del _tag_seq

    def add_style_of_type(self, name: str, style_type: WD_STYLE_TYPE, builtin: bool) -> CT_Style:
        """Return a newly added `w:style` element having `name` and `style_type`.

        `w:style/@customStyle` is set based on the value of `builtin`.
        """
        style = self.add_style()
        style.type = style_type
        style.customStyle = None if builtin else True
        style.styleId = styleId_from_name(name)
        style.name_val = name
        return style

    def default_for(self, style_type: WD_STYLE_TYPE) -> CT_Style | None:
        """Return `w:style[@w:type="*{style_type}*][-1]` or |None| if not found."""
        default_styles_for_type = [
            s for s in self._iter_styles() if s.type == style_type and s.default
        ]
        if not default_styles_for_type:
            return None
        # spec calls for last default in document order
        return default_styles_for_type[-1]

    def get_by_id(self, styleId: str) -> CT_Style | None:
        """`w:style` child where @styleId = `styleId`.

        |None| if not found.
        """
        xpath = f'w:style[@w:styleId="{styleId}"]'
        return next(iter(self.xpath(xpath)), None)

    def get_by_name(self, name: str) -> CT_Style | None:
        """`w:style` child with `w:name` grandchild having value `name`.

        |None| if not found.
        """
        xpath = 'w:style[w:name/@w:val="%s"]' % name
        return next(iter(self.xpath(xpath)), None)

    def _iter_styles(self) -> Iterator[CT_Style]:
        """Generate each of the `w:style` child elements in document order."""
        return (style for style in self.xpath("w:style"))
