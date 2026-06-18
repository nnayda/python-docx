"""Custom element classes related to footnotes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
)

if TYPE_CHECKING:
    from docx.oxml.text.paragraph import CT_P


class CT_Footnotes(BaseOxmlElement):
    """`<w:footnotes>` element, the root of the footnotes part."""

    _add_footnote: Callable[..., "CT_FtnEdn"]
    footnote_lst: List["CT_FtnEdn"]

    footnote = ZeroOrMore("w:footnote")

    def add_footnote(self, id: int) -> "CT_FtnEdn":
        """Return a newly-added `<w:footnote>` child having `w:id` of `id`.

        Overrides the generated `add_footnote()` (the `xmlchemy` metaclass skips its own
        version because this name is already defined on the class) to require the `id`
        attribute, which is mandatory on `w:footnote`. The private `_add_footnote(**attrs)`
        helper is still generated and used here.
        """
        return self._add_footnote(id=id)

    @property
    def _next_id(self) -> int:
        """Next available positive `w:id`, ignoring the separator (-1, 0) entries.

        Returns 1 when no user footnotes are present yet.
        """
        ids = [int(v) for v in self.xpath("./w:footnote/@w:id")]
        positive = [i for i in ids if i > 0]
        return max(positive) + 1 if positive else 1


class CT_FtnEdn(BaseOxmlElement):
    """`<w:footnote>` element, a single footnote.

    The same `CT_FtnEdn` complex type also backs the `<w:endnote>` element.
    """

    add_p: Callable[[], "CT_P"]
    p_lst: List["CT_P"]

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute("w:type", ST_String)  # pyright: ignore[reportAssignmentType]

    p = ZeroOrMore("w:p")


class CT_FtnEdnRef(BaseOxmlElement):
    """`<w:footnoteReference>` element, the in-text marker linking to a footnote."""

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]


class CT_Empty(BaseOxmlElement):
    """Generic class for empty elements (`w:footnoteRef`, `w:separator`,
    `w:continuationSeparator`) that carry no attributes or content."""
