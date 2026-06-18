"""Footnote-related proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx import types as t
from docx.oxml.parser import OxmlElement
from docx.shared import StoryChild
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    from docx.oxml.footnotes import CT_FtnEdn
    from docx.oxml.text.run import CT_R
    from docx.text.run import Run

#: Style id of the built-in character style applied to footnote reference marks.
FOOTNOTE_REFERENCE_STYLE_ID = "FootnoteReference"
#: Style id of the built-in paragraph style applied to footnote body paragraphs.
FOOTNOTE_TEXT_STYLE_ID = "FootnoteText"


class Footnote(StoryChild):
    """Proxy object wrapping a `<w:footnote>` element.

    A footnote is created via `Run.add_footnote()` or `Paragraph.add_footnote()` (or
    `DocumentPart.add_footnote()`); it is not constructed directly by callers.
    """

    def __init__(self, ftn: "CT_FtnEdn", parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._ftn = self._element = ftn

    def add_paragraph(self, text: str = "", style: str | None = None) -> Paragraph:
        """Append a paragraph to the footnote body and return it.

        Defaults to the built-in `FootnoteText` paragraph style.
        """
        p = self._ftn.add_p()
        paragraph = Paragraph(p, self._parent)
        # -- assign the style id directly at the element level; the proxy `.style`
        # -- setter does a name-based lookup that would emit a deprecation warning for a
        # -- style id whose UI name differs ("footnote text" vs id "FootnoteText"). --
        paragraph._p.style = (  # pyright: ignore[reportPrivateUsage]
            FOOTNOTE_TEXT_STYLE_ID if style is None else style
        )
        if text:
            paragraph.add_run(text)
        return paragraph

    def add_run(self, text: str = "", style: str | None = None) -> "Run":
        """Append a run containing `text` to the footnote's last paragraph.

        A first paragraph is created if the footnote body is empty.
        """
        p_lst = self._ftn.p_lst
        paragraph = Paragraph(p_lst[-1], self._parent) if p_lst else self.add_paragraph()
        return paragraph.add_run(text, style)

    @property
    def id(self) -> int:
        """The internal `w:id` link key for this footnote (not its displayed number)."""
        return self._ftn.id

    @property
    def text(self) -> str:
        """The concatenated text of the footnote body paragraphs."""
        return "".join(p.text for p in self._ftn.p_lst)

    @text.setter
    def text(self, text: str | None) -> None:
        for p in list(self._ftn.p_lst):
            self._ftn.remove(p)
        self._add_marked_paragraph(text)

    def _add_marked_paragraph(self, text: str | None) -> None:
        """Create the footnote's first paragraph, beginning with the in-note ref mark.

        Used at creation time and when replacing text. The mark is a run carrying the
        `FootnoteReference` character style and a `<w:footnoteRef/>` element.
        """
        paragraph = self.add_paragraph()
        paragraph._p.append(self._new_mark_run())  # pyright: ignore[reportPrivateUsage]
        if text:
            paragraph.add_run(text)

    @staticmethod
    def _new_mark_run() -> "CT_R":
        """Return a `<w:r>` containing the in-note footnote-reference mark."""
        r = cast("CT_R", OxmlElement("w:r"))
        r.style = FOOTNOTE_REFERENCE_STYLE_ID
        r.append(OxmlElement("w:footnoteRef"))
        return r


def new_footnote_reference_run(footnote_id: int) -> "CT_R":
    """Return a `<w:r>` containing a footnote-reference marker for `footnote_id`.

    The run carries the `FootnoteReference` character style and a
    `<w:footnoteReference w:id=.../>` child. Shared by `Run.add_footnote` and
    `Paragraph.add_footnote`.
    """
    r = cast("CT_R", OxmlElement("w:r"))
    r.style = FOOTNOTE_REFERENCE_STYLE_ID
    ref = r.add_footnoteReference()
    ref.id = footnote_id
    return r
