"""|FootnotesPart| and closely related objects."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.parser import parse_xml
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage


class FootnotesPart(StoryPart):
    """Proxy for the `word/footnotes.xml` part containing the document's footnotes.

    It is a story part because each footnote body is composed of paragraphs and runs,
    so it benefits from the `StoryPart` style-resolution and id helpers.
    """

    @classmethod
    def new(cls, package: OpcPackage) -> "FootnotesPart":
        """Return a newly-created footnotes part seeded with the separator and
        continuation-separator entries Word expects."""
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = parse_xml(cls._default_footnotes_xml())
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_footnotes_xml(cls):
        """Return bytes containing XML for a default footnotes part."""
        path = os.path.join(os.path.split(__file__)[0], "..", "templates", "default-footnotes.xml")
        with open(path, "rb") as f:
            return f.read()
