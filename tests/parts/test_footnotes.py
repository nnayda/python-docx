"""Test suite for the docx.parts.footnotes module."""

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.package import OpcPackage
from docx.oxml.footnotes import CT_Footnotes
from docx.parts.footnotes import FootnotesPart


class DescribeFootnotesPart:
    """Unit-test suite for FootnotesPart."""

    def it_is_created_with_the_correct_partname_and_content_type(self):
        package = OpcPackage()
        part = FootnotesPart.new(package)

        assert part.partname == "/word/footnotes.xml"
        assert part.content_type == CT.WML_FOOTNOTES
        assert part.package is package

    def it_seeds_separator_and_continuation_separator_entries(self):
        package = OpcPackage()
        part = FootnotesPart.new(package)

        footnotes = part.element
        assert isinstance(footnotes, CT_Footnotes)
        ids = [ftn.id for ftn in footnotes.footnote_lst]
        types = [ftn.type for ftn in footnotes.footnote_lst]
        assert ids == [-1, 0]
        assert types == ["separator", "continuationSeparator"]
