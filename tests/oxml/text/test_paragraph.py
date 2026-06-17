"""Test suite for the docx.oxml.text.paragraph module."""

from typing import cast

from docx.oxml.ns import qn
from docx.oxml.text.hyperlink import CT_Hyperlink
from docx.oxml.text.paragraph import CT_P

from ...unitutil.cxml import element


class DescribeCT_P:
    """Unit-test suite for the CT_P (<w:p>) element."""

    def it_can_add_a_hyperlink_carrying_an_external_relationship(self):
        p = cast(CT_P, element("w:p/w:r"))

        hyperlink = p.add_hyperlink("rId6")

        assert isinstance(hyperlink, CT_Hyperlink)
        assert hyperlink.rId == "rId6"
        assert p.hyperlink_lst == [hyperlink]
        # -- the new hyperlink is appended after the existing run --
        assert [child.tag for child in p] == [qn("w:r"), qn("w:hyperlink")]
