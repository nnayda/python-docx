# pyright: reportPrivateUsage=false

"""Test suite for the docx.oxml.styles module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.styles import CT_Styles

from ..unitutil.cxml import element, xml
from ..unitutil.mock import FixtureRequest


class DescribeCT_Styles:
    def it_can_add_a_style_of_type(
        self, add_fixture: tuple[CT_Styles, str, WD_STYLE_TYPE, bool, str]
    ) -> None:
        styles, name, style_type, builtin, expected_xml = add_fixture
        style = styles.add_style_of_type(name, style_type, builtin)
        assert styles.xml == expected_xml
        assert style is styles[-1]

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (
                "w:styles",
                "Foo Bar",
                WD_STYLE_TYPE.LIST,
                False,
                "w:styles/w:style{w:type=numbering,w:customStyle=1,w:styleId=FooBar"
                "}/w:name{w:val=Foo Bar}",
            ),
            (
                "w:styles",
                "heading 1",
                WD_STYLE_TYPE.PARAGRAPH,
                True,
                "w:styles/w:style{w:type=paragraph,w:styleId=Heading1}/w:name{w:val=heading 1}",
            ),
        ]
    )
    def add_fixture(
        self, request: FixtureRequest
    ) -> tuple[CT_Styles, str, WD_STYLE_TYPE, bool, str]:
        styles_cxml, name, style_type, builtin, expected_cxml = cast(
            "tuple[str, str, WD_STYLE_TYPE, bool, str]", request.param
        )
        styles = cast(CT_Styles, element(styles_cxml))
        expected_xml = xml(expected_cxml)
        return styles, name, style_type, builtin, expected_xml
