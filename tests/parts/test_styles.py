# pyright: reportPrivateUsage=false

"""Test suite for the docx.parts.styles module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.styles import CT_Styles
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.package import Package
from docx.parts.styles import StylesPart
from docx.styles.styles import Styles

from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock


class DescribeStylesPart:
    def it_provides_access_to_its_styles(
        self, styles_fixture: tuple[StylesPart, Mock, Mock]
    ) -> None:
        styles_part, Styles_, styles_ = styles_fixture
        styles = styles_part.styles
        Styles_.assert_called_once_with(styles_part.element)
        assert styles is styles_

    def it_can_construct_a_default_styles_part_to_help(self) -> None:
        package = Package()
        styles_part = StylesPart.default(package)
        assert isinstance(styles_part, StylesPart)
        assert styles_part.partname == "/word/styles.xml"
        assert styles_part.content_type == CT.WML_STYLES
        assert styles_part.package is package
        assert len(styles_part.element) == 6

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def styles_fixture(
        self, Styles_: Mock, styles_elm_: Mock, styles_: Mock
    ) -> tuple[StylesPart, Mock, Mock]:
        styles_part = StylesPart(
            cast(PackURI, None),
            cast(str, None),
            cast(BaseOxmlElement, styles_elm_),
            cast(Package, None),
        )
        return styles_part, Styles_, styles_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Styles_(self, request: FixtureRequest, styles_: Mock) -> Mock:
        return class_mock(request, "docx.parts.styles.Styles", return_value=styles_)

    @pytest.fixture
    def styles_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Styles)

    @pytest.fixture
    def styles_elm_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, CT_Styles)
