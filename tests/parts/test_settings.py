# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.parts.settings module"""

from __future__ import annotations

from typing import cast

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.settings import CT_Settings
from docx.package import Package
from docx.parts.settings import SettingsPart
from docx.settings import Settings

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, method_mock


class DescribeSettingsPart:
    def it_is_used_by_loader_to_construct_settings_part(
        self, load_: Mock, package_: Mock, settings_part_: Mock
    ) -> None:
        partname, blob = PackURI("/word/settings.xml"), b"blob"
        content_type = CT.WML_SETTINGS
        load_.return_value = settings_part_

        part = PartFactory(partname, content_type, cast(str, None), blob, package_)

        load_.assert_called_once_with(partname, content_type, blob, package_)
        assert part is settings_part_

    def it_provides_access_to_its_settings(
        self, settings_fixture: tuple[SettingsPart, Mock, Mock]
    ) -> None:
        settings_part, Settings_, settings_ = settings_fixture
        settings = settings_part.settings
        Settings_.assert_called_once_with(settings_part.element)
        assert settings is settings_

    def it_constructs_a_default_settings_part_to_help(self) -> None:
        package = Package()
        settings_part = SettingsPart.default(package)
        assert isinstance(settings_part, SettingsPart)
        assert settings_part.partname == "/word/settings.xml"
        assert settings_part.content_type == CT.WML_SETTINGS
        assert settings_part.package is package
        assert len(settings_part.element) == 6

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def settings_fixture(self, Settings_: Mock, settings_: Mock) -> tuple[SettingsPart, Mock, Mock]:
        settings_elm = cast(CT_Settings, element("w:settings"))
        settings_part = SettingsPart(
            cast(PackURI, None), cast(str, None), settings_elm, cast(Package, None)
        )
        return settings_part, Settings_, settings_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, SettingsPart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)

    @pytest.fixture
    def Settings_(self, request: FixtureRequest, settings_: Mock) -> Mock:
        return class_mock(request, "docx.parts.settings.Settings", return_value=settings_)

    @pytest.fixture
    def settings_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Settings)

    @pytest.fixture
    def settings_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, SettingsPart)
