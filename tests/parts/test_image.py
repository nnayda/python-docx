# pyright: reportPrivateUsage=false

"""Unit test suite for docx.parts.image module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.image.image import Image
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.package import Package
from docx.parts.image import ImagePart

from ..unitutil.file import test_file
from ..unitutil.mock import (
    ANY,
    FixtureRequest,
    Mock,
    initializer_mock,
    instance_mock,
    method_mock,
)


class DescribeImagePart:
    def it_is_used_by_PartFactory_to_construct_image_part(
        self,
        image_part_load_: Mock,
        partname_: Mock,
        blob_: Mock,
        package_: Mock,
        image_part_: Mock,
    ) -> None:
        content_type = CT.JPEG
        reltype = RT.IMAGE
        image_part_load_.return_value = image_part_

        part = PartFactory(
            cast(PackURI, partname_),
            content_type,
            reltype,
            cast(bytes, blob_),
            cast(Package, package_),
        )

        image_part_load_.assert_called_once_with(partname_, content_type, blob_, package_)
        assert part is image_part_

    def it_can_construct_from_an_Image_instance(
        self, image_: Mock, partname_: Mock, _init_: Mock
    ) -> None:
        image_part = ImagePart.from_image(cast(Image, image_), cast(PackURI, partname_))

        _init_.assert_called_once_with(ANY, partname_, image_.content_type, image_.blob, image_)
        assert isinstance(image_part, ImagePart)

    def it_knows_its_default_dimensions_in_EMU(
        self, dimensions_fixture: tuple[ImagePart, int, int]
    ) -> None:
        image_part, cx, cy = dimensions_fixture
        assert image_part.default_cx == cx
        assert image_part.default_cy == cy

    def it_knows_its_filename(self, filename_fixture: tuple[ImagePart, str]) -> None:
        image_part, expected_filename = filename_fixture
        assert image_part.filename == expected_filename

    def it_knows_the_sha1_of_its_image(self) -> None:
        blob = b"fO0Bar"
        image_part = ImagePart(cast(PackURI, None), cast(str, None), blob)
        assert image_part.sha1 == "4921e7002ddfba690a937d54bda226a7b8bdeb68"

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def blob_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, str)

    @pytest.fixture(params=["loaded", "new"])
    def dimensions_fixture(self, request: FixtureRequest) -> tuple[ImagePart, int, int]:
        image_file_path = test_file("monty-truth.png")
        image = Image.from_file(image_file_path)
        expected_cx, expected_cy = 1905000, 2717800

        # case 1: image part is loaded by PartFactory w/no Image inst
        if request.param == "loaded":
            partname = PackURI("/word/media/image1.png")
            content_type = CT.PNG
            image_part = ImagePart.load(partname, content_type, image.blob, cast(Package, None))
        # case 2: image part is newly created from image file
        else:
            image_part = ImagePart.from_image(image, cast(PackURI, None))

        return image_part, expected_cx, expected_cy

    @pytest.fixture(params=["loaded", "new"])
    def filename_fixture(self, request: FixtureRequest, image_: Mock) -> tuple[ImagePart, str]:
        partname = PackURI("/word/media/image666.png")
        if request.param == "loaded":
            image_part = ImagePart(partname, cast(str, None), cast(bytes, None), None)
            expected_filename = "image.png"
        else:
            image_.filename = "foobar.PXG"
            image_part = ImagePart(partname, cast(str, None), cast(bytes, None), image_)
            expected_filename = cast(str, image_.filename)
        return image_part, expected_filename

    @pytest.fixture
    def image_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Image)

    @pytest.fixture
    def _init_(self, request: FixtureRequest) -> Mock:
        return initializer_mock(request, ImagePart)

    @pytest.fixture
    def image_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def image_part_load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, ImagePart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)

    @pytest.fixture
    def partname_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, PackURI)
