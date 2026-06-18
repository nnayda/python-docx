# pyright: reportPrivateUsage=false

"""Unit test suite for docx.image package"""

from __future__ import annotations

import io
from typing import cast

import pytest

from docx.image.bmp import Bmp
from docx.image.exceptions import UnrecognizedImageError
from docx.image.gif import Gif
from docx.image.image import BaseImageHeader, Image, _ImageHeaderFactory
from docx.image.jpeg import Exif, Jfif
from docx.image.png import Png
from docx.image.tiff import Tiff
from docx.opc.constants import CONTENT_TYPE as CT
from docx.shared import Emu, Length

from ..unitutil.file import test_file
from ..unitutil.mock import (
    ANY,
    FixtureRequest,
    Mock,
    class_mock,
    function_mock,
    initializer_mock,
    instance_mock,
    method_mock,
    property_mock,
)


class DescribeImage:
    def it_can_construct_from_an_image_blob(
        self, blob_: Mock, BytesIO_: Mock, _from_stream_: Mock, stream_: Mock, image_: Mock
    ):
        image = Image.from_blob(blob_)

        BytesIO_.assert_called_once_with(blob_)
        _from_stream_.assert_called_once_with(stream_, blob_)
        assert image is image_

    def it_can_construct_from_an_image_path(
        self, from_path_fixture: tuple[str, Mock, Mock, bytes, str, Mock]
    ):
        image_path, _from_stream_, stream_, blob, filename, image_ = from_path_fixture
        image = Image.from_file(image_path)
        _from_stream_.assert_called_once_with(stream_, blob, filename)
        assert image is image_

    def it_can_construct_from_an_image_file_like(
        self, from_filelike_fixture: tuple[io.BytesIO, Mock, bytes, Mock]
    ):
        image_stream, _from_stream_, blob, image_ = from_filelike_fixture
        image = Image.from_file(image_stream)
        _from_stream_.assert_called_once_with(image_stream, blob, None)
        assert image is image_

    def it_can_construct_from_an_image_stream(
        self, from_stream_fixture: tuple[Mock, Mock, str | None, Mock, Mock, Mock, str]
    ):
        stream_, blob_, filename_in = from_stream_fixture[:3]
        _ImageHeaderFactory_, image_header_ = from_stream_fixture[3:5]
        Image__init_, filename_out = from_stream_fixture[5:]

        image = Image._from_stream(stream_, blob_, filename_in)

        _ImageHeaderFactory_.assert_called_once_with(stream_)
        Image__init_.assert_called_once_with(ANY, blob_, filename_out, image_header_)
        assert isinstance(image, Image)

    def it_provides_access_to_the_image_blob(self):
        blob = b"foobar"
        image = Image(blob, cast(str, None), cast(BaseImageHeader, None))
        assert image.blob == blob

    def it_knows_the_image_content_type(self, content_type_fixture: tuple[Mock, str]):
        image_header_, content_type = content_type_fixture
        image = Image(cast(bytes, None), cast(str, None), image_header_)
        assert image.content_type == content_type

    def it_knows_the_image_px_dimensions(self, dimensions_fixture: tuple[Mock, int, int]):
        image_header_, px_width, px_height = dimensions_fixture
        image = Image(cast(bytes, None), cast(str, None), image_header_)
        assert image.px_width == px_width
        assert image.px_height == px_height

    def it_knows_the_horz_and_vert_dpi_of_the_image(self, dpi_fixture: tuple[Mock, int, int]):
        image_header_, horz_dpi, vert_dpi = dpi_fixture
        image = Image(cast(bytes, None), cast(str, None), image_header_)
        assert image.horz_dpi == horz_dpi
        assert image.vert_dpi == vert_dpi

    def it_knows_the_image_native_size(self, size_fixture: tuple[Image, int, int]):
        image, width, height = size_fixture
        assert (image.width, image.height) == (width, height)
        assert isinstance(image.width, Length)
        assert isinstance(image.height, Length)

    def it_can_scale_its_dimensions(
        self, scale_fixture: tuple[Image, int | None, int | None, tuple[int, int]]
    ):
        image, width, height, expected_value = scale_fixture

        scaled_width, scaled_height = image.scaled_dimensions(width, height)

        assert (scaled_width, scaled_height) == expected_value
        assert isinstance(scaled_width, Length)
        assert isinstance(scaled_height, Length)

    def it_knows_the_image_filename(self):
        filename = "foobar.png"
        image = Image(cast(bytes, None), filename, cast(BaseImageHeader, None))
        assert image.filename == filename

    def it_knows_the_image_filename_extension(self):
        image = Image(cast(bytes, None), "foobar.png", cast(BaseImageHeader, None))
        assert image.ext == "png"

    def it_knows_the_sha1_of_its_image(self):
        blob = b"fO0Bar"
        image = Image(blob, cast(str, None), cast(BaseImageHeader, None))
        assert image.sha1 == "4921e7002ddfba690a937d54bda226a7b8bdeb68"

    def it_correctly_characterizes_known_images(
        self, known_image_fixture: tuple[str, tuple[str, str, int, int, int, int]]
    ):
        image_path, characteristics = known_image_fixture
        ext, content_type, px_width, px_height, horz_dpi, vert_dpi = characteristics
        with open(test_file(image_path), "rb") as stream:
            image = Image.from_file(stream)
            assert image.content_type == content_type
            assert image.ext == ext
            assert image.px_width == px_width
            assert image.px_height == px_height
            assert image.horz_dpi == horz_dpi
            assert image.vert_dpi == vert_dpi

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def content_type_fixture(self, image_header_: Mock) -> tuple[Mock, str]:
        content_type = "image/foobar"
        image_header_.content_type = content_type
        return image_header_, content_type

    @pytest.fixture
    def dimensions_fixture(self, image_header_: Mock) -> tuple[Mock, int, int]:
        px_width, px_height = 111, 222
        image_header_.px_width = px_width
        image_header_.px_height = px_height
        return image_header_, px_width, px_height

    @pytest.fixture
    def dpi_fixture(self, image_header_: Mock) -> tuple[Mock, int, int]:
        horz_dpi, vert_dpi = 333, 444
        image_header_.horz_dpi = horz_dpi
        image_header_.vert_dpi = vert_dpi
        return image_header_, horz_dpi, vert_dpi

    @pytest.fixture
    def from_filelike_fixture(
        self, _from_stream_: Mock, image_: Mock
    ) -> tuple[io.BytesIO, Mock, bytes, Mock]:
        image_path = test_file("python-icon.png")
        with open(image_path, "rb") as f:
            blob = f.read()
        image_stream = io.BytesIO(blob)
        return image_stream, _from_stream_, blob, image_

    @pytest.fixture
    def from_path_fixture(
        self, _from_stream_: Mock, BytesIO_: Mock, stream_: Mock, image_: Mock
    ) -> tuple[str, Mock, Mock, bytes, str, Mock]:
        filename = "python-icon.png"
        image_path = test_file(filename)
        with open(image_path, "rb") as f:
            blob = f.read()
        return image_path, _from_stream_, stream_, blob, filename, image_

    @pytest.fixture(params=["foobar.png", None])
    def from_stream_fixture(
        self,
        request: FixtureRequest,
        stream_: Mock,
        blob_: Mock,
        _ImageHeaderFactory_: Mock,
        image_header_: Mock,
        Image__init_: Mock,
    ) -> tuple[Mock, Mock, str | None, Mock, Mock, Mock, str]:
        filename_in = cast("str | None", request.param)
        filename_out = "image.png" if filename_in is None else filename_in
        return (
            stream_,
            blob_,
            filename_in,
            _ImageHeaderFactory_,
            image_header_,
            Image__init_,
            filename_out,
        )

    @pytest.fixture(params=[0, 1, 2, 3, 4, 5, 6, 7, 8])
    def known_image_fixture(
        self, request: FixtureRequest
    ) -> tuple[str, tuple[str, str, int, int, int, int]]:
        cases: tuple[tuple[str, tuple[str, str, int, int, int, int]], ...] = (
            ("python.bmp", ("bmp", CT.BMP, 211, 71, 96, 96)),
            ("sonic.gif", ("gif", CT.GIF, 290, 360, 72, 72)),
            ("python-icon.jpeg", ("jpg", CT.JPEG, 204, 204, 72, 72)),
            ("300-dpi.jpg", ("jpg", CT.JPEG, 1504, 1936, 300, 300)),
            ("monty-truth.png", ("png", CT.PNG, 150, 214, 72, 72)),
            ("150-dpi.png", ("png", CT.PNG, 901, 1350, 150, 150)),
            ("300-dpi.png", ("png", CT.PNG, 860, 579, 300, 300)),
            ("72-dpi.tiff", ("tiff", CT.TIFF, 48, 48, 72, 72)),
            ("300-dpi.TIF", ("tiff", CT.TIFF, 2464, 3248, 300, 300)),
            # ('CVS_LOGO.WMF',     ('wmf',  CT.X_WMF, 149,   59,  72,  72)),
        )
        image_filename, characteristics = cases[cast(int, request.param)]
        return image_filename, characteristics

    @pytest.fixture(
        params=[
            (None, None, 1000, 2000),
            (100, None, 100, 200),
            (None, 500, 250, 500),
            (1500, 1500, 1500, 1500),
        ]
    )
    def scale_fixture(
        self, request: FixtureRequest, width_prop_: Mock, height_prop_: Mock
    ) -> tuple[Image, int | None, int | None, tuple[int, int]]:
        width, height, scaled_width, scaled_height = cast(
            "tuple[int | None, int | None, int, int]", request.param
        )
        width_prop_.return_value = Emu(1000)
        height_prop_.return_value = Emu(2000)
        image = Image(cast(bytes, None), cast(str, None), cast(BaseImageHeader, None))
        return image, width, height, (scaled_width, scaled_height)

    @pytest.fixture
    def size_fixture(self, image_header_: Mock) -> tuple[Image, int, int]:
        image_header_.px_width, image_header_.px_height = 150, 75
        image_header_.horz_dpi, image_header_.vert_dpi = 72, 200
        image = Image(cast(bytes, None), cast(str, None), image_header_)
        return image, 1905000, 342900

    # fixture components ---------------------------------------------

    @pytest.fixture
    def blob_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, bytes)

    @pytest.fixture
    def BytesIO_(self, request: FixtureRequest, stream_: Mock) -> Mock:
        return class_mock(request, "docx.image.image.io.BytesIO", return_value=stream_)

    @pytest.fixture
    def filename_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, str)

    @pytest.fixture
    def _from_stream_(self, request: FixtureRequest, image_: Mock) -> Mock:
        return method_mock(request, Image, "_from_stream", autospec=False, return_value=image_)

    @pytest.fixture
    def height_prop_(self, request: FixtureRequest) -> Mock:
        return property_mock(request, Image, "height")

    @pytest.fixture
    def image_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Image)

    @pytest.fixture
    def _ImageHeaderFactory_(self, request: FixtureRequest, image_header_: Mock) -> Mock:
        return function_mock(
            request, "docx.image.image._ImageHeaderFactory", return_value=image_header_
        )

    @pytest.fixture
    def image_header_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, BaseImageHeader, default_ext="png")

    @pytest.fixture
    def Image__init_(self, request: FixtureRequest) -> Mock:
        return initializer_mock(request, Image)

    @pytest.fixture
    def stream_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, io.BytesIO)

    @pytest.fixture
    def width_prop_(self, request: FixtureRequest) -> Mock:
        return property_mock(request, Image, "width")


class Describe_ImageHeaderFactory:
    def it_constructs_the_right_class_for_a_given_image_stream(
        self, call_fixture: tuple[io.BytesIO, type[BaseImageHeader]]
    ):
        stream, expected_class = call_fixture
        image_header = _ImageHeaderFactory(stream)
        assert isinstance(image_header, expected_class)

    def it_raises_on_unrecognized_image_stream(self):
        stream = io.BytesIO(b"foobar 666 not an image stream")
        with pytest.raises(UnrecognizedImageError):
            _ImageHeaderFactory(stream)

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("python-icon.png", Png),
            ("python-icon.jpeg", Jfif),
            ("exif-420-dpi.jpg", Exif),
            ("sonic.gif", Gif),
            ("72-dpi.tiff", Tiff),
            ("little-endian.tif", Tiff),
            ("python.bmp", Bmp),
        ]
    )
    def call_fixture(self, request: FixtureRequest) -> tuple[io.BytesIO, type[BaseImageHeader]]:
        image_filename, expected_class = cast("tuple[str, type[BaseImageHeader]]", request.param)
        image_path = test_file(image_filename)
        with open(image_path, "rb") as f:
            blob = f.read()
        image_stream = io.BytesIO(blob)
        image_stream.seek(666)
        return image_stream, expected_class


class DescribeBaseImageHeader:
    def it_defines_content_type_as_an_abstract_property(self):
        base_image_header = BaseImageHeader(
            cast(int, None), cast(int, None), cast(int, None), cast(int, None)
        )
        with pytest.raises(NotImplementedError):
            base_image_header.content_type

    def it_defines_default_ext_as_an_abstract_property(self):
        base_image_header = BaseImageHeader(
            cast(int, None), cast(int, None), cast(int, None), cast(int, None)
        )
        with pytest.raises(NotImplementedError):
            base_image_header.default_ext

    def it_knows_the_image_dimensions(self):
        px_width, px_height = 42, 24
        image_header = BaseImageHeader(px_width, px_height, cast(int, None), cast(int, None))
        assert image_header.px_width == px_width
        assert image_header.px_height == px_height

    def it_knows_the_horz_and_vert_dpi_of_the_image(self):
        horz_dpi, vert_dpi = 42, 24
        image_header = BaseImageHeader(cast(int, None), cast(int, None), horz_dpi, vert_dpi)
        assert image_header.horz_dpi == horz_dpi
        assert image_header.vert_dpi == vert_dpi
