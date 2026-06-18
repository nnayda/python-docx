# pyright: reportPrivateUsage=false

"""Test suite for docx.image.bmp module."""

from __future__ import annotations

import io
from typing import cast

import pytest

from docx.image.bmp import Bmp
from docx.image.constants import MIME_TYPE

from ..unitutil.mock import ANY, FixtureRequest, Mock, initializer_mock


class DescribeBmp:
    def it_can_construct_from_a_bmp_stream(self, Bmp__init__: Mock):
        cx, cy, horz_dpi, vert_dpi = 26, 43, 200, 96
        bytes_ = (
            b"fillerfillerfiller\x1a\x00\x00\x00\x2b\x00\x00\x00"
            b"fillerfiller\xb8\x1e\x00\x00\x00\x00\x00\x00"
        )
        stream = io.BytesIO(bytes_)

        bmp = Bmp.from_stream(stream)

        Bmp__init__.assert_called_once_with(ANY, cx, cy, horz_dpi, vert_dpi)
        assert isinstance(bmp, Bmp)

    def it_knows_its_content_type(self):
        bmp = Bmp(cast(int, None), cast(int, None), cast(int, None), cast(int, None))
        assert bmp.content_type == MIME_TYPE.BMP

    def it_knows_its_default_ext(self):
        bmp = Bmp(cast(int, None), cast(int, None), cast(int, None), cast(int, None))
        assert bmp.default_ext == "bmp"

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def Bmp__init__(self, request: FixtureRequest):
        return initializer_mock(request, Bmp)
