# pyright: reportPrivateUsage=false

"""Test data builders shared by more than one other module."""

from __future__ import annotations

from typing_extensions import Self

from ...unitdata import BaseBuilder


class CT_OnOffBuilder(BaseBuilder):
    __nspfxs__ = ("w",)
    __attrs__ = ("w:val",)

    def __init__(self, tag: str) -> None:
        self.__tag__ = tag
        super(CT_OnOffBuilder, self).__init__()

    def with_val(self, value: object) -> Self:
        self._set_xmlattr("w:val", str(value))
        return self


class CT_StringBuilder(BaseBuilder):
    __nspfxs__ = ("w",)
    __attrs__ = ()

    def __init__(self, tag: str) -> None:
        self.__tag__ = tag
        super(CT_StringBuilder, self).__init__()

    def with_val(self, value: object) -> Self:
        self._set_xmlattr("w:val", str(value))
        return self
