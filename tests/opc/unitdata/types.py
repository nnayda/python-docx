"""XML test data builders for [Content_Types].xml elements."""

from __future__ import annotations

from typing_extensions import Self

from docx.opc.oxml import nsmap

from ...unitdata import BaseBuilder


class CT_DefaultBuilder(BaseBuilder):
    __tag__ = "Default"
    __nspfxs__ = ("ct",)
    __attrs__ = ("Extension", "ContentType")


class CT_OverrideBuilder(BaseBuilder):
    __tag__ = "Override"
    __nspfxs__ = ("ct",)
    __attrs__ = ("PartName", "ContentType")


class CT_TypesBuilder(BaseBuilder):
    __tag__ = "Types"
    __nspfxs__ = ("ct",)
    __attrs__ = ()

    def with_nsdecls(self, *nspfxs: str) -> Self:
        self._nsdecls = ' xmlns="%s"' % nsmap["ct"]
        return self


def a_Default() -> CT_DefaultBuilder:
    return CT_DefaultBuilder()


def a_Types() -> CT_TypesBuilder:
    return CT_TypesBuilder()


def an_Override() -> CT_OverrideBuilder:
    return CT_OverrideBuilder()
