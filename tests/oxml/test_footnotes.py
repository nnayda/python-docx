# pyright: reportPrivateUsage=false

"""Test suite for the docx.oxml.footnotes module."""

from __future__ import annotations

from typing import cast

from docx.oxml.footnotes import CT_Footnotes, CT_FtnEdn, CT_FtnEdnRef

from ..unitutil.cxml import element


class DescribeCT_Footnotes:
    """Unit-test suite for the CT_Footnotes (<w:footnotes>) element."""

    def it_calculates_the_next_id_ignoring_separators(self) -> None:
        footnotes = cast(
            CT_Footnotes,
            element("w:footnotes/(w:footnote{w:id=-1},w:footnote{w:id=0},w:footnote{w:id=1})"),
        )
        assert footnotes._next_id == 2

    def it_starts_ids_at_one_when_only_separators_present(self) -> None:
        footnotes = cast(
            CT_Footnotes,
            element("w:footnotes/(w:footnote{w:id=-1},w:footnote{w:id=0})"),
        )
        assert footnotes._next_id == 1

    def it_can_add_a_footnote_with_a_given_id(self) -> None:
        footnotes = cast(CT_Footnotes, element("w:footnotes"))

        ftn = footnotes.add_footnote(3)

        assert isinstance(ftn, CT_FtnEdn)
        assert ftn.id == 3
        assert footnotes.footnote_lst[-1] is ftn


class DescribeCT_FtnEdn:
    """Unit-test suite for the CT_FtnEdn (<w:footnote>) element."""

    def it_knows_its_id(self) -> None:
        ftn = cast(CT_FtnEdn, element("w:footnote{w:id=5}"))
        assert ftn.id == 5

    def it_knows_its_type_when_present(self) -> None:
        ftn = cast(CT_FtnEdn, element("w:footnote{w:id=-1,w:type=separator}"))
        assert ftn.type == "separator"

    def it_returns_None_type_when_absent(self) -> None:
        ftn = cast(CT_FtnEdn, element("w:footnote{w:id=1}"))
        assert ftn.type is None


class DescribeCT_FtnEdnRef:
    """Unit-test suite for the CT_FtnEdnRef (<w:footnoteReference>) element."""

    def it_knows_its_id(self) -> None:
        ref = cast(CT_FtnEdnRef, element("w:footnoteReference{w:id=4}"))
        assert ref.id == 4
