"""Provides a general interface to a `physical` OPC package, such as a zip file."""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING, cast
from zipfile import ZIP_DEFLATED, ZipFile, is_zipfile

from docx.opc.exceptions import PackageNotFoundError
from docx.opc.packuri import CONTENT_TYPES_URI

if TYPE_CHECKING:
    from docx.opc.packuri import PackURI


class PhysPkgReader:
    """Factory for physical package reader objects."""

    def __new__(cls, pkg_file: str | IO[bytes]) -> PhysPkgReader:
        # if `pkg_file` is a string, treat it as a path
        if isinstance(pkg_file, str):
            if os.path.isdir(pkg_file):
                reader_cls = _DirPkgReader
            elif is_zipfile(pkg_file):
                reader_cls = _ZipPkgReader
            else:
                raise PackageNotFoundError("Package not found at '%s'" % pkg_file)
        else:  # assume it's a stream and pass it to Zip reader to sort out
            reader_cls = _ZipPkgReader

        return object.__new__(reader_cls)

    def blob_for(self, pack_uri: PackURI) -> bytes:
        """Return the blob (binary contents) of the package item at `pack_uri`."""
        raise NotImplementedError  # pragma: no cover

    def close(self) -> None:
        """Release any resources held by this reader."""
        raise NotImplementedError  # pragma: no cover

    @property
    def content_types_xml(self) -> bytes:
        """Return the `[Content_Types].xml` blob from the package."""
        raise NotImplementedError  # pragma: no cover

    def rels_xml_for(self, source_uri: PackURI) -> bytes | None:
        """Return rels item XML for source with `source_uri`, or None if the item has no
        rels item."""
        raise NotImplementedError  # pragma: no cover


class PhysPkgWriter:
    """Factory for physical package writer objects."""

    def __new__(cls, pkg_file: str | IO[bytes]) -> PhysPkgWriter:
        return object.__new__(_ZipPkgWriter)

    def close(self) -> None:
        """Release any resources held by this writer."""
        raise NotImplementedError  # pragma: no cover

    def write(self, pack_uri: PackURI, blob: bytes) -> None:
        """Write `blob` to this package at `pack_uri`."""
        raise NotImplementedError  # pragma: no cover


class _DirPkgReader(PhysPkgReader):
    """Implements |PhysPkgReader| interface for an OPC package extracted into a
    directory."""

    def __init__(self, pkg_file: str | IO[bytes]):
        """`pkg_file` is the path to a directory containing an expanded package."""
        super(_DirPkgReader, self).__init__()
        # -- this reader is only ever created by the factory when `pkg_file` is a
        # -- directory path (a `str`); narrow it here. --
        self._path = os.path.abspath(cast(str, pkg_file))

    def blob_for(self, pack_uri: PackURI) -> bytes:
        """Return contents of file corresponding to `pack_uri` in package directory."""
        path = os.path.join(self._path, pack_uri.membername)
        with open(path, "rb") as f:
            blob = f.read()
        return blob

    def close(self) -> None:
        """Provides interface consistency with |ZipFileSystem|, but does nothing, a
        directory file system doesn't need closing."""
        pass

    @property
    def content_types_xml(self) -> bytes:
        """Return the `[Content_Types].xml` blob from the package."""
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri: PackURI) -> bytes | None:
        """Return rels item XML for source with `source_uri`, or None if the item has no
        rels item."""
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except IOError:
            rels_xml = None
        return rels_xml


class _ZipPkgReader(PhysPkgReader):
    """Implements |PhysPkgReader| interface for a zip file OPC package."""

    def __init__(self, pkg_file: str | IO[bytes]):
        super(_ZipPkgReader, self).__init__()
        self._zipf = ZipFile(pkg_file, "r")

    def blob_for(self, pack_uri: PackURI) -> bytes:
        """Return blob corresponding to `pack_uri`.

        Raises |ValueError| if no matching member is present in zip archive.
        """
        return self._zipf.read(pack_uri.membername)

    def close(self) -> None:
        """Close the zip archive, releasing any resources it is using."""
        self._zipf.close()

    @property
    def content_types_xml(self) -> bytes:
        """Return the `[Content_Types].xml` blob from the zip package."""
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri: PackURI) -> bytes | None:
        """Return rels item XML for source with `source_uri` or None if no rels item is
        present."""
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except KeyError:
            rels_xml = None
        return rels_xml


class _ZipPkgWriter(PhysPkgWriter):
    """Implements |PhysPkgWriter| interface for a zip file OPC package."""

    def __init__(self, pkg_file: str | IO[bytes]):
        super(_ZipPkgWriter, self).__init__()
        self._zipf = ZipFile(pkg_file, "w", compression=ZIP_DEFLATED)

    def close(self) -> None:
        """Close the zip archive, flushing any pending physical writes and releasing any
        resources it's using."""
        self._zipf.close()

    def write(self, pack_uri: PackURI, blob: bytes) -> None:
        """Write `blob` to this zip package with the membername corresponding to
        `pack_uri`."""
        self._zipf.writestr(pack_uri.membername, blob)
