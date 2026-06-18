# Changelog

All notable changes to this fork are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

For the release history inherited from upstream
[python-docx](https://github.com/python-openxml/python-docx), see
[HISTORY.rst](HISTORY.rst).

## [Unreleased]

### Added

- Fork packaging and project infrastructure: published to PyPI as
  `python-docx-nnayda` (import name remains `docx`).
- `AGENTS.md` / `CLAUDE.md`, `CONTRIBUTING.md`, `SECURITY.md`,
  `CODE_OF_CONDUCT.md`, and issue/PR templates.
- GitHub Actions: CI (ruff, pyright, pytest, behave), Conventional Commits
  PR-title check, CodeQL, docs publishing to GitHub Pages, and PyPI release via
  trusted publishing.
- CodeRabbit review configuration (`.coderabbit.yaml`).
- Dependabot for pip and GitHub Actions updates.

### Changed

- Documentation now builds with modern Sphinx and the Furo theme and is
  published to GitHub Pages.
- The full source tree and test suite now pass `pyright` in strict mode with
  zero errors. The CI `typecheck` job is now blocking (previously
  non-blocking), so new type errors fail the build. Minimum `pyright` bumped to
  `1.1.410`.

### Fixed

- `Document.add_comment` now normalizes a `None` `text` argument to an empty
  string instead of forwarding `None`.
- `OpcPackage.walk_parts` no longer uses a mutable default argument for its
  visited-parts accumulator.
- `BaseStyle` property accessors read from the always-present style element so
  they no longer mis-handle the post-`delete()` state.
- `tests/oxml/unitdata/numbering.py` `CT_NumBuilder.__attrs__` is now the
  intended one-tuple `("w:numId",)` rather than a bare string.
- Assigning `None` to `_Cell.bg_color` (or the underlying `CT_Tc`/`CT_TcPr`
  `bg_color`) now clears cell shading instead of raising `ValueError`, matching
  the getter, which already returns `None` for an unshaded cell.
