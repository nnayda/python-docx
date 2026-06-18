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
