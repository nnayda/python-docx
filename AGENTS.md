# AGENTS.md

`python-docx-nnayda` — a maintained fork of [python-docx](https://github.com/python-openxml/python-docx),
a Python library for creating, reading, and updating Microsoft Word `.docx`
files. The import name is still `docx`; only the PyPI distribution name differs
from upstream.

## Commands

This project uses [uv](https://docs.astral.sh/uv/) for environment and
dependency management.

```sh
uv sync                       # create .venv and install the project + dev group
uv run pytest                 # unit tests
uv run behave                 # acceptance/BDD tests (features/)
uv run ruff check .           # lint
uv run ruff format --check .  # formatting check (drop --check to auto-format)
uv run pyright                # static type checking (strict)
uv run tox                    # full test matrix across supported Pythons
```

Full pre-push gauntlet (mirrors CI):

```sh
uv run ruff check . && uv run ruff format --check . && uv run pyright && uv run pytest && uv run behave
```

Build the package and the docs:

```sh
uv build                                                   # sdist + wheel into dist/
uv run --with-requirements requirements-docs.txt \
  sphinx-build -b html docs docs/.build/html               # render docs locally
```

## Layout

- `src/docx/` — the library (import name `docx`). The public entry point is
  `docx.Document`, wired up in `src/docx/__init__.py`.
- `src/docx/oxml/` — lxml-backed objects mapping 1:1 to WordprocessingML XML
  elements. Everything above this layer manipulates these custom elements
  rather than raw lxml.
- `tests/` — pytest unit tests, mirroring the `src/docx/` tree.
- `features/` — `behave` Gherkin acceptance tests with `.docx` fixtures.
- `docs/` — Sphinx documentation (reStructuredText), published to GitHub Pages.
- `typings/` — local type stubs consumed by pyright.

## The one rule

The object layer (`docx.oxml`) owns all XML knowledge. Code outside `docx.oxml`
should manipulate the custom element objects, never reach into lxml or hardcode
XML strings/namespaces. If a change makes a high-level API class
namespace-aware or starts assembling raw XML outside the oxml layer, it's going
the wrong direction.

## Conventions

- **Versioning:** the single source of truth is `__version__` in
  `src/docx/__init__.py`; `pyproject.toml` reads it dynamically. Bump it there.
- **Commits/PRs:** PR titles follow [Conventional Commits](https://www.conventionalcommits.org/)
  (CI-enforced). PRs are **squash-merged**, so the PR title becomes the commit
  message on `main`. Keep one logical change per PR.
- **Style:** `ruff` is the arbiter for both lint and formatting — no manual
  style debates. `pyright` runs in strict mode; keep it green.
- **Tests:** add or update unit tests (`tests/`) for changed behavior; add a
  `behave` feature for user-visible capabilities.
- **Test naming:** pytest discovers `Test`/`Describe` classes and
  `it_`/`its_`/`they_`/`and_`/`but_` functions (see `[tool.pytest.ini_options]`).
- **Staying mergeable with upstream:** prefer changes that don't gratuitously
  diverge from `python-openxml/python-docx`, so upstream fixes keep merging
  cleanly. Fork-specific metadata lives in `pyproject.toml`, `README.md`, and
  `CHANGELOG.md`.

## Working with upstream

The `upstream` remote tracks `python-openxml/python-docx`. To pull in upstream
changes, merge `upstream/master` and resolve conflicts, keeping the fork's
packaging metadata.
