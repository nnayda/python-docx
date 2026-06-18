# Contributing

Thanks for your interest! This is a maintained fork of
[python-docx](https://github.com/python-openxml/python-docx). This document
covers the practicalities of contributing to the fork.

## Before you start

- For anything beyond a small fix, **open an issue first** (or comment on an
  existing one) so we can agree on the approach before you invest time.
- If your change is a general improvement or bug fix that isn't fork-specific,
  consider sending it to [upstream](https://github.com/python-openxml/python-docx)
  too — it benefits everyone and keeps the fork easy to rebase.

## Development setup

This project uses [uv](https://docs.astral.sh/uv/).

```sh
uv sync          # create .venv and install the project + dev dependencies
uv run pytest    # run the unit tests
```

See [AGENTS.md](AGENTS.md) for the full command list and repository layout.

## Pull requests

- Keep PRs focused: one logical change per PR.
- PR titles must follow [Conventional Commits](https://www.conventionalcommits.org/)
  (`feat: …`, `fix: …`, `docs: …`, `chore: …`, `refactor: …`, `test: …`,
  `ci: …`, `perf: …`, `build: …`). CI enforces this; PRs are **squash-merged**,
  so the PR title becomes the commit message on `main`.
- Add or update tests for behavior you add or change.
- CI must be green before merge: `ruff` (lint + format), `pyright`, `pytest`,
  and `behave`.

Run everything locally before pushing:

```sh
uv run ruff check . && uv run ruff format --check . && uv run pyright && uv run pytest && uv run behave
```

## Code style

- **Linting & formatting:** [ruff](https://docs.astral.sh/ruff/) is the arbiter
  for both — `uv run ruff format .` to auto-format, `uv run ruff check --fix .`
  to auto-fix lint. No manual style debates.
- **Types:** [pyright](https://github.com/microsoft/pyright) runs in strict
  mode and must stay green.
- Keep the object layer (`docx.oxml`) as the only place that knows about XML —
  see the architecture note in [AGENTS.md](AGENTS.md).

## Releasing (maintainers)

1. Bump `__version__` in `src/docx/__init__.py`.
2. Update `CHANGELOG.md`.
3. Merge to `main`, then create a GitHub Release with tag `vX.Y.Z`.
4. The release workflow builds the sdist/wheel and publishes to PyPI via
   trusted publishing.

## Reporting bugs and requesting features

Use the issue templates. For security vulnerabilities, follow
[SECURITY.md](SECURITY.md) instead of opening a public issue.

## Code of conduct

This project follows the [Contributor Covenant](CODE_OF_CONDUCT.md). Be kind.

## Licensing of contributions

This project is licensed under the [MIT License](LICENSE). By submitting a
contribution you agree that it is licensed under the same terms
(inbound = outbound). There is no CLA.
