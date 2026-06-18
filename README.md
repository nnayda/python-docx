# python-docx-nnayda

[![CI](https://github.com/nnayda/python-docx/actions/workflows/ci.yml/badge.svg)](https://github.com/nnayda/python-docx/actions/workflows/ci.yml)
[![Docs](https://github.com/nnayda/python-docx/actions/workflows/docs.yml/badge.svg)](https://nnayda.github.io/python-docx/)
[![PyPI](https://img.shields.io/pypi/v/python-docx-nnayda.svg)](https://pypi.org/project/python-docx-nnayda/)
[![Python versions](https://img.shields.io/pypi/pyversions/python-docx-nnayda.svg)](https://pypi.org/project/python-docx-nnayda/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

*python-docx-nnayda* is a maintained fork of
[python-docx](https://github.com/python-openxml/python-docx) — a Python library
for reading, creating, and updating Microsoft Word 2007+ (.docx) files.

> **Fork note:** this package installs as `python-docx-nnayda` but the import
> name is unchanged (`import docx`), so it remains a drop-in replacement. It
> adds fork-specific features (footnotes, hyperlink creation, table styling)
> on top of upstream. Install only one of `python-docx` or
> `python-docx-nnayda` into an environment, not both.

## Installation

```
pip install python-docx-nnayda
```

## Example

```python
>>> from docx import Document

>>> document = Document()
>>> document.add_paragraph("It was a dark and stormy night.")
<docx.text.paragraph.Paragraph object at 0x10f19e760>
>>> document.save("dark-and-stormy.docx")

>>> document = Document("dark-and-stormy.docx")
>>> document.paragraphs[0].text
'It was a dark and stormy night.'
```

## Documentation

Documentation for this fork is published at
<https://nnayda.github.io/python-docx/>. The upstream documentation also remains
a useful reference at <https://python-docx.readthedocs.org/en/latest/>.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) and [AGENTS.md](AGENTS.md) for setup,
conventions, and the development workflow. This fork uses
[Conventional Commits](https://www.conventionalcommits.org/) for PR titles and
squash-merges to `main`.

## License

MIT — see [LICENSE](LICENSE). This fork preserves the upstream license and
authorship.
