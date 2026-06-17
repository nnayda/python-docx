# Security Policy

## Supported versions

This is a fork; security fixes are applied to the latest released version only.

## Reporting a vulnerability

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, use GitHub's private vulnerability reporting:
[Report a vulnerability](https://github.com/nnayda/python-docx/security/advisories/new).

Please include:

- A description of the issue and its impact.
- Steps to reproduce (a minimal `.docx` and script if applicable).
- The version affected.

`python-docx` parses untrusted `.docx` files (ZIP archives containing XML), so
the most relevant classes of issue are XML parsing attacks (e.g. entity
expansion / XXE) and ZIP handling. When reporting, please redact any sensitive
content from sample files.

We aim to acknowledge reports within a few days and will coordinate disclosure
with you.
