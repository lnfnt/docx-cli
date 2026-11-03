# Changelog

## 0.1.4

- Packaging: vendored shared security modules now ship as `docx_cli/_secure` — internal
  library names and shims removed from the public tree (same code, neutral packaging).

## 0.1.3

- Audit wave w2 fixes: unquoted INCLUDETEXT detection, block-level instrText scanning, themeTint handling, numeric date formats, fail-closed error envelopes; 526 tests.

## 0.1.2

- LICENSE — теперь полный канонический текст Apache-2.0 (предыдущий был пересказом; права и ограничения те же).

## 0.1.1

- PyPI install: `pip install docx-cli`.

## 0.1.0

- First public release: deterministic .docx builder from JSON specs and Markdown,
  layout validation against gost/court profiles, `mcp` mode (8 tools), nested lists
  and pipe tables (B-10), Supreme Court and SOYU layout profiles (T-05), injection
  scanner for incoming text, write-guard: mutating tools answer `needs_confirmation`
  until confirmed.
