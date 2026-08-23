# Golden regression corpus

Per-language golden PDFs and their expected conversion snapshots. Layout:

```
golden/
  en/  Latin text: word-spacing thresholds, hyphenation, alignment cases
  zh/  Simplified Chinese: no inserted spaces, eastAsia font slot, mixed zh/en runs
  ja/  Japanese: kana/kanji mixing; vertical-text pages must degrade to bitmap
  ko/  Korean: real hangul inter-word spaces preserved
  ar/  Arabic: P1 expectation = page degrades to bitmap (no garbled text); P2 flips to real RTL output
```

Each case is a pair:

- `<name>.pdf` — the input document (small, redistribution-safe)
- `<name>.expected.json` — assertions: paragraph texts in order, per-paragraph
  align/indent, image count, expected `warnings`

Conventions:

- Negative examples from upstream pdf2docx issues (lost/multiplied spaces #103,
  RTL reversal #73) belong here as they get ported — name them `issue-<n>-*`.
- Keep PDFs generated (see `tests/helpers/fixtures.ts`) or freely licensed.
- P1 ships the structure with on-the-fly fixtures; committed samples grow with
  each regression we fix. A runner that walks these directories lands with the
  first committed sample.
