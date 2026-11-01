# @genoffice/doc2docx

Legacy Word 97-2003 (`.doc`) import, staged so every layer can be reviewed
and tested on its own. The converter reads the OLE2/CFB container locally and
rebuilds a `.docx` through the shared engine; no upload or external service is
involved.

## Stage 1 (this change): text foundation

Internal only — nothing is user-visible yet. This stage establishes the
binary-format foundation the later stages build on:

- **OLE2/CFB container reader**: validated wrapper over `cfb` with the OLE2
  signature/header checks, 256 MiB input and 128 MiB per-stream caps, and the
  `WordDocument` stream required.
- **FIB parser**: FibBase/FibRgW97/FibRgLw97/FibRgFcLcb97, the nFib 0x00C1–0x0112
  gate, encrypted-document refusal, `0Table`/`1Table` selection, and the CLX
  plus stage-2/3 table ranges.
- **CLX piece table**: Prc/Pcdt walk, Windows-1252 compressed and UTF-16LE
  pieces, CP monotonicity and per-piece WordDocument range validation.
- **Body text and paragraph boundaries**: Word control marks mapped to readable
  text (cell mark → tab, line/page break → newline, field instruction dropped
  with the result kept, Windows-1252 punctuation decoded).

API: `openLegacyDoc(bytes)` returns the FIB, piece table, streams and text;
`readLegacyDocText(bytes)` returns just `{ text, paragraphs }`.

## Roadmap

- **Stage 2**: CHPX/PAPX character and paragraph properties (bold, size,
  alignment, indents) layered on the piece table.
- **Stage 3**: stylesheet (STSH), tables (TAP) and images, then wire the
  converter into the docs import path (`.doc` → `.docx`) and the CLI.

## Format coverage (stage 1)

Word 97-2003 (nFib 0x00C1–0x0112), main document story only. Footnotes,
headers/footers, textboxes, tables, images and formatting are later stages.
Encrypted files and Word 6/95 documents are rejected with actionable
`LegacyDocError` codes.
