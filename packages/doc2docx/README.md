# @genoffice/doc2docx

Legacy Word 97-2003 (`.doc`) import, staged so every layer can be reviewed
and tested on its own. The converter reads the OLE2/CFB container locally and
rebuilds a `.docx` through the shared engine; no upload or external service is
involved.

## Stage 1 — text foundation

Internal only. Establishes the binary-format foundation the later stages build
on:

- **OLE2/CFB container reader**: validated wrapper over `cfb` with the OLE2
  signature/header checks, 256 MiB input and 128 MiB per-stream caps, and the
  `WordDocument` stream required.
- **FIB parser**: FibBase/FibRgW97/FibRgLw97/FibRgFcLcb97, the nFib 0x00C1–0x0112
  gate, encrypted-document refusal, `0Table`/`1Table` selection, and the CLX
  plus formatting table ranges.
- **CLX piece table**: Prc/Pcdt walk, Windows-1252 compressed and UTF-16LE
  pieces, CP monotonicity and per-piece WordDocument range validation.
- **Body text and paragraph boundaries**: Word control marks mapped to readable
  text (cell mark → tab, line/page break → newline, field instruction dropped
  with the result kept, Windows-1252 punctuation decoded).

API: `openLegacyDoc(bytes)` returns the FIB, piece table, streams and text;
`readLegacyDocText(bytes)` returns just `{ text, paragraphs }`.

## Stage 2 — formatting model

Internal only. Reads the direct character and paragraph formatting of the body
story on top of the stage-1 piece table:

- **PlcBteChpx / PlcBtePapx** — the FC → FKP page tables in the table stream.
- **CHPX / PAPX FKPs** — run FC ranges, the short (`cb != 0`) and long
  (`cb == 0`) PapxInFkp forms, paragraph style index (istd) and default
  paragraphs (`bOffset == 0`).
- **Sprm parsing** — all spra operand-size classes, variable-length operands,
  unknown sprms skipped structurally, huge PAPX read through the Data stream.
- **Model** — `readLegacyDocModel(bytes)` returns
  `{ fib, text, paragraphs, warnings }`; each paragraph has
  `ParagraphFormatting` and `FormattedRun[]` with `CharacterFormatting`.
  Only explicit sprms are surfaced — style-sheet inheritance is stage 3b.

## Stage 3a — .docx writer (core conversion)

Internal only (no CLI/UI wiring yet): `convertLegacyDocToDocx(bytes)` returns
`{ docx, warnings }`. The writer builds the document on docx-engine's blank
template (`buildBlankDocx`) and emits one `<w:p>` per model paragraph through
`generateParagraphXml`:

- **runs**: bold/italic/strike, underline kinds → `w:u`, font size, superscript/
  subscript, highlight colours, the 16 Word colour indices → hex;
- **paragraphs**: alignment, left/right/first-line indents, spacing before/
  after, line spacing (multiple → `w:line` in 240ths, otherwise an exact
  rule), keep-lines / keep-next / page-break-before;
- **warnings** name every known gap honestly: unresolved styles (stage 3b),
  flattened tables, and footnotes/endnotes/headers/text boxes that are not
  converted yet.

## Roadmap

- **Stage 3b**: stylesheet (STSH) resolution, so style-based formatting is
  applied instead of only direct sprms; tighter paragraph-count parity.
- **Stage 3c**: tables (TAP), images, headers/footers, section geometry
  (page size/margins), lists and hyperlinks.
- **Stage 3d**: docs import wiring (`.doc` → `.docx` in File → Open /
  drag-and-drop and the CLI), as core + UI PRs.

## Format coverage

Word 97-2003 (nFib 0x00C1–0x0112), main document story. Footnotes,
headers/footers, textboxes, tables, images and style-sheet formatting are later
stages. Encrypted files and Word 6/95 documents are rejected with actionable
`LegacyDocError` codes (`not_cfb`, `truncated`, `bad_fib`,
`unsupported_version`, `encrypted`, `missing_stream`, `bad_clx`,
`bad_piece_table`, `bad_fkp`, `bad_sprm`, `limits_exceeded`, `corrupt`).
