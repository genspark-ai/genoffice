# en — Latin golden cases

Target coverage: adaptive word-spacing thresholds (tight/loose tracking, large
type), end-of-line hyphenation joins, left/center/right alignment, first-line
indents. Port pdf2docx issue #103 (lost/extra word spaces) as `issue-103-*`.

P3 layout cases (fixtures generated on the fly by `tests/helpers/fixtures.ts`,
asserted end-to-end in `tests/integration.test.ts`):

- `buildStreamTablePdf` — borderless three-line (booktabs) table: 3 aligned
  rows framed by horizontal rules only → must become a real docx table.
- `buildTwoColumnPdf` — two balanced text columns → w:cols section, left
  column read before the right; must NEVER be detected as a table
  (pdf2docx issue #136, the "fake table" failure).
- `buildSpacedPdf` — low title with large whitespace above/below → the
  before_space chain restores it as paragraph spacingBefore.
- Negative family for the stream detector (plain paragraphs, poetry, TOC dot
  leaders, code blocks, 2-row label/value pairs without rules) lives in
  `tests/stream.test.ts` — port future upstream misfire reports there as
  `issue-<n>-*` cases.
