# ar — Arabic golden cases

P2 status: RTL pages no longer degrade to bitmaps. The pipeline covers
pdf2docx issue #73 (whole-string reversal), digit-run direction inside RTL
text, presentation-form NFKC folding and bracket mirroring — asserted by
`tests/rtl.test.ts` (unit, hand-built visual-order chars) and the
`integration: ar/he RTL` describe in `tests/integration.test.ts` (end-to-end
through PDFium with a system font).

No PDF binaries are committed yet: generating one requires embedding an
Arabic-capable font, and the fonts available locally (macOS Arial Unicode)
are not redistribution-safe. The end-to-end cases therefore build their
fixtures on the fly and skip when no suitable system font exists. A committed
sample can land here once a freely-licensed Arabic font (e.g. Amiri, SIL OFL)
is added to the fixture toolchain.
