# Round-trip guarantee

File format fidelity is the core product promise: for changes touching
open/save paths (docx/xlsx/pptx), include a round-trip test proving untouched
content survives byte-for-byte. This note restates the
[CONTRIBUTING](../CONTRIBUTING.md) round-trip section and points at the
precedent test. No new behavior is promised here.

## Per-format promise

- **docx.** Open archives the original by hash and never touches it;
  docx-engine parses the top-level elements of `word/document.xml` into a
  block tree anchored by `docxIndex` plus the original XML slice; save turns
  dirty blocks into OOXML fragments and splices them into the original
  `document.xml`; the zip is repacked with all other entries copied
  byte-for-byte. Untouched blocks keep their original bytes.
- **sheets and slides.** The same philosophy holds as in docs: the original
  file is the source of truth, edits are applied as narrow patches, and
  everything the editor did not touch survives the round trip untouched.

Concretely: saving with no edits produces the exact original file
(byte identical). Saving with edits changes only the edited structures; every
other zip entry and every other top-level element keeps its bytes.

## How to write the byte-identical test

Follow [packages/docx-engine/tests/roundtrip.test.ts](../packages/docx-engine/tests/roundtrip.test.ts):

1. Build a representative fixture (the precedent uses the
   `buildKitchenSinkDocx` helper) and parse it with the engine (`parseDocx`).
2. Save with no edits by passing every visible block back as `original`
   (`SaveBlock` entries with `kind: 'original'` and their `docxIndex`).
3. Assert the saved bytes equal the input bytes exactly.
4. For the edited case, replace one block with a `generated` block, save, and
   compare zip entries before and after (the precedent reads both with JSZip):
   expect the same entry names, expect every entry except `word/document.xml`
   to equal its original bytes, and expect the new `document.xml` to still
   contain the `originalXml` slice of every untouched block.
5. Reparse the saved file and expect the same block count and shapes back.

Keep the new test next to the engine it covers (`packages/*/tests` or
`apps/*/tests`, vitest), in English only, and name it so the open/save path
it guards is obvious.
