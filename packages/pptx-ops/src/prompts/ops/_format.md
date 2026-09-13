# Op documentation format (authoring guide — not loaded into any prompt)

The `ops/*.md` files are the single source of truth for what the model knows
about canonical edit ops. `shared/op-docs.ts` parses them at import time into
`OP_DOCS`; `opVocabulary()` / `opUsage()` / `opSignatureIndex()` / `opGuide()`
are all derived from the parse. Tests in `apps/slides/tests/op-docs*.test.ts`
assert that the docs track the op registry exactly, that every example
validates against a fixture deck, and that the files contain no CJK text.

## File layout

```
# <Group title>
> <one-line summary shown in the load_guide catalog>

<optional group-level prose: addressing, units, conventions>

### <opName>
`<signature>`

<prose, field table, examples, common mistakes, related ops>

### <nextOpName> (not-ai-callable)
`<signature>`
...
```

- One file per group; the file name is the group id (`text`, `element`,
  `insert`, `table`, `slide`, `deck`).
- A block starts at `### <opName>`; `opName` must match a registered op.
- The first non-empty line after the heading is the **signature**, wrapped in
  single backticks. It is emitted verbatim as `Usage: <opName> <signature>` in
  guided errors, so keep it compact and free of backticks.
- Heading suffixes in parentheses set flags: `(not-ai-callable)` hides the op
  from the vocabulary and the signature index (byte / clipboard / part-path
  payloads the model cannot produce); `(pending)` hides it entirely until the
  registering branch lands. Combine as `(not-ai-callable, pending)`.
- Every ```json fenced block inside an AI-callable op block is an example that
  the test suite runs through `runTxn(dryRun)`. Use the placeholder ids below;
  the test substitutes real ids from its fixture deck.

## Placeholder ids for examples

| Placeholder  | Fixture element                 |
| ------------ | ------------------------------- |
| `e_TEXT`     | a text box on slide 0           |
| `e_SHAPE`    | a rounded rectangle on slide 0  |
| `e_PICTURE`  | a picture on slide 0            |
| `e_TABLE`    | a 2x2 table on slide 0          |
| `e_CHART`    | a bar chart on slide 0          |
| `e_LINE`     | a straight connector on slide 0 |
| `e_GROUP`    | a group on slide 0              |
| `e_CHILD`    | a direct child of that group    |
| `SECTION_ID` | an existing section GUID        |

The fixture deck has exactly two slides (indices 0 and 1; durable ids `s_1`
and `s_2`).

## Writing rules

- English only. Example text content (titles, labels, categories) is English too.
- Units are document-space EMU unless a field name says otherwise (`...Pt`,
  `...Pct`, `...Deg`). Say so in the field table.
- Field tables list every field the op's `validate` / `apply` reads, not only
  the ones in the signature.
- "Common mistakes" entries describe what the model tends to get wrong and the
  corrective action, in the same voice as the guided errors.
- Keep each group file under roughly 8 KB; split a group into sub-files (and
  add a catalog entry) rather than exceeding that.
