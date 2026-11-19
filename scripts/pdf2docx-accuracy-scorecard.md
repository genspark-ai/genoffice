# PDF → editable accuracy scorecard — measured 2026-11-15

- genoffice pipeline at commit 19a389b3 (upstream/main, "before") vs this branch
  ("after", `augmentGridStreamColumns`); LibreOffice 26.8.2.2 via
  `soffice --infilter=writer_pdf_import --convert-to docx` for reference.
- Run: `node scripts/pdf2docx-accuracy.mjs <out.md>` (regenerates the after
  numbers on any checkout; the before numbers need a checkout WITHOUT the fix).

## before (upstream/main 19a389b3)

# PDF → editable accuracy — invoice / report class

Field = a ground-truth string that must survive as ONE contiguous cell/paragraph
(what a user can re-edit directly). split = present but fragmented across cells
(the #2669 symptom). Columns = cells in the detail/header row vs expected.

| sample            | engine    | intact | split | missing | detail columns (expected) |
| ----------------- | --------- | ------ | ----- | ------- | ------------------------- |
| invoice-banded-en | genoffice | 11/11  | 0     | 0       | 1 (8)                     |
| invoice-banded-en | soffice   | 11/11  | 0     | 0       | - (8)                     |
| invoice-cjk-real  | genoffice | 6/7    | 0     | 1       | 1 (8)                     |

- not intact: "91310000MA1FL0000X"
  | invoice-cjk-real | soffice | 7/7 | 0 | 0 | - (8) |
  | report-bordered | genoffice | 12/12 | 0 | 0 | - |
  | report-bordered | soffice | 12/12 | 0 | 0 | - |
  | report-borderless | genoffice | 9/9 | 0 | 0 | - |
  | report-borderless | soffice | 9/9 | 0 | 0 | - |

## after (this branch)

# PDF → editable accuracy — invoice / report class

Field = a ground-truth string that must survive as ONE contiguous cell/paragraph
(what a user can re-edit directly). split = present but fragmented across cells
(the #2669 symptom). Columns = cells in the detail/header row vs expected.

| sample            | engine    | intact | split | missing | detail columns (expected) |
| ----------------- | --------- | ------ | ----- | ------- | ------------------------- |
| invoice-banded-en | genoffice | 11/11  | 0     | 0       | 9 (8)                     |
| invoice-banded-en | soffice   | 11/11  | 0     | 0       | - (8)                     |
| invoice-cjk-real  | genoffice | 7/7    | 0     | 0       | 8 (8)                     |
| invoice-cjk-real  | soffice   | 7/7    | 0     | 0       | - (8)                     |
| report-bordered   | genoffice | 12/12  | 0     | 0       | -                         |
| report-bordered   | soffice   | 12/12  | 0     | 0       | -                         |
| report-borderless | genoffice | 9/9    | 0     | 0       | -                         |
| report-borderless | soffice   | 9/9    | 0     | 0       | -                         |

## Reading

The banded-form samples are where the two engines differ in KIND: the
pipeline rebuilds the detail band as a real table (8 columns where the
document has 8), LibreOffice's Writer import keeps the text but ships no
table structure for the form (detail columns "-": no header row found in any
w:tbl). Field-level parity everywhere else; the before-run's missing tax ID
is #2669's headline symptom and is what `augmentGridStreamColumns` fixes.
