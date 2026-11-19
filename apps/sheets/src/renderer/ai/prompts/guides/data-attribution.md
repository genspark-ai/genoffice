# Data source attribution guide (data-attribution)

When filling a sheet with external data (web search, scraping, APIs, etc.), you must leave traceable source citations. Implement them with `set_hyperlink` + `set_cell` + `format_range`.

## When attribution is required

| Data type | Required? | Recommended method |
|---|---|---|
| Financial data (SEC, earnings reports) | Required | Footer row or a dedicated Sources sheet |
| Web search results | Required | Source column |
| Data uploaded by the user | Not needed | — |
| Values derived by formulas/calculation | Not needed | — |
| Common knowledge (dates, unit conversions) | Not needed | — |

## Method 1: Source column

Add a "Source" column to the right of the data, with a hyperlink to the original data on each row. Suited to cases where each row has a different source.

```json
[
 {"op":"set_cell","sheetId":"s1","address":"G1","value":"Source"},
 {"op":"format_range","sheetId":"s1","range":"G1","format":{"bold":true}},
 {"op":"set_cell","sheetId":"s1","address":"G2","value":"SEC 10-K FY2024"},
 {"op":"set_hyperlink","sheetId":"s1","address":"G2","target":"https://sec.gov/..."},
 {"op":"format_range","sheetId":"s1","range":"G2:G20","format":{"fontColor":"#0563C1"}}
]
```

## Method 2: Footer source rows

Leave one blank row below the table, then write a "Sources:" section with one hyperlinked source per row. Suited to a whole table sharing a few common sources.

```json
[
 {"op":"set_cell","sheetId":"s1","address":"A23","value":"Sources:"},
 {"op":"format_range","sheetId":"s1","range":"A23","format":{"bold":true,"italic":true,"fontColor":"#666666"}},
 {"op":"set_cell","sheetId":"s1","address":"A24","value":"SEC EDGAR - NVIDIA 10-K FY2024"},
 {"op":"set_hyperlink","sheetId":"s1","address":"A24","target":"https://sec.gov/..."},
 {"op":"format_range","sheetId":"s1","range":"A24:A26","format":{"italic":true,"fontColor":"#0563C1"}}
]
```

## Method 3: Dedicated Sources sheet

When there are many sources and full traceability is needed, create a new "Sources" worksheet (structural operation, separate batch) with columns: Data point / Source / URL / Retrieved date.

```json
[
 {"op":"add_sheet","name":"Sources"}
]
```

After it applies, write into the new sheet:

```json
[
 {"op":"set_range","sheetId":"sSources","range":"A1:D1","values":[["Data point","Source","URL","Retrieved"]]},
 {"op":"format_range","sheetId":"sSources","range":"A1:D1","format":{"bold":true,"fillColor":"#F2F2F2"}},
 {"op":"set_range","sheetId":"sSources","range":"A2:D2","values":[["FY2024 revenue","SEC 10-K","https://sec.gov/...","2026-07-24"]]},
 {"op":"set_hyperlink","sheetId":"sSources","address":"C2","target":"https://sec.gov/..."}
]
```

## Discipline

- **Never fabricate data**: every factual cell needs a source (user-provided / already in the sheet / read via tools); attribution makes the source verifiable.
- Derived values are computed with formulas and their methodology explained — they do not need source tracing.
- When a hyperlink target is a bare domain the system prepends `https://` automatically; use "Sheet!A1" for in-workbook references.
