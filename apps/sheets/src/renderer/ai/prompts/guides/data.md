# Data tools guide (data)

All operations below are **layout-class**: they can share a batch with content/format operations, but not with structural operations (row/column or sheet insertion/deletion). They all work only on imported real xlsx files; once applied they are recorded in the edit journal and ⌘S saves them into the file.

**Before modifying existing settings, read_sheet_features first**: it reads back the current filter range and column criteria, all conditional formatting rules, data validation rules, defined names, and the shape/image inventory.

## Hyperlinks

- `{op:"set_hyperlink", sheetId, address:"B2", target:"https://example.com"}` — target can also be an in-workbook reference "Sheet1!A1" (# is added automatically); bare domains get https:// prepended automatically.
- `{op:"set_hyperlink", sheetId, address, target:null}` — removes the link (and clears the link styling).
- Setting a link gives the cell blue underlined styling; the cell content is unchanged.

## AutoFilter

- `{op:"set_filter", sheetId, range:"A1:D100"}` — creates an AutoFilter on the range (an existing filter is replaced); the first row of range is the filter header.
- `{op:"clear_filter", sheetId}` — removes the entire AutoFilter.
- `{op:"set_filter_criteria", sheetId, column:"B", values:["East","South"]}` — checkbox-style column criteria: rows whose **displayed cell text** is in values stay visible, other rows are hidden. Requires an existing AutoFilter on the sheet. values:null clears the column's criteria.
- read_range the column first to see which values actually exist before writing values; text must match verbatim.

## Conditional formatting

`{op:"add_conditional_format", sheetId, range, rule}` — rule discriminated by kind:

- `{kind:"number", operator:"greaterThan"|"greaterThanOrEqual"|"lessThan"|"lessThanOrEqual"|"equal"|"notEqual"|"between"|"notBetween", value, value2?, format}` — between/notBetween require value2
- `{kind:"text", operator:"contains"|"notContains"|"beginsWith"|"endsWith", text, format}`
- `{kind:"blank", blank:true|false, format}` — true highlights blank cells, false highlights non-blank
- `{kind:"duplicate", unique?, format}` — highlights duplicate values by default; unique:true highlights unique values instead
- `{kind:"top10", rank:10, percent?, bottom?, format}`
- `{kind:"formula", formula:"=B2>100", format}` — formula starts with =, written with relative references anchored at the range's top-left cell
- `{kind:"average", operator:"above"|"below"|"aboveOrEqual"|"belowOrEqual", format}` — compares against the range's own average
- `{kind:"colorScale", minColor:"#63BE7B", midColor?, maxColor:"#F8696B"}` — color scale (no format field)
- `{kind:"dataBar", color?}` — data bar (no format field)
- `{kind:"iconSet", set, reverse?, hideValue?}` — icon set (no format field); thresholds are spaced evenly. set is one of 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 3Symbols2, 4Arrows, 4ArrowsGray, 4RedToBlack, 4Rating, 4TrafficLights, 5Arrows, 5ArrowsGray, 5Quarters, 5Rating. reverse:true puts the best icon at the low end; hideValue:true shows only icons

format: `{fillColor?, fontColor?, bold?, italic?}` with at least one property. Highlight colors follow the low-saturation principle (light background, dark text), e.g. red #FFC7CE/#9C0006, green #C6EFCE/#006100, yellow #FFEB9C/#9C6500.

`{op:"clear_conditional_formats", sheetId}` — clears **all** conditional formatting rules on the sheet (including ones that came with the file); confirm with the user before doing this.

## Data validation

`{op:"set_data_validation", sheetId, range, validation}` — validation discriminated by kind:

- `{kind:"list", values:["Yes","No"]}` — dropdown list (≤100 items)
- `{kind:"listRef", range:"H1:H20"}` — dropdown options taken from a cell range's values
- `{kind:"numberBetween", min, max}`
- `{kind:"dateBetween", start:"2026-01-01", end:"2026-12-31"}` — ISO dates
- `{kind:"checkbox"}`
- `{kind:"formula", formula:"=LEN(A1)<=10"}`
- `validation:null` — clears validation rules on the range

## Notes

- `{op:"set_note", sheetId, address:"B2", text:"This figure includes tax"}` — adds/edits a cell note (yellow sticky, shown on hover).
- `{op:"set_note", sheetId, address, text:null}` — deletes the note.
- Saved as Excel legacy comments; visible in Excel when hovering the cell. Read existing notes with read_sheet_features first.

## Pivot tables

- `{op:"add_pivot", sheetId, sourceRange, targetCell, rowFields, columnField?, values:[{field, agg, showDataAs?}]}` — **creates a native pivot table** (real OOXML, interactive and refreshable when opened in Excel); rowFields supports 1–8 levels of row grouping, showDataAs supports percent-of-total/row/column. Usage and limits in the `pivot` guide (load_guide).
- `{op:"refresh_pivot", sheetId}` — **recomputes all pivot tables on the sheet**: using the layout recorded in the file, recomputes the data area from current source data and writes it back (on save it also makes Excel rebuild the cache when the file is opened). Use it after the source data has changed to bring pivot results up to date.
- If refresh finds **new category members** in the source data, the layout grows automatically (new members appended, output region expanded — the growth area must be empty); moved field columns, renamed headers, or pivots using calculated fields/grouping/value filters still fail with an explicit error (suggest the user refresh in Excel in that case).
- read_sheet_features lists each pivot table's row/column/value fields, source range, and whether it is recomputable.
- Modifying an existing pivot table's field layout is not supported yet; the output region is read-only (direct AI writes are rejected).
- When the data is irregular, live formulas are needed for further computation, or add_pivot limits are exceeded, fall back to a **formula aggregation table** (SUMIFS/COUNTIFS, see the pivot guide's fallback plan).
- Summary tables follow the formatting guide's table style rules (headers, alignment, total rows).

## Defined names

- `{op:"add_defined_name", name:"SalesRegion", ref:"Sheet1!$A$1:$B$10"}` — name starts with a letter/underscore; ref is an A1 reference (absolute $ references recommended) or a constant.
- `{op:"delete_defined_name", name}` — deletes the name (formula text referencing it is untouched, but those formulas will turn into #NAME? — confirm no formulas use it before deleting).
