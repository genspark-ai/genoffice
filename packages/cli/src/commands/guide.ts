import { flagBool } from '../args'
import type { CommandDef } from '../registry'
import { docsGuide } from '../formats/docx'
import { CliError, EXIT, type CommandResult } from '../result'

/** Same markdown the in-app AI loads on demand (load_guide); shipped inside the CLI so it never drifts from the ops. */
export const guideCommand: CommandDef = {
  name: 'guide',
  summary: 'Print the op reference and design guides an agent needs before writing ops or specs.',
  usage: 'guide <slides|docs|sheets> [group|design|spec] [--index]',
  options: [
    { name: 'index', description: 'one-line signature of every op instead of a group guide' },
  ],
  async run(args) {
    const [domain, group] = args.positionals
    if (domain === 'docs') return docsGuideResult()
    if (domain === 'sheets') return { summary: SHEETS_GUIDE }
    if (domain !== 'slides') {
      throw new CliError(EXIT.usage, 'guides available for: slides, docs, sheets', {
        usage: 'genoffice guide slides | genoffice guide docs | genoffice guide sheets',
      })
    }
    if (group === 'design' || group === 'spec') {
      const { SLIDES_GUIDES } = await import('@genoffice/pipelines/slides')
      return { summary: SLIDES_GUIDES[group].content }
    }
    const docs = await import('@genoffice/pptx-ops/op-docs')
    if (flagBool(args, 'index')) {
      return { summary: docs.opSignatureIndex() }
    }
    if (!group) {
      return {
        summary: [
          'Op groups (genoffice guide slides <group> prints one):',
          docs.opGuideCatalog(),
          '',
          'Building a new deck: `genoffice guide slides design` (the staged workflow: style sheet, outline, one page file at a time, build, QC) and `genoffice guide slides spec` (the outline and page spec JSON for `genoffice slides check` and `genoffice create --type pptx --spec`).',
          '',
          'Every op: { "op": "<name>", "target": { "slide": <index|"s_n">, "el"?: "e_*" }, ...fields }.',
          'Units are EMU (914400 per inch); font sizes are pt. `genoffice slides read <file>` lists ids and geometry.',
          'Vocabulary:',
          docs.opVocabulary(),
        ].join('\n'),
      }
    }
    const content = docs.opGuide(group)
    if (!content) {
      throw new CliError(EXIT.usage, `unknown group: ${group}`, { groups: [...docs.OP_GROUPS] })
    }
    return { summary: content }
  },
}

async function docsGuideResult(): Promise<CommandResult> {
  const g = await docsGuide()
  return {
    summary: [
      'Word ops (genoffice docs apply --ops): a JSON array; every entry has "op".',
      'Targets: { "blockIndexes": [..] } or { "nodeType": "docHeading"|"docParagraph"|"docListItem", "headingLevel"? }; get indexes from `genoffice docs read`.',
      '',
      ...g.signatures,
      '',
      'insert_content {html, afterBlockIndex?} — new blocks after an index (-1 = start of document; omitted = end of document)',
      'replace_blocks {html, startBlockIndex, endBlockIndex} — rewrite a block range',
      'insert_image {url, maxWidthPx?: 480, afterBlockIndex?} — url is a local path (absolute, or relative to the current directory then to the ops file), a data: URL or an http(s) URL; png/jpg/gif; scaled down to maxWidthPx; appended when afterBlockIndex is omitted',
      'insert_chart {kind: "bar"|"line"|"pie", title?, categories: string[], series: [{name?, values: (number|null)[]}], afterBlockIndex?} — native Word chart; values per series match the categories',
      'edit_chart {blockIndex, title?, categories?: (string|null)[], series?: [{index, name?, values?: (number|null)[]}]} — change the data of a chart block (`genoffice docs read` lists blocks with kind "chart"); counts must match the original, null keeps a position',
      'set_header_footer {kind: "header"|"footer", text, view?: "default"|"first"|"even"} — plain text, \\n between lines, {PAGE} / {NUMPAGES} become page-number fields, "" clears; view first/even switches the different-first-page / odd-even setting on; `genoffice docs read --header-footer` shows the current text',
      'reply_comment {parentId, text} / resolve_comment {id} — ids from `genoffice docs read --comments`; replies attach to the thread root',
      '',
      g.htmlRules,
    ].join('\n'),
  }
}

/** The headless sheets workbook DSL (the in-app propose_operations vocabulary, minus the editor-only ops). */
const SHEETS_GUIDE = [
  'Excel ops (genoffice sheet apply --ops): a JSON array; every entry has "op". Address a worksheet with',
  '"sheet": "<name>" (omitted = the active sheet; `genoffice sheet read` lists names). Ranges are A1:D9,',
  'rows are 1-based, columns are letters. Strings starting with "=" are formulas.',
  '',
  'Content:',
  '  set_cell {sheet?, address, value}            set_formula {sheet?, address, formula}',
  '  set_range {sheet?, range, values: [[..],..]}  clear_cell {sheet?, address}   clear_range {sheet?, range}',
  '  fill_range {sheet?, source, target}          copy_range {sheet?, source, target}',
  '  find_replace {sheet?, range?, find, replace}  sort_range {sheet?, range, byColumn: "B", order: "asc"|"desc", hasHeader?}',
  'Format:',
  '  format_range {sheet?, range, format: {bold?, italic?, underline?, strikethrough?, fontFamily?, fontSize?,',
  '    fontColor?, fillColor?, numberFormat?, horizontalAlign?: left|center|right, verticalAlign?: top|center|bottom,',
  '    wrapText?, textRotation?: -90..90|"vertical", indent?, border?: {type: all|top|bottom|left|right|none, color?}}}',
  '  (null clears a property; colors are #RRGGBB)',
  '  add_conditional_format {sheet?, range, rule}  rule = {kind: "number", operator: greaterThan|greaterThanOrEqual|lessThan|',
  '    lessThanOrEqual|equal|notEqual|between|notBetween, value, value2?, format} | {kind: "text", operator: contains|notContains|',
  '    beginsWith|endsWith, text, format} | {kind: "blank", blank: bool, format} | {kind: "duplicate", unique?, format} |',
  '    {kind: "top10", rank, percent?, bottom?, format} | {kind: "formula", formula: "=…", format} |',
  '    {kind: "colorScale", minColor, midColor?, maxColor} | {kind: "dataBar", color?}; format = {fillColor?, fontColor?, bold?, italic?}',
  '  clear_conditional_formats {sheet?}  (a sheet that already has rules must be cleared before new ones are added)',
  '  set_data_validation {sheet?, range, validation | null}  validation = {kind: "list", values: [..]} | {kind: "listRef", range} |',
  '    {kind: "numberBetween", min, max} | {kind: "dateBetween", start: "YYYY-MM-DD", end} | {kind: "checkbox"} | {kind: "formula", formula: "=…"}',
  'Layout:',
  '  merge_cells {sheet?, range}   unmerge_cells {sheet?, range}',
  '  set_row_height {sheet?, row, count?, heightPoints}   set_col_width {sheet?, column, count?, widthPx}',
  '  set_rows_hidden {sheet?, row, count?, hidden}   set_cols_hidden {sheet?, column, count?, hidden}',
  '  set_freeze {sheet?, rows, columns}  (0/0 unfreezes)',
  '  set_filter {sheet?, range}   set_filter_criteria {sheet?, column, values: [..] | null}   clear_filter {sheet?}',
  '  set_page_setup {sheet?, orientation?: portrait|landscape, paperSize?: 9 (A4) | 1 (Letter) …, scale? | fitToWidth?/fitToHeight?,',
  '    margins?: normal|wide|narrow, printGridlines?, printHeadings?, printArea?: "A1:F40" | null}',
  '  protect_sheet {sheet?, protected}',
  'Data:',
  '  set_hyperlink {sheet?, address, target: "https://…" | "Sheet1!A1" | null}   set_note {sheet?, address, text | null}',
  '  add_defined_name {name, ref: "Sheet1!$A$2:$A$9"}   delete_defined_name {name}  (not in a batch with sheet or row/column ops)',
  '  add_table {sheet?, range, name?, style?: "TableStyleMedium2", bandedRows?}  (first row = unique, non-empty column names)',
  'Visuals:',
  '  add_chart {sheet?, chartType: column|bar|line|area|pie|doughnut|scatter|radar|combo, dataRange, title?, anchorCell?}',
  '    (header row and a leading category column are detected; the chart lands two columns right of the data unless anchorCell)',
  '  edit_chart {chartPath: "xl/charts/chart1.xml", title?, chartType?, legend?: none|right|bottom|top|left,',
  '    dataLabels?: none|value|percent|category-percent, grouping?: clustered|stacked|percentStacked, seriesColors?: {"0": "#RRGGBB"},',
  '    axisTitles?: {category?, value?}}  (`genoffice sheet read` lists chart ids under features.charts)',
  '  add_image {sheet?, path: "/abs/or/relative.png" | "https://…", anchorCell}   add_shape {sheet?, shapeType | "textbox", anchorCell, fillColor?, text?}',
  'Structure (cannot share a batch with content ops):',
  '  insert_rows {sheet?, row, count}   delete_rows {sheet?, row, count}',
  '  insert_cols {sheet?, column, count}   delete_cols {sheet?, column, count}',
  '  add_sheet {name}   delete_sheet {sheet}   rename_sheet {sheet, name}   duplicate_sheet {sheet, name?}',
  '  move_sheet {sheet, position}   set_sheet_hidden {sheet, hidden}',
  '',
  'Not available headless (use the GenOffice app): pivots (add_pivot, refresh_pivot), sparklines, edits to tables or',
  'shapes created in an editor session (add/delete_table_row/column, delete_table, edit_shape, delete_visual), convert_to_values.',
  'Adding a conditional format or data validation to a sheet that already has rules is refused (the CLI cannot carry the',
  'existing rules over): clear_conditional_formats first, or use the app. set_filter_criteria needs a filter without criteria.',
  'Formulas written by a batch are evaluated by the workbook engine and stored with their results.',
].join('\n')
