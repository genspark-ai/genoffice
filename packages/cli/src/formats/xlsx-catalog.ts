import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { z } from 'zod'
import {
  renderGroups,
  signatureFromSchema,
  withFingerprint,
  type JsonSchema,
  type OpCatalog,
  type OpEntry,
} from '../op-catalog'
import { REFUSED_DSL_OPS } from './xlsx-gateway-ops'

/** Presentation order of the guide; a test asserts every supported op sits in exactly one group. */
export const SHEET_OP_GROUPS: Record<string, readonly string[]> = {
  content: [
    'set_cell',
    'set_formula',
    'set_range',
    'clear_cell',
    'clear_range',
    'fill_range',
    'copy_range',
    'find_replace',
    'sort_range',
    'convert_to_values',
  ],
  format: [
    'format_range',
    'add_conditional_format',
    'clear_conditional_formats',
    'set_data_validation',
  ],
  layout: [
    'merge_cells',
    'unmerge_cells',
    'set_row_height',
    'set_col_width',
    'set_rows_hidden',
    'set_cols_hidden',
    'set_freeze',
    'set_filter',
    'set_filter_criteria',
    'clear_filter',
    'set_page_setup',
    'protect_sheet',
  ],
  data: [
    'set_hyperlink',
    'set_note',
    'add_defined_name',
    'delete_defined_name',
    'add_table',
    'add_pivot',
  ],
  visuals: ['add_chart', 'edit_chart', 'add_image', 'add_shape', 'add_sparkline'],
  structure: [
    'insert_rows',
    'delete_rows',
    'insert_cols',
    'delete_cols',
    'add_sheet',
    'delete_sheet',
    'rename_sheet',
    'duplicate_sheet',
    'move_sheet',
    'set_sheet_hidden',
  ],
}

const GROUP_SUMMARIES: Record<string, string> = {
  content: 'cell values, formulas, fills, find/replace, sorting',
  format: 'cell formats, conditional formats, data validation',
  layout: 'merges, sizes, hidden rows/columns, panes, filters, page setup, protection',
  data: 'hyperlinks, notes, defined names, Excel tables, pivot tables',
  visuals: 'charts, images, shapes, sparklines',
  structure: 'rows, columns and sheets; later ops in the batch use the shifted addresses',
  'editor-only': 'accepted by the app, refused headless',
}

const NOTES: Record<string, string> = {
  set_cell:
    'a string value starting with "=" is written as a formula; type: "text" keeps such a string as literal cell text',
  set_range: 'start is the top-left cell; range, when given instead, must match the values grid',
  format_range:
    'null clears a property; colors are #RRGGBB or theme slots ("accent1", "accent1+40%", "dk2-25%", {theme, tint}) that follow the workbook theme; fill = {pattern, fg, bg?} or {gradient: {angle?, stops}}; border type none removes all borders',
  add_conditional_format:
    'adds to the rules the sheet already has (clear_conditional_formats first to start over); format = {fillColor?, fontColor?, bold?, italic?}',
  clear_conditional_formats:
    'removes every rule of the sheet; later add_conditional_format ops in the batch rebuild from scratch',
  set_data_validation:
    'adds to the rules the sheet already has; on a range that has one it replaces it, null removes it',
  set_freeze: '0/0 unfreezes',
  set_filter_criteria: 'needs a filter without criteria; null clears the column',
  set_page_setup:
    'paperSize 9 = A4, 1 = Letter; scale and fitToWidth/fitToHeight are exclusive; margins normal|wide|narrow; printArea/printTitles/header/footer null clears; rowBreaks/colBreaks replace the manual breaks',
  add_pivot:
    'sourceRange first row = field names; the output is written as cells at targetCell (must be empty) plus a native pivot Excel can refresh; targetSheet must already exist in the file; not in a batch with sheet or row/column ops',
  add_sparkline:
    'one sparkline per row of dataRange (at most 200 rows per op), in the column right of it unless targetCell; type line|column|stacked',
  convert_to_values:
    'replaces the formulas in range with the values the workbook engine computes; cells that evaluate to an error keep their formula',
  set_col_width: 'px',
  set_row_height: 'pt',
  add_defined_name: 'ref like "Sheet1!$A$2:$A$9"; not in a batch with sheet or row/column ops',
  delete_defined_name: 'not in a batch with sheet or row/column ops',
  add_table:
    'first row = unique, non-empty column names; not in a batch with row/column, merge or size ops on the same sheet',
  add_chart:
    'header row and a leading category column are detected; the chart lands two columns right of the data unless anchorCell',
  edit_chart:
    '`genoffice sheet read` lists chart paths under features.charts; repointing a series at other cells needs the app',
  add_image: 'path is a local file (absolute or relative) or an https URL',
  add_shape: 'shapeType "textbox" for a text box',
  sort_range: 'hasHeader keeps the first row in place',
  duplicate_sheet: 'the copy is cloned from the file; cell edits of the same batch are not in it',
}

const HEADER = [
  'Excel ops (genoffice sheet apply --ops): a JSON array; every entry has "op". Address a worksheet with',
  '"sheet": "<name>" (omitted = the active sheet; `genoffice sheet read` lists names). Ranges are A1:D9,',
  'rows are 1-based, columns are letters. Strings starting with "=" are formulas.',
  'Field notation: bare = string, n = number, bool = boolean, ? = optional, a|b = one of.',
]

const FOOTER = [
  'Ops run in order: after insert/delete rows or columns, the ops that follow address the shifted grid, and a',
  'sheet added or renamed earlier in the batch is addressed by its new name. Rules, links, notes, filters and',
  'visuals must come after any row/column op of their sheet; defined-name ops cannot share a batch with sheet',
  'or row/column ops. Formulas written by a batch are evaluated by the workbook engine and stored',
  'with their results; functions the engine lacks are left for Excel to compute on open (warning formulas_not_cached),',
  'and a formula the engine cannot parse is reported separately (warning formula_errors) so a bad reference is not',
  'mistaken for a missing function.',
]

function opSchemas(): Map<string, JsonSchema> {
  const json = z.toJSONSchema(workbookOperationSchema, {
    io: 'input',
    unrepresentable: 'any',
  }) as JsonSchema
  const out = new Map<string, JsonSchema>()
  for (const variant of json.oneOf ?? json.anyOf ?? []) {
    const name = variant.properties?.op?.const
    if (typeof name === 'string') out.set(name, cliSurface(variant))
  }
  return out
}

/** Fields the app's schema accepts but `sheet apply` rejects headless. */
const CLI_UNSUPPORTED_FIELDS: Record<string, readonly string[]> = {
  edit_chart: ['seriesData'],
}

/** The app addresses sheets by id; the CLI takes the name, defaulting to the active sheet. */
function cliSurface(schema: JsonSchema, nested = false): JsonSchema {
  const op = schema.properties?.op?.const
  const dropped = typeof op === 'string' ? (CLI_UNSUPPORTED_FIELDS[op] ?? []) : []
  const properties: Record<string, JsonSchema> = {}
  for (const [key, value] of Object.entries(schema.properties ?? {})) {
    if (dropped.includes(key)) continue
    if (key === 'sheetId') {
      properties.sheet = {
        type: 'string',
        description: nested ? 'worksheet name' : 'worksheet name; omitted = the active sheet',
      }
    } else if (key === 'sourceSheetId') {
      properties.sourceSheet = {
        type: 'string',
        description: 'worksheet of the source range; omitted = sheet',
      }
    } else if (key === 'targetSheetId') {
      properties.targetSheet = {
        type: 'string',
        description: 'worksheet that receives the output; omitted = sheet',
      }
    } else if (value.type === 'array' && value.items?.properties) {
      properties[key] = { ...value, items: cliSurface(value.items, true) }
    } else properties[key] = value
  }
  return {
    ...schema,
    properties,
    required: (schema.required ?? [])
      .filter((k) => k !== 'sheetId' && !dropped.includes(k))
      .map((k) =>
        k === 'sourceSheetId' ? 'sourceSheet' : k === 'targetSheetId' ? 'targetSheet' : k,
      ),
  }
}

let cached: OpCatalog | undefined

export function sheetCatalog(): OpCatalog {
  if (cached) return cached
  const schemas = opSchemas()
  const ops: OpEntry[] = []
  const groups = Object.entries(SHEET_OP_GROUPS).map(([name, names]) => {
    for (const op of names) {
      const schema = schemas.get(op)
      if (!schema) throw new Error(`sheet op catalog: ${op} has no schema`)
      ops.push({
        op,
        group: name,
        signature: signatureFromSchema(schema),
        ...(NOTES[op] ? { note: NOTES[op] } : {}),
        schema,
      })
    }
    return { name, summary: GROUP_SUMMARIES[name], ops: [...names] }
  })
  const refused = Object.keys(REFUSED_DSL_OPS)
  for (const op of refused) {
    const schema = schemas.get(op)
    ops.push({
      op,
      group: 'editor-only',
      signature: schema ? signatureFromSchema(schema) : '{}',
      available: false,
      reason: REFUSED_DSL_OPS[op]!,
    })
  }
  groups.push({ name: 'editor-only', summary: GROUP_SUMMARIES['editor-only'], ops: refused })
  cached = withFingerprint({ domain: 'sheets', groups, ops })
  return cached
}

export function sheetGuideText(group?: string): string {
  const lines = [...HEADER, '', ...renderGroups(sheetCatalog(), group)]
  if (!group) lines.push('', ...FOOTER)
  return lines.join('\n')
}
