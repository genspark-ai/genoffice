import type { Block, Run } from './types'

/**
 * Mail merge, engine half — the data source and the substitution walk.
 *
 * The MERGEFIELD itself already survives a round-trip (the parse whitelist
 * keeps it an instrField run), so what is missing is everything that gives the
 * field a value. This module reads a recipient list, substitutes the fields
 * with one record's data and honours Word's control fields (NEXT / SKIPIF / IF).
 * genoffice#1984.
 *
 * Deliberately no UI here: the panel, the recipient list, the preview and the
 * batch export are all built on top of these two functions, and none of them
 * can be tested without a DOM.
 */

/** one recipient row; keys are the column headers */
export type MergeRecord = Record<string, string>

export interface MergeData {
  /** column headers, in file order */
  readonly fields: readonly string[]
  readonly records: readonly MergeRecord[]
}

/** the guillemets Word puts around a field name in an unmerged template */
const PLACEHOLDER = /«([^»]*)»/g

/**
 * Split a field instruction into its keyword and the rest: `MERGEFIELD "Postal
 * Address" \* MERGEFORMAT` -> `['MERGEFIELD', '"Postal Address" \\* MERGEFORMAT']`.
 */
export function splitInstruction(instr: string): [keyword: string, rest: string] {
  const trimmed = instr.trim()
  const space = trimmed.search(/\s/)
  return space === -1
    ? [trimmed.toUpperCase(), '']
    : [trimmed.slice(0, space).toUpperCase(), trimmed.slice(space + 1).trim()]
}

/**
 * The column name a MERGEFIELD asks for. Quoted so a header with a space
 * survives; a bare name stops at the first whitespace, which keeps the
 * `\* MERGEFORMAT`-style switches out of the name.
 */
export function mergeFieldName(instr: string): string | null {
  const [keyword, rest] = splitInstruction(instr)
  if (keyword !== 'MERGEFIELD' || !rest) return null
  const quoted = /^"([^"]*)"/.exec(rest)
  return quoted ? quoted[1]! : (rest.split(/\s/)[0] ?? null)
}

/** "Ada" from `\* MERGEFORMAT` instructions is not the field name — this is the raw form */
export function placeholderText(name: string): string {
  return `«${name}»`
}

/**
 * The field code a merge field carries in a template.
 *
 * The name is quoted so a header with a space survives, and the MERGEFORMAT
 * switch is the one Word writes: a template built here is then the same field
 * a template built in Word carries, and both read back through mergeFieldName.
 * genoffice#2097.
 */
export function mergeFieldInstruction(name: string): string {
  return `MERGEFIELD "${name.replace(/"/g, '')}" \\* MERGEFORMAT`
}

/** the file name characters Windows and macOS refuse, plus the path separators */
const ILLEGAL_FILE_NAME_CHARS = /[\\/:*?"<>|]/g

/** the name a merged record is written under, derived from the record itself.
 *
 * `pattern` names a column of the data source; an empty pattern takes the
 * first one. A pattern that is no column is a literal, and `{Column}` inside
 * it is filled from the record. The result is reduced to what a file name can
 * carry, so a value like "A/B" cannot escape the folder it is written into.
 */
export function mergeOutputFileName(
  pattern: string,
  record: MergeRecord,
  fields: readonly string[],
): string {
  const trimmed = pattern.trim()
  const source = trimmed === '' ? (fields[0] ?? '') : trimmed
  // a pattern naming a column takes that column's value; anything else is literal
  const literal = fields.includes(source) ? valueOf(record, source) : source
  const filled = literal.replace(/\{([^{}]*)\}/g, (_match, name: string) =>
    valueOf(record, name.trim()),
  )
  const cleaned = filled
    .replace(ILLEGAL_FILE_NAME_CHARS, '_')
    // control characters a pasted value can carry
    .replace(/\p{Cc}/gu, '')
    .replace(/\s+/g, ' ')
    .trim()
    // a name that is only dots is not a name on any file system
    .replace(/^\.+/, '')
    .trim()
  if (cleaned === '') return 'document'
  return cleaned.slice(0, 120)
}

export interface MergeFieldSwitches {
  /** `\b "…"`: text placed before the value, only when the value is non-empty */
  before: string
  /** `\f "…"`: text placed after the value, only when the value is non-empty */
  after: string
  /** `\* Upper | Lower | FirstCap | Caps` */
  caseFormat: string | null
}

/**
 * Word's Address Block and Greeting Line write `\f " "` on every name part, so
 * dropping the switches runs the parts together ("AdaLovelace").
 */
export function mergeFieldSwitches(instr: string): MergeFieldSwitches {
  const out: MergeFieldSwitches = { before: '', after: '', caseFormat: null }
  const re = /\\([bf*])\s+(?:"([^"]*)"|(\S+))/g
  for (const m of instr.matchAll(re)) {
    const arg = m[2] ?? m[3] ?? ''
    if (m[1] === 'b') out.before = arg
    else if (m[1] === 'f') out.after = arg
    else if (!/^MERGEFORMAT$/i.test(arg) && !/^CHARFORMAT$/i.test(arg)) out.caseFormat = arg
  }
  return out
}

function applyCaseFormat(value: string, format: string | null): string {
  switch (format?.toLowerCase()) {
    case 'upper':
      return value.toUpperCase()
    case 'lower':
      return value.toLowerCase()
    case 'firstcap':
      return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase()
    case 'caps':
      return value
        .toLowerCase()
        .replace(/(^|[\s-])(\p{L})/gu, (_, sep, ch) => sep + ch.toUpperCase())
    default:
      return value
  }
}

/** the text a MERGEFIELD run shows for one record, switches applied */
export function mergeFieldText(instr: string, record: MergeRecord): string {
  const name = mergeFieldName(instr)
  if (!name) return ''
  const value = applyCaseFormat(valueOf(record, name), mergeFieldSwitches(instr).caseFormat)
  if (value === '') return ''
  const { before, after } = mergeFieldSwitches(instr)
  return before + value + after
}

/** the record a cursor position reads; past the last row every field is empty,
 *  the way the trailing labels of Word's last sheet come out blank */
function recordAt(data: MergeData, index: number): MergeRecord {
  return data.records[Math.max(index, 0)] ?? {}
}

/** the value a record holds for `name`, or '' when the column is absent */
export function valueOf(record: MergeRecord, name: string): string {
  return record[name] ?? ''
}

/**
 * An IF with no comparison tests one operand: `IF «Paid» "yes" "no"` is true
 * when the column holds anything. A bare word with no guillemets is a literal,
 * so it is true exactly when it is non-empty text.
 */
function isTruthyOperand(text: string, record: MergeRecord): boolean {
  const value = resolveOperand(text, record)
  return value !== '' && value !== '0'
}

/**
 * Evaluate an IF / SKIPIF comparison. Word writes `«Field» = "literal"` or
 * `«Field» = «Other»`; both sides are resolved against the record first, so a
 * MERGEFIELD wrapper around a name is invisible here.
 */
export function compareOperands(
  left: string,
  op: string,
  right: string,
  record: MergeRecord,
): boolean {
  const a = resolveOperand(left, record)
  const b = resolveOperand(right, record)
  // Word compares numerically when both sides are numbers: `IF 9 < 50` is true
  const na = numeric(a)
  const nb = numeric(b)
  if (na !== null && nb !== null) return compareValues(na, nb, op)
  return compareValues(a, b, op)
}

function numeric(text: string): number | null {
  const t = text.trim()
  if (t === '' || !/^[-+]?(\d+\.?\d*|\.\d+)$/.test(t)) return null
  return Number(t)
}

function compareValues<T extends string | number>(a: T, b: T, op: string): boolean {
  switch (op) {
    case '=':
    case '<>':
      return (a === b) === (op === '=')
    case '<':
      return a < b
    case '<=':
      return a <= b
    case '>':
      return a > b
    case '>=':
      return a >= b
    default:
      return false
  }
}

/**
 * Resolve one comparison operand against the record.
 *
 * A bare word is a LITERAL, not a field reference: in `SKIPIF Status = "yes"`
 * `Status` is the text "Status" and `"yes"` is the text "yes", so the test is
 * false for every record. Word only treats an operand as a data column when it
 * is wrapped in guillemets (`«Status»`) or carries the MERGEFIELD keyword.
 */
function resolveOperand(token: string, record: MergeRecord): string {
  const trimmed = token.trim()
  const quoted = /^"(.*)"$/s.exec(trimmed)
  if (quoted) return quoted[1] ?? ''
  const guillemet = /^«([^»]*)»$/.exec(trimmed)
  if (guillemet) return valueOf(record, guillemet[1] ?? '')
  const instruction = /^MERGEFIELD\s+(?:"([^"]*)"|(\S+))/i.exec(trimmed)
  if (instruction) return valueOf(record, instruction[1] ?? instruction[2] ?? '')
  return trimmed
}

/**
 * The first comparison an IF/SKIPIF instruction carries, plus whatever follows
 * it. `right` must stop at the operand's own end: in
 * `Skip = "yes" "hidden" "shown"` the right operand is `"yes"` and the two
 * branch literals are what is left, so the trailing group has to be lazy about
 * where `right` stops and greedy about what it keeps.
 */
function parseComparison(
  instr: string,
): { left: string; op: string; right: string; rest: string } | null {
  const m = /^(.+?)\s*(<>|=|<=|>=|<|>)\s*("(?:[^"]*)"|\S+)(?:\s+([\s\S]*))?$/.exec(instr.trim())
  return m ? { left: m[1]!, op: m[2]!, right: m[3]!, rest: m[4] ?? '' } : null
}

/** true when an IF/SKIPIF condition holds for this record */
export function conditionHolds(instr: string, record: MergeRecord): boolean {
  const [, rest] = splitInstruction(instr)
  const cmp = parseComparison(rest)
  if (!cmp) {
    // no comparison: the FIRST operand is the test (`IF «Paid» "got" "missed"`),
    // not the whole instruction — that is non-empty text for every record, so
    // the false branch would be unreachable
    return isTruthyOperand(tokenizeOperands(rest)[0] ?? '', record)
  }
  return compareOperands(cmp.left, cmp.op, cmp.right, record)
}

/**
 * The branch an IF instruction selects, unwrapped from its quotes.
 *
 * Word writes `IF «Field» = "yes" "hidden" "shown"` — the condition comes
 * first, so the branches are whatever follows the comparison, not the first two
 * tokens. An IF with no comparison (`IF «Field» "yes" "no"`) uses the field as
 * the test and takes the two operands as the branches.
 */
export function ifBranch(instr: string, record: MergeRecord): string {
  const [, rest] = splitInstruction(instr)
  const branches = branchesAfterCondition(rest)
  const chosen = conditionHolds(instr, record) ? (branches[0] ?? '') : (branches[1] ?? '')
  return chosen.replace(/^"(.*)"$/s, '$1').replace(PLACEHOLDER, (_, name) => valueOf(record, name))
}

/** the two branch literals of an IF, with the condition stripped off the front */
function branchesAfterCondition(rest: string): string[] {
  // parseComparison returns everything after the right operand, which for an IF
  // is exactly the two branches; without a comparison the first operand is the
  // test itself, so drop it and take what follows
  const cmp = parseComparison(rest)
  const operands = tokenizeOperands(cmp ? cmp.rest : rest)
  return cmp ? operands : operands.slice(1)
}

/**
 * Split the IF operands on whitespace but keep a quoted run (with spaces)
 * together: `Skip = "yes" "a b" "c d"` -> 5 tokens.
 */
function tokenizeOperands(rest: string): string[] {
  const out: string[] = []
  let current = ''
  let quoted = false
  for (let i = 0; i < rest.length; i += 1) {
    const ch = rest[i]!
    if (ch === '"') {
      quoted = !quoted
      current += ch
      continue
    }
    if (!quoted && /\s/.test(ch)) {
      if (current) out.push(current)
      current = ''
      continue
    }
    current += ch
  }
  if (current) out.push(current)
  return out
}

/** what a control field does to the run that carries it */
export type ControlEffect =
  | { readonly kind: 'next' }
  | { readonly kind: 'skip' }
  | { readonly kind: 'skipNoop' }
  | { readonly kind: 'if'; readonly text: string }

export function controlEffect(instr: string, record: MergeRecord): ControlEffect | null {
  const [keyword] = splitInstruction(instr)
  if (keyword === 'NEXT') return { kind: 'next' }
  if (keyword === 'SKIPIF')
    return conditionHolds(instr, record) ? { kind: 'skip' } : { kind: 'skipNoop' }
  if (keyword === 'IF') return { kind: 'if', text: ifBranch(instr, record) }
  return null
}

/** substitute one run: a MERGEFIELD becomes its value, a control field becomes its
 *  effect, and any other field is reported untouched so the caller keeps it */
function mergeRun(
  run: Run,
  record: MergeRecord,
): { run: Run; effect: ControlEffect | null; untouched: boolean } {
  const instr = run.instrField
  if (!instr) return { run, effect: null, untouched: false }
  const [keyword] = splitInstruction(instr)
  // keep the run's formatting, drop the instruction — the value replaces both
  const { instrField: _dropped, fldDirty: _dirty, ...rest } = run
  if (keyword === 'MERGEFIELD') {
    return {
      run: { ...rest, text: mergeFieldText(instr, record) },
      effect: null,
      untouched: false,
    }
  }
  const effect = controlEffect(instr, record)
  // an unrelated field (PAGE, DATE, a Zotero citation): leave the run alone
  if (!effect) return { run, effect: null, untouched: true }
  return {
    run: { ...rest, text: effect.kind === 'if' ? effect.text : '' },
    effect,
    untouched: false,
  }
}

/** the run-position cursor a single template pass walks with */
export interface MergeCursorState {
  /** index into data.records; NEXT moves it forward */
  index: number
  /** a NEXT fired during this pass */
  next: boolean
  /** a SKIPIF tested true during this pass */
  skipped: boolean
}

/** true when every run in the paragraph was consumed by a control field, so the
 *  paragraph itself must not survive into the merged output */
function consumedByControlField(runs: readonly Run[] | undefined): boolean {
  if (!runs || runs.length === 0) return false
  return runs.every((run) => {
    if (!run.instrField) return false
    const [keyword] = splitInstruction(run.instrField)
    return keyword === 'NEXT' || keyword === 'SKIPIF'
  })
}

function mergeRuns(
  runs: readonly Run[] | undefined,
  data: MergeData,
  cursor: MergeCursorState,
): Run[] {
  if (!runs || runs.length === 0) return []
  const out: Run[] = []
  for (const run of runs) {
    // a run without a field is the template's own content: an image or an
    // index entry carries its payload beside an empty text and must survive
    // every record's copy
    if (!run.instrField) {
      out.push(run)
      continue
    }
    const record = recordAt(data, cursor.index)
    const { run: merged, effect, untouched } = mergeRun(run, record)
    // a field this merge knows nothing about (a PAGE, a DATE) is left exactly as
    // it was — it still evaluates in Word, and dropping it would lose the field
    if (untouched) {
      out.push(run)
      continue
    }
    if (effect === null) {
      // a substituted MERGEFIELD keeps its slot even when the column is empty
      // or absent — the position is the user's layout; only a control field
      // takes its paragraph's content away
      out.push(merged)
      continue
    }
    if (effect.kind === 'next') {
      cursor.next = true
      cursor.index += 1
      continue
    }
    if (effect.kind === 'skip') {
      cursor.skipped = true
      continue
    }
    // skipNoop: a SKIPIF that tested false — contributes nothing, marks nothing
    if (effect.kind === 'skipNoop') continue
    // IF contributes its chosen branch inline, exactly where the field sat
    if (effect.text) out.push(merged)
  }
  return out
}

/**
 * Walk the template once, substituting fields from `data` as the cursor moves.
 *
 * The tree is walked structurally rather than through a serializer, so a field
 * inside a table cell is substituted exactly like one in the body. Floating
 * textboxes are display-only (the save keeps their original bytes), so a field
 * inside one stays as it was — for the same reason as the nested tables below.
 */
export function mergeRecords(
  blocks: readonly Block[],
  data: MergeData,
  startIndex = 0,
): { blocks: Block[]; cursor: MergeCursorState } {
  const cursor: MergeCursorState = { index: startIndex, next: false, skipped: false }
  // merge FIRST, drop the paragraph afterwards: a paragraph holding only a NEXT
  // is invisible in the output, but its field still has to advance the cursor —
  // filtering before the walk would leave every record reading the same row
  const walk = (list: readonly Block[]): Block[] => {
    const out: Block[] = []
    for (const block of list) {
      const controlOnly = consumedByControlField(block.runs)
      const runs = mergeRuns(block.runs, data, cursor)
      if (controlOnly) continue
      const copy: Block = { ...block, runs }
      // the table is merged with the SAME cursor object, so a NEXT inside a
      // cell advances the body walk too — that is how a label sheet stays in step
      if (block.table) copy.table = mergeTable(block.table, data, cursor)
      out.push(copy)
    }
    return out
  }
  return { blocks: walk(blocks), cursor }
}

/** table cells, their rich paragraphs and any nested table */
function mergeTable(
  table: NonNullable<Block['table']>,
  data: MergeData,
  cursor: MergeCursorState,
): NonNullable<Block['table']> {
  return {
    ...table,
    // TableModel.rows is a grid of cells (w:tr -> w:tc), not row objects
    rows: table.rows.map((row) =>
      row.map((cell) => {
        const next = { ...cell }
        if (cell.richParas) {
          next.richParas = cell.richParas.map((para) => ({
            ...para,
            runs: mergeRuns(para.runs, data, cursor),
          }))
          next.paras = next.richParas.map((para) => (para.runs ?? []).map((r) => r.text).join(''))
        }
        // nestedTables are display-only: the outer table keeps their original
        // bytes on save, so substituting inside one would change what is shown
        // but not what is written — the field would come back on the next open
        return next
      }),
    ),
  }
}

/**
 * Merge every record into its own copy of the block tree.
 *
 * A template with no NEXT yields one set per record — the letter-per-recipient
 * case. A NEXT hands the remainder of the template to the next record, which is
 * what a directory or a label sheet does. A SKIPIF removes the whole record.
 */
export function mergeBlockTree(blocks: readonly Block[], data: MergeData): Block[][] {
  const out: Block[][] = []
  let index = 0
  for (let round = 0; index < data.records.length && round <= data.records.length; round += 1) {
    const { blocks: merged, cursor } = mergeRecords(blocks, data, index)
    if (!cursor.skipped) out.push(merged)
    // a NEXT-driven pass (a label sheet) consumed several records; the next
    // sheet starts after the last one read
    index = cursor.index + 1
  }
  return out
}

/**
 * Build a recipient list from already-parsed spreadsheet rows.
 *
 * Takes `string[][]` rather than CSV text on purpose: the CSV text→rows step
 * belongs to xlsx-gateway, which already owns `parseCsv` (RFC 4180 quotes, the
 * BOM, CRLF, the `sep=` declaration) and is what Sheets uses on open. The app
 * layer wires them together, so a list that opens correctly in Sheets opens
 * correctly here, and this package keeps no CSV parser of its own.
 *
 * The first row names the columns; the rest are records.
 */
export function mergeDataFromRows(rows: readonly (readonly string[])[]): MergeData {
  if (rows.length === 0) return { fields: [], records: [] }
  const header = (rows[0] ?? []).map((h) => h.trim())
  const records = rows.slice(1).map((row) => {
    const record: MergeRecord = {}
    header.forEach((name, i) => {
      record[name] = row[i] ?? ''
    })
    return record
  })
  return { fields: header, records }
}
