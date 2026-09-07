/** Cap on document / selection plain text returned to the model / UI. */
export const PLAIN_TEXT_MAX_CHARS = 80_000
/** Short selection preview attached to every user turn. */
export const SELECTION_PREVIEW_CHARS = 400
export const PLAIN_TEXT_UNAVAILABLE = 'plain text unavailable'
export const PARAGRAPH_PREPARE_UNAVAILABLE = 'paragraph prepare is unavailable'
export const PARAGRAPH_NOT_EDITABLE = 'paragraph is not editable'
export const SELECTION_NOT_IN_PARAGRAPH = 'nothing is selected in this paragraph'
export const PARAGRAPH_INDEX_OUT_OF_RANGE = 'paragraph index out of range'
export const INSERT_CONTENT_EMPTY = 'insert text must not be empty'
export const FIELD_NOT_FOUND = 'field not found'
export const TABLE_NOT_FOUND = 'table cell not found'
export const TABLE_SIZE_INVALID = 'table size must be positive integers'
export const TABLE_CELLS_UNFILLED = 'table inserted but cell writes failed'
export const FORMAT_EMPTY = 'format must include at least one property'
export const FORMAT_EMPTY_RANGE = 'no characters to format in this paragraph'
export const TABLE_EDIT_EMPTY = 'table edit must include a valid action'
export const TABLE_STYLE_EMPTY = 'table style must include fill, valign, border, or width'
export const PAGE_SETUP_EMPTY = 'page setup must include orientation, paper, margin, or columns'
export const TABLE_MAX_ROWS = 20
export const TABLE_MAX_COLS = 10
const HWP_PER_MM = 7200 / 25.4
const PAGE_MARGIN_MAX_MM = 80
const TABLE_WIDTH_MIN_MM = 20
const TABLE_WIDTH_MAX_MM = 300
const PAGE_COLUMNS_MAX = 4
const PAPER_SIZES = {
  A4: [59528, 84188],
  A3: [84188, 119055],
  B4: [72850, 103040],
  B5: [51502, 72850],
  Letter: [62208, 80496],
  Legal: [62208, 102816],
} as const
const FONT_SIZE_MIN_PT = 8
const FONT_SIZE_MAX_PT = 72
/** Hangul paragraph dialog: pt → stored indent/margin (same as studio LS()). */
const PARA_PT_TO_UNIT = 200
const INDENT_MAX_PT = 100
const LINE_SPACING_MIN = 80
const LINE_SPACING_MAX = 300
/** v1 applyTextCommand replacement cap (Unicode code points). */
export const PARAGRAPH_MAX_CODE_POINTS = 4000
export const FIELD_MAX_CODE_POINTS = 8000

export interface HangulSelectionState {
  page: number | null
  hasSelection: boolean
}

export interface HangulParagraphPreview {
  index: number
  editable: boolean
  reason: string | null
  section: number
  paragraph: number
  text: string
}

export interface HangulField {
  name: string
  value: string
  type: string | null
}

export interface HangulTableCell {
  index: number
  row: number
  col: number
  text: string
}

export interface HangulTable {
  index: number
  section: number
  paragraph: number
  control: number
  rows: number
  cols: number
  cells: HangulTableCell[]
}

export type HangulAlign = 'left' | 'center' | 'right' | 'justify'
export type HangulListStyle = 'none' | 'bullet' | 'number'

export interface HangulFormatSpec {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  fontSize?: number
  color?: string
  font?: string
  align?: HangulAlign
  list?: HangulListStyle
  lineSpacing?: number
  indentLeft?: number
  indentRight?: number
  indentFirstLine?: number
}

export interface HangulStudioFacade {
  pageCount(): Promise<number>
  currentPage(): Promise<number | null>
  readSelectionState(): Promise<HangulSelectionState>
  getPlainText(): Promise<string>
  getSelectionText(): Promise<string | null>
  hasSelection(): Promise<boolean>
  listParagraphs(): Promise<HangulParagraphPreview[]>
  insertContent(text: string, afterIndex?: number): Promise<{ count: number; start: number }>
  replaceParagraph(text: string, index?: number): Promise<{ before: string; after: string }>
  replaceSelection(text: string): Promise<{ before: string; after: string }>
  listFields(): Promise<HangulField[]>
  setField(name: string, value: string): Promise<{ name: string; before: string; after: string }>
  listTables(): Promise<HangulTable[]>
  replaceCell(
    table: number,
    row: number,
    col: number,
    text: string,
  ): Promise<{ before: string; after: string }>
  insertTable(
    rows: number,
    cols: number,
    cells?: string[][],
    afterIndex?: number,
  ): Promise<{ table: number; rows: number; cols: number; unfilled: string[] }>
  applyFormat(
    format: HangulFormatSpec,
    index?: number,
    indexes?: number[],
    cell?: HangulCellFormatTarget,
  ): Promise<{ indexes: number[]; applied: string[]; table?: number; row?: number }>
  editTable(spec: HangulTableEditSpec): Promise<HangulTableEditResult>
  styleTable(spec: HangulTableStyleSpec): Promise<{ table: number; applied: string[] }>
  setPage(spec: HangulPageSetupSpec): Promise<{ applied: string[] }>
}

export interface HangulCellFormatTarget {
  table: number
  row: number
  col?: number
}

export type HangulTableEditAction =
  | 'insert_row'
  | 'insert_column'
  | 'delete_row'
  | 'delete_column'
  | 'merge'
  | 'split'

export interface HangulTableEditSpec {
  action: HangulTableEditAction
  table: number
  row?: number
  col?: number
  after?: boolean
  endRow?: number
  endCol?: number
  splitRows?: number
  splitCols?: number
}

export interface HangulTableEditResult {
  table: number
  action: HangulTableEditAction
  detail: string
  rows?: number
  cols?: number
  cells?: Array<{ row: number; col: number }>
}

export type HangulVAlign = 'top' | 'center' | 'bottom'
export type HangulPaper = keyof typeof PAPER_SIZES
export type HangulOrientation = 'portrait' | 'landscape'

export interface HangulTableStyleSpec {
  table: number
  row?: number
  col?: number
  fill?: string
  valign?: HangulVAlign
  border?: string | false
  width?: number
}

export interface HangulPageSetupSpec {
  orientation?: HangulOrientation
  paper?: HangulPaper
  marginTop?: number
  marginBottom?: number
  marginLeft?: number
  marginRight?: number
  columns?: number
  columnSpacing?: number
}

export interface StudioTextSource {
  pageCount(): Promise<number>
  exportHml(): Promise<Uint8Array>
  getHmlSaveState?(): Promise<{ hmlSavable: boolean }>
  getSelectionContext(): Promise<{
    collapsed: boolean
    selectedTextSha256: string | null
    page?: number
  }>
  hwpctrl: { call(method: string, args?: unknown[]): Promise<unknown> }
  getDocumentState?(): Promise<{
    documentEpoch: number
    changeSeq: number
    documentSha256: string
  }>
  applyTextCommand?(command: {
    schemaVersion: 1
    commandId: string
    expectedDocumentEpoch: number
    expectedChangeSeq: number
    expectedDocumentSha256: string
    target: HangulParagraphTarget
    expectedBeforeSha256: string
    expectedFormatSha256: string
    expectedAdjacentContextSha256: string
    replacement: string
  }): Promise<{ target: HangulParagraphTarget }>
  focusTarget?(target: HangulParagraphTarget): Promise<unknown>
  _request?(method: string, params?: Record<string, unknown>): Promise<unknown>
}

export interface HangulParagraphTarget {
  kind: 'body_paragraph'
  section: number
  paragraph: number
  charOffset: 0
  length: number
}

export interface PreparedParagraph {
  editable: boolean
  reason: string | null
  target: HangulParagraphTarget | null
  text: string | null
  textSha256: string | null
  formatSha256: string | null
  adjacentContextSha256: string | null
  selectionStart: number | null
  selectionEnd: number | null
}

export function clipPlainText(text: string, max = PLAIN_TEXT_MAX_CHARS): string {
  if (text.length <= max) return text
  return `${text.slice(0, max)}\n\n[truncated]`
}

function decodeCodePoint(value: number): string {
  if (!Number.isInteger(value) || value < 0 || value > 0x10ffff) return ''
  if (value >= 0xd800 && value <= 0xdfff) return ''
  return String.fromCodePoint(value)
}

export function stripHmlToPlainText(xml: string): string {
  return xml
    .replace(/<\?xml[\s\S]*?\?>/g, '')
    .replace(/<!(?:DOCTYPE|-- )[\s\S]*?>/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, n: string) => decodeCodePoint(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n: string) => decodeCodePoint(parseInt(n, 16)))
    .replace(/&amp;/gi, '&')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim()
}

function asText(value: unknown): string | null {
  if (typeof value === 'string' && value.length > 0) return value
  if (value && typeof value === 'object' && 'result' in value) {
    const result = (value as { result: unknown }).result
    if (typeof result === 'string' && result.length > 0) return result
  }
  return null
}

async function tryHwpctrl(
  studio: StudioTextSource,
  method: string,
  args?: unknown[],
): Promise<string | null> {
  try {
    return asText(await studio.hwpctrl.call(method, args))
  } catch {
    return null
  }
}

async function plainTextFromHml(studio: StudioTextSource): Promise<string | null> {
  if (studio.getHmlSaveState) {
    try {
      const state = await studio.getHmlSaveState()
      if (!state.hmlSavable) return null
    } catch {
      /* still try exportHml — some hosts omit the state API */
    }
  }
  try {
    return stripHmlToPlainText(new TextDecoder('utf-8').decode(await studio.exportHml()))
  } catch {
    return null
  }
}

export async function getPlainText(studio: StudioTextSource): Promise<string> {
  const fromCtrl = await tryHwpctrl(studio, 'GetTextFile', ['TEXT', ''])
  if (fromCtrl) return clipPlainText(fromCtrl)
  const fromHml = await plainTextFromHml(studio)
  if (fromHml !== null) return clipPlainText(fromHml)
  throw new Error(PLAIN_TEXT_UNAVAILABLE)
}

export async function getSelectionText(studio: StudioTextSource): Promise<string | null> {
  const fromCtrl = await tryHwpctrl(studio, 'GetTextFile', ['TEXT', 'saveblock'])
  return fromCtrl ? clipPlainText(fromCtrl) : null
}

export async function readSelectionState(studio: StudioTextSource): Promise<HangulSelectionState> {
  try {
    const sel = await studio.getSelectionContext()
    const page = typeof sel.page === 'number' && sel.page > 0 ? sel.page : null
    return { page, hasSelection: !sel.collapsed && Boolean(sel.selectedTextSha256) }
  } catch {
    return { page: null, hasSelection: false }
  }
}

export async function hasSelection(studio: StudioTextSource): Promise<boolean> {
  return (await readSelectionState(studio)).hasSelection
}

export async function currentPage(studio: StudioTextSource): Promise<number | null> {
  return (await readSelectionState(studio)).page
}

export function fileNameOf(path: string | null): string {
  if (!path) return 'untitled.hwp'
  const parts = path.split(/[\\/]/)
  return parts[parts.length - 1] ?? path
}

export function splitInsertParagraphs(text: string): string[] {
  const normalized = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  const lines = normalized.split('\n')
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop()
  if (lines.length === 0 || (lines.length === 1 && lines[0] === '')) {
    throw new Error(INSERT_CONTENT_EMPTY)
  }
  return lines.map((line) => normalizeReplacement(line))
}

export function normalizeReplacement(
  text: string,
  max = PARAGRAPH_MAX_CODE_POINTS,
): string {
  if (/[\u0000-\u001f\u007f]/u.test(text)) {
    throw new Error('replacement must not contain control characters')
  }
  if (Array.from(text).length > max) {
    throw new Error(`replacement must be at most ${max} characters`)
  }
  return text
}

export function spliceParagraphText(
  paragraph: string,
  start: number,
  end: number,
  insert: string,
): string {
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > paragraph.length
  ) {
    throw new Error(SELECTION_NOT_IN_PARAGRAPH)
  }
  return `${paragraph.slice(0, start)}${insert}${paragraph.slice(end)}`
}

function requestStudio(
  studio: StudioTextSource,
  method: string,
  params?: Record<string, unknown>,
): Promise<unknown> {
  if (typeof studio._request !== 'function') throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  return studio._request(method, params)
}

function asPrepared(value: unknown): PreparedParagraph {
  if (!value || typeof value !== 'object') throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  const raw = value as Partial<PreparedParagraph>
  const target = raw.target
  const okTarget =
    target &&
    target.kind === 'body_paragraph' &&
    Number.isSafeInteger(target.section) &&
    Number.isSafeInteger(target.paragraph)
  return {
    editable: raw.editable === true,
    reason: typeof raw.reason === 'string' ? raw.reason : null,
    target: okTarget
      ? {
          kind: 'body_paragraph',
          section: target.section,
          paragraph: target.paragraph,
          charOffset: 0,
          length: Number.isSafeInteger(target.length) ? target.length : 0,
        }
      : null,
    text: typeof raw.text === 'string' ? raw.text : null,
    textSha256: typeof raw.textSha256 === 'string' ? raw.textSha256 : null,
    formatSha256: typeof raw.formatSha256 === 'string' ? raw.formatSha256 : null,
    adjacentContextSha256:
      typeof raw.adjacentContextSha256 === 'string' ? raw.adjacentContextSha256 : null,
    selectionStart: Number.isInteger(raw.selectionStart) ? (raw.selectionStart as number) : null,
    selectionEnd: Number.isInteger(raw.selectionEnd) ? (raw.selectionEnd as number) : null,
  }
}

export async function prepareCurrentParagraph(studio: StudioTextSource): Promise<PreparedParagraph> {
  return asPrepared(await requestStudio(studio, 'prepareTextCommand'))
}

async function applyPrepared(
  studio: StudioTextSource,
  prepared: PreparedParagraph,
  replacement: string,
): Promise<{ before: string; after: string }> {
  const after = normalizeReplacement(replacement)
  if (!studio.getDocumentState || !studio.applyTextCommand) {
    throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  }
  if (
    !prepared.editable ||
    !prepared.target ||
    !prepared.textSha256 ||
    !prepared.formatSha256 ||
    !prepared.adjacentContextSha256
  ) {
    throw new Error(prepared.reason || PARAGRAPH_NOT_EDITABLE)
  }
  const state = await studio.getDocumentState()
  const receipt = await studio.applyTextCommand({
    schemaVersion: 1,
    commandId: crypto.randomUUID(),
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target: prepared.target,
    expectedBeforeSha256: prepared.textSha256,
    expectedFormatSha256: prepared.formatSha256,
    expectedAdjacentContextSha256: prepared.adjacentContextSha256,
    replacement: after,
  })
  await focusReplacedParagraph(studio, receipt.target ?? prepared.target, after)
  return { before: prepared.text ?? '', after }
}

async function focusReplacedParagraph(
  studio: StudioTextSource,
  target: HangulParagraphTarget,
  after: string,
): Promise<void> {
  if (!studio.focusTarget) return
  try {
    // applyTextCommand's receipt.target keeps the pre-apply length. After the
    // write, studio's exact-length check throws TARGET_NOT_FOUND.
    await studio.focusTarget({ ...target, length: after.length })
  } catch {
    // Apply already committed. Losing caret focus must not fail the tool.
  }
}

export async function replaceCurrentParagraph(
  studio: StudioTextSource,
  replacement: string,
): Promise<{ before: string; after: string }> {
  return applyPrepared(studio, await prepareCurrentParagraph(studio), replacement)
}

export async function replaceCurrentSelection(
  studio: StudioTextSource,
  replacement: string,
): Promise<{ before: string; after: string }> {
  const insert = normalizeReplacement(replacement)
  const prepared = await prepareCurrentParagraph(studio)
  if (
    prepared.selectionStart == null ||
    prepared.selectionEnd == null ||
    prepared.selectionEnd <= prepared.selectionStart
  ) {
    throw new Error(SELECTION_NOT_IN_PARAGRAPH)
  }
  const paragraph = prepared.text ?? ''
  const selected = paragraph.slice(prepared.selectionStart, prepared.selectionEnd)
  const next = spliceParagraphText(
    paragraph,
    prepared.selectionStart,
    prepared.selectionEnd,
    insert,
  )
  await applyPrepared(studio, prepared, next)
  return { before: selected, after: insert }
}

export async function listBodyParagraphs(
  studio: StudioTextSource,
): Promise<HangulParagraphPreview[]> {
  const raw = await requestStudio(studio, 'listBodyParagraphs')
  if (!Array.isArray(raw)) throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  return raw.map((item, index) => {
    const prepared = asPrepared(item)
    return {
      index,
      editable: prepared.editable,
      reason: prepared.reason,
      section: prepared.target?.section ?? -1,
      paragraph: prepared.target?.paragraph ?? -1,
      text: prepared.text ?? '',
    }
  })
}

export async function replaceParagraphAt(
  studio: StudioTextSource,
  index: number,
  replacement: string,
): Promise<{ before: string; after: string }> {
  const raw = await requestStudio(studio, 'listBodyParagraphs')
  if (!Array.isArray(raw)) throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  if (!raw[index]) throw new Error(PARAGRAPH_INDEX_OUT_OF_RANGE)
  return applyPrepared(studio, asPrepared(raw[index]), replacement)
}

function paragraphIsEmpty(item: HangulParagraphPreview): boolean {
  return !item.text.trim()
}

async function caretListIndex(
  studio: StudioTextSource,
  items: HangulParagraphPreview[],
): Promise<number | null> {
  try {
    const prepared = await prepareCurrentParagraph(studio)
    if (!prepared.target) return null
    const caret = items.findIndex(
      (item) =>
        item.section === prepared.target!.section && item.paragraph === prepared.target!.paragraph,
    )
    return caret >= 0 ? caret : null
  } catch {
    return null
  }
}

/** One studio RPC: insert empties and applyTextCommand each line inside the agent. */
async function fillInsertedParagraphs(
  studio: StudioTextSource,
  section: number,
  insertAt: number,
  writeFrom: number,
  lines: string[],
): Promise<number> {
  const inserted = await requestStudio(studio, 'insertFilledParagraphs', {
    section,
    index: insertAt,
    texts: lines,
  })
  if (!inserted || typeof inserted !== 'object') throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  return writeFrom
}

async function resolveInsertAnchor(
  studio: StudioTextSource,
  afterIndex?: number,
): Promise<{
  items: HangulParagraphPreview[]
  fillIndex: number | null
  section: number
  insertAt: number
  writeFrom: number
}> {
  const items = await listBodyParagraphs(studio)
  let fillIndex: number | null = null
  let section = items[0]?.section ?? 0
  let insertAt = items.length > 0 ? items[items.length - 1]!.paragraph + 1 : 0
  let writeFrom = items.length

  if (afterIndex === -1) {
    if (items[0]?.editable && paragraphIsEmpty(items[0])) {
      fillIndex = 0
      section = items[0].section
      insertAt = items[0].paragraph + 1
      writeFrom = 1
    } else {
      section = items[0]?.section ?? 0
      insertAt = 0
      writeFrom = 0
    }
  } else if (afterIndex == null) {
    const caret = await caretListIndex(studio, items)
    if (caret != null && items[caret]!.editable && paragraphIsEmpty(items[caret]!)) {
      fillIndex = caret
      section = items[caret]!.section
      insertAt = items[caret]!.paragraph + 1
      writeFrom = caret + 1
    } else if (caret != null) {
      section = items[caret]!.section
      insertAt = items[caret]!.paragraph + 1
      writeFrom = caret + 1
    } else {
      const last = items[items.length - 1]
      section = last?.section ?? 0
      insertAt = last != null ? last.paragraph + 1 : 0
      writeFrom = items.length
    }
  } else {
    if (!Number.isInteger(afterIndex) || afterIndex < 0 || afterIndex >= items.length) {
      throw new Error(PARAGRAPH_INDEX_OUT_OF_RANGE)
    }
    section = items[afterIndex]!.section
    insertAt = items[afterIndex]!.paragraph + 1
    writeFrom = afterIndex + 1
  }

  return { items, fillIndex, section, insertAt, writeFrom }
}

export async function insertContent(
  studio: StudioTextSource,
  text: string,
  afterIndex?: number,
): Promise<{ count: number; start: number }> {
  const lines = splitInsertParagraphs(text)
  const { fillIndex, section, insertAt, writeFrom } = await resolveInsertAnchor(studio, afterIndex)

  let remaining = lines
  let start = fillIndex ?? writeFrom
  if (fillIndex != null) {
    await replaceParagraphAt(studio, fillIndex, lines[0]!)
    remaining = lines.slice(1)
  }

  if (remaining.length > 0) {
    const filledStart = await fillInsertedParagraphs(
      studio,
      section,
      insertAt,
      writeFrom,
      remaining,
    )
    if (fillIndex == null) start = filledStart
  }

  return { count: lines.length, start }
}

function asFields(value: unknown): HangulField[] {
  if (!Array.isArray(value)) throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  return value
    .map((item) => {
      const raw = item && typeof item === 'object' ? (item as Partial<HangulField>) : {}
      return {
        name: typeof raw.name === 'string' ? raw.name : '',
        value: typeof raw.value === 'string' ? raw.value : '',
        type: typeof raw.type === 'string' ? raw.type : null,
      }
    })
    .filter((field) => field.name)
}

export async function listDocumentFields(studio: StudioTextSource): Promise<HangulField[]> {
  return asFields(await requestStudio(studio, 'listFields'))
}

export async function setDocumentField(
  studio: StudioTextSource,
  name: string,
  value: string,
): Promise<{ name: string; before: string; after: string }> {
  const fieldName = name.trim()
  if (!fieldName) throw new Error(FIELD_NOT_FOUND)
  const after = normalizeReplacement(value, FIELD_MAX_CODE_POINTS)
  const fields = await listDocumentFields(studio)
  const current = fields.find((field) => field.name === fieldName)
  try {
    await requestStudio(studio, 'setField', { name: fieldName, value: after })
  } catch {
    try {
      await studio.hwpctrl.call('PutFieldText', [fieldName, after])
    } catch {
      throw new Error(current ? PARAGRAPH_PREPARE_UNAVAILABLE : FIELD_NOT_FOUND)
    }
  }
  return { name: fieldName, before: current?.value ?? '', after }
}

function asTables(value: unknown): HangulTable[] {
  if (!Array.isArray(value)) throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  return value.map((item, index) => {
    const raw = item && typeof item === 'object' ? (item as Record<string, unknown>) : {}
    const cells = Array.isArray(raw.cells)
      ? raw.cells.map((cell, cellIndex) => {
          const row = cell && typeof cell === 'object' ? (cell as Record<string, unknown>) : {}
          return {
            index: Number.isInteger(row.index) ? (row.index as number) : cellIndex,
            row: Number.isInteger(row.row) ? (row.row as number) : 0,
            col: Number.isInteger(row.col) ? (row.col as number) : 0,
            text: typeof row.text === 'string' ? row.text : '',
          }
        })
      : []
    return {
      index,
      section: Number.isInteger(raw.section) ? (raw.section as number) : 0,
      paragraph: Number.isInteger(raw.paragraph) ? (raw.paragraph as number) : 0,
      control: Number.isInteger(raw.control) ? (raw.control as number) : 0,
      rows: Number.isInteger(raw.rows) ? (raw.rows as number) : 0,
      cols: Number.isInteger(raw.cols) ? (raw.cols as number) : 0,
      cells,
    }
  })
}

export async function listDocumentTables(studio: StudioTextSource): Promise<HangulTable[]> {
  return asTables(await requestStudio(studio, 'listTables'))
}

export async function replaceTableCell(
  studio: StudioTextSource,
  tableIndex: number,
  row: number,
  col: number,
  replacement: string,
): Promise<{ before: string; after: string }> {
  const after = normalizeReplacement(replacement, FIELD_MAX_CODE_POINTS)
  const tables = await listDocumentTables(studio)
  const table = tables[tableIndex]
  const cell = table?.cells.find((item) => item.row === row && item.col === col)
  if (!table || !cell) throw new Error(TABLE_NOT_FOUND)
  await requestStudio(studio, 'replaceCell', {
    section: table.section,
    paragraph: table.paragraph,
    control: table.control,
    cellIndex: cell.index,
    text: after,
  })
  return { before: cell.text, after }
}

function requireTableSize(value: unknown, label: string, max: number): number {
  if (!Number.isInteger(value) || (value as number) < 1) throw new Error(TABLE_SIZE_INVALID)
  if ((value as number) > max) throw new Error(`table ${label} must be at most ${max}`)
  return value as number
}

function asTableCells(value: unknown): string[][] | undefined {
  if (value == null) return undefined
  if (!Array.isArray(value)) throw new Error('table cells must be an array of rows')
  return value.map((row) => {
    if (!Array.isArray(row)) throw new Error('table cells must be an array of rows')
    return row.map((cell) => (cell == null ? '' : String(cell)))
  })
}

function asInsertedTableLoc(
  value: unknown,
  fallbackSection: number,
  fallbackParagraph: number,
): { section: number; paragraph: number; control: number } {
  const raw = value && typeof value === 'object' ? (value as Record<string, unknown>) : {}
  const section = Number.isInteger(raw.section) ? (raw.section as number) : fallbackSection
  const paragraph = Number.isInteger(raw.paragraph)
    ? (raw.paragraph as number)
    : Number.isInteger(raw.paraIdx)
      ? (raw.paraIdx as number)
      : fallbackParagraph
  const control = Number.isInteger(raw.control)
    ? (raw.control as number)
    : Number.isInteger(raw.controlIdx)
      ? (raw.controlIdx as number)
      : 0
  return { section, paragraph, control }
}

export async function insertDocumentTable(
  studio: StudioTextSource,
  rows: number,
  cols: number,
  cells?: string[][],
  afterIndex?: number,
): Promise<{ table: number; rows: number; cols: number; unfilled: string[] }> {
  const rowCount = requireTableSize(rows, 'rows', TABLE_MAX_ROWS)
  const colCount = requireTableSize(cols, 'cols', TABLE_MAX_COLS)
  const fill = asTableCells(cells)
  const { fillIndex, section, insertAt } = await resolveInsertAnchor(studio, afterIndex)
  let paragraph = fillIndex != null ? insertAt - 1 : insertAt
  if (fillIndex == null) {
    await requestStudio(studio, 'insertBodyParagraphs', {
      section,
      index: insertAt,
      count: 1,
    })
    paragraph = insertAt
  }
  const inserted = asInsertedTableLoc(
    await requestStudio(studio, 'insertTable', {
      section,
      index: paragraph,
      rows: rowCount,
      cols: colCount,
    }),
    section,
    paragraph,
  )
  const tables = await listDocumentTables(studio)
  const created = tables.find(
    (table) =>
      table.section === inserted.section &&
      table.paragraph === inserted.paragraph &&
      table.control === inserted.control,
  )
  if (!created) throw new Error(TABLE_NOT_FOUND)
  const unfilled: string[] = []
  let attempted = 0
  let firstError = ''
  if (fill) {
    for (let row = 0; row < Math.min(fill.length, rowCount); row += 1) {
      const line = fill[row] ?? []
      for (let col = 0; col < Math.min(line.length, colCount); col += 1) {
        const text = line[col] ?? ''
        if (!text) continue
        attempted += 1
        const cell = created.cells.find((item) => item.row === row && item.col === col)
        if (!cell) {
          unfilled.push(`r${row}c${col}`)
          if (!firstError) firstError = TABLE_NOT_FOUND
          continue
        }
        try {
          await requestStudio(studio, 'replaceCell', {
            section: created.section,
            paragraph: created.paragraph,
            control: created.control,
            cellIndex: cell.index,
            text: normalizeReplacement(text, FIELD_MAX_CODE_POINTS),
          })
        } catch (err) {
          unfilled.push(`r${row}c${col}`)
          if (!firstError) firstError = err instanceof Error ? err.message : String(err)
        }
      }
    }
  }
  if (attempted > 0 && unfilled.length === attempted) {
    throw new Error(
      `${TABLE_CELLS_UNFILLED} (${unfilled.join(', ')}): ${firstError || 'unknown error'}. The table is already in the document — do not insert another.`,
    )
  }
  return { table: created.index, rows: rowCount, cols: colCount, unfilled }
}

const ALIGN_VALUES = new Set<HangulAlign>(['left', 'center', 'right', 'justify'])
const LIST_VALUES = new Set<HangulListStyle>(['none', 'bullet', 'number'])

function parseColor(value: string): string {
  const raw = value.trim()
  const hex = raw.startsWith('#') ? raw.slice(1) : raw
  if (/^[0-9a-fA-F]{3}$/.test(hex)) {
    const [r, g, b] = hex
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase()
  }
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#${hex.toLowerCase()}`
  throw new Error('color must be a 3- or 6-digit hex value')
}

function parsePointMeasure(value: unknown, label: string, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`${label} must be a number of points`)
  }
  if (value < min || value > max) {
    throw new Error(`${label} must be between ${min} and ${max}`)
  }
  return value
}

function parseLineSpacingPercent(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    throw new Error('lineSpacing must be a multiplier (1.5) or a percent (150)')
  }
  const percent = value <= 5 ? Math.round(value * 100) : Math.round(value)
  if (percent < LINE_SPACING_MIN || percent > LINE_SPACING_MAX) {
    throw new Error(`lineSpacing must be between ${LINE_SPACING_MIN / 100} and ${LINE_SPACING_MAX / 100}`)
  }
  return percent
}

function charFormatPayload(format: HangulFormatSpec): Record<string, unknown> | null {
  const payload: Record<string, unknown> = {}
  if (typeof format.bold === 'boolean') payload.bold = format.bold
  if (typeof format.italic === 'boolean') payload.italic = format.italic
  if (typeof format.underline === 'boolean') payload.underline = format.underline
  if (typeof format.strikethrough === 'boolean') payload.strikethrough = format.strikethrough
  if (format.fontSize != null) {
    const pt = parsePointMeasure(format.fontSize, 'fontSize', FONT_SIZE_MIN_PT, FONT_SIZE_MAX_PT)
    payload.fontSize = Math.round(pt * 100)
  }
  if (format.color != null) {
    if (typeof format.color !== 'string') throw new Error('color must be a 3- or 6-digit hex value')
    payload.textColor = parseColor(format.color)
  }
  if (format.font != null) {
    const name = format.font.trim()
    if (!name) throw new Error('font must not be empty')
    payload.fontName = name
  }
  return Object.keys(payload).length > 0 ? payload : null
}

function paraFormatPayload(format: HangulFormatSpec): Record<string, unknown> | null {
  const payload: Record<string, unknown> = {}
  if (format.align != null) {
    if (!ALIGN_VALUES.has(format.align)) throw new Error('align must be left, center, right, or justify')
    payload.alignment = format.align
  }
  if (format.list != null) {
    if (!LIST_VALUES.has(format.list)) throw new Error('list must be none, bullet, or number')
    if (format.list === 'none') payload.headType = 'None'
    else if (format.list === 'bullet') payload.headType = 'Bullet'
    else payload.headType = 'Number'
  }
  if (format.lineSpacing != null) {
    payload.lineSpacing = parseLineSpacingPercent(format.lineSpacing)
    payload.lineSpacingType = 'Percent'
  }
  if (format.indentLeft != null) {
    payload.marginLeft = Math.round(
      parsePointMeasure(format.indentLeft, 'indentLeft', 0, INDENT_MAX_PT) * PARA_PT_TO_UNIT,
    )
  }
  if (format.indentRight != null) {
    payload.marginRight = Math.round(
      parsePointMeasure(format.indentRight, 'indentRight', 0, INDENT_MAX_PT) * PARA_PT_TO_UNIT,
    )
  }
  if (format.indentFirstLine != null) {
    const pt = parsePointMeasure(format.indentFirstLine, 'indentFirstLine', -INDENT_MAX_PT, INDENT_MAX_PT)
    payload.indent = Math.round(pt * PARA_PT_TO_UNIT)
  }
  return Object.keys(payload).length > 0 ? payload : null
}

function formatAppliedLabels(format: HangulFormatSpec): string[] {
  const applied: string[] = []
  if (typeof format.bold === 'boolean') applied.push(`bold=${format.bold}`)
  if (typeof format.italic === 'boolean') applied.push(`italic=${format.italic}`)
  if (typeof format.underline === 'boolean') applied.push(`underline=${format.underline}`)
  if (typeof format.strikethrough === 'boolean') applied.push(`strikethrough=${format.strikethrough}`)
  if (format.fontSize != null) {
    applied.push(
      `fontSize=${parsePointMeasure(format.fontSize, 'fontSize', FONT_SIZE_MIN_PT, FONT_SIZE_MAX_PT)}`,
    )
  }
  if (format.color) applied.push(`color=${parseColor(format.color)}`)
  if (format.font) applied.push(`font=${format.font}`)
  if (format.align) applied.push(`align=${format.align}`)
  if (format.list) applied.push(`list=${format.list}`)
  if (format.lineSpacing != null) applied.push(`lineSpacing=${parseLineSpacingPercent(format.lineSpacing) / 100}`)
  if (format.indentLeft != null) {
    applied.push(`indentLeft=${parsePointMeasure(format.indentLeft, 'indentLeft', 0, INDENT_MAX_PT)}`)
  }
  if (format.indentRight != null) {
    applied.push(`indentRight=${parsePointMeasure(format.indentRight, 'indentRight', 0, INDENT_MAX_PT)}`)
  }
  if (format.indentFirstLine != null) {
    applied.push(
      `indentFirstLine=${parsePointMeasure(format.indentFirstLine, 'indentFirstLine', -INDENT_MAX_PT, INDENT_MAX_PT)}`,
    )
  }
  return applied
}

function assertFormatable(prepared: PreparedParagraph): HangulParagraphTarget {
  if (!prepared.target) throw new Error(prepared.reason || PARAGRAPH_NOT_EDITABLE)
  if (prepared.editable) return prepared.target
  const reason = prepared.reason ?? ''
  if (/mixed/i.test(reason)) return prepared.target
  throw new Error(reason || PARAGRAPH_NOT_EDITABLE)
}

async function applyFormatToParagraph(
  studio: StudioTextSource,
  prepared: PreparedParagraph,
  format: HangulFormatSpec,
  useSelection: boolean,
): Promise<void> {
  const target = assertFormatable(prepared)
  const char = charFormatPayload(format)
  const para = paraFormatPayload(format)
  let start = 0
  let end = target.length
  if (
    useSelection &&
    prepared.selectionStart != null &&
    prepared.selectionEnd != null &&
    prepared.selectionEnd > prepared.selectionStart
  ) {
    start = prepared.selectionStart
    end = prepared.selectionEnd
  }
  if (char && end > start) {
    await requestStudio(studio, 'applyBodyCharFormat', {
      section: target.section,
      paragraph: target.paragraph,
      start,
      end,
      format: char,
    })
  }
  if (para) {
    await requestStudio(studio, 'applyBodyParaFormat', {
      section: target.section,
      paragraph: target.paragraph,
      format: para,
    })
  }
  if ((!char || end <= start) && !para) throw new Error(FORMAT_EMPTY_RANGE)
}

export async function applyParagraphFormat(
  studio: StudioTextSource,
  format: HangulFormatSpec,
  index?: number,
  indexes?: number[],
): Promise<{ indexes: number[]; applied: string[] }> {
  const applied = formatAppliedLabels(format)
  if (applied.length === 0) throw new Error(FORMAT_EMPTY)
  const raw = await requestStudio(studio, 'listBodyParagraphs')
  if (!Array.isArray(raw)) throw new Error(PARAGRAPH_PREPARE_UNAVAILABLE)
  const items = raw.map((item) => asPrepared(item))
  let targets: number[]
  if (indexes != null) {
    if (!Array.isArray(indexes) || indexes.length === 0) {
      throw new Error('indexes must be a non-empty array of integers')
    }
    targets = indexes
  } else if (index != null) {
    targets = [index]
  } else {
    const current = await prepareCurrentParagraph(studio)
    await applyFormatToParagraph(studio, current, format, true)
    const caret = items.findIndex(
      (item) =>
        item.target?.section === current.target!.section &&
        item.target?.paragraph === current.target!.paragraph,
    )
    return { indexes: [caret >= 0 ? caret : 0], applied }
  }
  const formatted: number[] = []
  let lockedReason: string | null = null
  for (const targetIndex of targets) {
    if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= items.length) {
      throw new Error(PARAGRAPH_INDEX_OUT_OF_RANGE)
    }
    const prepared = items[targetIndex]!
    try {
      await applyFormatToParagraph(studio, prepared, format, false)
      formatted.push(targetIndex)
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err)
      if (isLockedFormatError(message) || (indexes != null && message === FORMAT_EMPTY_RANGE)) {
        lockedReason = message
        continue
      }
      throw err
    }
  }
  if (formatted.length === 0) {
    const first = items.findIndex((item) => {
      try {
        assertFormatable(item)
        return true
      } catch {
        return false
      }
    })
    throw new Error(
      `${lockedReason || PARAGRAPH_NOT_EDITABLE}. First editable paragraph is [${first < 0 ? 'none' : first}]`,
    )
  }
  return { indexes: formatted, applied }
}

export async function applyTableCellFormat(
  studio: StudioTextSource,
  format: HangulFormatSpec,
  target: HangulCellFormatTarget,
): Promise<{ table: number; row: number; cols: number[]; applied: string[] }> {
  if (format.list != null) throw new Error('list is not supported in table cells')
  const applied = formatAppliedLabels(format)
  if (applied.length === 0) throw new Error(FORMAT_EMPTY)
  const tables = await listDocumentTables(studio)
  const table = tables[target.table]
  if (!table) throw new Error(TABLE_NOT_FOUND)
  if (!Number.isInteger(target.row) || target.row < 0 || target.row >= table.rows) {
    throw new Error(TABLE_NOT_FOUND)
  }
  if (target.col != null && (!Number.isInteger(target.col) || target.col < 0 || target.col >= table.cols)) {
    throw new Error(TABLE_NOT_FOUND)
  }
  const cells = table.cells.filter(
    (cell) => cell.row === target.row && (target.col == null || cell.col === target.col),
  )
  if (cells.length === 0) throw new Error(TABLE_NOT_FOUND)
  const char = charFormatPayload(format)
  const para = paraFormatPayload(format)
  const cols: number[] = []
  for (const cell of cells) {
    const first = cell.text.split('\n')[0] ?? ''
    const end = first.length
    if (char && end > 0) {
      await requestStudio(studio, 'applyCellCharFormat', {
        section: table.section,
        paragraph: table.paragraph,
        control: table.control,
        cellIndex: cell.index,
        cellPara: 0,
        start: 0,
        end,
        format: char,
      })
    }
    if (para) {
      await requestStudio(studio, 'applyCellParaFormat', {
        section: table.section,
        paragraph: table.paragraph,
        control: table.control,
        cellIndex: cell.index,
        cellPara: 0,
        format: para,
      })
    }
    if ((char && end > 0) || para) cols.push(cell.col)
  }
  if (cols.length === 0) throw new Error(FORMAT_EMPTY_RANGE)
  return { table: target.table, row: target.row, cols, applied }
}

async function requireTable(studio: StudioTextSource, tableIndex: number): Promise<HangulTable> {
  const tables = await listDocumentTables(studio)
  const table = tables[tableIndex]
  if (!table) throw new Error(TABLE_NOT_FOUND)
  return table
}

async function tableEditResult(
  studio: StudioTextSource,
  table: number,
  action: HangulTableEditAction,
  detail: string,
): Promise<HangulTableEditResult> {
  try {
    const next = await requireTable(studio, table)
    return {
      table,
      action,
      detail,
      rows: next.rows,
      cols: next.cols,
      cells: next.cells.map((cell) => ({ row: cell.row, col: cell.col })),
    }
  } catch {
    return { table, action, detail }
  }
}

function mmToHwp(mm: number, label: string, min: number, max: number): number {
  if (!Number.isFinite(mm)) throw new Error(`${label} must be a number of millimeters`)
  if (mm < min || mm > max) throw new Error(`${label} must be between ${min} and ${max} mm`)
  return Math.round(mm * HWP_PER_MM)
}

export async function editDocumentTable(
  studio: StudioTextSource,
  spec: HangulTableEditSpec,
): Promise<HangulTableEditResult> {
  const table = await requireTable(studio, spec.table)
  const loc = { section: table.section, paragraph: table.paragraph, control: table.control }
  const after = spec.after !== false
  if (spec.action === 'insert_row') {
    if (!Number.isInteger(spec.row) || spec.row! < 0 || spec.row! >= table.rows) {
      throw new Error(TABLE_NOT_FOUND)
    }
    if (table.rows + 1 > TABLE_MAX_ROWS) throw new Error(`table rows must be at most ${TABLE_MAX_ROWS}`)
    await requestStudio(studio, 'insertTableRow', { ...loc, row: spec.row, after })
    return tableEditResult(
      studio,
      spec.table,
      spec.action,
      after ? `row after ${spec.row}` : `row before ${spec.row}`,
    )
  }
  if (spec.action === 'insert_column') {
    if (!Number.isInteger(spec.col) || spec.col! < 0 || spec.col! >= table.cols) {
      throw new Error(TABLE_NOT_FOUND)
    }
    if (table.cols + 1 > TABLE_MAX_COLS) throw new Error(`table columns must be at most ${TABLE_MAX_COLS}`)
    await requestStudio(studio, 'insertTableColumn', { ...loc, col: spec.col, after })
    return tableEditResult(
      studio,
      spec.table,
      spec.action,
      after ? `column after ${spec.col}` : `column before ${spec.col}`,
    )
  }
  if (spec.action === 'delete_row') {
    if (!Number.isInteger(spec.row) || spec.row! < 0 || spec.row! >= table.rows) {
      throw new Error(TABLE_NOT_FOUND)
    }
    if (table.rows <= 1) throw new Error('cannot delete the last table row')
    await requestStudio(studio, 'deleteTableRow', { ...loc, row: spec.row })
    return tableEditResult(studio, spec.table, spec.action, `row ${spec.row}`)
  }
  if (spec.action === 'delete_column') {
    if (!Number.isInteger(spec.col) || spec.col! < 0 || spec.col! >= table.cols) {
      throw new Error(TABLE_NOT_FOUND)
    }
    if (table.cols <= 1) throw new Error('cannot delete the last table column')
    await requestStudio(studio, 'deleteTableColumn', { ...loc, col: spec.col })
    return tableEditResult(studio, spec.table, spec.action, `column ${spec.col}`)
  }
  if (spec.action === 'merge') {
    const startRow = spec.row
    const startCol = spec.col
    const endRow = spec.endRow
    const endCol = spec.endCol
    if (
      !Number.isInteger(startRow) ||
      !Number.isInteger(startCol) ||
      !Number.isInteger(endRow) ||
      !Number.isInteger(endCol)
    ) {
      throw new Error('merge needs row, col, endRow, and endCol')
    }
    const r1 = Math.min(startRow!, endRow!)
    const r2 = Math.max(startRow!, endRow!)
    const c1 = Math.min(startCol!, endCol!)
    const c2 = Math.max(startCol!, endCol!)
    if (r1 < 0 || c1 < 0 || r2 >= table.rows || c2 >= table.cols) throw new Error(TABLE_NOT_FOUND)
    if (r1 === r2 && c1 === c2) throw new Error('merge needs more than one cell')
    await requestStudio(studio, 'mergeTableCells', {
      ...loc,
      startRow: r1,
      startCol: c1,
      endRow: r2,
      endCol: c2,
    })
    return tableEditResult(studio, spec.table, spec.action, `r${r1}c${c1}:r${r2}c${c2}`)
  }
  if (spec.action === 'split') {
    if (!Number.isInteger(spec.row) || !Number.isInteger(spec.col)) {
      throw new Error('split needs row and col')
    }
    if (spec.row! < 0 || spec.col! < 0 || spec.row! >= table.rows || spec.col! >= table.cols) {
      throw new Error(TABLE_NOT_FOUND)
    }
    const rows = spec.splitRows ?? 2
    const cols = spec.splitCols ?? 1
    if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1 || rows > 10 || cols > 10) {
      throw new Error('split rows/cols must be integers from 1 to 10')
    }
    if (rows === 1 && cols === 1) throw new Error('split needs rows or cols greater than 1')
    await requestStudio(studio, 'splitTableCellInto', {
      ...loc,
      row: spec.row,
      col: spec.col,
      rows,
      cols,
      equalHeight: true,
      mergeFirst: false,
    })
    return tableEditResult(
      studio,
      spec.table,
      spec.action,
      `r${spec.row}c${spec.col} into ${rows}x${cols}`,
    )
  }
  throw new Error(TABLE_EDIT_EMPTY)
}

function cellStylePayload(spec: HangulTableStyleSpec): Record<string, unknown> | null {
  const props: Record<string, unknown> = {}
  if (spec.fill != null) {
    const color = parseColor(spec.fill)
    props.fillType = 'solid'
    props.fillColor = color
  }
  if (spec.valign != null) {
    if (spec.valign !== 'top' && spec.valign !== 'center' && spec.valign !== 'bottom') {
      throw new Error('valign must be top, center, or bottom')
    }
    props.verticalAlign = spec.valign === 'top' ? 'Top' : spec.valign === 'center' ? 'Center' : 'Bottom'
  }
  if (spec.border === false) {
    const none = { type: 0, width: 0, color: '#000000' }
    props.borderLeft = none
    props.borderRight = none
    props.borderTop = none
    props.borderBottom = none
  } else if (spec.border != null) {
    const side = { type: 1, width: 8, color: parseColor(spec.border) }
    props.borderLeft = side
    props.borderRight = side
    props.borderTop = side
    props.borderBottom = side
  }
  return Object.keys(props).length > 0 ? props : null
}

export async function styleDocumentTable(
  studio: StudioTextSource,
  spec: HangulTableStyleSpec,
): Promise<{ table: number; applied: string[] }> {
  const table = await requireTable(studio, spec.table)
  const loc = { section: table.section, paragraph: table.paragraph, control: table.control }
  const applied: string[] = []
  const cellProps = cellStylePayload(spec)
  if (cellProps) {
    if (spec.fill != null) applied.push(`fill=${parseColor(spec.fill)}`)
    if (spec.valign != null) applied.push(`valign=${spec.valign}`)
    if (spec.border === false) applied.push('border=none')
    else if (spec.border != null) applied.push(`border=${parseColor(spec.border)}`)
    const cells = table.cells.filter(
      (cell) =>
        (spec.row == null || cell.row === spec.row) && (spec.col == null || cell.col === spec.col),
    )
    if (cells.length === 0) throw new Error(TABLE_NOT_FOUND)
    for (const cell of cells) {
      await requestStudio(studio, 'setCellProperties', {
        ...loc,
        cellIndex: cell.index,
        props: cellProps,
      })
    }
  }
  if (spec.width != null) {
    const current = await requestStudio(studio, 'getTableProperties', loc)
    const props =
      current && typeof current === 'object' ? { ...(current as Record<string, unknown>) } : {}
    props.tableWidth = mmToHwp(spec.width, 'width', TABLE_WIDTH_MIN_MM, TABLE_WIDTH_MAX_MM)
    await requestStudio(studio, 'setTableProperties', { ...loc, props })
    applied.push(`width=${spec.width}`)
  }
  if (applied.length === 0) throw new Error(TABLE_STYLE_EMPTY)
  return { table: spec.table, applied }
}

export async function setDocumentPage(
  studio: StudioTextSource,
  spec: HangulPageSetupSpec,
): Promise<{ applied: string[] }> {
  const applied: string[] = []
  const pageTouched =
    spec.orientation != null ||
    spec.paper != null ||
    spec.marginTop != null ||
    spec.marginBottom != null ||
    spec.marginLeft != null ||
    spec.marginRight != null
  if (pageTouched) {
    const current = await requestStudio(studio, 'getPageDef', { section: 0 })
    const props =
      current && typeof current === 'object' ? { ...(current as Record<string, unknown>) } : {}
    if (spec.paper != null) {
      const size = PAPER_SIZES[spec.paper]
      if (!size) throw new Error('paper must be A4, A3, B4, B5, Letter, or Legal')
      const landscape = spec.orientation === 'landscape' || (spec.orientation == null && props.landscape === true)
      props.width = landscape ? size[1] : size[0]
      props.height = landscape ? size[0] : size[1]
      props.landscape = landscape
      applied.push(`paper=${spec.paper}`)
    } else if (spec.orientation != null) {
      if (spec.orientation !== 'portrait' && spec.orientation !== 'landscape') {
        throw new Error('orientation must be portrait or landscape')
      }
      const landscape = spec.orientation === 'landscape'
      const width = Number(props.width) || PAPER_SIZES.A4[0]
      const height = Number(props.height) || PAPER_SIZES.A4[1]
      const already = props.landscape === true
      if (already !== landscape) {
        props.width = height
        props.height = width
      }
      props.landscape = landscape
      applied.push(`orientation=${spec.orientation}`)
    }
    if (spec.paper != null && spec.orientation != null) applied.push(`orientation=${spec.orientation}`)
    for (const [key, label] of [
      ['marginTop', 'marginTop'],
      ['marginBottom', 'marginBottom'],
      ['marginLeft', 'marginLeft'],
      ['marginRight', 'marginRight'],
    ] as const) {
      const value = spec[key]
      if (value == null) continue
      props[key] = mmToHwp(value, label, 0, PAGE_MARGIN_MAX_MM)
      applied.push(`${label}=${value}`)
    }
    await requestStudio(studio, 'setPageDef', { section: 0, props })
  }
  if (spec.columns != null || spec.columnSpacing != null) {
    const current = await requestStudio(studio, 'getColumnDef', { section: 0 })
    const prev = current && typeof current === 'object' ? (current as Record<string, unknown>) : {}
    const count = spec.columns ?? (Number.isInteger(prev.columnCount) ? (prev.columnCount as number) : 1)
    if (!Number.isInteger(count) || count < 1 || count > PAGE_COLUMNS_MAX) {
      throw new Error(`columns must be an integer from 1 to ${PAGE_COLUMNS_MAX}`)
    }
    const spacing =
      spec.columnSpacing != null
        ? mmToHwp(spec.columnSpacing, 'columnSpacing', 0, PAGE_MARGIN_MAX_MM)
        : Number(prev.spacing) || Math.round(5 * HWP_PER_MM)
    await requestStudio(studio, 'setColumnDef', {
      section: 0,
      count,
      columnType: Number(prev.columnType) || 0,
      sameWidth: prev.sameWidth !== false,
      spacing,
    })
    if (spec.columns != null) applied.push(`columns=${count}`)
    if (spec.columnSpacing != null) applied.push(`columnSpacing=${spec.columnSpacing}`)
  }
  if (applied.length === 0) throw new Error(PAGE_SETUP_EMPTY)
  return { applied }
}

function isLockedFormatError(message: string): boolean {
  return (
    message === PARAGRAPH_NOT_EDITABLE || /table|control|field|locked/i.test(message)
  )
}

export function createStudioFacade(studio: StudioTextSource): HangulStudioFacade {
  return {
    pageCount: () => studio.pageCount(),
    currentPage: () => currentPage(studio),
    readSelectionState: () => readSelectionState(studio),
    getPlainText: () => getPlainText(studio),
    getSelectionText: () => getSelectionText(studio),
    hasSelection: () => hasSelection(studio),
    listParagraphs: () => listBodyParagraphs(studio),
    insertContent: (text, afterIndex) => insertContent(studio, text, afterIndex),
    replaceParagraph: (text, index) =>
      index == null ? replaceCurrentParagraph(studio, text) : replaceParagraphAt(studio, index, text),
    replaceSelection: (text) => replaceCurrentSelection(studio, text),
    listFields: () => listDocumentFields(studio),
    setField: (name, value) => setDocumentField(studio, name, value),
    listTables: () => listDocumentTables(studio),
    replaceCell: (table, row, col, text) => replaceTableCell(studio, table, row, col, text),
    insertTable: (rows, cols, cells, afterIndex) =>
      insertDocumentTable(studio, rows, cols, cells, afterIndex),
    applyFormat: (format, index, indexes, cell) =>
      cell
        ? applyTableCellFormat(studio, format, cell).then((result) => ({
            indexes: result.cols,
            applied: result.applied,
            table: result.table,
            row: result.row,
          }))
        : applyParagraphFormat(studio, format, index, indexes),
    editTable: (spec) => editDocumentTable(studio, spec),
    styleTable: (spec) => styleDocumentTable(studio, spec),
    setPage: (spec) => setDocumentPage(studio, spec),
  }
}
