/** Cap on document / selection plain text returned to the model / UI. */
export const PLAIN_TEXT_MAX_CHARS = 80_000
/** Short selection preview attached to every user turn. */
export const SELECTION_PREVIEW_CHARS = 400
export const PLAIN_TEXT_UNAVAILABLE = 'plain text unavailable'
export const PARAGRAPH_PREPARE_UNAVAILABLE = 'paragraph prepare is unavailable'
export const PARAGRAPH_NOT_EDITABLE = 'paragraph is not editable'
export const SELECTION_NOT_IN_PARAGRAPH = 'nothing is selected in this paragraph'
export const PARAGRAPH_INDEX_OUT_OF_RANGE = 'paragraph index out of range'
export const FIELD_NOT_FOUND = 'field not found'
export const TABLE_NOT_FOUND = 'table cell not found'
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

export interface HangulStudioFacade {
  pageCount(): Promise<number>
  currentPage(): Promise<number | null>
  readSelectionState(): Promise<HangulSelectionState>
  getPlainText(): Promise<string>
  getSelectionText(): Promise<string | null>
  hasSelection(): Promise<boolean>
  listParagraphs(): Promise<HangulParagraphPreview[]>
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
  await studio.focusTarget?.(receipt.target)
  return { before: prepared.text ?? '', after }
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

export function createStudioFacade(studio: StudioTextSource): HangulStudioFacade {
  return {
    pageCount: () => studio.pageCount(),
    currentPage: () => currentPage(studio),
    readSelectionState: () => readSelectionState(studio),
    getPlainText: () => getPlainText(studio),
    getSelectionText: () => getSelectionText(studio),
    hasSelection: () => hasSelection(studio),
    listParagraphs: () => listBodyParagraphs(studio),
    replaceParagraph: (text, index) =>
      index == null ? replaceCurrentParagraph(studio, text) : replaceParagraphAt(studio, index, text),
    replaceSelection: (text) => replaceCurrentSelection(studio, text),
    listFields: () => listDocumentFields(studio),
    setField: (name, value) => setDocumentField(studio, name, value),
    listTables: () => listDocumentTables(studio),
    replaceCell: (table, row, col, text) => replaceTableCell(studio, table, row, col, text),
  }
}
