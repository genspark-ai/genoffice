/**
 * The engine behind F7 / Review ▸ Spelling: the frame's own spell checker
 * (Chromium Hunspell / the OS dictionary — the same engine docs' squiggles
 * run on), driven in batch over cell texts. Nothing leaves the machine: the
 * renderer tokenises each string cell, asks the frame which words are
 * misspelled, and takes its suggestions. The pure halves (collection,
 * tokenising, white-lists, replacement, rich-text rewriting) are
 * unit-tested without Univer or the frame.
 */
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

/** The frame's spell engine, as the preload bridges it from webFrame. */
export interface SpellEngine {
  isWordMisspelled(word: string): boolean
  getWordSuggestions(word: string): string[]
}

/** One string cell candidate for the checker (raw in-memory snapshot value). */
export interface SpellCellRef {
  readonly sheetId: string
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly column: number
  readonly text: string
}

/** One flagged word occurrence in one cell. */
export interface SpellFinding {
  readonly cell: SpellCellRef
  readonly word: string
  readonly suggestions: readonly string[]
}

/** Minimal structural slice of Univer's workbook snapshot (keeps tests Univer-free). */
export interface SpellSnapshotCell {
  readonly v?: unknown
  readonly f?: unknown
  readonly p?: unknown
}

export interface SpellSnapshotSheet {
  readonly name?: string | undefined
  readonly cellData?:
    | Readonly<
        Record<string, Readonly<Record<string, SpellSnapshotCell | null | undefined>> | undefined>
      >
    | undefined
}

export interface SpellSnapshot {
  readonly sheets: Readonly<Record<string, SpellSnapshotSheet | undefined>>
}

/** A text must have a letter for spelling to be a question at all. */
const HAS_LETTER = /\p{L}/u
/** Hard cap on cells per run; beyond it the dialog reports a truncation note. */
export const MAX_CELLS_PER_RUN = 5000
/** Words per suggestion list, Excel's dialog shows about this many. */
export const MAX_SUGGESTIONS = 5
/** A word longer than this is a code fragment, not a typo. */
const MAX_WORD_CHARS = 64

function cellText(cell: SpellSnapshotCell | null | undefined): string | null {
  // formula cells are skipped: the cached result is not text the user retypes
  if (!cell || (cell.f !== undefined && cell.f !== null && cell.f !== '')) return null
  const value = cell.v
  if (typeof value !== 'string') return null
  if (value.trim().length === 0 || !HAS_LETTER.test(value)) return null
  return value
}

/**
 * Every checkable string cell of the whole workbook in sheet order,
 * row-major within a sheet: plain string values only (a formula's cached
 * result is not text the user retypes), no withheld cells. Rich-text cells
 * take part — their runs survive a replacement (see replaceInRichDoc).
 */
export function collectWorkbookTextCells(
  snapshot: SpellSnapshot,
  sheetOrder: readonly { sheetId: string; sheetName: string }[],
  isWithheld?: (sheetId: string, row: number, column: number) => boolean,
  maxCells = MAX_CELLS_PER_RUN,
): { cells: SpellCellRef[]; truncated: boolean } {
  const cells: SpellCellRef[] = []
  let truncated = false
  for (const sheet of sheetOrder) {
    const cellData = snapshot.sheets[sheet.sheetId]?.cellData
    if (!cellData) continue
    const rows = Object.keys(cellData)
      .map(Number)
      .filter((row) => Number.isInteger(row) && row >= 0)
      .sort((a, b) => a - b)
    for (const row of rows) {
      const rowCells = cellData[String(row)]
      if (!rowCells) continue
      const columns = Object.keys(rowCells)
        .map(Number)
        .filter((column) => Number.isInteger(column) && column >= 0)
        .sort((a, b) => a - b)
      for (const column of columns) {
        const text = cellText(rowCells[String(column)])
        if (text === null) continue
        if (isWithheld?.(sheet.sheetId, row, column)) continue
        if (cells.length >= maxCells) {
          truncated = true
          return { cells, truncated }
        }
        cells.push({
          sheetId: sheet.sheetId,
          sheetName: sheet.sheetName,
          address: formatAddress(row, column),
          row,
          column,
          text,
        })
      }
    }
  }
  return { cells, truncated }
}

/**
 * The distinct words of a cell text, first-seen order: letters (plus the
 * apostrophe-dash tail every dictionary spelling carries), lower-cased so
 * a repeated word costs one engine call. Numbers, codes and URLs mostly
 * never make it in — a token must start with a letter.
 */
export function tokenizeCellText(text: string): string[] {
  const seen = new Set<string>()
  const words: string[] = []
  for (const match of text.matchAll(/\p{L}[\p{L}\p{M}'’-]*/gu)) {
    const raw = match[0]
    if (raw.length < 2 || raw.length > MAX_WORD_CHARS) continue
    const key = raw.toLowerCase()
    if (seen.has(key)) continue
    seen.add(key)
    words.push(key)
  }
  return words
}

/**
 * A word known-correct (dictionary) or dismissed for this session. Compared
 * lower-cased — the sets are fed from stored dictionary entries and session
 * dismissals, whose casing is not guaranteed.
 */
export function isWhiteListed(
  word: string,
  learned: ReadonlySet<string>,
  ignored: ReadonlySet<string>,
): boolean {
  const key = word.toLowerCase()
  const contains = (set: ReadonlySet<string>): boolean => {
    for (const entry of set) {
      if (entry.toLowerCase() === key) return true
    }
    return false
  }
  return contains(learned) || contains(ignored)
}

/** Applies one suggestion to a cell text: the first occurrence, case-preserved. */
export function replaceSuggestion(text: string, word: string, suggestion: string): string {
  const at = indexOfWordOf(text, word)
  if (at === -1) return text
  // match the flagged word's casing shape when it is obvious (All Caps / Capitalized)
  let replacement = suggestion
  if (/^[\p{Lu}][\p{Ll}]*$/u.test(word) && /^[\p{Ll}]/u.test(suggestion)) {
    replacement = suggestion[0]!.toUpperCase() + suggestion.slice(1)
  } else if (/^[\p{Lu}\p{Lt}]+$/u.test(word) && /^[\p{Ll}]/u.test(suggestion)) {
    replacement = suggestion.toUpperCase()
  }
  return text.slice(0, at) + replacement + text.slice(at + word.length)
}

/** The first standalone occurrence of `word`: a letter or digit must not touch it. */
export function indexOfWordOf(text: string, word: string): number {
  const isWordChar = (ch: string | undefined): boolean =>
    ch !== undefined && /[\p{L}\p{N}\p{M}'’_-]/u.test(ch)
  let at = text.indexOf(word)
  while (at !== -1) {
    if (!isWordChar(text[at - 1]) && !isWordChar(text[at + word.length])) return at
    at = text.indexOf(word, at + 1)
  }
  return -1
}

// ─── Rich text (Univer cell.p document) ────────────────────────────────────

/** Structural slice of Univer's rich-text document (IExtendedDocument). */
export interface RichDoc {
  readonly body?: {
    readonly dataStream?: string | undefined
    readonly textRuns?: readonly { st: number; ed: number; readonly ts?: unknown }[] | undefined
  } | null
}

/**
 * Rewrites the first standalone occurrence of `word` inside a rich-text
 * cell's document, keeping every run: the dataStream carries the text, the
 * runs carry the formatting as [st, ed) offsets into it. A run crossing the
 * replaced span is stretched over the suggestion (the shape Excel's own
 * corrector leaves behind); a run wholly inside it collapses away.
 */
export function replaceInRichDoc(doc: RichDoc, word: string, suggestion: string): RichDoc | null {
  const stream = doc.body?.dataStream
  if (stream === undefined) return null
  // the dataStream joins paragraphs with \r\n; a word never spans one, so
  // the plain-text search is valid
  const at = indexOfWordOf(stream, word)
  if (at === -1) return null
  const end = at + word.length
  const delta = suggestion.length - word.length
  const nextStream = stream.slice(0, at) + suggestion + stream.slice(end)
  const runs: { st: number; ed: number; ts?: unknown }[] = []
  for (const run of doc.body?.textRuns ?? []) {
    let { st, ed } = run
    if (st > at && st >= end) st += delta
    else if (st > at) st = at
    if (ed >= end) ed += delta
    else if (ed > at) ed = at
    if (ed <= st) continue
    runs.push({ st, ed, ...(run.ts !== undefined ? { ts: run.ts } : {}) })
  }
  return { body: { ...(doc.body ?? {}), dataStream: nextStream, textRuns: runs } as never }
}

// ─── The run ───────────────────────────────────────────────────────────────

/**
 * Checks every cell against the frame engine, synchronously — Hunspell over
 * a few thousand words costs milliseconds, so there is nothing to stream or
 * cancel. Duplicate words cost one engine call each; the map caches
 * `null` for a correct word and the suggestions for a misspelled one.
 */
export function runSheetSpellCheck(
  cells: readonly SpellCellRef[],
  engine: SpellEngine,
  learnedWords: ReadonlySet<string>,
  ignoredWords: ReadonlySet<string>,
): SpellFinding[] {
  const verdicts = new Map<string, string[] | null>()
  const findings: SpellFinding[] = []
  for (const cell of cells) {
    for (const word of tokenizeCellText(cell.text)) {
      if (isWhiteListed(word, learnedWords, ignoredWords)) continue
      let cached = verdicts.get(word)
      if (cached === undefined) {
        if (!engine.isWordMisspelled(word)) {
          verdicts.set(word, null)
          continue
        }
        const seen = new Set<string>([word])
        const suggestions: string[] = []
        for (const item of engine.getWordSuggestions(word)) {
          if (item.length === 0 || seen.has(item)) continue
          seen.add(item)
          suggestions.push(item)
          if (suggestions.length >= MAX_SUGGESTIONS) break
        }
        cached = suggestions
        verdicts.set(word, cached)
      }
      if (cached === null) continue
      // the finding carries the word as written, so the list reads like the sheet
      const asWritten = new RegExp(escapeRegExp(word), 'i').exec(cell.text)?.[0] ?? word
      findings.push({ cell, word: asWritten, suggestions: cached })
    }
  }
  return findings
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
