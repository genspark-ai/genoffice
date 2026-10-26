/**
 * The engine behind F7 / Review ▸ Spelling (see spellcheck-pref.ts for why it
 * exists and why it is not the Chromium native checker).
 *
 * Detection runs cell-text batches through the configured AI provider — the
 * same backend docs' F7 proofread uses — while everything around it (cell
 * traversal, white-lists, suggestions, replacements) stays local and
 * deterministic. The pure halves (collection, batching, response parsing,
 * replacement) are unit-tested without Univer or the network.
 */
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

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
export interface SpellSnapshotSheet {
  readonly name?: string | undefined
  readonly cellData?:
    | Readonly<
        Record<
          string,
          Readonly<Record<string, { v?: unknown; f?: unknown } | null | undefined>> | undefined
        >
      >
    | undefined
}

export interface SpellSnapshot {
  readonly sheets: Readonly<Record<string, SpellSnapshotSheet | undefined>>
}

/** A text must have a letter for spelling to be a question at all. */
const HAS_LETTER = /\p{L}/u
/** Cells longer than this are checked on their prefix only (prompt budget). */
export const MAX_TEXT_CHARS = 1000
/** Texts per AI request: large enough to amortize the round trip, small */
/** enough that one response stays small and the progress bar keeps moving. */
export const SPELL_BATCH_TEXTS = 30
/** Hard cap on texts per run; beyond it the dialog reports a truncation note. */
export const MAX_TEXTS_PER_RUN = 2000

function cellText(cell: { v?: unknown; f?: unknown } | null | undefined): string | null {
  // formula cells are skipped: the cached result is not text the user retypes
  if (!cell || (cell.f !== undefined && cell.f !== null && cell.f !== '')) return null
  const value = cell.v
  if (typeof value !== 'string') return null
  if (value.trim().length === 0 || !HAS_LETTER.test(value)) return null
  // the raw (untrimmed) string: a Change must preserve the cell's own spacing
  return value
}

/**
 * Every checkable string cell of one sheet in row-major order: plain string
 * values only (a formula's cached result is not text the user retypes), no
 * withheld cells, empty/number-only strings dropped before they can cost a
 * prompt token. Formula checks and grammar stay out of scope, like the rest
 * of the feature.
 */
export function collectSheetTextCells(
  snapshot: SpellSnapshot,
  sheetId: string,
  sheetName: string,
  isWithheld?: (sheetId: string, row: number, column: number) => boolean,
): SpellCellRef[] {
  const cellData = snapshot.sheets[sheetId]?.cellData
  if (!cellData) return []
  const rows = Object.keys(cellData)
    .map(Number)
    .filter((row) => Number.isInteger(row) && row >= 0)
    .sort((a, b) => a - b)
  const cells: SpellCellRef[] = []
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
      if (isWithheld?.(sheetId, row, column)) continue
      cells.push({ sheetId, sheetName, address: formatAddress(row, column), row, column, text })
    }
  }
  return cells
}

/** Dedupes cell texts into the batches the AI sees; order preserved. */
export function buildTextBatches(
  cells: readonly SpellCellRef[],
  batchSize = SPELL_BATCH_TEXTS,
  maxTexts = MAX_TEXTS_PER_RUN,
): { batches: string[][]; cellsByText: Map<string, SpellCellRef[]>; truncated: boolean } {
  const cellsByText = new Map<string, SpellCellRef[]>()
  const texts: string[] = []
  let truncated = false
  for (const cell of cells) {
    const key = cell.text.length > MAX_TEXT_CHARS ? cell.text.slice(0, MAX_TEXT_CHARS) : cell.text
    const existing = cellsByText.get(key)
    if (existing) {
      existing.push(cell)
      continue
    }
    if (texts.length >= maxTexts) {
      truncated = true
      // the cell still maps under its (clipped) key so later findings resolve
      cellsByText.set(key, [cell])
      continue
    }
    texts.push(key)
    cellsByText.set(key, [cell])
  }
  const batches: string[][] = []
  for (let index = 0; index < texts.length; index += batchSize) {
    batches.push(texts.slice(index, index + batchSize))
  }
  return { batches, cellsByText, truncated }
}

/** One model reply finding, still in text-index space. */
interface RawFinding {
  readonly index: number
  readonly word: string
  readonly suggestions: readonly string[]
}

/**
 * Parses one checker reply. Only minified JSON is expected; fences and stray
 * prose around the object are tolerated (stripped) because small models do
 * add them. A finding survives only when its word really occurs in its text —
 * a hallucinated index or a "corrected" word that is not in the cell would
 * otherwise surface a change the user cannot map onto their data.
 */
export function parseSpellReply(
  content: string,
  texts: readonly string[],
): { findings: RawFinding[] } | null {
  const start = content.indexOf('{')
  const end = content.lastIndexOf('}')
  if (start === -1 || end <= start) return null
  let parsed: unknown
  try {
    parsed = JSON.parse(content.slice(start, end + 1))
  } catch {
    return null
  }
  const rawFindings = (parsed as { findings?: unknown }).findings
  if (!Array.isArray(rawFindings)) return null
  const findings: RawFinding[] = []
  for (const entry of rawFindings) {
    if (!entry || typeof entry !== 'object') continue
    const record = entry as { i?: unknown; w?: unknown; s?: unknown; suggestions?: unknown }
    const index = typeof record.i === 'number' ? Math.trunc(record.i) : -1
    const word = typeof record.w === 'string' ? record.w.trim() : ''
    if (index < 1 || index > texts.length || word.length === 0) continue
    const text = texts[index - 1] ?? ''
    if (!text.includes(word)) continue
    const rawSuggestions = Array.isArray(record.s)
      ? record.s
      : Array.isArray(record.suggestions)
        ? record.suggestions
        : []
    const seenSuggestions = new Set<string>()
    const suggestions: string[] = []
    for (const item of rawSuggestions) {
      if (typeof item !== 'string') continue
      const suggestion = item.trim()
      if (suggestion.length === 0 || suggestion === word || seenSuggestions.has(suggestion))
        continue
      seenSuggestions.add(suggestion)
      suggestions.push(suggestion)
      if (suggestions.length >= 5) break
    }
    findings.push({ index, word, suggestions })
  }
  return { findings }
}

/**
 * A word known-correct (dictionary) or dismissed for this session. Both sides
 * are compared lower-cased — the sets are fed from stored dictionary entries
 * and session dismissals, whose casing is not guaranteed.
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

/** Applies one suggestion to a cell text: the flagged occurrence, case-preserved. */
export function replaceSuggestion(text: string, word: string, suggestion: string): string {
  const at = text.indexOf(word)
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

export interface ChatLike {
  (
    system: string,
    user: string,
  ): Promise<{ ok: boolean; content?: string | undefined; error?: string | undefined }>
}

export interface SpellCheckRunOptions {
  readonly cells: readonly SpellCellRef[]
  readonly learnedWords: ReadonlySet<string>
  readonly ignoredWords: ReadonlySet<string>
  readonly requestChat: ChatLike
  readonly onProgress?: (checkedTexts: number, totalTexts: number) => void
  /** streamed per batch so the list fills while the run continues */
  readonly onFindings?: (findings: SpellFinding[]) => void
  /** polled between batches; false stops the run as "canceled" */
  readonly shouldContinue?: () => boolean
  readonly batchSize?: number
  readonly maxTexts?: number
}

export type SpellCheckRunResult =
  | { readonly status: 'done'; readonly truncated: boolean }
  | { readonly status: 'canceled' }
  | { readonly status: 'error'; readonly error: string }

const SYSTEM_PROMPT = [
  'You are a strict spelling checker for spreadsheet cell texts.',
  'Flag misspelled words only: no grammar, no style, no capitalization complaints.',
  'Never flag proper nouns, product/company/person/place names you cannot prove wrong,',
  'numbers, dates, times, currency amounts, codes, identifiers, emails, URLs or file paths.',
  'Reply with ONLY minified JSON of the shape {"findings":[{"i":<line number>,"w":"<misspelled word exactly as written>","s":["<correction>", ...]}]}',
  'where "i" is the 1-based index into the given array and "s" holds up to 3 corrections.',
  'A text without misspellings produces no entry. No findings at all: {"findings":[]}.',
].join(' ')

/** Runs the whole sheet check: dedupe → batches → model → white-list filter. */
export async function runSheetSpellCheck(
  options: SpellCheckRunOptions,
): Promise<SpellCheckRunResult> {
  const { batches, cellsByText, truncated } = buildTextBatches(
    options.cells,
    options.batchSize,
    options.maxTexts,
  )
  let checked = 0
  const total = batches.reduce((count, batch) => count + batch.length, 0)
  options.onProgress?.(checked, total)
  for (const batch of batches) {
    if (options.shouldContinue && !options.shouldContinue()) {
      return { status: 'canceled' }
    }
    const known = [...options.learnedWords, ...options.ignoredWords]
    const user = JSON.stringify({
      knownCorrectWords: known.length > 0 ? known.slice(0, 400) : undefined,
      texts: batch.map((text) =>
        text.length > MAX_TEXT_CHARS ? text.slice(0, MAX_TEXT_CHARS) : text,
      ),
    })
    let reply: Awaited<ReturnType<ChatLike>>
    try {
      reply = await options.requestChat(SYSTEM_PROMPT, user)
    } catch (error) {
      return { status: 'error', error: error instanceof Error ? error.message : String(error) }
    }
    if (!reply.ok) {
      return { status: 'error', error: reply.error ?? 'The spelling check request failed.' }
    }
    const parsed = reply.content ? parseSpellReply(reply.content, batch) : null
    if (parsed) {
      const findings: SpellFinding[] = []
      for (const raw of parsed.findings) {
        if (isWhiteListed(raw.word, options.learnedWords, options.ignoredWords)) continue
        for (const cell of cellsByText.get(batch[raw.index - 1] ?? '') ?? []) {
          findings.push({ cell, word: raw.word, suggestions: raw.suggestions })
        }
      }
      if (findings.length > 0) options.onFindings?.(findings)
    }
    checked += batch.length
    options.onProgress?.(checked, total)
  }
  return { status: 'done', truncated }
}
