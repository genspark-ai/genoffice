import { extname } from 'node:path'
import type { ScriptFormat } from '@genoffice/agent-core'
import type { DeckSummary } from '@genoffice/cli/headless-formats'

/**
 * What the script kind's format modules share: the snapshot union the sandbox is
 * built from, the caps that keep one format's snapshot bounded, and the errors
 * they throw. The sandbox itself lives in script-host.ts; these types describe
 * what it is handed and what it hands back.
 */

/** One block of the read-only snapshot a docs script sees. */
export interface ScriptBlock {
  index: number
  type: string
  level?: number
  kind?: string
  text: string
}

export interface DocsSnapshot {
  format?: 'docs'
  blocks: ScriptBlock[]
  text: string
}

/** The per-cell format echo a sheets snapshot holds (plain data for the sandbox). */
export interface ScriptCellFormat {
  numberFormat?: string
  bold?: boolean
  italic?: boolean
  fillColor?: string
  fontColor?: string
}

/** One sheet of the read-only snapshot a sheets script sees. */
export interface SheetSnapshotSheet {
  name: string
  /** A1 of the first and last non-empty cell */
  first: string
  last: string
  rows: number
  cols: number
  /** the workbook's values (the raw model value where the file stores one), by address */
  cells: Record<string, string | number | boolean | null>
  formulas: Record<string, string>
  /** number format, bold, italic, fill of the styled cells, by address */
  formats: Record<string, ScriptCellFormat>
}

export interface SheetsSnapshot {
  format: 'sheets'
  active: string | null
  sheets: SheetSnapshotSheet[]
}

export interface TextSnapshot {
  format: 'text'
  text: string
}

export interface DeckSnapshot {
  format: 'deck'
  /** the deck as describeDeck reports it: slides, element boxes (EMU), text previews */
  deck: DeckSummary
}

export interface PdfSnapshot {
  format: 'pdf'
  pdf: {
    pages: number
    encrypted: boolean
    /** 1-based pages the snapshot actually holds text for */
    pagesRead: number
    /** the page count the host reads into a snapshot; $.pdf.pages refuses beyond it */
    pageCap: number
    title?: string
    author?: string
    subject?: string
    creator?: string
    producer?: string
    pageTexts: {
      page: number
      width_pt: number
      height_pt: number
      text: string
      truncated: boolean
    }[]
  }
}

export type ScriptSnapshot =
  DocsSnapshot | SheetsSnapshot | TextSnapshot | DeckSnapshot | PdfSnapshot

/** What the sandbox produced, before the format's engine has judged it. */
export interface SandboxOutcome {
  /** the operations the script proposed, in order; the caller validates them */
  ops: Record<string, unknown>[]
  log: string[]
  /** reasons a hook gave for stopping the event; empty for a script */
  vetoes: string[]
  ms: number
  /** the full replacement text a text script returned, when it returned one */
  text?: string
  /** the report a read-only script returned, JSON-encoded by the sandbox */
  result?: string
}

/** A failed run, with the one-line next step the tool should show. */
export class ScriptRunError extends Error {
  readonly suggestion?: string | undefined
  constructor(message: string, suggestion?: string) {
    super(message)
    this.name = 'ScriptRunError'
    this.suggestion = suggestion
  }
}

/** The extensions each script format serves; a file outside these cannot be scripted. */
export const SCRIPT_FORMAT_EXTS: Record<ScriptFormat, readonly string[]> = {
  docs: ['docx', 'docm', 'dotx', 'dotm'],
  sheets: ['xlsx', 'xlsm', 'xltx', 'xltm'],
  text: ['md', 'markdown', 'html', 'htm'],
  deck: ['pptx', 'pptm', 'potx', 'potm'],
  pdf: ['pdf'],
}

/** The script format a document path runs as; undefined when no format serves it. */
export function scriptFormatFor(path: string): ScriptFormat | undefined {
  const ext = extname(path).slice(1).toLowerCase()
  for (const format of Object.keys(SCRIPT_FORMAT_EXTS) as ScriptFormat[]) {
    if (SCRIPT_FORMAT_EXTS[format].includes(ext)) return format
  }
  return undefined
}

// ── caps: every read a format hands a script is bounded, refusing over the cap
// with a reason (docs/mods-implementation-report.md §8.3.2) ─────────────────────

/** non-empty cells a sheets snapshot may hold before the run is refused */
export const SHEETS_SNAPSHOT_CELL_CAP = 50_000
/** cells one $.sheets.range / $.sheets.formats call may cover */
export const SHEETS_READ_CELL_CAP = 20_000
/** cells $.sheets.formats may echo back */
export const SHEETS_FORMAT_CAP = 20_000
/** matches one $.sheets.find call returns before it reports the rest as truncated */
export const SHEETS_FIND_CAP = 200
/** bytes of source text a text snapshot may hold before the run is refused */
export const TEXT_SNAPSHOT_BYTES = 4 * 1024 * 1024
/** pages a pdf snapshot reads; $.pdf.pages refuses beyond it with the page count */
export const PDF_SNAPSHOT_PAGE_CAP = 200
