/**
 * Find over a streamed workbook's file cells, executed by the sidecar's
 * `find_cells` over its chunk index instead of paging cell JSON into the
 * renderer. The sidecar sees the file only: journal edits are overlaid by
 * the callers (session-edited coordinates are shadowed and searched from the
 * journal), and structural ops are mapped here from file to screen space.
 */
import type { WorkbookFindCellsMatch, WorkbookFindCellsRequest } from '../shared/desktop-api'
import { MAX_FIND_CELLS_PAGE } from '../shared/desktop-api'
import { lazyFileSheetId, lazySheetMeta, type LazyWorkbookState } from './univer-state'
import { fileToScreen } from './view-transform'

/** Matches one Find session keeps before telling the user to narrow it. */
export const MAX_FIND_MATCHES = 100_000

export type SidecarFindQuery = Omit<
  WorkbookFindCellsRequest,
  'sessionId' | 'sheetId' | 'resumeAt' | 'limit'
>

/** A file hit already translated to screen coordinates. */
export interface SidecarFindHit {
  readonly row: number
  readonly column: number
  readonly value: WorkbookFindCellsMatch['value']
  readonly valueText: string | null
  readonly formula: string | undefined
}

export interface SidecarFindSheetOptions {
  /** Abort between pages (dialog closed, workbook switched). */
  readonly alive?: () => boolean
  /** Stop after this many kept hits (as reported by onPage). */
  readonly maxMatches?: number
  /** Pages still waiting on worksheet indexing give up after this long. */
  readonly indexingWaitLimitMs?: number
  /** Receives a page; returns how many of its hits the caller kept. */
  readonly onPage: (hits: SidecarFindHit[]) => number
}

export interface SidecarFindSheetOutcome {
  /** The whole sheet was scanned. */
  readonly complete: boolean
  /** Stopped at maxMatches. */
  readonly capped: boolean
  readonly matches: number
}

export function sidecarFindAvailable(): boolean {
  if (typeof window === 'undefined') return false
  const api = (window as { desktopApi?: Partial<Window['desktopApi']> }).desktopApi
  return typeof api?.findWorkbookCells === 'function'
}

const INDEXING_RETRY_MS = 50

/**
 * Pages the sidecar over one (screen) sheet's file cells. Resolves when the
 * sheet is exhausted, the caller stopped it, or the cap was hit; rejects only
 * when the very first page fails, so callers can fall back to another scan
 * before anything was emitted.
 */
export async function findSheetCellsInFile(
  state: LazyWorkbookState,
  sheetId: string,
  query: SidecarFindQuery,
  options: SidecarFindSheetOptions,
): Promise<SidecarFindSheetOutcome> {
  const alive = options.alive ?? (() => true)
  const maxMatches = options.maxMatches ?? MAX_FIND_MATCHES
  if (maxMatches <= 0) return { complete: false, capped: true, matches: 0 }
  const meta = lazySheetMeta(state, sheetId)
  if (!meta || meta.rowCount <= 0 || meta.columnCount <= 0) {
    return { complete: true, capped: false, matches: 0 }
  }
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const fileSheetId = lazyFileSheetId(state, sheetId)
  let resumeAt: WorkbookFindCellsRequest['resumeAt']
  let matches = 0
  let pages = 0
  const started = Date.now()
  for (;;) {
    if (!alive()) return { complete: false, capped: false, matches }
    let page
    try {
      page = await window.desktopApi.findWorkbookCells({
        sessionId: state.file.sessionId,
        sheetId: fileSheetId,
        ...query,
        ...(resumeAt === undefined ? {} : { resumeAt }),
        limit: Math.max(1, Math.min(MAX_FIND_CELLS_PAGE, maxMatches - matches)),
      })
    } catch (error) {
      if (pages === 0) throw error
      return { complete: false, capped: false, matches }
    }
    pages += 1
    if (!alive()) return { complete: false, capped: false, matches }
    const hits: SidecarFindHit[] = []
    for (const match of page.matches) {
      const row = ops.length === 0 ? match.row : fileToScreen(ops, 'row', match.row)
      const column = ops.length === 0 ? match.column : fileToScreen(ops, 'column', match.column)
      if (row === null || column === null) continue
      hits.push({
        row,
        column,
        value: match.value,
        valueText: match.valueText,
        formula: match.formulaText,
      })
    }
    if (hits.length > 0) matches += options.onPage(hits)
    if (page.complete) return { complete: true, capped: false, matches }
    if (matches >= maxMatches) return { complete: false, capped: true, matches }
    if (!page.indexingComplete) {
      if (Date.now() - started > (options.indexingWaitLimitMs ?? Number.POSITIVE_INFINITY)) {
        return { complete: false, capped: false, matches }
      }
      await new Promise((resolve) => setTimeout(resolve, INDEXING_RETRY_MS))
    }
    resumeAt = page.nextCursor
    if (resumeAt === undefined) return { complete: false, capped: false, matches }
  }
}
