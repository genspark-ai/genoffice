/**
 * Per-file reading position persistence (WPS-style "return to where you left
 * off"). All entries live under one localStorage key as a path-keyed map with
 * LRU pruning, so the store cannot grow without bound. Pure module — storage
 * is injectable for tests.
 */

export type PdfFitMode = 'width' | 'page' | null

export type PdfViewState = {
  /** 1-based page number of the row at the viewport top */
  page: number
  /** Viewport-top offset within that row, as a fraction of the row height.
      Fraction (not pixels) so the position survives a different window size
      or fit-computed scale on the next open. */
  frac: number
  scale: number
  fitMode: PdfFitMode
}

type StoredEntry = PdfViewState & { at: number }

export const VIEW_STATE_KEY = 'genoffice-pdf-view-state'
export const MAX_VIEW_ENTRIES = 100

type StorageLike = Pick<Storage, 'getItem' | 'setItem'>

function readAll(storage: StorageLike): Record<string, unknown> {
  try {
    const parsed: unknown = JSON.parse(storage.getItem(VIEW_STATE_KEY) ?? '{}')
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed)
      ? (parsed as Record<string, unknown>)
      : {}
  } catch {
    return {}
  }
}

/** Entries come from disk and may be from an older build or corrupted — never trust them. */
function sanitize(entry: unknown): StoredEntry | null {
  if (!entry || typeof entry !== 'object') return null
  const e = entry as Record<string, unknown>
  if (typeof e.page !== 'number' || !Number.isFinite(e.page) || e.page < 1) return null
  if (typeof e.scale !== 'number' || !Number.isFinite(e.scale) || e.scale <= 0) return null
  if (e.fitMode !== 'width' && e.fitMode !== 'page' && e.fitMode !== null) return null
  return {
    page: Math.floor(e.page),
    frac:
      typeof e.frac === 'number' && Number.isFinite(e.frac) ? Math.min(Math.max(e.frac, 0), 1) : 0,
    scale: e.scale,
    fitMode: e.fitMode,
    at: typeof e.at === 'number' && Number.isFinite(e.at) ? e.at : 0,
  }
}

export function loadViewState(
  path: string,
  storage: StorageLike = localStorage,
): PdfViewState | null {
  if (!path) return null
  const entry = sanitize(readAll(storage)[path])
  if (!entry) return null
  const { at: _at, ...state } = entry
  return state
}

export function saveViewState(
  path: string,
  state: PdfViewState,
  storage: StorageLike = localStorage,
  now: number = Date.now(),
): void {
  if (!path) return
  const all = readAll(storage)
  all[path] = { ...state, at: now } satisfies StoredEntry
  const paths = Object.keys(all)
  if (paths.length > MAX_VIEW_ENTRIES) {
    const age = (p: string) => sanitize(all[p])?.at ?? -1
    paths.sort((a, b) => age(b) - age(a))
    for (const stale of paths.slice(MAX_VIEW_ENTRIES)) delete all[stale]
  }
  try {
    storage.setItem(VIEW_STATE_KEY, JSON.stringify(all))
  } catch {
    // Quota exceeded / storage unavailable: restoring is best-effort, never break viewing
  }
}

/** Derive the persistable view state from the current scroll geometry.
 *  Mirrors the row layout in App.tsx: rows stack vertically with a leading
 *  and inter-row gap; `rowHeights` are the on-screen row heights in px. */
export function captureViewState(args: {
  scrollTop: number
  rowHeights: number[]
  /** 1-based page number of each row's first page */
  rowPages: number[]
  gap: number
  scale: number
  fitMode: PdfFitMode
}): PdfViewState {
  const { scrollTop, rowHeights, rowPages, gap, scale, fitMode } = args
  let top = gap
  let rowIdx = 0
  let rowTop = top
  for (let i = 0; i < rowHeights.length; i++) {
    if (top <= scrollTop) {
      rowIdx = i
      rowTop = top
    } else break
    top += rowHeights[i]! + gap
  }
  const height = rowHeights[rowIdx] ?? 0
  return {
    page: rowPages[rowIdx] ?? 1,
    frac: height > 0 ? Math.min(Math.max((scrollTop - rowTop) / height, 0), 1) : 0,
    scale,
    fitMode,
  }
}
