import { statSync } from 'node:fs'
import { stat } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import type { RecentEntry, RecentPage, RecentQuery } from '../shared/home-api'

const RECENT_PAGE_DEFAULT = 50
const RECENT_PAGE_MAX = 200

function toRecentEntry(
  path: string,
  starredPaths: ReadonlySet<string>,
  info?: { mtimeMs: number; size: number },
): RecentEntry {
  return {
    path,
    name: basename(path),
    ext: extname(path).slice(1).toLowerCase(),
    mtimeMs: info?.mtimeMs ?? 0,
    sizeBytes: info?.size ?? 0,
    starred: starredPaths.has(path),
    ...(info ? {} : { missing: true }),
  }
}

/** Explicit folder browsing retains its existing synchronous metadata contract. */
export function statPathEntries(
  paths: readonly string[],
  starredPaths: ReadonlySet<string>,
): RecentEntry[] {
  return paths.map((path) => {
    try {
      return toRecentEntry(path, starredPaths, statSync(path))
    } catch {
      return toRecentEntry(path, starredPaths)
    }
  })
}

export function statPathEntriesAsync(
  paths: readonly string[],
  starredPaths: ReadonlySet<string>,
): Promise<RecentEntry[]> {
  return Promise.all(
    paths.map(async (path) => {
      try {
        return toRecentEntry(path, starredPaths, await stat(path))
      } catch {
        // Keep unavailable paths listed, just as for a disconnected local drive.
        return toRecentEntry(path, starredPaths)
      }
    }),
  )
}

export const STAT_PATHS_MAX = RECENT_PAGE_MAX

/** A renderer-supplied list is bounded like every recents page. */
export function capStatPaths(paths: readonly string[]): string[] {
  return paths.slice(0, STAT_PATHS_MAX)
}

export function normalizeRecentQuery(
  raw: unknown,
): Required<Omit<RecentQuery, 'ext' | 'group' | 'query'>> & { ext?: string; query?: string } {
  const query = (raw ?? {}) as RecentQuery
  // offset/limit cross the preload boundary, so an IPC caller may send "10" rather
  // than 10; without coercion every page silently restarted at the first page.
  const rawOffset = Number(query.offset)
  const rawLimit = Number(query.limit)
  const offset = Number.isFinite(rawOffset) ? Math.max(0, Math.floor(rawOffset)) : 0
  const limit = Number.isFinite(rawLimit)
    ? Math.min(RECENT_PAGE_MAX, Math.max(0, Math.floor(rawLimit)))
    : RECENT_PAGE_DEFAULT
  // Sidebar keys are bare extensions ("xlsx"), but IPC callers may send
  // ".xlsx", " XLSX ", or "..." — normalize so openable files cannot hide
  // behind a filter that only differs in dots/case/whitespace.
  const rawExt =
    typeof query.ext === 'string' ? query.ext.trim().toLowerCase().replace(/^\.+/, '') : ''
  const ext = rawExt ? rawExt : undefined
  const text = typeof query.query === 'string' ? query.query.trim().toLowerCase() : ''
  return { offset, limit, ext, query: text || undefined }
}

/** sidebar filter keys that stand for a family of extensions, not one exact ext */
export const EXT_FAMILY: Record<string, readonly string[]> = {
  // mirrors Home's FILTER_FAMILY and SEARCH_EXT_FAMILY
  docx: ['docx', 'doc', 'dot', 'odt', 'ott', 'hwpx', 'hwp'],
  // delimited text belongs to the sheets family: Home's own FILTER_FAMILY and
  // the shell's open routing both treat .csv/.tsv as spreadsheets, so a
  // sidebar filtered on "xlsx" must page them in too (csv was missing here).
  xlsx: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'ods', 'ots'],
  pptx: ['pptx', 'ppt', 'odp', 'otp'],
  // the text app opens txt/json as source too, so the sidebar "md" filter has
  // to page them in the same way Home's FILTER_FAMILY already does — otherwise
  // the two views disagree about which files the filter means.
  md: ['md', 'markdown', 'txt', 'json'],
  html: ['html', 'htm'],
}

/** Family-aware extension match for sidebar filters (recents and starred share it). */
export function matchesExtFamily(entryExt: string, filterExt: string): boolean {
  const family = EXT_FAMILY[filterExt]
  return family ? family.includes(entryExt) : entryExt === filterExt
}

/** Page over the recents paths, preserving the source's newest-first order (unavailable paths stay, flagged missing). */
export async function pageRecentPaths(
  paths: readonly string[],
  raw: unknown,
  starredPaths: ReadonlySet<string>,
): Promise<RecentPage> {
  const { offset, limit, ext, query } = normalizeRecentQuery(raw)
  // Probe only the returned page, without blocking main on unavailable mounts.
  let filtered = ext
    ? paths.filter((p) => matchesExtFamily(extname(p).slice(1).toLowerCase(), ext))
    : paths
  if (query) {
    filtered = filtered.filter((p) => p.toLowerCase().includes(query))
  }
  const page = limit === 0 ? [] : filtered.slice(offset, offset + limit)
  return {
    entries: await statPathEntriesAsync(page, starredPaths),
    total: filtered.length,
    totalAll: paths.length,
  }
}
