import { statSync } from 'node:fs'
import { basename, extname } from 'node:path'
import type { RecentEntry, RecentPage, RecentQuery } from '../shared/home-api'

const RECENT_PAGE_DEFAULT = 50
const RECENT_PAGE_MAX = 200

function toRecentEntry(path: string, starredPaths: ReadonlySet<string>): RecentEntry {
  try {
    const stat = statSync(path)
    return {
      path,
      name: basename(path),
      ext: extname(path).slice(1).toLowerCase(),
      mtimeMs: stat.mtimeMs,
      sizeBytes: stat.size,
      starred: starredPaths.has(path),
    }
  } catch {
    // A failed stat is often transient (disconnected drive, pending mount,
    // cloud placeholder) — dropping the entry made the recents list silently
    // lose files until a later reload (r158). Word keeps unavailable recents
    // listed; the row is flagged so the UI can dim it and offer removal.
    return {
      path,
      name: basename(path),
      ext: extname(path).slice(1).toLowerCase(),
      mtimeMs: 0,
      sizeBytes: 0,
      starred: starredPaths.has(path),
      missing: true,
    }
  }
}

export function statPathEntries(
  paths: readonly string[],
  starredPaths: ReadonlySet<string>,
): RecentEntry[] {
  return paths.map((path) => toRecentEntry(path, starredPaths))
}

/**
 * How many paths one statPaths call may stat. stat is synchronous, so the
 * renderer-supplied list has to be bounded like every other recents list: the
 * page bound, which is already the largest one the UI asks for. statting an
 * unbounded list blocked the main process once it grew past a few hundred
 * entries (see pageRecentPaths).
 */
export const STAT_PATHS_MAX = RECENT_PAGE_MAX

/** statPaths crosses the IPC boundary, so the caller's list is capped before any stat. */
export function capStatPaths(paths: readonly string[]): string[] {
  return paths.slice(0, STAT_PATHS_MAX)
}

export function normalizeRecentQuery(
  raw: unknown,
): Required<Omit<RecentQuery, 'ext'>> & { ext?: string } {
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
  return { offset, limit, ext }
}

/** sidebar filter keys that stand for a family of extensions, not one exact ext */
export const EXT_FAMILY: Record<string, readonly string[]> = {
  // mirrors Home's FILTER_FAMILY and the search-side SEARCH_EXT_FAMILY
  docx: ['docx', 'doc'],
  // delimited text belongs to the sheets family: Home's own FILTER_FAMILY and
  // the shell's open routing both treat .csv/.tsv as spreadsheets, so a
  // sidebar filtered on "xlsx" must page them in too (csv was missing here).
  xlsx: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv'],
  pptx: ['pptx', 'ppt'],
  md: ['md', 'markdown'],
  html: ['html', 'htm'],
}

/** Family-aware extension match for sidebar filters (recents and starred share it). */
export function matchesExtFamily(entryExt: string, filterExt: string): boolean {
  const family = EXT_FAMILY[filterExt]
  return family ? family.includes(entryExt) : entryExt === filterExt
}

/** Page over the recents paths, preserving the source's newest-first order (unavailable paths stay, flagged missing). */
export function pageRecentPaths(
  paths: readonly string[],
  raw: unknown,
  starredPaths: ReadonlySet<string>,
): RecentPage {
  const { offset, limit, ext } = normalizeRecentQuery(raw)
  // The extension filter is pure string work (extname needs no stat), so filter
  // and count first and stat only the page being returned — statting every
  // path of a long recents list on every page turn blocked the main process
  // once the list grew past a few hundred entries.
  const filtered = ext
    ? paths.filter((p) => matchesExtFamily(extname(p).slice(1).toLowerCase(), ext))
    : paths
  const page = limit === 0 ? [] : filtered.slice(offset, offset + limit)
  return {
    entries: statPathEntries(page, starredPaths),
    total: filtered.length,
    totalAll: paths.length,
  }
}
