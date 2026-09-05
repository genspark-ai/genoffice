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

export function normalizeRecentQuery(
  raw: unknown,
): Required<Omit<RecentQuery, 'ext'>> & { ext?: string } {
  const query = (raw ?? {}) as RecentQuery
  const offset = Number.isFinite(query.offset) ? Math.max(0, Math.floor(query.offset!)) : 0
  const limit = Number.isFinite(query.limit)
    ? Math.min(RECENT_PAGE_MAX, Math.max(0, Math.floor(query.limit!)))
    : RECENT_PAGE_DEFAULT
  const ext = typeof query.ext === 'string' && query.ext ? query.ext.toLowerCase() : undefined
  return { offset, limit, ext }
}

/** sidebar filter keys that stand for a family of extensions, not one exact ext.
    Families mirror the extensions the shell router actually opens (XLSX_RE /
    MD_RE in index.ts) so filtered-out-but-openable files cannot hide. */
export const EXT_FAMILY: Record<string, readonly string[]> = {
  // routeDocumentPath opens xlsx/xlsm/xls/csv alike
  xlsx: ['xlsx', 'xlsm', 'xls', 'csv'],
  // legacy .doc opens via the shell's extract-to-text converter, so it belongs
  // under the Word filter with .docx
  docx: ['docx', 'doc'],
  // routeDocumentPath opens .md and .markdown alike
  md: ['md', 'markdown'],
}

/** True when a recents/starred entry extension passes the sidebar filter key. */
export function matchesExtFamily(filterExt: string | undefined, entryExt: string): boolean {
  if (!filterExt) return true
  return (EXT_FAMILY[filterExt] ?? [filterExt]).includes(entryExt)
}

/** Page over the recents paths, preserving the source's newest-first order (unavailable paths stay, flagged missing). */
export function pageRecentPaths(
  paths: readonly string[],
  raw: unknown,
  starredPaths: ReadonlySet<string>,
): RecentPage {
  const { offset, limit, ext } = normalizeRecentQuery(raw)
  const all = statPathEntries(paths, starredPaths)
  const filtered = all.filter((entry) => matchesExtFamily(ext, entry.ext))
  return {
    entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
    total: filtered.length,
    totalAll: all.length,
  }
}
