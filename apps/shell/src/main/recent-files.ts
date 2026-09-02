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

/** sidebar filter keys that stand for a family of extensions, not one exact ext */
const EXT_FAMILY: Record<string, readonly string[]> = { xlsx: ['xlsx', 'xlsm'] }

/** Page over the recents paths, preserving the source's newest-first order (unavailable paths stay, flagged missing). */
export function pageRecentPaths(
  paths: readonly string[],
  raw: unknown,
  starredPaths: ReadonlySet<string>,
): RecentPage {
  const { offset, limit, ext } = normalizeRecentQuery(raw)
  const all = statPathEntries(paths, starredPaths)
  const family = ext ? (EXT_FAMILY[ext] ?? [ext]) : undefined
  const filtered = family ? all.filter((entry) => family.includes(entry.ext)) : all
  return {
    entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
    total: filtered.length,
    totalAll: all.length,
  }
}
