import { existsSync, readFileSync } from 'node:fs'
import { writeJsonAtomic } from '@genoffice/electron-utils/atomic-write'

/**
 * The starred.json store behind the home screen's favorites. The path is the
 * key of every entry; a path appears at most once. Shell's home IPC and the
 * docs menu both read through docs-main's wrappers around these helpers.
 *
 * On-disk shape (since starred groups): { version: 1, items: StarredItem[] }.
 * The legacy shape was a flat string[]; it is migrated on read — every entry
 * keeps its star but lands in no group — and the next write persists the
 * versioned shape. There is no separate groups list: a group exists while at
 * least one entry carries its name and disappears with its last member.
 */

/** one starred file; `group` is the user-named group it belongs to (absent = ungrouped) */
export interface StarredItem {
  path: string
  group?: string
}

/** starred.json shape since groups landed */
export interface StarredStore {
  version: 1
  items: StarredItem[]
}

function readJson(path: string): unknown {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8'))
  } catch {
    /* corrupted state file: fall back to an empty store */
  }
  return null
}

/** one entry of either on-disk shape: a bare legacy path or an item object */
function parseStarredItem(entry: unknown): StarredItem | null {
  if (typeof entry === 'string') return { path: entry }
  if (!entry || typeof entry !== 'object' || typeof (entry as StarredItem).path !== 'string') {
    return null
  }
  const { path, group } = entry as StarredItem
  return typeof group === 'string' && group.trim() ? { path, group: group.trim() } : { path }
}

/** Parse either on-disk shape into entries. Unknown shapes (corrupt file) and
 *  entries degrade to dropped/ungrouped rather than throwing: a broken
 *  starred.json must not take the whole home screen down. */
export function parseStarredItems(raw: unknown): StarredItem[] {
  let entries: unknown[] | null = null
  if (Array.isArray(raw)) {
    // legacy flat list: bare paths, and object entries degrade to ungrouped
    entries = raw
  } else if (raw && typeof raw === 'object' && Array.isArray((raw as { items?: unknown }).items)) {
    entries = (raw as { items: unknown[] }).items
  }
  if (!entries) return []
  return entries.flatMap((entry) => {
    const item = parseStarredItem(entry)
    return item ? [item] : []
  })
}

export function readStarredItems(filePath: string): StarredItem[] {
  return parseStarredItems(readJson(filePath))
}

export function writeStarredItems(filePath: string, items: StarredItem[]): void {
  const store: StarredStore = { version: 1, items }
  writeJsonAtomic(filePath, store)
}

/** star / unstar one path (unstar drops its group with it; re-starring starts ungrouped) */
export function toggleStarredItem(items: StarredItem[], filePath: string): StarredItem[] {
  return items.some((item) => item.path === filePath)
    ? items.filter((item) => item.path !== filePath)
    : [...items, { path: filePath }]
}

/** remove paths; null when nothing matched so callers skip the write */
export function dropStarredItems(items: StarredItem[], filePaths: string[]): StarredItem[] | null {
  const drop = new Set(filePaths)
  if (drop.size === 0) return null
  const next = items.filter((item) => !drop.has(item.path))
  return next.length !== items.length ? next : null
}

/** assign each path to `group` (null/'' clears the group); null when no entry changed */
export function assignStarredGroup(
  items: StarredItem[],
  filePaths: string[],
  group: string | null,
): StarredItem[] | null {
  const targets = new Set(filePaths)
  if (targets.size === 0) return null
  const name = group?.trim() ?? ''
  let changed = false
  const next = items.map((item) => {
    if (!targets.has(item.path)) return item
    if ((item.group ?? '') === name) return item
    changed = true
    return name ? { path: item.path, group: name } : { path: item.path }
  })
  return changed ? next : null
}

/** re-key one entry after a disk rename, keeping its group; null when absent */
export function renameStarredItem(
  items: StarredItem[],
  oldPath: string,
  newPath: string,
): StarredItem[] | null {
  if (!items.some((item) => item.path === oldPath)) return null
  return items.map((item) => (item.path === oldPath ? { ...item, path: newPath } : item))
}

/** group names in first-seen order (drives the Starred view's filter pills) */
export function starredGroupNames(items: StarredItem[]): string[] {
  const names: string[] = []
  for (const item of items) {
    if (item.group && !names.includes(item.group)) names.push(item.group)
  }
  return names
}

/** path → group for the grouped entries only (ungrouped paths are absent) */
export function starredGroupMap(items: StarredItem[]): Map<string, string> {
  const map = new Map<string, string>()
  for (const item of items) {
    if (item.group) map.set(item.path, item.group)
  }
  return map
}
