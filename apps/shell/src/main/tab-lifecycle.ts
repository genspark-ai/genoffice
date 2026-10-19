import type { TabKind } from '../shared/tabs-api'

export const MAX_RECENTLY_CLOSED = 20

export interface TabCloseTarget {
  id: string
  kind?: TabKind
  closable?: boolean
}

export interface TabCloseManager {
  list(): readonly TabCloseTarget[]
  closeTab(id: string): Promise<void>
}

/**
 * Closes tabs sequentially one after another:
 * - Skips tabs that no longer exist, are marked unclosable, or are 'home'.
 * - If closing a tab is cancelled (e.g. user selects Cancel on an unsaved-changes prompt,
 *   leaving the tab open in the manager's list), sequence halts immediately to preserve unsaved work.
 */
export async function closeTabsInOrder(
  tabs: TabCloseManager,
  ids: readonly string[],
): Promise<void> {
  for (const id of ids) {
    const list = tabs.list()
    const target = list.find((tab) => tab.id === id)
    if (!target) continue
    if (target.id === 'home' || target.closable === false) continue

    await tabs.closeTab(id)

    // If the tab remains in the list after closeTab (e.g. cancelled close confirmation),
    // stop the chain immediately and leave all remaining tabs open.
    if (tabs.list().some((tab) => tab.id === id)) {
      return
    }
  }
}

export interface ClosedTabEntry {
  kind: TabKind
  filePath: string
}

export interface RecentlyClosedTracker {
  push(closed: { kind?: TabKind; filePath?: string }): void
  reopenNext(
    openDocument: (filePath: string) => boolean,
    fileExists?: (filePath: string) => boolean,
  ): boolean
  getEntries(): readonly ClosedTabEntry[]
  clear(): void
}

/**
 * Tracks recently closed document files with a bounded history stack.
 * - Filters out untitled documents, 'home', and 'help'.
 * - Deduplicates on repeated closures by moving the path to the top of the stack.
 * - When reopening, skips files that no longer exist on disk.
 */
export function createRecentlyClosedTracker(
  maxEntries = MAX_RECENTLY_CLOSED,
): RecentlyClosedTracker {
  const stack: ClosedTabEntry[] = []

  return {
    push(closed) {
      if (!closed.filePath) return
      if (closed.kind === 'home' || closed.kind === 'help') return
      const at = stack.findIndex((entry) => entry.filePath === closed.filePath)
      if (at >= 0) stack.splice(at, 1)
      stack.push({ kind: closed.kind ?? 'docs', filePath: closed.filePath })
      if (stack.length > maxEntries) stack.shift()
    },
    reopenNext(openDocument, fileExists) {
      while (stack.length > 0) {
        const closed = stack.pop()!
        if (fileExists && !fileExists(closed.filePath)) continue
        if (openDocument(closed.filePath)) return true
      }
      return false
    },
    getEntries() {
      return stack
    },
    clear() {
      stack.length = 0
    },
  }
}
