import type { DocsTabInfo } from '../shared/ipc'

/** the id of the document window after or before the focused one, round the end; null when there is no other */
export function neighbourWindow(tabs: readonly DocsTabInfo[], dir: 1 | -1): string | null {
  const at = tabs.findIndex((tab) => tab.focused)
  if (at < 0 || tabs.length < 2) return null
  return tabs[(at + dir + tabs.length) % tabs.length].id
}
