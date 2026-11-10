import { useSyncExternalStore } from 'react'
import { SHORTCUTS_CHANGED_EVENT, formatChord, getShortcutMap } from './shortcuts'
import type { ShortcutId } from './shortcuts'

function subscribe(cb: () => void): () => void {
  window.addEventListener(SHORTCUTS_CHANGED_EVENT, cb)
  return () => window.removeEventListener(SHORTCUTS_CHANGED_EVENT, cb)
}

/** first binding of an action as display text ('Ctrl+K'), '' when unbound; follows rebinding live */
export function useShortcutHint(id: ShortcutId): string {
  const map = useSyncExternalStore(subscribe, getShortcutMap, getShortcutMap)
  const first = map[id][0]
  return first ? formatChord(first) : ''
}
