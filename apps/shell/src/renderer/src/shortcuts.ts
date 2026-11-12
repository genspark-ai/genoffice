import type { ShortcutId } from '../../shared/shortcuts-api'
import * as chords from '../../shared/chords'
import type { KeyLike } from '../../shared/chords'

/**
 * Keyboard shortcut registry for the GenOffice shell.
 *
 * Chords are stored as canonical strings: modifiers in the fixed order
 * `Mod`, `Ctrl`, `Alt`, `Shift` followed by one key, joined with '+'
 * (e.g. 'Mod+Shift+T', 'Alt+ArrowUp'). `Mod` is Ctrl on Windows/Linux and
 * Cmd on macOS, so a binding stays portable; `Ctrl` is the literal Control key
 * and only differs from `Mod` on macOS.
 */

export type { ShortcutId }

export type ShortcutGroup = 'navigation' | 'general'

export interface ShortcutDef {
  id: ShortcutId
  group: ShortcutGroup
  defaults: readonly string[]
}

export const SHORTCUT_DEFS: readonly ShortcutDef[] = [
  { id: 'openConversationPicker', group: 'navigation', defaults: ['Mod+K'] },
  { id: 'openFileSearch', group: 'navigation', defaults: ['Mod+P'] },
  { id: 'newConversation', group: 'navigation', defaults: ['Mod+N', 'Mod+Shift+O', 'Mod+Shift+L'] },
  { id: 'goBack', group: 'navigation', defaults: ['Mod+['] },
  { id: 'goForward', group: 'navigation', defaults: ['Mod+]'] },
  { id: 'newTab', group: 'navigation', defaults: ['Mod+T'] },
  { id: 'reopenClosedTab', group: 'navigation', defaults: ['Mod+Shift+T'] },
  { id: 'previousTab', group: 'navigation', defaults: ['Alt+ArrowUp'] },
  { id: 'nextTab', group: 'navigation', defaults: ['Alt+ArrowDown'] },
  { id: 'openSettings', group: 'general', defaults: ['Mod+,'] },
  { id: 'openRecycleBin', group: 'general', defaults: ['Mod+Shift+Delete'] },
]

export const IS_MAC = typeof navigator !== 'undefined' && /mac/i.test(navigator.platform ?? '')

// platform-defaulted wrappers over the shared pure helpers
export const normalizeChord = chords.normalizeChord
export const isUsableChord = chords.isUsableChord
export const isReservedChord = chords.isReservedChord
export function chordFromEvent(e: KeyLike, isMac: boolean = IS_MAC): string | null {
  return chords.chordFromEvent(e, isMac)
}
export function formatChordParts(chord: string, isMac: boolean = IS_MAC) {
  return chords.formatChordParts(chord, isMac)
}
export function formatChord(chord: string, isMac: boolean = IS_MAC): string {
  return chords.formatChord(chord, isMac)
}

// ── persisted overrides ──────────────────────────────────────

export type ShortcutMap = Record<ShortcutId, string[]>

const STORAGE_KEY = 'genoffice.shortcuts'
export const SHORTCUTS_CHANGED_EVENT = 'genoffice:shortcuts-changed'

export function defaultShortcutMap(): ShortcutMap {
  const map = {} as ShortcutMap
  for (const def of SHORTCUT_DEFS) map[def.id] = [...def.defaults]
  return map
}

/** merge stored overrides over the defaults; unknown ids and malformed chords are dropped */
export function loadShortcutMap(storage: Pick<Storage, 'getItem'> | undefined): ShortcutMap {
  const map = defaultShortcutMap()
  try {
    const raw = storage?.getItem(STORAGE_KEY)
    if (!raw) return map
    const data = JSON.parse(raw) as Record<string, unknown>
    for (const def of SHORTCUT_DEFS) {
      const stored = data[def.id]
      if (!Array.isArray(stored)) continue
      const chords: string[] = []
      for (const item of stored) {
        const norm = typeof item === 'string' ? normalizeChord(item) : null
        if (norm && isUsableChord(norm) && !isReservedChord(norm) && !chords.includes(norm))
          chords.push(norm)
      }
      map[def.id] = chords
    }
  } catch {
    return defaultShortcutMap()
  }
  return resolveConflicts(map)
}

/** a chord may drive one action only: the first action in registry order keeps it */
function resolveConflicts(map: ShortcutMap): ShortcutMap {
  const seen = new Set<string>()
  for (const def of SHORTCUT_DEFS) {
    map[def.id] = map[def.id].filter((c) => {
      if (seen.has(c)) return false
      seen.add(c)
      return true
    })
  }
  return map
}

export function saveShortcutMap(
  storage: Pick<Storage, 'setItem' | 'removeItem'> | undefined,
  map: ShortcutMap,
): void {
  try {
    const defaults = defaultShortcutMap()
    const overrides: Partial<ShortcutMap> = {}
    for (const def of SHORTCUT_DEFS) {
      if (JSON.stringify(map[def.id]) !== JSON.stringify(defaults[def.id]))
        overrides[def.id] = map[def.id]
    }
    if (Object.keys(overrides).length === 0) storage?.removeItem(STORAGE_KEY)
    else storage?.setItem(STORAGE_KEY, JSON.stringify(overrides))
  } catch {
    /* storage unavailable: the change lasts until the window closes */
  }
}

/** which action currently owns `chord`, if any */
export function findOwner(map: ShortcutMap, chord: string): ShortcutId | null {
  for (const def of SHORTCUT_DEFS) if (map[def.id].includes(chord)) return def.id
  return null
}

/** action bound to a chord, or null */
export function resolveShortcut(map: ShortcutMap, chord: string): ShortcutId | null {
  return findOwner(map, chord)
}

// ── live store ───────────────────────────────────────────────

let current: ShortcutMap | null = null

export function getShortcutMap(): ShortcutMap {
  if (!current)
    current = loadShortcutMap(typeof localStorage === 'undefined' ? undefined : localStorage)
  return current
}

export function setShortcutMap(next: ShortcutMap): void {
  current = resolveConflicts({ ...next })
  saveShortcutMap(typeof localStorage === 'undefined' ? undefined : localStorage, current)
  if (typeof window !== 'undefined') window.dispatchEvent(new Event(SHORTCUTS_CHANGED_EVENT))
}

export function resetShortcutMap(): void {
  setShortcutMap(defaultShortcutMap())
}

// ── action bus: UI pieces register handlers, the dispatcher fires them ──

type Handler = (id: ShortcutId) => boolean | void
const handlers = new Set<Handler>()

/** register a handler; return true from it to mark the shortcut as handled */
export function onShortcut(handler: Handler): () => void {
  handlers.add(handler)
  return () => handlers.delete(handler)
}

export function runShortcut(id: ShortcutId): boolean {
  for (const h of [...handlers]) if (h(id) === true) return true
  return false
}

/** while Settings is recording a new chord the dispatcher must stay quiet */
let recording = false
export function setShortcutRecording(on: boolean): void {
  recording = on
}
export function isShortcutRecording(): boolean {
  return recording
}
