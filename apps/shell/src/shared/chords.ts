/**
 * Pure chord helpers shared by the shell renderer and the main process.
 *
 * Chords are canonical strings: modifiers in the fixed order `Mod`, `Ctrl`,
 * `Alt`, `Shift` followed by one key, joined with '+' (e.g. 'Mod+Shift+T',
 * 'Alt+ArrowUp'). `Mod` is Ctrl on Windows/Linux and Cmd on macOS, so a binding
 * stays portable; `Ctrl` is the literal Control key and only differs from
 * `Mod` on macOS. Nothing here touches the DOM or Electron.
 */

/** chords the app (or the OS text fields) already owns — they can't be rebound */
const RESERVED = new Set([
  'Mod+C',
  'Mod+V',
  'Mod+X',
  'Mod+A',
  'Mod+Z',
  'Mod+Y',
  'Mod+Shift+Z',
  'Mod+F',
  'Mod+W',
  'Mod+Q',
])

const MOD_ORDER = ['Mod', 'Ctrl', 'Alt', 'Shift'] as const

/** physical-key → character for punctuation, so Shift/AltGr layouts still match */
const CODE_CHARS: Record<string, string> = {
  BracketLeft: '[',
  BracketRight: ']',
  Comma: ',',
  Period: '.',
  Slash: '/',
  Backslash: '\\',
  Semicolon: ';',
  Quote: "'",
  Backquote: '`',
  Minus: '-',
  Equal: '=',
}

const NAMED_KEYS = new Set([
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
  'Enter',
  'Escape',
  'Tab',
  'Space',
  'Backspace',
  'Delete',
  'Home',
  'End',
  'PageUp',
  'PageDown',
])

export interface KeyLike {
  key: string
  code?: string
  ctrlKey: boolean
  metaKey: boolean
  altKey: boolean
  shiftKey: boolean
}

/** the key part of a keyboard event, layout-independent; null for bare modifiers / unknown keys */
function keyName(e: Pick<KeyLike, 'key' | 'code'>): string | null {
  const { key, code } = e
  if (key === 'Control' || key === 'Shift' || key === 'Alt' || key === 'Meta' || key === 'AltGraph')
    return null
  if (code?.startsWith('Key')) return code.slice(3).toUpperCase()
  if (code?.startsWith('Digit')) return code.slice(5)
  if (code && CODE_CHARS[code]) return CODE_CHARS[code]!
  if (key === ' ') return 'Space'
  if (/^F\d{1,2}$/.test(key)) return key
  if (NAMED_KEYS.has(key)) return key
  if (key.length === 1) return key.toUpperCase()
  return null
}

/** canonical chord for a keyboard event, or null when it is only a modifier press */
export function chordFromEvent(e: KeyLike, isMac: boolean): string | null {
  const name = keyName(e)
  if (!name) return null
  const mods: string[] = []
  if (isMac ? e.metaKey : e.ctrlKey) mods.push('Mod')
  // on macOS the literal Control key is its own modifier; elsewhere meta is the OS key and is ignored
  if (isMac && e.ctrlKey) mods.push('Ctrl')
  if (e.altKey) mods.push('Alt')
  if (e.shiftKey) mods.push('Shift')
  return [...mods, name].join('+')
}

/** normalise a stored/typed chord (modifier order, key case); null when malformed */
export function normalizeChord(chord: string): string | null {
  const parts = chord.split('+')
  // a trailing '+' key ("Mod++") splits into an empty last part
  const key = parts[parts.length - 1] === '' ? '+' : parts[parts.length - 1]!
  const rawMods = parts.slice(0, key === '+' ? -2 : -1)
  const mods = new Set<string>()
  for (const m of rawMods) {
    if (!(MOD_ORDER as readonly string[]).includes(m)) return null
    mods.add(m)
  }
  if (!key) return null
  const k = key.length === 1 ? key.toUpperCase() : key
  if (k.length > 1 && !NAMED_KEYS.has(k) && !/^F\d{1,2}$/.test(k)) return null
  return [...MOD_ORDER.filter((m) => mods.has(m)), k].join('+')
}

/** a binding must carry a modifier (or be an F-key) so ordinary typing is never hijacked */
export function isUsableChord(chord: string): boolean {
  const parts = chord.split('+')
  const key = parts[parts.length - 1]!
  if (/^F\d{1,2}$/.test(key)) return true
  return parts.length > 1 && parts.slice(0, -1).some((m) => m === 'Mod' || m === 'Ctrl' || m === 'Alt')
}

export function isReservedChord(chord: string): boolean {
  return RESERVED.has(chord)
}

const SYMBOLS_MAC: Record<string, string> = { Mod: '⌘', Ctrl: '⌃', Alt: '⌥', Shift: '⇧' }
const KEY_GLYPHS: Record<string, string> = {
  ArrowUp: '↑',
  ArrowDown: '↓',
  ArrowLeft: '←',
  ArrowRight: '→',
  Escape: 'Esc',
  Delete: 'Del',
  Space: 'Space',
}

/** display parts for a chord: ['Ctrl+', 'K'] style, matching the settings keycaps */
export function formatChordParts(chord: string, isMac: boolean): { mods: string; key: string } {
  const parts = chord.split('+')
  const key = parts[parts.length - 1] === '' ? '+' : parts[parts.length - 1]!
  const mods = parts.slice(0, key === '+' ? -2 : -1)
  const label = (m: string): string => (isMac ? SYMBOLS_MAC[m]! : m === 'Mod' ? 'Ctrl' : m)
  const modText = mods.map(label).join(isMac ? '' : '+')
  return {
    mods: modText ? (isMac ? modText : `${modText}+`) : '',
    key: KEY_GLYPHS[key] ?? key,
  }
}

export function formatChord(chord: string, isMac: boolean): string {
  const { mods, key } = formatChordParts(chord, isMac)
  return `${mods}${key}`
}

