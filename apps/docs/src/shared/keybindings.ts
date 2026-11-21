/**
 * User-assigned keyboard shortcuts (Tools ▸ Customize Keyboard…). A chord is
 * written the way Electron writes an accelerator with the modifiers spelled
 * out for this machine ('Cmd+Shift+K', 'Ctrl+Alt+F5'), so the same string
 * serves the menu, the key handling and the stored file.
 */

/** binding id (a command id, or `command:payload`) -> every chord it answers to, in place of its defaults */
export type KeyOverrides = Record<string, string[]>

const MODIFIERS = ['Ctrl', 'Alt', 'Shift', 'Cmd'] as const
type Modifier = (typeof MODIFIERS)[number]

const NAMED_KEYS = [
  'Space',
  'Enter',
  'Tab',
  'Up',
  'Down',
  'Left',
  'Right',
  'Home',
  'End',
  'PageUp',
  'PageDown',
  'Delete',
  'Backspace',
]
const KEY = new RegExp(`^(?:[A-Z0-9]|F\\d{1,2}|[-=,./;'\`\\[\\]\\\\]|${NAMED_KEYS.join('|')})$`)

export function makeChord(modifiers: Iterable<Modifier>, key: string): string {
  const held = new Set(modifiers)
  return [...MODIFIERS.filter((m) => held.has(m)), key].join('+')
}

/** a chord in canonical form, or null when the text is not one */
export function toChord(text: string, isMac: boolean): string | null {
  const parts = text.split('+')
  let key = parts.pop() ?? ''
  // a trailing '+' is the plus key itself, which sits on '='
  if (key === '' && parts.at(-1) === '') {
    parts.pop()
    key = '='
  }
  if (key === 'Plus') key = '='
  if (key === 'Return') key = 'Enter'
  if (key.length === 1) key = key.toUpperCase()
  if (!KEY.test(key)) return null
  const held: Modifier[] = []
  for (const part of parts) {
    const modifier =
      part === 'CmdOrCtrl' || part === 'CommandOrControl'
        ? isMac
          ? 'Cmd'
          : 'Ctrl'
        : part === 'Command'
          ? 'Cmd'
          : part === 'Control'
            ? 'Ctrl'
            : part === 'Option'
              ? 'Alt'
              : part
    if (!MODIFIERS.includes(modifier as Modifier)) return null
    held.push(modifier as Modifier)
  }
  if (held.includes('Cmd') && !isMac) return null
  return makeChord(held, key)
}

/** only what a user may pick: Ctrl or Cmd held, or a function key; typing and caret keys stay the document's */
export function isAssignableChord(chord: string): boolean {
  const parts = chord.split('+')
  const key = parts.pop() ?? ''
  return parts.includes('Ctrl') || parts.includes('Cmd') || /^F\d{1,2}$/.test(key)
}

/** `command` or `command:payload`; a payload may be a font or style name, so any printable text */
function isBindingId(id: string): boolean {
  const at = id.indexOf(':')
  const payload = at < 0 ? null : id.slice(at + 1)
  if (!/^[a-z0-9-]+$/i.test(at < 0 ? id : id.slice(0, at))) return false
  if (payload === null) return true
  return payload.length >= 1 && payload.length <= 120 && !/\p{Cc}/u.test(payload)
}

const MAX_BINDINGS = 2000
const MAX_CHORDS = 8

/** keeps only well-formed entries from a value read from disk or sent over IPC */
export function sanitizeKeyOverrides(value: unknown, isMac: boolean): KeyOverrides {
  const out: KeyOverrides = {}
  if (!value || typeof value !== 'object' || Array.isArray(value)) return out
  for (const [id, chords] of Object.entries(value).slice(0, MAX_BINDINGS)) {
    if (!isBindingId(id) || !Array.isArray(chords)) continue
    const kept: string[] = []
    for (const chord of chords.slice(0, MAX_CHORDS)) {
      const canonical = typeof chord === 'string' ? toChord(chord, isMac) : null
      if (canonical && !kept.includes(canonical)) kept.push(canonical)
    }
    out[id] = kept
  }
  return out
}

export function bindingId(command: string, payload?: string): string {
  return payload ? `${command}:${payload}` : command
}

/**
 * The chord a menu item shows: the one assigned last for a customized
 * binding, the built-in accelerator otherwise. A chord on a typing key
 * (Enter, Tab, an arrow) is left to the key handling in the document: as a
 * menu accelerator it would also fire from a comment box, which submits on
 * Cmd/Ctrl+Enter.
 */
export function menuAccelerator(
  overrides: KeyOverrides,
  id: string,
  builtIn: string | undefined,
): string | undefined {
  if (!Object.hasOwn(overrides, id)) return builtIn
  return overrides[id].filter((chord) => !NAMED_KEYS.includes(chord.split('+').at(-1) ?? '')).at(-1)
}

/** whether a binding still answers to one of its built-in accelerators */
export function keepsAccelerator(
  overrides: KeyOverrides,
  id: string,
  accelerator: string,
  isMac: boolean,
): boolean {
  if (!Object.hasOwn(overrides, id)) return true
  const chord = toChord(accelerator, isMac)
  return !!chord && overrides[id].includes(chord)
}
