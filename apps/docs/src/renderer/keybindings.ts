/**
 * Which chord runs which command once the user's assignments (Tools ▸
 * Customize Keyboard…) are laid over the built-in ones. A binding nobody
 * customized is left to the code that has always handled it; a customized one
 * is run from here, and a built-in chord it gave up is swallowed.
 */
import { COMMANDS, NONE, type ActionCommandId, type CommandDef } from '../shared/commands'
import {
  bindingId,
  isAssignableChord,
  makeChord,
  toChord,
  type KeyOverrides,
} from '../shared/keybindings'
import type { StringKey } from './i18n/locale'
import { SHORTCUTS, type ShortcutDef } from './shortcuts'

export const IS_MAC = typeof navigator !== 'undefined' && /mac/i.test(navigator.platform)

const GLYPH_MODIFIERS: Record<string, 'Cmd' | 'Ctrl' | 'Alt' | 'Shift'> = {
  '⌘': 'Cmd',
  '⌃': 'Ctrl',
  '⌥': 'Alt',
  '⇧': 'Shift',
}
const KEY_GLYPHS: Record<string, string> = {
  Enter: '⏎',
  Space: '␣',
  Up: '↑',
  Down: '↓',
  Left: '←',
  Right: '→',
  Backspace: '⌫',
  Delete: '⌦',
  NumEnter: '⌤',
  NumPlus: 'Num+',
  NumMinus: 'Num-',
}
const GLYPH_KEYS = Object.fromEntries(
  Object.entries(KEY_GLYPHS).map(([key, glyph]) => [glyph, key]),
)

/** one chord of the shortcut list, which writes Mac glyphs ('⇧⌘S') or 'Ctrl+Shift+S' */
function chordFromNotation(text: string, isMac: boolean): string | null {
  if (!/[⌘⌃⌥⇧]/.test(text)) return toChord(text, isMac)
  const held: ('Cmd' | 'Ctrl' | 'Alt' | 'Shift')[] = []
  let rest = text
  while (rest[0] in GLYPH_MODIFIERS) {
    const modifier = GLYPH_MODIFIERS[rest[0]]
    held.push(modifier === 'Cmd' && !isMac ? 'Ctrl' : modifier)
    rest = rest.slice(1)
  }
  return toChord(makeChord(held, GLYPH_KEYS[rest] ?? rest), isMac)
}

const chordsOf = (text: string, isMac: boolean): string[] =>
  text.split(' / ').flatMap((part) => chordFromNotation(part, isMac) ?? [])

export function shortcutChords(def: ShortcutDef, isMac: boolean): string[] {
  return chordsOf(isMac ? def.keys : (def.win ?? def.keys), isMac)
}

/** the chord as the user reads it: glyphs on the Mac, Ctrl+Alt+Shift+key elsewhere */
export function formatChord(chord: string, isMac = IS_MAC): string {
  if (!isMac) return chord
  const parts = chord.split('+')
  // a trailing '+' cannot occur: the plus key is written '='
  const key = parts.pop() ?? ''
  const glyph = (name: string, mark: string) => (parts.includes(name) ? mark : '')
  return (
    glyph('Ctrl', '⌃') +
    glyph('Alt', '⌥') +
    glyph('Shift', '⇧') +
    glyph('Cmd', '⌘') +
    (KEY_GLYPHS[key] ?? key)
  )
}

const CODE_KEYS: Record<string, string> = {
  Minus: '-',
  Equal: '=',
  Comma: ',',
  Period: '.',
  Slash: '/',
  Semicolon: ';',
  Quote: "'",
  Backquote: '`',
  BracketLeft: '[',
  BracketRight: ']',
  Backslash: '\\',
  Space: 'Space',
  Enter: 'Enter',
  Tab: 'Tab',
  ArrowUp: 'Up',
  ArrowDown: 'Down',
  ArrowLeft: 'Left',
  ArrowRight: 'Right',
  Home: 'Home',
  End: 'End',
  PageUp: 'PageUp',
  PageDown: 'PageDown',
  Delete: 'Delete',
  Backspace: 'Backspace',
  NumpadEnter: 'NumEnter',
  NumpadAdd: 'NumPlus',
  NumpadSubtract: 'NumMinus',
}

export type ChordEvent = Pick<
  KeyboardEvent,
  'key' | 'code' | 'ctrlKey' | 'altKey' | 'shiftKey' | 'metaKey'
>

/** the chord a key press spells, or null for a key no chord is made of (a bare modifier, a keypad digit) */
export function chordFromEvent(event: ChordEvent, isMac = IS_MAC): string | null {
  // a Latin letter is the layout's own answer (Dvorak); anything else falls back to the key's position
  const key = /^[a-z]$/i.test(event.key)
    ? event.key.toUpperCase()
    : /^F\d{1,2}$/.test(event.key)
      ? event.key
      : event.key === 'Clear'
        ? 'Clear'
        : (/^Key([A-Z])$/.exec(event.code)?.[1] ??
          /^Digit(\d)$/.exec(event.code)?.[1] ??
          CODE_KEYS[event.code])
  if (!key) return null
  const held: ('Cmd' | 'Ctrl' | 'Alt' | 'Shift')[] = []
  if (event.ctrlKey) held.push('Ctrl')
  if (event.altKey) held.push('Alt')
  if (event.shiftKey) held.push('Shift')
  if (event.metaKey) {
    if (!isMac) return null
    held.push('Cmd')
  }
  return makeChord(held, key)
}

/** the command and payload a binding id stands for, or null when it names no runnable command */
export function bindingCommand(id: string): { cmd: ActionCommandId; payload?: string } | null {
  const at = id.indexOf(':')
  const cmd = at < 0 ? id : id.slice(0, at)
  if (!Object.hasOwn(COMMANDS, cmd)) return null
  const def: CommandDef = COMMANDS[cmd as ActionCommandId]
  if (def.role || def.planned) return null
  return { cmd: cmd as ActionCommandId, ...(at < 0 ? {} : { payload: id.slice(at + 1) }) }
}

interface Defaults {
  /** binding id -> its built-in chords */
  chords: Map<string, string[]>
  /** chords of built-in bindings that run no command, with the shortcut row that names them */
  fixed: Map<string, ShortcutDef | null>
}

// taken by the system or the window before a document command could see them
const RESERVED = {
  mac: ['Cmd+Q', 'Cmd+H', 'Alt+Cmd+H', 'Cmd+M', 'Cmd+,', 'Ctrl+Cmd+F', 'Cmd+`', 'Cmd+Tab'],
  other: ['Alt+F4', 'F11', 'Ctrl+Tab', 'Ctrl+Shift+Tab'],
}

function buildDefaults(isMac: boolean): Defaults {
  const chords = new Map<string, string[]>()
  const add = (id: string, chord: string | null) => {
    const list = chords.get(id) ?? []
    if (chord && !list.includes(chord)) list.push(chord)
    chords.set(id, list)
  }
  for (const [id, def] of Object.entries<CommandDef>(COMMANDS)) {
    if (!def.accelerator || def.role || def.planned) continue
    const accelerator =
      typeof def.accelerator === 'string'
        ? def.accelerator
        : def.accelerator[isMac ? 'mac' : 'other']
    if (accelerator) add(id, toChord(accelerator, isMac))
  }
  const fixed = new Map<string, ShortcutDef | null>()
  const taken = new Set([...chords.values()].flat())
  for (const def of SHORTCUTS) {
    for (const chord of shortcutChords(def, isMac)) {
      // a chord the command table gives to another command is that command's (Redo's second key is Repeat)
      if (!def.cmd) fixed.set(chord, def)
      else if (!taken.has(chord) || chords.get(bindingId(def.cmd, def.payload))?.includes(chord))
        add(bindingId(def.cmd, def.payload), chord)
    }
  }
  for (const text of RESERVED[isMac ? 'mac' : 'other']) {
    const chord = toChord(text, isMac)
    if (chord && !fixed.has(chord)) fixed.set(chord, null)
  }
  return { chords, fixed }
}

const defaultsByPlatform = new Map<boolean, Defaults>()
function defaults(isMac: boolean): Defaults {
  let built = defaultsByPlatform.get(isMac)
  if (!built) defaultsByPlatform.set(isMac, (built = buildDefaults(isMac)))
  return built
}

export function defaultChords(id: string, isMac = IS_MAC): readonly string[] {
  return defaults(isMac).chords.get(id) ?? []
}

/**
 * Commands whose built-in chord the window answers as well as the menu bar:
 * the menu takes the key first where it can, and a key press that never
 * reaches the menu still runs the command.
 */
const WINDOW_KEY_COMMANDS: readonly ActionCommandId[] = [
  'file-new-from-template',
  'edit-paste-special',
  'edit-find-next',
  'edit-find-previous',
  'edit-toggle-drawing',
  'view-focus',
  'view-ribbon',
  'insert-equation',
  'format-object',
  'format-selection-pane',
  'tools-thesaurus',
  'table-delete-rows',
  'replace',
]

const windowKeysByPlatform = new Map<boolean, Map<string, ActionCommandId>>()

export function windowKeyCommand(
  overrides: KeyOverrides,
  chord: string,
  isMac = IS_MAC,
): ActionCommandId | null {
  let keys = windowKeysByPlatform.get(isMac)
  if (!keys) {
    keys = new Map()
    for (const id of WINDOW_KEY_COMMANDS)
      for (const built of defaultChords(id, isMac)) keys.set(built, id)
    windowKeysByPlatform.set(isMac, keys)
  }
  const id = keys.get(chord)
  // a customized command is the key layer's to run, on the chords the user left it
  return id && !Object.hasOwn(overrides, id) ? id : null
}

export interface KeyRun {
  cmd: ActionCommandId
  payload?: string
}

let documentKeys: Map<string, KeyRun & { id: string }> | null = null

/**
 * The command a chord of Word for Mac runs while the document has the keys.
 * It is decided ahead of the editor keymaps and Chromium's own editing keys,
 * which give several of these chords another meaning.
 */
export function documentKeyCommand(
  overrides: KeyOverrides,
  chord: string,
  isMac = IS_MAC,
): KeyRun | null {
  if (!isMac) return null
  if (!documentKeys) {
    documentKeys = new Map()
    for (const def of SHORTCUTS) {
      if (!def.cmd || !def.docKeys) continue
      const run = { id: bindingId(def.cmd, def.payload), cmd: def.cmd, payload: def.payload }
      for (const built of chordsOf(def.docKeys, true)) documentKeys.set(built, run)
    }
  }
  const run = documentKeys.get(chord)
  if (!run || Object.hasOwn(overrides, run.id)) return null
  return { cmd: run.cmd, ...(run.payload ? { payload: run.payload } : {}) }
}

export function effectiveChords(
  overrides: KeyOverrides,
  id: string,
  isMac = IS_MAC,
): readonly string[] {
  return Object.hasOwn(overrides, id) ? overrides[id] : defaultChords(id, isMac)
}

export type ChordOwner =
  | { kind: 'binding'; id: string }
  /** built in and not a command; `row` is null for a chord the system keeps */
  | { kind: 'fixed'; row: ShortcutDef | null }

export function chordOwner(
  overrides: KeyOverrides,
  chord: string,
  isMac = IS_MAC,
): ChordOwner | null {
  for (const [id, chords] of Object.entries(overrides)) {
    if (chords.includes(chord)) return { kind: 'binding', id }
  }
  const { chords, fixed } = defaults(isMac)
  for (const [id, list] of chords) {
    if (!Object.hasOwn(overrides, id) && list.includes(chord)) return { kind: 'binding', id }
  }
  const row = fixed.get(chord)
  return row !== undefined ? { kind: 'fixed', row } : null
}

function withChords(
  overrides: KeyOverrides,
  id: string,
  chords: string[],
  isMac: boolean,
): KeyOverrides {
  const next = { ...overrides }
  const builtIn = defaultChords(id, isMac)
  // back to exactly the built-in chords is no customization at all
  if (chords.length === builtIn.length && chords.every((chord, i) => chord === builtIn[i]))
    delete next[id]
  else next[id] = chords
  return next
}

export function removeChord(
  overrides: KeyOverrides,
  id: string,
  chord: string,
  isMac = IS_MAC,
): KeyOverrides {
  const kept = effectiveChords(overrides, id, isMac).filter((c) => c !== chord)
  return withChords(overrides, id, kept, isMac)
}

/** gives the chord to the binding, taking it from the one that had it; null when it cannot be assigned */
export function assignChord(
  overrides: KeyOverrides,
  id: string,
  chord: string,
  isMac = IS_MAC,
): KeyOverrides | null {
  if (!isAssignableChord(chord) || !bindingCommand(id)) return null
  const owner = chordOwner(overrides, chord, isMac)
  if (owner?.kind === 'fixed') return null
  if (owner?.id === id) return overrides
  const freed = owner ? removeChord(overrides, owner.id, chord, isMac) : overrides
  return withChords(freed, id, [...effectiveChords(freed, id, isMac), chord], isMac)
}

export type KeyVerdict =
  | { run: { cmd: ActionCommandId; payload?: string } }
  | 'swallow'
  /** nothing customized touches this chord: the built-in handling takes it */
  | 'pass'

export interface KeyTable {
  run: Map<string, { cmd: ActionCommandId; payload?: string }>
  /** built-in chords their binding no longer answers to */
  dropped: Set<string>
  /** every chord some built-in binding still answers to */
  builtIn: Set<string>
}

export function buildKeyTable(overrides: KeyOverrides, isMac = IS_MAC): KeyTable {
  const table: KeyTable = { run: new Map(), dropped: new Set(), builtIn: new Set() }
  const { chords, fixed } = defaults(isMac)
  for (const [id, list] of Object.entries(overrides)) {
    const command = bindingCommand(id)
    if (!command) continue
    for (const chord of list) table.run.set(chord, command)
    for (const chord of defaultChords(id, isMac))
      if (!list.includes(chord)) table.dropped.add(chord)
  }
  for (const [id, list] of chords) {
    if (!Object.hasOwn(overrides, id)) for (const chord of list) table.builtIn.add(chord)
  }
  for (const chord of fixed.keys()) table.builtIn.add(chord)
  return table
}

export function keyVerdict(table: KeyTable, chord: string, isMac = IS_MAC): KeyVerdict {
  const run = table.run.get(chord)
  if (run) return { run }
  if (table.dropped.has(chord)) return 'swallow'
  // the built-in handlers take Control for Command on the Mac
  if (isMac && !table.builtIn.has(chord) && chord.startsWith('Ctrl+') && !chord.includes('Cmd+')) {
    const parts = chord.slice('Ctrl+'.length).split('+')
    const key = parts.pop() ?? ''
    const twin = makeChord([...(parts as ('Alt' | 'Shift')[]), 'Cmd'], key)
    if (table.dropped.has(twin)) return 'swallow'
  }
  return 'pass'
}

/** a command as the Customize Keyboard dialog lists it */
export interface KeyCommand {
  id: string
  label: string
  /** the menu that holds it; null for a command only the ribbon or a shortcut reaches */
  menu: string | null
  /** where to find it: 'Format ▸ Font…' */
  path: string
}

/** what the main process knows of a command in the menu bar */
export interface MenuCatalogEntry {
  id: string
  label: string
  menu: string
  path: string
}

/** ribbon commands with no menu item and no built-in chord, offered so they can be given one */
const RIBBON_ONLY: readonly { cmd: ActionCommandId; labelKey: StringKey }[] = [
  { cmd: 'strikethrough', labelKey: 'ribbonStrikethrough' },
  { cmd: 'bullets', labelKey: 'ribbonBullets' },
  { cmd: 'numbering', labelKey: 'ribbonNumbering' },
  { cmd: 'insert-blank-page', labelKey: 'ribbonBlankPage' },
  { cmd: 'revision-previous', labelKey: 'ribbonPrevChange' },
  { cmd: 'revision-next', labelKey: 'ribbonNextChange' },
  { cmd: 'read-mode', labelKey: 'ribbonReadMode' },
]

/** ribbon commands that take a value: each value is a binding of its own (`font-size:12`) */
export const VALUE_COMMANDS: readonly { cmd: ActionCommandId; labelKey: StringKey }[] = [
  { cmd: 'font-size', labelKey: 'ribbonFontSizeTip' },
  { cmd: 'font-color', labelKey: 'ribbonFontColor' },
  { cmd: 'highlight', labelKey: 'ribbonTextHighlightColor' },
]

/**
 * The name of a binding whose payload is a value picked in the dialog (a
 * font, a style, a size, a colour); null for any other binding.
 */
export function valueBindingLabel(
  id: string,
  t: (key: StringKey) => string,
  styleName?: (styleId: string) => string | undefined,
): string | null {
  const binding = bindingCommand(id)
  if (!binding?.payload) return null
  const { cmd, payload } = binding
  if (cmd === 'font-family') return payload
  if (cmd === 'apply-style') return styleName?.(payload) ?? payload
  const valued = VALUE_COMMANDS.find((entry) => entry.cmd === cmd)
  if (!valued) return null
  const value =
    payload !== NONE ? payload : t(cmd === 'highlight' ? 'ribbonNoColor' : 'ribbonAutomatic')
  return `${t(valued.labelKey)}: ${value}`
}

const menuLabel = (label: string) => label.replace(/…$|\.\.\.$/, '').replace(/&/g, '')

export function listKeyCommands(
  catalog: readonly MenuCatalogEntry[],
  t: (key: StringKey) => string,
): KeyCommand[] {
  const byId = new Map<string, KeyCommand>()
  for (const entry of catalog) {
    if (bindingCommand(entry.id) && !byId.has(entry.id))
      byId.set(entry.id, { ...entry, label: menuLabel(entry.label) })
  }
  const offer = (id: string, label: string) => {
    if (!byId.has(id)) byId.set(id, { id, label, menu: null, path: '' })
  }
  for (const def of SHORTCUTS) {
    if (def.cmd) offer(bindingId(def.cmd, def.payload), t(def.labelKey) + (def.labelSuffix ?? ''))
  }
  for (const extra of RIBBON_ONLY) offer(extra.cmd, t(extra.labelKey))
  return [...byId.values()]
}
