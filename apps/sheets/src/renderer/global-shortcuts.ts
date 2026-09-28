/**
 * Window-level Excel key chords that map onto ribbon commands (Univer-native
 * ones live in excel-shortcuts.ts). Pure so the table can be unit-tested;
 * ExcelShell's keydown handler resolves and dispatches.
 *
 * Gates: `grid` fires whenever a text field does not own focus (dialog
 * openers work on a selected cell); `sheet` additionally requires no in-cell
 * edit because the command writes to the range; `idle` only checks editing;
 * `formulaBar` also accepts Univer's formula bar, which `grid` skips.
 * Nothing fires while a modal is open: the grid keeps focus behind the
 * overlay, so the grid-target guard alone cannot see the dialog.
 */
import { formatShortcutCommand } from './excel-format-shortcuts'

export const MODAL_MASK_SELECTOR = '.dialog-backdrop'

export function isModalOpen(root: Pick<Document, 'querySelector'> | null = document): boolean {
  if (!root) return false
  return root.querySelector(MODAL_MASK_SELECTOR) !== null
}

export type GlobalShortcutAction =
  | { readonly kind: 'dialog'; readonly dialog: 'formatCells' | 'goTo' }
  | { readonly kind: 'command'; readonly command: string }

export interface GlobalShortcutKeyEvent {
  readonly code: string
  readonly metaKey: boolean
  readonly ctrlKey: boolean
  readonly altKey: boolean
  readonly shiftKey: boolean
  readonly defaultPrevented: boolean
}

export interface GlobalShortcutGuards {
  readonly isMac: boolean
  readonly modalOpen: boolean
  readonly cellEditing: boolean
  readonly gridTarget: boolean
  readonly formulaBarTarget: boolean
}

type Gate = 'grid' | 'sheet' | 'idle' | 'formulaBar'

interface Chord {
  /** Ctrl or Cmd (either is accepted, like the rest of the app) */
  readonly mod?: boolean
  readonly shift?: boolean
  readonly alt?: boolean
  /** Excel-for-Mac binding: Cmd on mac only */
  readonly cmdMac?: boolean
}

interface Binding {
  readonly code: string
  readonly chord: Chord
  readonly gate: Gate
  readonly action: GlobalShortcutAction
}

const command = (command: string): GlobalShortcutAction => ({ kind: 'command', command })
const FORMAT_CELLS: GlobalShortcutAction = { kind: 'dialog', dialog: 'formatCells' }
const GO_TO: GlobalShortcutAction = { kind: 'dialog', dialog: 'goTo' }

const BINDINGS: readonly Binding[] = [
  { code: 'Digit1', chord: { mod: true }, gate: 'grid', action: FORMAT_CELLS },
  { code: 'KeyG', chord: { mod: true }, gate: 'grid', action: GO_TO },
  { code: 'F5', chord: {}, gate: 'grid', action: GO_TO },
  {
    code: 'Backquote',
    chord: { mod: true },
    gate: 'grid',
    action: command('toggle-show-formulas'),
  },
  { code: 'KeyK', chord: { mod: true }, gate: 'grid', action: command('link-open') },
  { code: 'F2', chord: { shift: true }, gate: 'grid', action: command('note-open') },
  { code: 'F3', chord: { shift: true }, gate: 'grid', action: command('insert-function-open') },
  { code: 'F3', chord: { mod: true }, gate: 'grid', action: command('name-manager-open') },
  {
    code: 'KeyU',
    chord: { mod: true, shift: true },
    gate: 'formulaBar',
    action: command('formula-bar-toggle'),
  },
  { code: 'Digit5', chord: { mod: true }, gate: 'sheet', action: command('strike') },
  { code: 'Equal', chord: { alt: true }, gate: 'sheet', action: command('autofn:SUM') },
  { code: 'Semicolon', chord: { mod: true }, gate: 'sheet', action: command('insert-now:date') },
  {
    code: 'Semicolon',
    chord: { mod: true, shift: true },
    gate: 'sheet',
    action: command('insert-now:time'),
  },
  { code: 'F11', chord: { shift: true }, gate: 'sheet', action: command('insert-sheet') },
  { code: 'KeyE', chord: { mod: true }, gate: 'sheet', action: command('flash-fill') },
  {
    code: 'KeyV',
    chord: { mod: true, shift: true },
    gate: 'sheet',
    action: command('paste-special:value'),
  },
  { code: 'BracketLeft', chord: { mod: true }, gate: 'sheet', action: command('trace-precedents') },
  {
    code: 'BracketRight',
    chord: { mod: true },
    gate: 'sheet',
    action: command('trace-dependents'),
  },
  {
    code: 'ArrowRight',
    chord: { alt: true, shift: true },
    gate: 'sheet',
    action: command('outline-group:rows'),
  },
  {
    code: 'ArrowLeft',
    chord: { alt: true, shift: true },
    gate: 'sheet',
    action: command('outline-ungroup:rows'),
  },
  {
    code: 'Period',
    chord: { mod: true, shift: true },
    gate: 'sheet',
    action: command('font-size-step:1'),
  },
  {
    code: 'Comma',
    chord: { mod: true, shift: true },
    gate: 'sheet',
    action: command('font-size-step:-1'),
  },
  {
    code: 'KeyF',
    chord: { cmdMac: true, shift: true },
    gate: 'sheet',
    action: command('filter-toggle'),
  },
  {
    code: 'KeyT',
    chord: { cmdMac: true, shift: true },
    gate: 'sheet',
    action: command('autofn:SUM'),
  },
  { code: 'F9', chord: {}, gate: 'idle', action: command('calculate-now') },
  { code: 'F9', chord: { shift: true }, gate: 'idle', action: command('calculate-sheet') },
  { code: 'PageDown', chord: {}, gate: 'grid', action: command('page-row:1') },
  { code: 'PageUp', chord: {}, gate: 'grid', action: command('page-row:-1') },
  { code: 'PageDown', chord: { alt: true }, gate: 'grid', action: command('page-col:1') },
  { code: 'PageUp', chord: { alt: true }, gate: 'grid', action: command('page-col:-1') },
]

function chordMatches(event: GlobalShortcutKeyEvent, chord: Chord, isMac: boolean): boolean {
  if (event.shiftKey !== Boolean(chord.shift) || event.altKey !== Boolean(chord.alt)) return false
  if (chord.cmdMac) return isMac && event.metaKey && !event.ctrlKey
  return (event.metaKey || event.ctrlKey) === Boolean(chord.mod)
}

function gateOpen(gate: Gate, guards: GlobalShortcutGuards): boolean {
  switch (gate) {
    case 'grid':
      return guards.gridTarget
    case 'sheet':
      return !guards.cellEditing && guards.gridTarget
    case 'idle':
      return !guards.cellEditing
    case 'formulaBar':
      return guards.gridTarget || guards.formulaBarTarget
  }
}

/**
 * Action for a keydown, or null. A key Univer's dispatcher already
 * consumed arrives with defaultPrevented (its own Alt+= QuickSum on
 * win/linux, for one) and must not fire a second time here.
 */
export function resolveGlobalShortcut(
  event: GlobalShortcutKeyEvent,
  guards: GlobalShortcutGuards,
): GlobalShortcutAction | null {
  if (guards.modalOpen || event.defaultPrevented) return null
  for (const binding of BINDINGS) {
    if (binding.code !== event.code || !chordMatches(event, binding.chord, guards.isMac)) continue
    return gateOpen(binding.gate, guards) ? binding.action : null
  }
  // number-format / border chords (Ctrl+Shift+digit, Ctrl+Shift+&/_, mac ⌘⌥
  // borders) write to the range: same gate as the table's `sheet` rows
  const formatCommand = formatShortcutCommand(event)
  if (formatCommand && gateOpen('sheet', guards)) return { kind: 'command', command: formatCommand }
  return null
}
