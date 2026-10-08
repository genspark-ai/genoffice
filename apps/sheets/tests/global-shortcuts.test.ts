/**
 * Window-level Excel chords: routing per command, mac-only variants, and the
 * guards (text field, in-cell edit, open modal, already-handled key). Nothing
 * may reach the sheet while a modal is open: the grid keeps focus behind the
 * overlay, so the grid-target guard alone cannot see the dialog.
 */
import { describe, expect, it } from 'vitest'
import {
  isModalOpen,
  resolveGlobalShortcut,
  type GlobalShortcutAction,
  type GlobalShortcutGuards,
  type GlobalShortcutKeyEvent,
} from '../src/renderer/global-shortcuts'

type Mods = Partial<Omit<GlobalShortcutKeyEvent, 'code'>>

function keyEvent(code: string, mods: Mods = {}): GlobalShortcutKeyEvent {
  return {
    code,
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    shiftKey: false,
    defaultPrevented: false,
    ...mods,
  }
}

const ctrl = { ctrlKey: true }
const meta = { metaKey: true }
const alt = { altKey: true }
const shift = { shiftKey: true }
const ctrlShift = { ctrlKey: true, shiftKey: true }
const metaShift = { metaKey: true, shiftKey: true }
const altShift = { altKey: true, shiftKey: true }

const GRID: GlobalShortcutGuards = {
  isMac: false,
  modalOpen: false,
  cellEditing: false,
  gridTarget: true,
  formulaBarTarget: false,
}
const MAC: GlobalShortcutGuards = { ...GRID, isMac: true }
const TEXT_FIELD: GlobalShortcutGuards = { ...GRID, gridTarget: false }
const FORMULA_BAR: GlobalShortcutGuards = { ...TEXT_FIELD, formulaBarTarget: true }
const EDITING: GlobalShortcutGuards = { ...GRID, cellEditing: true }
const BEHIND_MODAL: GlobalShortcutGuards = { ...GRID, modalOpen: true }
const FORMAT_CELLS: GlobalShortcutAction = { kind: 'dialog', dialog: 'formatCells' }
const GO_TO: GlobalShortcutAction = { kind: 'dialog', dialog: 'goTo' }
const cmd = (command: string): GlobalShortcutAction => ({ kind: 'command', command })

describe('isModalOpen', () => {
  const stub = (hit: unknown) => ({ querySelector: () => hit }) as unknown as Document

  it('is true while a dialog backdrop is mounted', () => {
    expect(isModalOpen(stub({ className: 'dialog-backdrop' }))).toBe(true)
  })

  it('is false with no dialog and false for a null root', () => {
    expect(isModalOpen(stub(null))).toBe(false)
    expect(isModalOpen(null)).toBe(false)
  })
})

describe('resolveGlobalShortcut routing', () => {
  it('maps Excel chords onto ribbon commands with Ctrl or Cmd', () => {
    const cases: Array<[GlobalShortcutKeyEvent, GlobalShortcutGuards, GlobalShortcutAction]> = [
      [keyEvent('Digit1', ctrl), GRID, FORMAT_CELLS],
      [keyEvent('Digit1', meta), MAC, FORMAT_CELLS],
      [keyEvent('KeyG', ctrl), GRID, GO_TO],
      [keyEvent('F5'), GRID, GO_TO],
      [keyEvent('Backquote', ctrl), GRID, cmd('toggle-show-formulas')],
      [keyEvent('KeyK', meta), MAC, cmd('link-open')],
      [keyEvent('KeyE', ctrl), GRID, cmd('flash-fill')],
      [keyEvent('F2', shift), GRID, cmd('note-open')],
      [keyEvent('F3', shift), GRID, cmd('insert-function-open')],
      [keyEvent('F3', ctrl), GRID, cmd('name-manager-open')],
      [keyEvent('F11', shift), GRID, cmd('insert-sheet')],
      [keyEvent('KeyV', ctrlShift), GRID, cmd('paste-special:value')],
      [keyEvent('BracketLeft', ctrl), GRID, cmd('trace-precedents')],
      [keyEvent('BracketRight', meta), MAC, cmd('trace-dependents')],
      [keyEvent('ArrowRight', altShift), GRID, cmd('outline-group:rows')],
      [keyEvent('ArrowLeft', altShift), GRID, cmd('outline-ungroup:rows')],
      [keyEvent('Period', ctrlShift), GRID, cmd('font-size-step:1')],
      [keyEvent('Comma', metaShift), MAC, cmd('font-size-step:-1')],
      [keyEvent('KeyU', ctrlShift), GRID, cmd('formula-bar-toggle')],
      [keyEvent('Digit5', ctrl), GRID, cmd('strike')],
      [keyEvent('Equal', alt), GRID, cmd('autofn:SUM')],
      [keyEvent('Semicolon', ctrl), GRID, cmd('insert-now:date')],
      [keyEvent('Semicolon', ctrlShift), GRID, cmd('insert-now:time')],
      [keyEvent('F9'), GRID, cmd('calculate-now')],
      [keyEvent('F9', shift), GRID, cmd('calculate-sheet')],
      [keyEvent('PageDown'), GRID, cmd('page-row:1')],
      [keyEvent('PageUp'), GRID, cmd('page-row:-1')],
      [keyEvent('PageDown', alt), GRID, cmd('page-col:1')],
    ]
    for (const [event, guards, action] of cases) {
      expect(resolveGlobalShortcut(event, guards), event.code).toEqual(action)
    }
  })

  it('requires the exact modifier set', () => {
    // Ctrl+Shift+1 is Excel's number-format chord (excel-format-shortcuts): pick an unbound one
    expect(resolveGlobalShortcut(keyEvent('KeyG', ctrlShift), GRID)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('KeyV', ctrl), GRID)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('F11'), GRID)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('ArrowRight', alt), GRID)).toBeNull()
    expect(
      resolveGlobalShortcut(keyEvent('Equal', { altKey: true, metaKey: true }), MAC),
    ).toBeNull()
  })

  it('binds the Excel-for-Mac Cmd variant of the filter on mac only', () => {
    expect(resolveGlobalShortcut(keyEvent('KeyF', metaShift), MAC)).toEqual(cmd('filter-toggle'))
    expect(resolveGlobalShortcut(keyEvent('KeyF', ctrlShift), GRID)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('KeyF', ctrlShift), MAC)).toBeNull()
  })

  it('binds AutoSum cross-platform; the filter Ctrl/⌘+Shift+L chord belongs to Univer', () => {
    expect(resolveGlobalShortcut(keyEvent('KeyL', ctrlShift), GRID)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('KeyL', metaShift), MAC)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('KeyT', ctrlShift), GRID)).toEqual(cmd('autofn:SUM'))
    expect(resolveGlobalShortcut(keyEvent('KeyT', metaShift), MAC)).toEqual(cmd('autofn:SUM'))
  })

  it('leaves unrelated keys to the rest of the app', () => {
    expect(resolveGlobalShortcut(keyEvent('KeyA'), GRID)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('KeyG'), GRID)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('Enter'), GRID)).toBeNull()
  })
})

describe('guards', () => {
  const combos: Array<[string, GlobalShortcutKeyEvent]> = [
    ['Ctrl+1 (Format Cells)', keyEvent('Digit1', ctrl)],
    ['Ctrl+G (Go To)', keyEvent('KeyG', ctrl)],
    ['F5 (Go To)', keyEvent('F5')],
    ['Ctrl+K (Link)', keyEvent('KeyK', ctrl)],
    ['Ctrl+` (Show Formulas)', keyEvent('Backquote', ctrl)],
    ['F9 (recalculate)', keyEvent('F9')],
    ['Shift+F9 (calculate sheet)', keyEvent('F9', shift)],
    ['Ctrl+5 (strikethrough)', keyEvent('Digit5', ctrl)],
    ['Alt+= (AutoSum)', keyEvent('Equal', alt)],
    ['Ctrl+; (insert date)', keyEvent('Semicolon', ctrl)],
    ['Shift+F11 (insert sheet)', keyEvent('F11', shift)],
    ['PageDown', keyEvent('PageDown')],
  ]

  for (const [name, event] of combos) {
    it(`ignores ${name} while a modal is open`, () => {
      expect(resolveGlobalShortcut(event, BEHIND_MODAL)).toBeNull()
    })
  }

  it('never fires from a text field, dialog openers included', () => {
    for (const [name, event] of combos) {
      if (name.startsWith('F9') || name.startsWith('Shift+F9')) continue
      expect(resolveGlobalShortcut(event, TEXT_FIELD), name).toBeNull()
    }
  })

  it('toggles the formula bar from the formula bar itself, never from app fields', () => {
    expect(resolveGlobalShortcut(keyEvent('KeyU', ctrlShift), FORMULA_BAR)).toEqual(
      cmd('formula-bar-toggle'),
    )
    expect(resolveGlobalShortcut(keyEvent('KeyU', ctrlShift), GRID)).toEqual(
      cmd('formula-bar-toggle'),
    )
    expect(resolveGlobalShortcut(keyEvent('KeyU', ctrlShift), TEXT_FIELD)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('Digit1', ctrl), FORMULA_BAR)).toBeNull()
  })

  it('keeps range-writing commands off while a cell is being edited', () => {
    expect(resolveGlobalShortcut(keyEvent('Digit5', ctrl), EDITING)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('KeyE', ctrl), EDITING)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('F9'), EDITING)).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('Digit1', ctrl), EDITING)).toEqual(FORMAT_CELLS)
    expect(resolveGlobalShortcut(keyEvent('KeyG', ctrl), EDITING)).toEqual(GO_TO)
  })

  it('bails once something already handled the key', () => {
    expect(
      resolveGlobalShortcut(keyEvent('Equal', { ...alt, defaultPrevented: true }), GRID),
    ).toBeNull()
    expect(resolveGlobalShortcut(keyEvent('PageDown', { defaultPrevented: true }), GRID)).toBeNull()
  })
})

describe('number-format and border shortcuts', () => {
  const percent = keyEvent('Digit5', ctrlShift)
  const outline = keyEvent('Digit7', metaShift)

  it('route through the ribbon format and border commands from the grid', () => {
    expect(resolveGlobalShortcut(percent, GRID)).toEqual({ kind: 'command', command: 'format:0%' })
    expect(resolveGlobalShortcut(outline, GRID)).toEqual({
      kind: 'command',
      command: 'border:outer',
    })
  })

  it('stay out of text fields, cell editing and modals', () => {
    expect(resolveGlobalShortcut(percent, { ...GRID, gridTarget: false })).toBeNull()
    expect(resolveGlobalShortcut(percent, { ...GRID, cellEditing: true })).toBeNull()
    expect(resolveGlobalShortcut(outline, { ...GRID, modalOpen: true })).toBeNull()
  })

  it('keep the mac-only border chords off Windows and Linux', () => {
    const macBorder = keyEvent('Digit0', { metaKey: true, altKey: true })
    expect(resolveGlobalShortcut(macBorder, MAC)).toEqual({
      kind: 'command',
      command: 'border:outer',
    })
    expect(resolveGlobalShortcut(macBorder, GRID)).toBeNull()
  })
})
