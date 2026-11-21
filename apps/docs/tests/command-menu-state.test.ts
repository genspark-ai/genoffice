/**
 * The renderer reports which commands are checked or unavailable, and the
 * main process applies that to the menu bar. These pin the report's shape and
 * that nothing but real command ids survives the IPC boundary.
 */
import { describe, expect, it } from 'vitest'
import {
  COMMANDS,
  toCommandMenuState,
  type CommandDef,
  type CommandId,
} from '../src/shared/commands'
import { menuCommandState, type MenuStateInput } from '../src/renderer/commands/menu-state'

const BODY: MenuStateInput = {
  aiSidebar: true,
  darkPage: false,
  hasDoc: true,
  editable: true,
  inTable: false,
  inTextBox: false,
  spellcheck: true,
  readOnlyRecommended: false,
  onDisk: true,
  view: {
    mode: 'print',
    readMode: false,
    ruler: false,
    grid: false,
    navigation: false,
    thumbnails: false,
    markup: false,
    ribbon: true,
    focus: false,
    reviewPane: false,
  },
  cellsSelected: false,
  hasSelection: true,
  pictureSelected: false,
  hasNotes: true,
  hasChart: true,
  canSplitTable: true,
  canComment: true,
  inEmptyParagraph: false,
  hasComments: true,
  hasTableOfFigures: true,
  hasIndex: true,
  hasTableOfAuthorities: true,
  hasMergeSource: true,
  windowsDetached: true,
  labelSheet: true,
  canUndo: true,
  canRedo: true,
  canRepeat: false,
  target: { hasSelection: true, editable: true, field: false },
}
const state = (over: Partial<MenuStateInput>) => menuCommandState({ ...BODY, ...over })
/** commands that only hold a place in the menu are always greyed out: left aside here */
const PLANNED = new Set(
  Object.entries<CommandDef>(COMMANDS)
    .filter(([, def]) => def.planned)
    .map(([id]) => id),
)
const disabled = (over: Partial<MenuStateInput>) =>
  new Set(state(over).disabled.filter((id) => !PLANNED.has(id)))

describe('menu command state', () => {
  it('checks the AI sidebar and dark page toggles from app state', () => {
    const toggles = (over: Partial<MenuStateInput>) =>
      state(over).checked.filter((id) => id.startsWith('toggle-a') || id === 'toggle-dark')
    expect(toggles({})).toEqual(['toggle-ai'])
    expect(toggles({ aiSidebar: false, darkPage: true })).toEqual(['toggle-dark'])
    expect(toggles({ aiSidebar: false })).toEqual([])
  })

  it('checks the current layout and the panes that are showing in the View menu', () => {
    expect(state({}).checked).toContain('view-print-layout')
    const outline = state({
      view: { ...BODY.view, mode: 'outline', ruler: true, navigation: true },
    })
    expect(outline.checked).toEqual(
      expect.arrayContaining(['view-outline', 'toggle-ruler', 'toggle-navigation']),
    )
    expect(outline.checked).not.toContain('view-print-layout')
    expect(outline.checked).not.toContain('view-thumbnails')
    const draft = state({ view: { ...BODY.view, mode: 'draft' } })
    expect(draft.checked).toContain('view-draft')
    expect(draft.checked).not.toContain('view-print-layout')
    expect(draft.disabled).not.toContain('view-draft')
    const thumbs = state({ view: { ...BODY.view, navigation: true, thumbnails: true } })
    expect(thumbs.checked).not.toContain('view-markup')
    expect(state({ view: { ...BODY.view, markup: true } }).checked).toContain('view-markup')
    expect(thumbs.checked).toEqual(expect.arrayContaining(['toggle-navigation', 'view-thumbnails']))
    expect(outline.checked).not.toContain('toggle-grid')
    // View > Ribbon is ticked while the band shows; Always Open Read-Only follows the document
    expect(state({}).checked).toContain('view-ribbon')
    expect(state({ view: { ...BODY.view, ribbon: false } }).checked).not.toContain('view-ribbon')
    expect(state({}).checked).not.toContain('file-always-read-only')
    expect(state({ readOnlyRecommended: true }).checked).toContain('file-always-read-only')
    expect(state({ view: { ...BODY.view, focus: true } }).checked).toContain('view-focus')
    // Read Mode is a view of its own: no layout is ticked while it is on
    const reading = state({ view: { ...BODY.view, readMode: true } })
    expect(reading.checked).not.toContain('view-print-layout')
  })

  it('Hide Spelling Errors is ticked while spell checking is off', () => {
    expect(state({}).checked).not.toContain('tools-hide-spelling')
    expect(state({ spellcheck: false }).checked).toContain('tools-hide-spelling')
  })

  it('a command that only holds its place in the menu is always greyed out', () => {
    expect(PLANNED.size).toBeGreaterThan(0)
    for (const over of [{}, { hasDoc: false }, { inTable: true }] as Partial<MenuStateInput>[])
      for (const id of PLANNED) expect(state(over).disabled, id).toContain(id)
  })

  it('with the caret in body text, only the table commands are greyed out', () => {
    const off = disabled({})
    expect([...off].every((id) => id.startsWith('table-'))).toBe(true)
    for (const id of [
      'table-insert-rows-below',
      'table-select-row',
      'table-merge-cells',
      'table-properties',
    ] as const)
      expect(off.has(id)).toBe(true)
    // Insert Table, View Gridlines and Formula stay available outside a table, as in Word
    expect(off.has('insert-table')).toBe(false)
    expect(off.has('table-gridlines')).toBe(false)
    expect(off.has('table-formula')).toBe(false)
  })

  it('Formula in body text needs a caret in text: not a text box, not a selected picture', () => {
    expect(disabled({ inTextBox: true }).has('table-formula')).toBe(true)
    expect(disabled({ pictureSelected: true }).has('table-formula')).toBe(true)
    expect(disabled({ inTable: true, pictureSelected: true }).has('table-formula')).toBe(false)
    expect(disabled({ editable: false }).has('table-formula')).toBe(true)
  })

  it('at a bare caret the commands that act on a selection are greyed out', () => {
    const needs = ['format-fit-text', 'insert-autotext-new'] as const
    const caret = disabled({ hasSelection: false })
    for (const id of needs) {
      expect(caret.has(id), id).toBe(true)
      expect(disabled({}).has(id), id).toBe(false)
    }
    expect([...caret].filter((id) => !id.startsWith('table-')).sort()).toEqual([...needs])
    // Convert Text to Table takes the caret's paragraph when nothing is selected
    expect(caret.has('convert-text-to-table')).toBe(false)
  })

  it('with a picture selected the commands that work on text or put some at the caret are greyed out', () => {
    const off = disabled({ pictureSelected: true })
    expect([...off].filter((id) => !id.startsWith('table-')).sort()).toEqual([
      'format-combine-characters',
      'format-drop-cap',
      'format-enclose-characters',
      'format-fit-text',
      'format-horizontal-in-vertical',
      'format-phonetic-guide',
      'format-two-lines-in-one',
      'insert-advanced-symbol',
      'insert-date-time',
      'insert-field-dialog',
      'tools-thesaurus',
      'tools-translate-selection',
    ])
    // what is selected can still be kept as AutoText
    expect(off.has('insert-autotext-new')).toBe(false)
  })

  it('Footnotes waits for a note in the document, Select Data for a chart', () => {
    expect(disabled({}).has('view-footnotes')).toBe(false)
    expect(disabled({ hasNotes: false }).has('view-footnotes')).toBe(true)
    expect(disabled({}).has('edit-select-data')).toBe(false)
    expect(disabled({ hasChart: false }).has('edit-select-data')).toBe(true)
  })

  it('Split Table follows what the command can do where the caret is', () => {
    expect(disabled({ inTable: true }).has('table-split-table')).toBe(false)
    expect(disabled({ inTable: true, canSplitTable: false }).has('table-split-table')).toBe(true)
  })

  it('with the caret in a table, Merge Cells waits for a cell selection and there is no text to convert or drop', () => {
    expect([...disabled({ inTable: true })].sort()).toEqual([
      'convert-text-to-table',
      'format-drop-cap',
      'table-merge-cells',
    ])
  })

  it('with several cells selected, Merge Cells is on and what goes in at a caret or works on a run of text is off', () => {
    expect([...disabled({ inTable: true, cellsSelected: true })].sort()).toEqual([
      'convert-text-to-table',
      'format-combine-characters',
      'format-drop-cap',
      'format-enclose-characters',
      'format-horizontal-in-vertical',
      'format-phonetic-guide',
      'format-two-lines-in-one',
      'insert-advanced-symbol',
      'insert-date-time',
      'insert-field-dialog',
      'insert-link',
      'insert-table',
      'tools-thesaurus',
      'tools-translate-selection',
    ])
    // Fit Text and a new AutoText entry take the cells' text, as in Word
    const cells = disabled({ inTable: true, cellsSelected: true })
    expect(cells.has('format-fit-text')).toBe(false)
    expect(cells.has('insert-autotext-new')).toBe(false)
  })

  it('Rename and Move wait for the document to have a file, Text Direction for a table', () => {
    expect(disabled({}).has('file-rename')).toBe(false)
    const unsaved = disabled({ onDisk: false })
    expect(unsaved.has('file-rename')).toBe(true)
    expect(unsaved.has('file-move')).toBe(true)
    expect(unsaved.has('file-properties')).toBe(false)
    // Text Direction turns the section in body text, the cells in a table
    expect(disabled({}).has('format-text-direction')).toBe(false)
    expect(disabled({ inTable: true }).has('format-text-direction')).toBe(false)
    expect(disabled({ editable: false }).has('format-text-direction')).toBe(true)
    expect(disabled({ hasDoc: false }).has('format-text-direction')).toBe(true)
  })

  it('with no document, only what works without one stays on', () => {
    const off = disabled({ hasDoc: false, editable: false })
    const on = Object.keys(COMMANDS).filter((id) => !off.has(id as CommandId) && !PLANNED.has(id))
    expect(on.sort()).toEqual(
      [
        'autocorrect-options',
        'file-new-from-template',
        'ink-tool',
        'new',
        'new-window',
        'open',
        'open-path',
        'preferences',
        'shortcuts',
        'table-gridlines',
        'toggle-ai',
        'toggle-dark',
        'tools-customize-keyboard',
        'view-ribbon',
      ].sort(),
    )
  })

  it('a read-only document keeps reading commands and table selection, not edits', () => {
    const off = disabled({ editable: false, inTable: true })
    for (const id of [
      'bold',
      'font-dialog',
      'paragraph-dialog',
      'page-setup',
      'format-document',
      'format-equation-options',
      'view-header-footer',
      'insert-comment',
      'ai-proofread',
      'align-center',
      'insert-image',
      'insert-page-break',
      'table-delete-rows',
      'table-properties',
      'table-formula',
      'tools-translate-selection',
    ] as const)
      expect(off.has(id)).toBe(true)
    for (const id of [
      'save-as',
      'print',
      'find',
      // Replace opens the Find panel, which offers no replacing in a read-only document
      'replace',
      // Thesaurus only reads the selection; Translate Selection would rewrite it
      'tools-thesaurus',
      'word-count',
      'zoom-in',
      'table-select-row',
      'export-pdf',
    ] as const)
      expect(off.has(id)).toBe(false)
  })

  it('Undo and Redo wait for something to undo or redo', () => {
    const off = disabled({ inTable: true, canUndo: false, canRedo: true })
    expect(off.has('undo')).toBe(true)
    expect(off.has('redo')).toBe(false)
    expect(disabled({ canUndo: true, canRedo: false }).has('redo')).toBe(true)
  })

  it('Repeat is Redo while there is something to redo, else the last command, else greyed out', () => {
    expect(state({ canRedo: true, canRepeat: true }).repeat).toBe('redo')
    expect(state({ canRedo: false, canRepeat: true }).repeat).toBe('repeat')
    const nothing = state({ canRedo: false, canRepeat: false })
    expect(nothing.repeat).toBeUndefined()
    expect(nothing.disabled).toContain('repeat')
    const inField = state({
      canRepeat: true,
      target: { hasSelection: false, editable: true, field: true },
    })
    expect(inField.repeat).toBeUndefined()
    expect(inField.disabled).toContain('repeat')
    expect(toCommandMenuState({ checked: [], disabled: [], repeat: 'repeat' }).repeat).toBe(
      'repeat',
    )
    expect(toCommandMenuState({ checked: [], disabled: [], repeat: 'x' })).toEqual({
      checked: [],
      disabled: [],
    })
  })

  it('Cut and Copy wait for a selection; Cut and Paste for an editable target', () => {
    const caret = disabled({ target: { hasSelection: false, editable: true, field: false } })
    expect(['cut', 'copy'].every((id) => caret.has(id as CommandId))).toBe(true)
    expect(caret.has('paste')).toBe(false)

    const readOnly = disabled({
      editable: false,
      target: { hasSelection: true, editable: false, field: false },
    })
    expect(readOnly.has('copy')).toBe(false)
    for (const id of ['cut', 'paste', 'paste-match-style'] as const)
      expect(readOnly.has(id)).toBe(true)
    expect(readOnly.has('select-all')).toBe(false)
  })

  it('a focused text field decides the clipboard for itself, whatever the document allows', () => {
    const inField = disabled({
      hasDoc: false,
      editable: false,
      target: { hasSelection: true, editable: true, field: true },
    })
    for (const id of ['cut', 'copy', 'paste', 'select-all'] as const)
      expect(inField.has(id)).toBe(false)
    // Paste Special opens on the document: a field, or no document, leaves it nothing to do
    expect(inField.has('edit-paste-special')).toBe(true)
    const fieldOverDoc = disabled({ target: { hasSelection: false, editable: true, field: true } })
    expect(fieldOverDoc.has('paste')).toBe(false)
    expect(fieldOverDoc.has('edit-paste-special')).toBe(true)
    const inDoc = disabled({ target: { hasSelection: false, editable: true, field: false } })
    expect(inDoc.has('edit-paste-special')).toBe(false)

    const emptyField = disabled({ target: { hasSelection: false, editable: true, field: true } })
    expect(emptyField.has('copy')).toBe(true)
    expect(emptyField.has('paste')).toBe(false)
  })

  it('Undo and Redo step aside for a focused text field, so its own undo gets the key', () => {
    const inField = disabled({ target: { hasSelection: false, editable: true, field: true } })
    expect(inField.has('undo')).toBe(true)
    expect(inField.has('redo')).toBe(true)
  })

  it('keeps only real command ids from an IPC payload', () => {
    expect(
      toCommandMenuState({
        checked: ['toggle-ai', 'no-such-command', 7, null, 'toString'],
        disabled: ['save', '__proto__'],
      }),
    ).toEqual({ checked: ['toggle-ai'], disabled: ['save'] })
  })

  it('reads a malformed payload as nothing checked or disabled', () => {
    for (const value of [null, undefined, 'x', 3, {}, { checked: 'toggle-ai', disabled: {} }]) {
      expect(toCommandMenuState(value)).toEqual({ checked: [], disabled: [] })
    }
  })

  it('Update Labels waits for a document that is a sheet of labels', () => {
    expect(disabled({ labelSheet: false }).has('mailings-update-labels')).toBe(true)
    expect(disabled({}).has('mailings-update-labels')).toBe(false)
  })

  it('Insert > Comment is greyed where a comment has nothing to hold on to; Format > Drop Cap in an empty paragraph', () => {
    expect(disabled({}).has('insert-comment')).toBe(false)
    expect(disabled({ canComment: false }).has('insert-comment')).toBe(true)
    expect(disabled({}).has('format-drop-cap')).toBe(false)
    expect(disabled({ inEmptyParagraph: true }).has('format-drop-cap')).toBe(true)
  })

  it('the commands of a recipient list, of a table to update and of a window to merge wait for one', () => {
    const waits = {
      hasMergeSource: ['mailings-edit-recipients', 'mailings-rules', 'mailings-finish-merge'],
      hasIndex: ['update-index'],
      hasTableOfFigures: ['update-table-of-figures'],
      hasTableOfAuthorities: ['toa-update'],
      hasComments: ['show-comments'],
      windowsDetached: ['view-merge-windows'],
    } as const
    for (const [input, ids] of Object.entries(waits))
      for (const id of ids) {
        expect(disabled({ [input]: false }).has(id), id).toBe(true)
        expect(disabled({}).has(id), id).toBe(false)
      }
  })
})
