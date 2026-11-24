/**
 * The shortcut registry is what the Keyboard Shortcuts sheet
 * renders, so it has to stay honest: no chord listed twice on either platform,
 * every row translatable, and no menu accelerator missing from the list.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  SHORTCUTS,
  SHORTCUT_GROUPS,
  shortcutKeys,
  type ShortcutDef,
} from '../src/renderer/shortcuts'
import { strings } from '../src/renderer/i18n/strings'
import { COMMANDS, type CommandDef } from '../src/shared/commands'
import { toChord } from '../src/shared/keybindings'
import {
  assignChord,
  buildKeyTable,
  chordFromEvent,
  chordOwner,
  formatChord,
  defaultChords,
  keyVerdict,
  shortcutChords,
  documentKeyCommand,
  windowKeyCommand,
} from '../src/renderer/keybindings'

// alternates are listed as "A / B"; the separator has to keep ⌘/ itself intact
const chords = (def: ShortcutDef, isMac: boolean) =>
  shortcutKeys(def, isMac).split(' / ').filter(Boolean)

describe('shortcut registry', () => {
  it('has unique ids and known groups', () => {
    const ids = SHORTCUTS.map((def) => def.id)
    expect(new Set(ids).size).toBe(ids.length)
    const groups = new Set(SHORTCUT_GROUPS.map((group) => group.id))
    for (const def of SHORTCUTS) expect(groups.has(def.group)).toBe(true)
  })

  it.each([
    ['macOS', true],
    ['Windows/Linux', false],
  ])('binds every chord once on %s', (_platform, isMac) => {
    const seen = new Map<string, string>()
    for (const def of SHORTCUTS) {
      for (const chord of chords(def, isMac as boolean)) {
        expect(seen.get(chord), `${chord} is claimed by ${seen.get(chord)} and ${def.id}`).toBe(
          undefined,
        )
        seen.set(chord, def.id)
      }
    }
  })

  it('labels every row in every language', () => {
    for (const [lang, dict] of Object.entries(strings)) {
      for (const entry of [...SHORTCUTS, ...SHORTCUT_GROUPS]) {
        const value = (dict as Record<string, string>)[entry.labelKey]
        expect(value, `${entry.labelKey} missing in ${lang}`).toBeTruthy()
      }
    }
  })

  it('rewrites Mac notation for Windows/Linux', () => {
    const byId = new Map(SHORTCUTS.map((def) => [def.id, def]))
    const win = (id: string) => shortcutKeys(byId.get(id)!, false)
    expect(win('save-as')).toBe('Ctrl+Shift+S')
    expect(win('page-break')).toBe('Ctrl+Enter')
    expect(win('nbsp')).toBe('Ctrl+Shift+Space')
    expect(win('style-h1')).toBe('Ctrl+Alt+1')
    expect(win('proofread')).toBe('F7')
    // platform-specific keys come from the explicit override, not a rewrite
    expect(win('endnote')).toBe('Ctrl+Alt+D')
    expect(shortcutKeys(byId.get('endnote')!, true)).toBe('⌥⌘E')
  })

  it('lists every menu accelerator', () => {
    // most live in the command table; the menu builder still spells a few itself
    const source = readFileSync(join(__dirname, '../src/main/docs-main.ts'), 'utf8')
    const accelerators = [
      ...Object.values<CommandDef>(COMMANDS).flatMap((def) =>
        typeof def.accelerator === 'string' ? [def.accelerator] : [],
      ),
      ...[...source.matchAll(/accelerator: '([^']+)'/g)].map((m) => m[1]),
    ]
    expect(accelerators.length).toBeGreaterThan(15)
    const listed = new Set(SHORTCUTS.flatMap((def) => chords(def, false)))
    for (const accelerator of accelerators) {
      // Electron writes modifiers in any order; the sheet renders Ctrl+Alt+Shift+key
      const parts = accelerator.split('+')
      const key = parts.pop()!
      const mods = new Set(parts)
      const normalized = [
        mods.has('CmdOrCtrl') || mods.has('Ctrl') ? 'Ctrl' : null,
        mods.has('Alt') ? 'Alt' : null,
        mods.has('Shift') ? 'Shift' : null,
        key,
      ]
        .filter(Boolean)
        .join('+')
      expect(listed.has(normalized), `${accelerator} is not in the shortcut sheet`).toBe(true)
    }
  })

  it.each([
    ['macOS', true],
    ['Windows/Linux', false],
  ])('lists every per-platform menu accelerator on %s', (_platform, isMac) => {
    const listed = new Map(
      SHORTCUTS.flatMap((def) =>
        def.cmd ? shortcutChords(def, isMac).map((chord) => [chord, def.cmd] as const) : [],
      ),
    )
    let checked = 0
    for (const [id, def] of Object.entries<CommandDef>(COMMANDS)) {
      if (typeof def.accelerator !== 'object') continue
      const accelerator = def.accelerator[isMac ? 'mac' : 'other']
      if (!accelerator) continue
      checked++
      expect(listed.get(toChord(accelerator, isMac)!), `${accelerator} of ${id}`).toBe(id)
    }
    expect(checked).toBe(isMac ? 14 : 1)
  })

  it('gives Shift+Cmd+G to Find Previous on the Mac and keeps Word Count on Ctrl+Shift+G elsewhere', () => {
    expect(chordOwner({}, 'Shift+Cmd+G', true)).toEqual({
      kind: 'binding',
      id: 'edit-find-previous',
    })
    expect(defaultChords('word-count', true)).toEqual([])
    expect(chordOwner({}, 'Ctrl+Shift+G', false)).toEqual({ kind: 'binding', id: 'word-count' })
    expect(defaultChords('edit-find-previous', false)).toEqual([])
  })

  it('runs the Word for Mac menu chords from the window on the Mac only', () => {
    expect(windowKeyCommand({}, 'Cmd+G', true)).toBe('edit-find-next')
    expect(windowKeyCommand({}, 'Ctrl+Cmd+X', true)).toBe('table-delete-rows')
    expect(windowKeyCommand({}, 'Ctrl+=', true)).toBe('insert-equation')
    expect(windowKeyCommand({}, 'Cmd+F', true)).toBe(null)
    expect(windowKeyCommand({}, 'Ctrl+G', false)).toBe(null)
    expect(windowKeyCommand({}, 'Ctrl+Shift+G', false)).toBe(null)
    // a customized command is run by the key layer, with the chords the user gave it
    const moved = assignChord({}, 'word-count', 'Shift+Cmd+G', true)!
    expect(moved).toEqual({ 'word-count': ['Shift+Cmd+G'], 'edit-find-previous': [] })
    expect(windowKeyCommand(moved, 'Shift+Cmd+G', true)).toBe(null)
    expect(windowKeyCommand(moved, 'Cmd+G', true)).toBe('edit-find-next')
    expect(keyVerdict(buildKeyTable(moved, true), 'Shift+Cmd+G', true)).toEqual({
      run: { cmd: 'word-count' },
    })
  })

  it('keeps the chords of the document key table among the Mac chords of their row', () => {
    let rows = 0
    for (const def of SHORTCUTS) {
      if (!def.docKeys) continue
      rows++
      expect(def.cmd, def.id).toBeTruthy()
      const all = chords(def, true)
      for (const chord of def.docKeys.split(' / '))
        expect(all, `${def.id} ${chord}`).toContain(chord)
    }
    expect(rows).toBeGreaterThan(50)
  })

  it("gives the chords Word for Mac uses differently Word's meaning", () => {
    const run = (chord: string) => documentKeyCommand({}, chord, true)
    const owner = (chord: string) => chordOwner({}, chord, true)
    // indent / outdent
    expect(owner('Ctrl+Shift+M')).toEqual({ kind: 'binding', id: 'indent-increase' })
    expect(owner('Shift+Cmd+M')).toEqual({ kind: 'binding', id: 'indent-decrease' })
    expect(owner('Ctrl+M')).toEqual({ kind: 'binding', id: 'indent-increase' })
    // Option+Shift+arrows extend by paragraph; moving a paragraph is on Control+Shift
    expect(run('Alt+Shift+Up')).toEqual({ cmd: 'caret-move', payload: 'para-up-extend' })
    expect(run('Alt+Shift+Down')).toEqual({ cmd: 'caret-move', payload: 'para-down-extend' })
    expect(owner('Ctrl+Shift+Up')).toMatchObject({ kind: 'fixed', row: { id: 'move-up' } })
    expect(owner('Ctrl+Shift+Down')).toMatchObject({ kind: 'fixed', row: { id: 'move-down' } })
    // Preferences, Subscript, the non-breaking hyphen
    expect(owner('Cmd+,')).toEqual({ kind: 'fixed', row: null })
    expect(owner('Shift+Cmd+-')).toEqual({ kind: 'binding', id: 'subscript' })
    expect(owner('Alt+Shift+Cmd+-')).toMatchObject({ kind: 'fixed', row: { id: 'nb-hyphen' } })
    expect(owner('Cmd+.')).toBe(null)
    // Normal style, with no chord left for New Window
    expect(owner('Shift+Cmd+N')).toEqual({ kind: 'binding', id: 'style-normal' })
    // Paste Text Only, Paste Formatting, Replace, Spelling and Grammar
    expect(run('Shift+Cmd+V')).toEqual({ cmd: 'paste-as', payload: 'text' })
    expect(owner('Alt+Cmd+V')).toMatchObject({ kind: 'fixed', row: { id: 'paste-format' } })
    expect(windowKeyCommand({}, 'Shift+Cmd+H', true)).toBe('replace')
    expect(windowKeyCommand({}, 'Ctrl+H', true)).toBe('replace')
    expect(run('F7')).toEqual({ cmd: 'tools-spelling' })
    expect(defaultChords('ai-proofread', true)).toEqual([])
  })

  it('leaves Windows and Linux with the chords they had', () => {
    const owner = (chord: string) => chordOwner({}, chord, false)
    expect(owner('Ctrl+Shift+M')).toEqual({ kind: 'binding', id: 'indent-decrease' })
    expect(owner('Ctrl+,')).toEqual({ kind: 'binding', id: 'subscript' })
    expect(owner('Ctrl+.')).toEqual({ kind: 'binding', id: 'superscript' })
    expect(owner('Ctrl+Shift+-')).toMatchObject({ kind: 'fixed', row: { id: 'nb-hyphen' } })
    expect(owner('Alt+Shift+Up')).toMatchObject({ kind: 'fixed', row: { id: 'move-up' } })
    expect(owner('Ctrl+Shift+N')).toMatchObject({ kind: 'fixed', row: { id: 'new-window' } })
    expect(owner('F7')).toEqual({ kind: 'binding', id: 'ai-proofread' })
    expect(owner('Ctrl+Alt+Shift+V')).toMatchObject({ kind: 'fixed', row: { id: 'paste-format' } })
    for (const chord of ['Ctrl+Shift+M', 'Home', 'Ctrl+A', 'F7', 'Ctrl+Shift+D'])
      expect(documentKeyCommand({}, chord, false)).toBe(null)
  })

  it('runs the chords Word for Mac adds from the document key table', () => {
    const run = (chord: string) => documentKeyCommand({}, chord, true)
    const expected: [string, string, string?][] = [
      ['Home', 'caret-move', 'line-start'],
      ['Ctrl+Down', 'caret-move', 'line-end'],
      ['Cmd+Up', 'caret-move', 'para-up'],
      ['Alt+Down', 'caret-move', 'para-down'],
      ['Cmd+Home', 'caret-move', 'doc-start'],
      ['Alt+End', 'caret-move', 'doc-end'],
      ['Shift+End', 'caret-move', 'line-end-extend'],
      ['Shift+Cmd+Home', 'caret-move', 'doc-start-extend'],
      ['Shift+Cmd+Down', 'caret-move', 'para-down-extend'],
      ['Cmd+Backspace', 'delete-word', 'back'],
      ['Cmd+Delete', 'delete-word', 'forward'],
      ['Ctrl+A', 'edit-text', 'select-all'],
      ['F2', 'edit-text', 'cut'],
      ['Ctrl+C', 'edit-text', 'copy'],
      ['F4', 'edit-text', 'paste'],
      ['Ctrl+Z', 'undo'],
      ['Alt+Enter', 'repeat'],
      ['Ctrl+N', 'new'],
      ['Ctrl+P', 'print'],
      ['F5', 'goto'],
      ['Ctrl+Shift+B', 'bold'],
      ['Shift+Cmd+I', 'italic'],
      ['Ctrl+U', 'underline'],
      ['Shift+Cmd+D', 'double-underline'],
      ['Shift+Cmd+W', 'word-underline'],
      ['Shift+Cmd+X', 'strikethrough'],
      ['Shift+Cmd+A', 'all-caps'],
      ['Shift+Cmd+K', 'small-caps'],
      ['Ctrl+Shift+H', 'hidden-text'],
      ['Shift+Cmd+L', 'bullets', 'on'],
      ['Ctrl+Shift+D', 'insert-field', 'DATE'],
      ['Ctrl+Shift+T', 'insert-field', 'TIME'],
      ['Ctrl+Shift+P', 'insert-field', 'PAGE'],
      ['Shift+Cmd+F5', 'insert-bookmark'],
      ['Alt+F3', 'insert-autotext-new'],
      ['Alt+Shift+Cmd+X', 'mark-index-entry'],
      ['Alt+Shift+Cmd+I', 'toa-mark-citation'],
      ['Ctrl+Cmd+I', 'table-insert-rows-above'],
      ['Ctrl+Alt+Cmd+T', 'tools-translate-selection'],
      ['Ctrl+Shift+E', 'mailings-edit-recipients'],
      ['Ctrl+Shift+N', 'mailings-finish-merge'],
      ['Alt+Cmd+P', 'view-print-layout'],
      ['Alt+Cmd+O', 'view-outline'],
      ['Alt+Cmd+N', 'view-draft'],
      ['Alt+Cmd+S', 'split-view'],
      ['Alt+Shift+Cmd+S', 'styles-pane'],
    ]
    for (const [chord, cmd, payload] of expected)
      expect(run(chord), chord).toEqual(payload ? { cmd, payload } : { cmd })
    // chords with handlers of their own stay with them
    for (const chord of ['Cmd+B', 'Cmd+Left', 'Alt+Backspace', 'Cmd+Z', 'Cmd+A', 'Cmd+,'])
      expect(run(chord), chord).toBe(null)
    expect(windowKeyCommand({}, 'Shift+F4', true)).toBe('edit-find-next')
    expect(windowKeyCommand({}, 'Shift+F7', true)).toBe('tools-thesaurus')
  })

  it('runs the editing commands behind Word for Mac chords from the document key table', () => {
    const expected: [string, string, string?][] = [
      ['F8', 'extend-selection', 'grow'],
      ['Shift+F8', 'extend-selection', 'shrink'],
      ['Shift+F5', 'go-back'],
      ['Alt+Cmd+Z', 'go-back'],
      ['Ctrl+Home', 'table-cell-go', 'row-start'],
      ['Alt+Shift+Home', 'table-cell-go', 'row-start'],
      ['Ctrl+End', 'table-cell-go', 'row-end'],
      ['Alt+Shift+End', 'table-cell-go', 'row-end'],
      ['Ctrl+PageUp', 'table-cell-go', 'column-start'],
      ['Alt+Shift+PageUp', 'table-cell-go', 'column-start'],
      ['Ctrl+PageDown', 'table-cell-go', 'column-end'],
      ['Alt+Shift+PageDown', 'table-cell-go', 'column-end'],
      ['Alt+Clear', 'table-select-table'],
      ['Cmd+F3', 'spike', 'cut'],
      ['Shift+Cmd+F3', 'spike', 'insert'],
      ['Ctrl+Shift+Left', 'outline-level', 'promote'],
      ['Ctrl+Shift+Right', 'outline-level', 'demote'],
      ['Ctrl+Shift+=', 'outline-fold', 'expand'],
      ['Ctrl+Shift+-', 'outline-fold', 'collapse'],
      ['Cmd+F11', 'field-lock', 'lock'],
      ['Cmd+3', 'field-lock', 'lock'],
      ['Shift+Cmd+F11', 'field-lock', 'unlock'],
      ['Cmd+4', 'field-lock', 'unlock'],
      ['Shift+Cmd+F9', 'field-unlink'],
      ['Cmd+6', 'field-unlink'],
      ['F11', 'field-go', 'next'],
      ['Shift+F11', 'field-go', 'prev'],
    ]
    for (const [chord, cmd, payload] of expected) {
      expect(documentKeyCommand({}, chord, true), chord).toEqual(
        payload ? { cmd, payload } : { cmd },
      )
      expect(documentKeyCommand({}, chord, false), chord).toBe(null)
    }
    // Lock Fields given another key: the chords it had are swallowed, the new one runs it
    const moved = assignChord({}, 'field-lock:lock', 'Ctrl+Cmd+L', true)!
    expect(documentKeyCommand(moved, 'Cmd+3', true)).toBe(null)
    const table = buildKeyTable(moved, true)
    expect(keyVerdict(table, 'Ctrl+Cmd+L', true)).toEqual({
      run: { cmd: 'field-lock', payload: 'lock' },
    })
    expect(keyVerdict(table, 'Cmd+4', true)).toBe('pass')
    // selecting to the end of a row has no chord of Word's and can be given one
    const select = SHORTCUTS.find((def) => def.id === 'select-row-end')!
    expect([shortcutKeys(select, true), shortcutKeys(select, false)]).toEqual(['', ''])
    const given = assignChord({}, 'table-cell-go:row-end-extend', 'Ctrl+Shift+End', true)!
    expect(keyVerdict(buildKeyTable(given, true), 'Ctrl+Shift+End', true)).toEqual({
      run: { cmd: 'table-cell-go', payload: 'row-end-extend' },
    })
  })

  it('runs the window, browse, Outline level and keypad chords of Word for Mac from the document key table', () => {
    const expected: [string, string, string?][] = [
      ['PageUp', 'caret-move', 'screen-up'],
      ['PageDown', 'caret-move', 'screen-down'],
      ['Shift+PageUp', 'caret-move', 'screen-up-extend'],
      ['Shift+PageDown', 'caret-move', 'screen-down-extend'],
      ['Alt+Cmd+PageUp', 'caret-move', 'window-start'],
      ['Alt+Cmd+PageDown', 'caret-move', 'window-end'],
      ['Alt+Shift+Cmd+PageUp', 'caret-move', 'window-start-extend'],
      ['Alt+Shift+Cmd+PageDown', 'caret-move', 'window-end-extend'],
      ['Cmd+PageUp', 'browse', 'prev'],
      ['Cmd+PageDown', 'browse', 'next'],
      ['Alt+PageUp', 'section-go', 'prev'],
      ['Alt+PageDown', 'section-go', 'next'],
      ['Cmd+F6', 'window-go', 'next'],
      ['Alt+F6', 'window-go', 'next'],
      ['Shift+Cmd+F6', 'window-go', 'prev'],
      ['Alt+Shift+F6', 'window-go', 'prev'],
      ['Ctrl+Shift+C', 'close-pane'],
      ['Ctrl+Shift+A', 'outline-show', 'all'],
      ...Array.from({ length: 9 }, (_, i): [string, string, string] => [
        `Ctrl+Shift+${i + 1}`,
        'outline-show',
        String(i + 1),
      ]),
      ['Ctrl+Shift+NumPlus', 'outline-fold', 'expand'],
      ['Ctrl+Shift+NumMinus', 'outline-fold', 'collapse'],
      ['Cmd+F9', 'field-insert-chars'],
      ['Cmd+NumEnter', 'insert-section-break', 'nextPage'],
      ['Cmd+NumMinus', 'insert-symbol', '\u2013'],
      ['Alt+Cmd+NumMinus', 'insert-symbol', '\u2014'],
      ['Clear', 'edit-clear-contents'],
    ]
    for (const [chord, cmd, payload] of expected) {
      expect(documentKeyCommand({}, chord, true), chord).toEqual(
        payload ? { cmd, payload } : { cmd },
      )
      expect(documentKeyCommand({}, chord, false), chord).toBe(null)
    }
    // the main block's Enter keeps the page break
    expect(documentKeyCommand({}, 'Cmd+Enter', true)).toBe(null)
    // Next Window given another key: its chords are swallowed and the new one runs it
    const moved = assignChord({}, 'window-go:next', 'Ctrl+Cmd+F6', true)!
    expect(documentKeyCommand(moved, 'Cmd+F6', true)).toBe(null)
    expect(keyVerdict(buildKeyTable(moved, true), 'Ctrl+Cmd+F6', true)).toEqual({
      run: { cmd: 'window-go', payload: 'next' },
    })
  })

  it('names the keypad keys Word for Mac binds apart from the main block', () => {
    const press = (code: string, key: string, held: Partial<KeyboardEventInit> = {}) =>
      chordFromEvent(
        { key, code, ctrlKey: false, altKey: false, shiftKey: false, metaKey: false, ...held },
        true,
      )
    expect(press('NumpadEnter', 'Enter', { metaKey: true })).toBe('Cmd+NumEnter')
    expect(press('Enter', 'Enter', { metaKey: true })).toBe('Cmd+Enter')
    expect(press('NumpadSubtract', '-', { metaKey: true, altKey: true })).toBe('Alt+Cmd+NumMinus')
    expect(press('NumpadAdd', '+', { ctrlKey: true, shiftKey: true })).toBe('Ctrl+Shift+NumPlus')
    expect(press('Numpad5', '5')).toBe(null)
    expect(formatChord('Cmd+NumEnter', true)).toBe('\u2318\u2324')
    expect(formatChord('Ctrl+Shift+NumPlus', true)).toBe('\u2303\u21e7Num+')
  })

  it('lets a Customize Keyboard assignment win over a Word for Mac chord', () => {
    // Control+Shift+M given to Word Count: Increase Indent keeps Control+M only
    const moved = assignChord({}, 'word-count', 'Ctrl+Shift+M', true)!
    expect(moved).toEqual({ 'word-count': ['Ctrl+Shift+M'], 'indent-increase': ['Ctrl+M'] })
    expect(keyVerdict(buildKeyTable(moved, true), 'Ctrl+Shift+M', true)).toEqual({
      run: { cmd: 'word-count' },
    })
    // a chord taken from a caret move is the new command's, and the row's other chord is the key layer's
    const para = assignChord({}, 'strikethrough', 'Cmd+Up', true)!
    expect(documentKeyCommand(para, 'Cmd+Up', true)).toBe(null)
    expect(documentKeyCommand(para, 'Alt+Up', true)).toBe(null)
    const table = buildKeyTable(para, true)
    expect(keyVerdict(table, 'Cmd+Up', true)).toEqual({ run: { cmd: 'strikethrough' } })
    expect(keyVerdict(table, 'Alt+Up', true)).toEqual({
      run: { cmd: 'caret-move', payload: 'para-up' },
    })
    // an untouched command keeps its table chords
    expect(documentKeyCommand(para, 'Cmd+Down', true)).toEqual({
      cmd: 'caret-move',
      payload: 'para-down',
    })
  })
})
