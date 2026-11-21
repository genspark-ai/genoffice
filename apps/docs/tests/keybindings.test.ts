/**
 * The layer that lays the user's keyboard assignments over the built-in
 * shortcuts: what a chord resolves to, and that nothing changes until an
 * assignment is made.
 */
import { describe, expect, it } from 'vitest'
import {
  isAssignableChord,
  menuAccelerator,
  sanitizeKeyOverrides,
  toChord,
} from '../src/shared/keybindings'
import {
  assignChord,
  bindingCommand,
  buildKeyTable,
  chordFromEvent,
  chordOwner,
  defaultChords,
  effectiveChords,
  formatChord,
  keyVerdict,
  listKeyCommands,
  removeChord,
  shortcutChords,
  valueBindingLabel,
  VALUE_COMMANDS,
} from '../src/renderer/keybindings'
import { ownsItsKeys } from '../src/renderer/key-layer'
import { SHORTCUTS } from '../src/renderer/shortcuts'
import { COMMANDS, type CommandDef } from '../src/shared/commands'
import { t } from '../src/renderer/i18n/locale'

const press = (key: string, code: string, mods: Partial<Record<string, boolean>> = {}) => ({
  key,
  code,
  ctrlKey: !!mods.ctrl,
  altKey: !!mods.alt,
  shiftKey: !!mods.shift,
  metaKey: !!mods.meta,
})

describe('chords', () => {
  it('normalizes accelerators for the platform', () => {
    expect(toChord('Shift+CmdOrCtrl+S', true)).toBe('Shift+Cmd+S')
    expect(toChord('Shift+CmdOrCtrl+S', false)).toBe('Ctrl+Shift+S')
    expect(toChord('CmdOrCtrl+=', true)).toBe('Cmd+=')
    expect(toChord('Alt+Cmd+G', false)).toBe(null)
    expect(toChord('Ctrl+Nope', true)).toBe(null)
  })

  it('reads every chord of the shortcut list on both platforms', () => {
    for (const isMac of [true, false]) {
      for (const def of SHORTCUTS) {
        const expected = (isMac ? def.keys : (def.win ?? def.keys)).split(' / ').length
        expect(shortcutChords(def, isMac).length, def.id).toBe(expected)
      }
    }
    const byId = new Map(SHORTCUTS.map((def) => [def.id, def]))
    expect(shortcutChords(byId.get('superscript')!, true)).toEqual(['Shift+Cmd+=', 'Cmd+.'])
    expect(shortcutChords(byId.get('page-break')!, false)).toEqual(['Ctrl+Enter'])
    expect(shortcutChords(byId.get('endnote')!, false)).toEqual(['Ctrl+Alt+D'])
  })

  it('spells a key press, by letter where the layout types one and by position otherwise', () => {
    expect(chordFromEvent(press('k', 'KeyK', { meta: true, shift: true }), true)).toBe(
      'Shift+Cmd+K',
    )
    // Dvorak types 'o' on the S key
    expect(chordFromEvent(press('o', 'KeyS', { ctrl: true }), false)).toBe('Ctrl+O')
    // Option on the Mac and a Cyrillic layout type no Latin letter
    expect(chordFromEvent(press('©', 'KeyG', { meta: true, alt: true }), true)).toBe('Alt+Cmd+G')
    expect(chordFromEvent(press('ы', 'KeyS', { ctrl: true }), false)).toBe('Ctrl+S')
    expect(chordFromEvent(press('F5', 'F5', { shift: true }), false)).toBe('Shift+F5')
    expect(chordFromEvent(press('Shift', 'ShiftLeft', { shift: true }), false)).toBe(null)
  })

  it('accepts only chords that cannot be typing', () => {
    expect(isAssignableChord('Ctrl+Shift+K')).toBe(true)
    expect(isAssignableChord('F6')).toBe(true)
    expect(isAssignableChord('Shift+K')).toBe(false)
    expect(isAssignableChord('Alt+K')).toBe(false)
    expect(isAssignableChord('Tab')).toBe(false)
  })

  it('shows glyphs on the Mac', () => {
    expect(formatChord('Ctrl+Alt+Shift+Cmd+K', true)).toBe('⌃⌥⇧⌘K')
    expect(formatChord('Ctrl+Shift+K', false)).toBe('Ctrl+Shift+K')
  })
})

describe('built-in bindings', () => {
  it.each([true, false])('give every chord to one binding (mac: %s)', (isMac) => {
    const seen = new Map<string, string>()
    const ids = [
      ...Object.keys(COMMANDS),
      ...SHORTCUTS.flatMap((def) => (def.cmd && def.payload ? `${def.cmd}:${def.payload}` : [])),
    ]
    for (const id of ids) {
      for (const chord of defaultChords(id, isMac)) {
        expect(seen.get(chord), `${chord}: ${seen.get(chord)} and ${id}`).toBe(undefined)
        seen.set(chord, id)
        expect(chordOwner({}, chord, isMac)).toEqual({ kind: 'binding', id })
      }
    }
    expect(seen.size).toBeGreaterThan(40)
  })

  it('cover every accelerator of the command table', () => {
    for (const [id, def] of Object.entries<CommandDef>(COMMANDS)) {
      if (def.accelerator && !def.role)
        expect(defaultChords(id, true).length, id).toBeGreaterThan(0)
    }
  })

  it('leave a chord the command table gives to another command with that command', () => {
    for (const isMac of [true, false]) {
      const mod = isMac ? 'Cmd' : 'Ctrl'
      expect(defaultChords('repeat', isMac)).toEqual([`${mod}+Y`])
      expect(defaultChords('redo', isMac)).not.toContain(`${mod}+Y`)
      expect(defaultChords('redo', isMac).length).toBe(1)
    }
    expect(defaultChords('open', true)).toEqual(['Cmd+O', 'Shift+Cmd+O'])
  })

  it('link shortcut rows only to commands that can run', () => {
    for (const def of SHORTCUTS) {
      if (def.cmd) expect(bindingCommand(def.cmd), def.id).not.toBe(null)
    }
    expect(bindingCommand('cut')).toBe(null)
    expect(bindingCommand('no-such-command')).toBe(null)
    expect(bindingCommand('line-spacing:1.5')).toEqual({ cmd: 'line-spacing', payload: '1.5' })
  })
})

describe('assignments', () => {
  it('leave every key alone until one is made', () => {
    const table = buildKeyTable({}, true)
    expect(table.run.size).toBe(0)
    expect(table.dropped.size).toBe(0)
    for (const chord of ['Cmd+B', 'Cmd+S', 'Ctrl+D', 'Tab', 'Cmd+K'])
      expect(keyVerdict(table, chord, true)).toBe('pass')
  })

  it('add a chord to a command and keep its built-in one', () => {
    const next = assignChord({}, 'strikethrough', 'Ctrl+Shift+X', false)!
    expect(next).toEqual({ strikethrough: ['Ctrl+Shift+X'] })
    const bold = assignChord(next, 'bold', 'Ctrl+Shift+Y', false)!
    expect(effectiveChords(bold, 'bold', false)).toEqual(['Ctrl+B', 'Ctrl+Shift+Y'])
    const table = buildKeyTable(bold, false)
    expect(keyVerdict(table, 'Ctrl+Shift+X', false)).toEqual({ run: { cmd: 'strikethrough' } })
    expect(keyVerdict(table, 'Ctrl+B', false)).toEqual({ run: { cmd: 'bold' } })
    expect(keyVerdict(table, 'Ctrl+I', false)).toBe('pass')
  })

  it('move a chord from the command that had it', () => {
    expect(chordOwner({}, 'Cmd+B', true)).toEqual({ kind: 'binding', id: 'bold' })
    const next = assignChord({}, 'strikethrough', 'Cmd+B', true)!
    expect(next).toEqual({ bold: [], strikethrough: ['Cmd+B'] })
    expect(chordOwner(next, 'Cmd+B', true)).toEqual({ kind: 'binding', id: 'strikethrough' })
    expect(keyVerdict(buildKeyTable(next, true), 'Cmd+B', true)).toEqual({
      run: { cmd: 'strikethrough' },
    })
  })

  it('run a command with its payload', () => {
    const next = assignChord({}, 'insert-section-break:continuous', 'Ctrl+F6', false)!
    expect(keyVerdict(buildKeyTable(next, false), 'Ctrl+F6', false)).toEqual({
      run: { cmd: 'insert-section-break', payload: 'continuous' },
    })
  })

  it('swallow a built-in chord that was removed, the Control twin on the Mac included', () => {
    const next = removeChord({}, 'font-dialog', 'Cmd+D', true)
    expect(next).toEqual({ 'font-dialog': [] })
    const table = buildKeyTable(next, true)
    expect(keyVerdict(table, 'Cmd+D', true)).toBe('swallow')
    expect(keyVerdict(table, 'Ctrl+D', true)).toBe('swallow')
    // Control+M is a binding of its own on the Mac, not Command+M's twin
    expect(keyVerdict(table, 'Ctrl+M', true)).toBe('pass')
  })

  it('refuse typing keys, built-in keys that are not commands, and system keys', () => {
    expect(assignChord({}, 'bold', 'Shift+K', false)).toBe(null)
    expect(chordOwner({}, 'Cmd+W', true)).toMatchObject({ kind: 'fixed' })
    expect(assignChord({}, 'bold', 'Cmd+W', true)).toBe(null)
    expect(chordOwner({}, 'Cmd+Q', true)).toEqual({ kind: 'fixed', row: null })
    expect(assignChord({}, 'bold', 'Cmd+Q', true)).toBe(null)
    expect(assignChord({}, 'cut', 'Ctrl+F6', false)).toBe(null)
  })

  it('forget a binding that is back at its built-in chords', () => {
    const added = assignChord({}, 'bold', 'Ctrl+F6', false)!
    expect(removeChord(added, 'bold', 'Ctrl+F6', false)).toEqual({})
  })
})

describe('the key layer', () => {
  it('leaves dialogs and text fields outside the document their keys', () => {
    document.body.innerHTML =
      '<div class="ProseMirror" contenteditable="true"><p id="doc">x</p></div>' +
      '<table><tbody><tr><td id="cell" contenteditable="true">x</td></tr></tbody></table>' +
      '<textarea id="comment"></textarea><input id="find" />' +
      '<div class="modal-backdrop"><button id="modal">x</button></div><button id="ribbon">x</button>'
    const owns = (id: string) => ownsItsKeys(document.getElementById(id)!)
    expect(owns('doc')).toBe(false)
    expect(owns('cell')).toBe(false)
    expect(owns('ribbon')).toBe(false)
    expect(owns('comment')).toBe(true)
    expect(owns('find')).toBe(true)
    expect(owns('modal')).toBe(true)
    document.body.innerHTML = ''
  })
})

describe('stored assignments', () => {
  it('keep only well-formed entries', () => {
    expect(
      sanitizeKeyOverrides(
        {
          bold: ['Ctrl+Shift+Y', 'ctrl+shift+y', 7, 'Ctrl+Shift+Y'],
          'line-spacing:1.5': ['Ctrl+F6'],
          '../x': ['Ctrl+F7'],
          italic: 'Ctrl+F8',
        },
        false,
      ),
    ).toEqual({ bold: ['Ctrl+Shift+Y'], 'line-spacing:1.5': ['Ctrl+F6'] })
    expect(sanitizeKeyOverrides(null, false)).toEqual({})
    expect(sanitizeKeyOverrides([['Ctrl+F6']], false)).toEqual({})
  })

  it('show the chord assigned last on the menu item, the built-in one otherwise', () => {
    expect(menuAccelerator({}, 'bold', 'CmdOrCtrl+B')).toBe('CmdOrCtrl+B')
    expect(menuAccelerator({ bold: ['Ctrl+B', 'Ctrl+F6'] }, 'bold', 'CmdOrCtrl+B')).toBe('Ctrl+F6')
    expect(menuAccelerator({ bold: [] }, 'bold', 'CmdOrCtrl+B')).toBe(undefined)
    // a typing key stays off the menu, where it would fire from any text field
    expect(menuAccelerator({ bold: ['Ctrl+F6', 'Ctrl+Enter'] }, 'bold', undefined)).toBe('Ctrl+F6')
    expect(menuAccelerator({ bold: ['Ctrl+Enter'] }, 'bold', undefined)).toBe(undefined)
    expect(menuAccelerator({ bold: ['Ctrl+F6'] }, 'toString', undefined)).toBe(undefined)
  })
})

describe('the command list', () => {
  it('names menu commands by their menu label and adds the ones with no menu item', () => {
    const commands = listKeyCommands(
      [
        { id: 'font-dialog', label: 'Font…', menu: 'Format', path: 'Format ▸ Font…' },
        { id: 'cut', label: 'Cut', menu: 'Edit', path: 'Edit ▸ Cut' },
      ],
      t,
    )
    const byId = new Map(commands.map((command) => [command.id, command]))
    expect(byId.get('font-dialog')).toMatchObject({ label: 'Font', menu: 'Format' })
    expect(byId.has('cut')).toBe(false)
    expect(byId.get('strikethrough')).toMatchObject({ menu: null })
    expect(byId.get('line-spacing:1.5')?.label).toContain('1.5')
    expect(new Set(commands.map((command) => command.id)).size).toBe(commands.length)
  })
})

describe('commands that take a value', () => {
  it('keep a font or style name as the payload of a stored binding', () => {
    expect(
      sanitizeKeyOverrides(
        {
          'font-family:Times New Roman': ['Ctrl+F6'],
          'font-family:\u6e38\u660e\u671d': ['Ctrl+F7'],
          'apply-style:Heading1': ['Ctrl+F8'],
          'font-color:FF0000': ['Ctrl+F9'],
          'font-family:bad\u0007name': ['Ctrl+F10'],
          [`font-family:${'x'.repeat(121)}`]: ['Ctrl+F11'],
          'font-family:': ['Ctrl+F12'],
        },
        false,
      ),
    ).toEqual({
      'font-family:Times New Roman': ['Ctrl+F6'],
      'font-family:\u6e38\u660e\u671d': ['Ctrl+F7'],
      'apply-style:Heading1': ['Ctrl+F8'],
      'font-color:FF0000': ['Ctrl+F9'],
    })
  })

  it('run with the value once a chord is assigned', () => {
    const id = 'font-family:Times New Roman'
    expect(bindingCommand(id)).toEqual({ cmd: 'font-family', payload: 'Times New Roman' })
    const overrides = assignChord({}, id, 'Ctrl+Shift+F9', false)
    expect(overrides).toEqual({ [id]: ['Ctrl+Shift+F9'] })
    expect(keyVerdict(buildKeyTable(overrides!, false), 'Ctrl+Shift+F9', false)).toEqual({
      run: { cmd: 'font-family', payload: 'Times New Roman' },
    })
    expect(chordOwner(overrides!, 'Ctrl+Shift+F9', false)).toEqual({ kind: 'binding', id })
    // the chord moves from one value to another of the same command
    const moved = assignChord(overrides!, 'font-family:Arial', 'Ctrl+Shift+F9', false)
    expect(moved).toEqual({ 'font-family:Arial': ['Ctrl+Shift+F9'] })
    // a style name may hold a colon: the command ends at the first one
    expect(bindingCommand('apply-style:a:b')).toEqual({ cmd: 'apply-style', payload: 'a:b' })
  })

  it('are named by their value', () => {
    expect(valueBindingLabel('font-family:Times New Roman', t)).toBe('Times New Roman')
    expect(valueBindingLabel('apply-style:Heading1', t)).toBe('Heading1')
    expect(valueBindingLabel('apply-style:Heading1', t, () => 'Heading 1')).toBe('Heading 1')
    expect(valueBindingLabel('font-size:12', t)).toBe(`${t('ribbonFontSizeTip')}: 12`)
    expect(valueBindingLabel('font-color:FF0000', t)).toBe(`${t('ribbonFontColor')}: FF0000`)
    expect(valueBindingLabel('font-color:none', t)).toBe(
      `${t('ribbonFontColor')}: ${t('ribbonAutomatic')}`,
    )
    expect(valueBindingLabel('highlight:none', t)).toBe(
      `${t('ribbonTextHighlightColor')}: ${t('ribbonNoColor')}`,
    )
    expect(valueBindingLabel('highlight:yellow', t)).toBe(
      `${t('ribbonTextHighlightColor')}: yellow`,
    )
    // a binding with no value, or of a command that is not picked by value, is not one of these
    expect(valueBindingLabel('bold', t)).toBe(null)
    expect(valueBindingLabel('font-size', t)).toBe(null)
    expect(valueBindingLabel('line-spacing:1.5', t)).toBe(null)
    expect(valueBindingLabel('nope:1', t)).toBe(null)
  })

  it('are commands of the table with a handler', () => {
    for (const { cmd } of VALUE_COMMANDS) expect(bindingCommand(`${cmd}:x`)?.cmd).toBe(cmd)
    // the bare command is offered in the list but is no binding of its own
    const listed = listKeyCommands([], t).map((command) => command.id)
    for (const { cmd } of VALUE_COMMANDS) expect(listed).not.toContain(cmd)
  })
})
