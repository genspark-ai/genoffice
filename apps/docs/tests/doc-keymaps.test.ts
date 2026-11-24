/** Keyboard assignments a document carries: which run here, and how they sit over the machine's. */
import { describe, expect, it } from 'vitest'
import { readKeymaps } from '@genoffice/docx-engine'
import {
  DOCUMENT_BINDINGS,
  assignDocKey,
  chordToKcm,
  docKeysFromXml,
  docKeysXml,
  isLiveDocKey,
  kcmToChord,
  layerDocKeys,
  removeDocKey,
  resetDocKeys,
} from '../src/renderer/doc-keymaps'
import {
  bindingCommand,
  buildKeyTable,
  defaultChords,
  effectiveChords,
  keyVerdict,
} from '../src/renderer/keybindings'

const MAC = true
const WNE = 'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"'
const keymap = (kcm: string, inner: string, more = '') =>
  `<wne:keymap wne:kcmPrimary="${kcm}"${more}>${inner}</wne:keymap>`
const fci = (name: string) => `<wne:fci wne:fciName="${name}" wne:swArg="0000"/>`
const part = (...entries: string[]) =>
  `<wne:tcg ${WNE}><wne:keymaps>${entries.join('')}</wne:keymaps></wne:tcg>`

describe('a key as a document writes it', () => {
  it('is the modifier byte and the virtual key code, as the standard gives them', () => {
    // [MS-OE376] 3.1.2.1.1.8: Control-K is 024B, Alt-F9 0478, Shift-Control-Alt-C 0743
    expect(chordToKcm('Ctrl+K')).toBe(0x024b)
    expect(chordToKcm('Alt+F9')).toBe(0x0478)
    expect(chordToKcm('Ctrl+Alt+Shift+C')).toBe(0x0743)
    expect(chordToKcm('Shift+Cmd+B')).toBe(0x0942)
    expect(chordToKcm('Ctrl+7')).toBe(0x0237)
    expect(chordToKcm('Cmd+=')).toBe(0x08bb)
    // the keypad's Enter has no key code of its own
    expect(chordToKcm('Cmd+NumEnter')).toBeNull()
  })

  it('reads back to the same chord, and to nothing for a key there is no chord for', () => {
    for (const chord of ['Ctrl+K', 'Alt+F9', 'Ctrl+Alt+Shift+C', 'Shift+Cmd+B', 'Cmd+=', 'Cmd+Up'])
      expect(kcmToChord(chordToKcm(chord)!, MAC)).toBe(chord)
    expect(kcmToChord(0x0842, false)).toBeNull()
    expect(kcmToChord(0x1042, MAC)).toBeNull()
    expect(kcmToChord(0x02e7, MAC)).toBeNull()
  })

  it('names only commands this app has', () => {
    expect(DOCUMENT_BINDINGS.filter((id) => !bindingCommand(id))).toEqual([])
  })
})

describe('the keys of an open document', () => {
  const xml = part(
    keymap('094B', fci('Italic')),
    keymap('094D', '<wne:macro wne:macroName="PROJECT.MODULE1.RUN"/>'),
    keymap('094E', fci('NoSuchCommandHere')),
    keymap('0942', fci('bold'), ' wne:kcmSecondary="0441"'),
    '<wne:keymap wne:mask="1" wne:kcmPrimary="0849"/>',
  )
  const keys = docKeysFromXml(xml, MAC)

  it('run the fixed commands this app has and mask built-in keys; the rest is only carried', () => {
    expect(keys.map((key) => [key.chord, key.id, isLiveDocKey(key)])).toEqual([
      ['Shift+Cmd+K', 'italic', true],
      ['Shift+Cmd+M', null, false],
      ['Shift+Cmd+N', null, false],
      // a two-key sequence is not a chord here, whatever it runs
      [null, 'bold', false],
      ['Cmd+I', null, true],
    ])
  })

  it('sit over the assignments of the machine', () => {
    const machine = { bold: ['Cmd+B', 'Ctrl+Shift+Y'] }
    const layered = layerDocKeys(machine, keys, MAC)
    expect(effectiveChords(layered, 'italic', MAC)).toContain('Shift+Cmd+K')
    expect(effectiveChords(layered, 'italic', MAC)).not.toContain('Cmd+I')
    expect(effectiveChords(layered, 'bold', MAC)).toEqual(['Cmd+B', 'Ctrl+Shift+Y'])
    const table = buildKeyTable(layered, MAC)
    expect(keyVerdict(table, 'Shift+Cmd+K', MAC)).toEqual({ run: { cmd: 'italic' } })
    // the masked built-in key runs nothing
    expect(keyVerdict(table, 'Cmd+I', MAC)).toBe('swallow')
    expect(keyVerdict(table, 'Shift+Cmd+M', MAC)).toBe('pass')
    expect(layerDocKeys(machine, [], MAC)).toBe(machine)
  })

  it('take a new key as a key map of the command Word names, and give up the chord elsewhere', () => {
    const next = assignDocKey({}, keys, 'bold', 'Shift+Cmd+K', MAC)!
    expect(next.filter(isLiveDocKey).map((key) => [key.chord, key.id])).toEqual([
      ['Cmd+I', null],
      ['Shift+Cmd+K', 'bold'],
    ])
    const out = docKeysXml(xml, next)!
    expect(out).toContain(keymap('094D', '<wne:macro wne:macroName="PROJECT.MODULE1.RUN"/>'))
    expect(out).toContain(keymap('094E', fci('NoSuchCommandHere')))
    expect(out).toContain(keymap('0942', fci('bold'), ' wne:kcmSecondary="0441"'))
    expect(out).toContain(keymap('094B', fci('Bold')))
    expect(out).not.toContain(fci('Italic'))
    expect(readKeymaps(out)).toHaveLength(5)
    // a command Word has no name for, and a key no key code spells, stay out of the document
    expect(assignDocKey({}, keys, 'toggle-ai', 'Shift+Cmd+K', MAC)).toBeNull()
    expect(assignDocKey({}, keys, 'bold', 'Cmd+NumEnter', MAC)).toBeNull()
    // a chord the system keeps cannot be given away
    expect(assignDocKey({}, keys, 'bold', 'Cmd+Q', MAC)).toBeNull()
  })

  it('drop their own key map on Remove, and mask a key that comes from elsewhere', () => {
    const own = removeDocKey({}, keys, 'italic', 'Shift+Cmd+K', MAC)!
    expect(own.some((key) => key.chord === 'Shift+Cmd+K')).toBe(false)
    expect(own).toHaveLength(keys.length - 1)

    expect(defaultChords('bold', MAC)).toContain('Cmd+B')
    const masked = removeDocKey({}, keys, 'bold', 'Cmd+B', MAC)!
    expect(masked.at(-1)).toMatchObject({ chord: 'Cmd+B', id: null, entry: { mask: true } })
    expect(effectiveChords(layerDocKeys({}, masked, MAC), 'bold', MAC)).not.toContain('Cmd+B')
    expect(docKeysXml(xml, masked)).toContain('<wne:keymap wne:mask="1" wne:kcmPrimary="0842"/>')

    // taking the mask off gives the built-in key back
    const unmasked = removeDocKey({}, keys, 'italic', 'Cmd+I', MAC)!
    expect(unmasked.filter((key) => key.chord === 'Cmd+I')).toEqual([
      expect.objectContaining({ id: null }),
    ])
  })

  it('reset to the key maps this app does not run', () => {
    const left = resetDocKeys(keys)
    expect(left.map((key) => key.chord)).toEqual(['Shift+Cmd+M', 'Shift+Cmd+N', null])
    expect(readKeymaps(docKeysXml(xml, left)!)).toHaveLength(3)
  })

  it('are none for a document without the part, which gets one with its first key', () => {
    expect(docKeysFromXml(undefined, MAC)).toEqual([])
    expect(docKeysXml(undefined, [])).toBeNull()
    const made = assignDocKey({}, [], 'bold', 'Ctrl+Shift+Y', MAC)!
    expect(docKeysXml(undefined, made)).toContain(keymap('0359', fci('Bold')))
  })

  it('leave Save on its key: a file can neither mask it nor give it to another command', () => {
    const taken = docKeysFromXml(
      part(keymap('0853', fci('Italic')), '<wne:keymap wne:mask="1" wne:kcmPrimary="0853"/>'),
      MAC,
    )
    const layered = layerDocKeys({}, taken, MAC)
    expect(layered).toEqual({})
    expect(effectiveChords(layered, 'save', MAC)).toContain('Cmd+S')
    expect(assignDocKey({}, [], 'italic', 'Cmd+S', MAC)).toBeNull()
    expect(removeDocKey({}, [], 'save', 'Cmd+S', MAC)).toBeNull()
  })
})
