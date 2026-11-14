import { describe, expect, it } from 'vitest'
import { isSaveShortcut } from '../src/renderer/save-shortcut'

const cases: [string, KeyboardEventInit, boolean][] = [
  ['Latin Ctrl+S', { key: 's', code: 'KeyS', ctrlKey: true }, true],
  ['Latin Cmd+S', { key: 's', code: 'KeyS', metaKey: true }, true],
  ['uppercase Latin on a remapped key', { key: 'S', code: 'KeyO', ctrlKey: true }, true],
  ['Latin s on a remapped key', { key: 's', code: 'KeyO', metaKey: true }, true],
  ['Dvorak Open', { key: 'o', code: 'KeyS', ctrlKey: true }, false],
  ['Colemak R', { key: 'r', code: 'KeyS', ctrlKey: true }, false],
  ['uppercase Dvorak Open', { key: 'O', code: 'KeyS', metaKey: true }, false],
  ['Cyrillic physical Save', { key: 'я', code: 'KeyS', metaKey: true }, true],
  ['Cyrillic Ctrl Save', { key: 'а', code: 'KeyS', ctrlKey: true }, true],
  ['Shift Save route', { key: 'S', code: 'KeyS', metaKey: true, shiftKey: true }, true],
  ['Polish Ctrl+Alt', { key: 'ś', code: 'KeyS', ctrlKey: true, altKey: true }, false],
  ['Latvian Ctrl+Alt', { key: 'š', code: 'KeyS', ctrlKey: true, altKey: true }, false],
  ['Romanian Ctrl+Alt', { key: 'ș', code: 'KeyS', ctrlKey: true, altKey: true }, false],
  ['Latin Ctrl+Alt', { key: 's', code: 'KeyS', ctrlKey: true, altKey: true }, false],
  ['Cmd+Alt', { key: 'я', code: 'KeyS', metaKey: true, altKey: true }, false],
  [
    'AltGraph without altKey',
    { key: 'ś', code: 'KeyS', ctrlKey: true, modifierAltGraph: true },
    false,
  ],
  ['AltGraph Latin', { key: 's', code: 'KeyS', metaKey: true, modifierAltGraph: true }, false],
  ['no command modifier', { key: 'я', code: 'KeyS' }, false],
  ['Alt only', { key: 'ś', code: 'KeyS', altKey: true }, false],
  ['non-Latin wrong position', { key: 'я', code: 'KeyZ', metaKey: true }, false],
  ['other Latin letter', { key: 'f', code: 'KeyF', ctrlKey: true }, false],
  ['ASCII punctuation is not a fallback', { key: ';', code: 'KeyS', ctrlKey: true }, false],
  [
    'unidentified key is not a character',
    { key: 'Unidentified', code: 'KeyS', ctrlKey: true },
    false,
  ],
]

describe('Save shortcut matching', () => {
  it.each(cases)('%s', (_name, init, expected) => {
    expect(isSaveShortcut(new KeyboardEvent('keydown', init))).toBe(expected)
  })
})
