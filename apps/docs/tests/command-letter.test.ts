/**
 * Keyboard layouts emit their own character in `event.key`, so Docs' command
 * shortcuts have to match the physical key (`event.code`) as well — under a
 * Cyrillic layout Ctrl+S used to save nothing and fall through to undo.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { isCommandLetter } from '../src/renderer/command-letter'

const press = (init: KeyboardEventInit) => new KeyboardEvent('keydown', init)

describe('command letters under a non-Latin layout', () => {
  it('matches the physical key when the layout reports another character', () => {
    // Bulgarian Phonetic: the S key types 'с', and Ctrl+S must still save
    expect(isCommandLetter(press({ key: 'с', code: 'KeyS' }), 's')).toBe(true)
    // Bulgarian BDS reports 'я' on the same physical key
    expect(isCommandLetter(press({ key: 'я', code: 'KeyS' }), 's')).toBe(true)
    expect(isCommandLetter(press({ key: 'σ', code: 'KeyS' }), 's')).toBe(true)
    expect(isCommandLetter(press({ key: 'ד', code: 'KeyD' }), 'd')).toBe(true)
  })

  it('still matches the letter a QWERTY-like layout reports', () => {
    expect(isCommandLetter(press({ key: 's', code: 'KeyS' }), 's')).toBe(true)
    expect(isCommandLetter(press({ key: 'S', code: 'KeyS' }), 's')).toBe(true)
    // a non-QWERTY mapping without a code (synthetic events, older shells)
    expect(isCommandLetter(press({ key: 's' }), 's')).toBe(true)
  })

  it('does not match another key that happens to report the letter', () => {
    expect(isCommandLetter(press({ key: 'd', code: 'KeyD' }), 's')).toBe(false)
    expect(isCommandLetter(press({ key: 'с', code: 'KeyC' }), 's')).toBe(false)
  })

  it('follows a Latin letter even when it sits on another physical key', () => {
    // Dvorak: Ctrl+O lands on the key at the QWERTY S position; the layout's
    // answer ('o') decides, so Open does not turn into Save
    expect(isCommandLetter(press({ key: 'o', code: 'KeyS' }), 'o')).toBe(true)
    expect(isCommandLetter(press({ key: 'o', code: 'KeyS' }), 's')).toBe(false)
    // Dvorak: the D key is at the H position — Font, not Replace
    expect(isCommandLetter(press({ key: 'd', code: 'KeyH' }), 'd')).toBe(true)
    expect(isCommandLetter(press({ key: 'd', code: 'KeyH' }), 'h')).toBe(false)
    // Colemak: the R key is at the S position — neither Save nor a command
    expect(isCommandLetter(press({ key: 'r', code: 'KeyS' }), 's')).toBe(false)
  })

  it('lets AltGr characters reach the document', () => {
    // Windows sends Ctrl+Alt for AltGr: 'ś' on the Polish layout, 'ó' on the
    // UK-extended one. Both sit on the S/O keys but are text, not shortcuts.
    expect(
      isCommandLetter(press({ key: 'ś', code: 'KeyS', ctrlKey: true, altKey: true }), 's'),
    ).toBe(false)
    expect(
      isCommandLetter(press({ key: 'ó', code: 'KeyO', ctrlKey: true, altKey: true }), 'o'),
    ).toBe(false)
  })

  it('leaves an IME composition alone', () => {
    expect(isCommandLetter({ key: 'Process', code: 'KeyS', isComposing: true }, 's')).toBe(false)
  })
})

describe('the Docs window handler', () => {
  const source = readFileSync(join(__dirname, '../src/renderer/App.tsx'), 'utf8')

  it('never compares a command letter against event.key alone', () => {
    // one `e.key === '<letter>'` short-circuit is exactly the Cyrillic bug:
    // the character the layout reports is not the letter the shortcut needs.
    // A new letter shortcut should call isCommandLetter(e, 'x') instead.
    expect(source).not.toMatch(/\be\.key === '[a-zA-Z]'/)
  })

  it('runs Save before the editor keymaps see the keystroke', () => {
    // bubble-phase handling could not stop the editor reverting the edit the
    // report describes, so the Save branch is registered in the capture phase
    const at = source.indexOf("isCommandLetter(e, 's')")
    expect(at).toBeGreaterThan(-1)
    expect(source.slice(at)).toMatch(/window\.addEventListener\('keydown', handler, true\)/)
  })
})
