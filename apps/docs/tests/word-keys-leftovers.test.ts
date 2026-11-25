/**
 * Chords of Word for Mac whose commands came last: Normal style on the keypad,
 * Link to Previous, Insert Merge Field, Mark Table of Contents Entry, Next
 * Misspelling and Copy Text; and the menu bar following a document's own keys.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { inlineToRuns } from '../src/renderer/editor/convert'
import { copyTextKey, copyTextModeOn, startCopyText } from '../src/renderer/editor/copy-text'
import { markTocEntry, tocEntryInstruction } from '../src/renderer/editor/index-entry'
import { tocEntryField } from '../src/renderer/editor/index-tables'
import { ignoreSpellingOnce, misspellingAfter } from '../src/renderer/editor/spell-menu'
import { MENU_COMMAND_HANDLERS, type CommandHost } from '../src/renderer/commands/menu-commands'
import {
  assignChord,
  buildKeyTable,
  defaultChords,
  documentKeyCommand,
  keyVerdict,
} from '../src/renderer/keybindings'
import { loadDocumentKeys } from '../src/renderer/key-layer'

type Json = Record<string, unknown>
let editor: Editor

const para = (value: string): Json => ({
  type: 'docParagraph',
  content: [{ type: 'text', text: value }],
})

function make(content: Json[]): Editor {
  const element = document.createElement('div')
  document.body.appendChild(element)
  editor = new Editor({
    element,
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  return editor
}

afterEach(() => {
  editor?.destroy()
  document.body.replaceChildren()
})

const paragraphText = (index = 0) => editor.state.doc.child(index).textContent

describe('the chords', () => {
  it('run their commands from the document key table on the Mac only', () => {
    const expected: [string, string][] = [
      ['Alt+Shift+Clear', 'style-normal'],
      ['Ctrl+Shift+R', 'hf-link-previous'],
      ['Ctrl+Shift+F', 'insert-merge-field'],
      ['Alt+Shift+Cmd+O', 'mark-toc-entry'],
      ['Alt+F7', 'spell-next'],
      ['Shift+F2', 'copy-text'],
    ]
    for (const [chord, cmd] of expected) {
      expect(documentKeyCommand({}, chord, true), chord).toEqual({ cmd })
      expect(documentKeyCommand({}, chord, false), chord).toBe(null)
    }
    // F2 alone stays Cut, F7 alone Spelling and Grammar
    expect(documentKeyCommand({}, 'F2', true)).toEqual({ cmd: 'edit-text', payload: 'cut' })
    expect(documentKeyCommand({}, 'F7', true)).toEqual({ cmd: 'tools-spelling' })
  })

  it('make Normal style a command whose three keys Customize Keyboard can move', () => {
    expect(defaultChords('style-normal', true)).toEqual([
      'Shift+Cmd+N',
      'Alt+Cmd+0',
      'Alt+Shift+Clear',
    ])
    expect(defaultChords('style-normal', false)).toEqual(['Ctrl+Alt+0'])
    const moved = assignChord({}, 'word-count', 'Shift+Cmd+N', true)!
    expect(moved['style-normal']).toEqual(['Alt+Cmd+0', 'Alt+Shift+Clear'])
    expect(keyVerdict(buildKeyTable(moved, true), 'Shift+Cmd+N', true)).toEqual({
      run: { cmd: 'word-count' },
    })
    // the keypad chord goes with the customized binding, not the document key table
    expect(documentKeyCommand(moved, 'Alt+Shift+Clear', true)).toBe(null)
    expect(keyVerdict(buildKeyTable(moved, true), 'Alt+Shift+Clear', true)).toEqual({
      run: { cmd: 'style-normal' },
    })
  })
})

describe('the commands', () => {
  it('reach the host, and do nothing without a document to change', () => {
    make([para('one')])
    const fns = {
      applyHeadingLevel: vi.fn(),
      headerFooterAction: vi.fn(),
      showMergeFields: vi.fn(),
      openMarkTocEntry: vi.fn(),
      nextMisspelling: vi.fn(),
    }
    const host = { editor, hasDoc: true, ...fns } as unknown as CommandHost
    const all = [
      'style-normal',
      'hf-link-previous',
      'insert-merge-field',
      'mark-toc-entry',
      'spell-next',
    ] as const
    for (const id of all) MENU_COMMAND_HANDLERS[id](host, undefined)
    expect(fns.applyHeadingLevel).toHaveBeenCalledWith(0)
    expect(fns.headerFooterAction).toHaveBeenCalledWith({ type: 'toggleLinkToPrevious' })
    for (const fn of Object.values(fns)) expect(fn).toHaveBeenCalledTimes(1)
    for (const id of all) MENU_COMMAND_HANDLERS[id]({ ...host, hasDoc: false }, undefined)
    editor.setEditable(false)
    for (const id of all) MENU_COMMAND_HANDLERS[id](host, undefined)
    MENU_COMMAND_HANDLERS['copy-text'](host, undefined)
    expect(copyTextModeOn()).toBe(false)
    for (const fn of Object.values(fns)) expect(fn).toHaveBeenCalledTimes(1)
  })
})

describe('Copy Text (Shift+F2)', () => {
  const press = (key: string) => {
    const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
    let taken = false
    const listener = (e: KeyboardEvent) => {
      taken = copyTextKey(e)
      if (taken) e.stopImmediatePropagation()
    }
    window.addEventListener('keydown', listener, true)
    editor.view.dom.dispatchEvent(event)
    window.removeEventListener('keydown', listener, true)
    return taken
  }

  it('copies the selection to where the caret is when Return is pressed', () => {
    make([para('one two'), para('end')])
    editor.commands.setTextSelection({ from: 1, to: 4 })
    expect(startCopyText(editor)).toBe('to')
    editor.commands.setTextSelection(13)
    expect(press('a')).toBe(false)
    expect(press('Enter')).toBe(true)
    expect(paragraphText(0)).toBe('one two')
    expect(paragraphText(1)).toBe('endone')
    expect(copyTextModeOn()).toBe(false)
    expect(press('Enter')).toBe(false)
  })

  it('with a bare caret asks what to copy, and brings it to the caret', () => {
    make([para('one two'), para('end')])
    editor.commands.setTextSelection(13)
    expect(startCopyText(editor)).toBe('from')
    editor.commands.setTextSelection({ from: 5, to: 8 })
    expect(press('Enter')).toBe(true)
    expect(paragraphText(1)).toBe('endtwo')
    expect(paragraphText(0)).toBe('one two')
  })

  it('is left by Escape and by an edit', () => {
    make([para('one two')])
    editor.commands.setTextSelection({ from: 1, to: 4 })
    startCopyText(editor)
    expect(press('Escape')).toBe(true)
    expect(copyTextModeOn()).toBe(false)
    // the question stays in the status bar until the mode is left
    const copyTextPrompt = vi.fn()
    const host = { editor, hasDoc: true, copyTextPrompt } as unknown as CommandHost
    MENU_COMMAND_HANDLERS['copy-text'](host, undefined)
    expect(copyTextPrompt.mock.calls).toEqual([['to']])
    editor.commands.setTextSelection(8)
    expect(press('Enter')).toBe(true)
    expect(copyTextPrompt.mock.calls).toEqual([['to'], [null]])
    expect(paragraphText()).toBe('one twoone')
    editor.commands.setTextSelection({ from: 1, to: 4 })
    startCopyText(editor)
    editor.commands.insertContentAt(8, '!')
    expect(copyTextModeOn()).toBe(false)
    expect(paragraphText()).toBe('one two!one')
  })
})

describe('Next Misspelling (Option+F7)', () => {
  const wrong = (word: string) => word.startsWith('x')

  it('takes the word at the caret, then the ones after it, and goes round the end', () => {
    make([para('good xone fine'), para('xtwo ok')])
    editor.commands.setTextSelection(7)
    const first = misspellingAfter(editor.state, wrong)!
    expect(first.word).toBe('xone')
    editor.commands.setTextSelection({ from: first.from, to: first.to })
    const second = misspellingAfter(editor.state, wrong)!
    expect(second.word).toBe('xtwo')
    editor.commands.setTextSelection({ from: second.from, to: second.to })
    expect(misspellingAfter(editor.state, wrong)!.word).toBe('xone')
  })

  it('passes a word set aside by Ignore Once and says when nothing is left', () => {
    make([para('xone xtwo')])
    editor.commands.setTextSelection(1)
    ignoreSpellingOnce(editor, { from: 1, to: 5 })
    expect(misspellingAfter(editor.state, wrong)!.word).toBe('xtwo')
    expect(misspellingAfter(editor.state, () => false)).toBe(null)
  })
})

describe('Mark Table of Contents Entry', () => {
  it('writes a TC field the table of contents reads back', () => {
    expect(tocEntryInstruction({ text: 'A "quoted" title', id: 'C', level: 2 })).toBe(
      '"A quoted title" \\l 2',
    )
    expect(tocEntryInstruction({ text: 'Fig', id: 'f', level: 12 })).toBe('"Fig" \\f F \\l 9')
    expect(tocEntryField(tocEntryInstruction({ text: 'Fig', id: 'F', level: 3 }))).toEqual({
      text: 'Fig',
      id: 'F',
      level: 3,
    })
    expect(tocEntryField(tocEntryInstruction({ text: 'Plain', id: 'C', level: 1 }))).toEqual({
      text: 'Plain',
      level: 1,
    })
  })

  it('puts the hidden marker behind the selection and saves it as a TC field', () => {
    make([para('Intro text')])
    editor.commands.setTextSelection({ from: 1, to: 6 })
    markTocEntry(editor, 6, { text: 'Intro', id: 'C', level: 1 })
    const runs = inlineToRuns((editor.getJSON().content![0].content ?? []) as never)
    expect(runs.map((run) => run.tcInstr ?? run.text)).toEqual(['Intro', '"Intro" \\l 1', ' text'])
    expect(editor.state.selection.empty).toBe(true)
    markTocEntry(editor, 1, { text: '  ', id: 'C', level: 1 })
    expect(editor.state.doc.child(0).childCount).toBe(3)
  })
})

describe('the menu bar while a document carries keys', () => {
  it('is told the keys in force, and told again when the document has none', () => {
    const setDocumentKeyOverrides = vi.fn()
    ;(window as unknown as { desktop: unknown }).desktop = { setDocumentKeyOverrides }
    const WNE = 'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"'
    loadDocumentKeys(
      `<wne:tcg ${WNE}><wne:keymaps><wne:keymap wne:kcmPrimary="024B">` +
        `<wne:fci wne:fciName="Bold" wne:swArg="0000"/></wne:keymap></wne:keymaps></wne:tcg>`,
    )
    expect(setDocumentKeyOverrides).toHaveBeenCalledTimes(1)
    const sent = setDocumentKeyOverrides.mock.calls[0][0] as Record<string, string[]>
    expect(sent.bold).toContain('Ctrl+K')
    loadDocumentKeys(undefined)
    expect(setDocumentKeyOverrides).toHaveBeenLastCalledWith(null)
    loadDocumentKeys(undefined)
    expect(setDocumentKeyOverrides).toHaveBeenCalledTimes(2)
    delete (window as unknown as { desktop?: unknown }).desktop
  })
})
