/**
 * The second batch of commands behind chords of Word for Mac: field braces
 * typed in place, Browse, sections, windows, the split pane, Outline levels.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  closeFieldCodes,
  insertFieldChars,
  openFieldCodes,
} from '../src/renderer/editor/field-chars'
import { collectFields, toggleAllFieldCodes } from '../src/renderer/editor/field-codes'
import { inlineToRuns } from '../src/renderer/editor/convert'
import {
  foldsForLevel,
  outlineViewKey,
  showOutlineLevel,
} from '../src/renderer/editor/outline-view'
import { currentBrowseObject, goToPosition, setBrowseObject } from '../src/renderer/editor/goto'
import { parseCaretMove } from '../src/renderer/editor/caret-keys'
import { neighbourWindow } from '../src/renderer/window-order'
import { MENU_COMMAND_HANDLERS, type CommandHost } from '../src/renderer/commands/menu-commands'

type Json = Record<string, unknown>
let editor: Editor

const text = (value: string, marks?: Json[]): Json => ({
  type: 'text',
  text: value,
  ...(marks ? { marks } : {}),
})
const para = (...content: Json[]): Json => ({ type: 'docParagraph', content })
const heading = (level: number, value: string): Json => ({
  type: 'docHeading',
  attrs: { level },
  content: [text(value)],
})

function make(content: Json[]): Editor {
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  return editor
}

afterEach(() => {
  setBrowseObject('page')
  editor?.destroy()
})

const paragraphText = (index = 0) => editor.state.doc.child(index).textContent
const runsOf = (index = 0) =>
  inlineToRuns((editor.getJSON().content![index].content ?? []) as never)

describe('field braces typed in place (Cmd+F9)', () => {
  it('puts the caret between the braces and takes what is typed as the code', () => {
    make([para(text('ab'))])
    editor.commands.setTextSelection(2)
    expect(insertFieldChars(editor)).toBe(true)
    expect(paragraphText()).toBe('a  b')
    expect(editor.state.selection.from).toBe(3)
    editor.commands.insertContent('DATE')
    const [open] = openFieldCodes(editor.state)
    expect(editor.state.doc.textBetween(open.from, open.to)).toBe(' DATE ')
    // typing after the closing brace is ordinary text again
    editor.commands.setTextSelection(open.to)
    editor.commands.insertContent('!')
    expect(openFieldCodes(editor.state)[0].to).toBe(open.to)
  })

  it('wraps a selection as the code and refuses a field inside a field', () => {
    make([para(text('see PAGE here'))])
    editor.commands.setTextSelection({ from: 5, to: 9 })
    expect(insertFieldChars(editor)).toBe(true)
    expect(paragraphText()).toBe('see  PAGE  here')
    expect(insertFieldChars(editor)).toBe(false)
    editor.setEditable(false)
    editor.commands.setTextSelection(1)
    expect(insertFieldChars(editor)).toBe(false)
  })

  it('becomes a field with a result on Update Field', () => {
    make([para(text('ab'))])
    editor.commands.setTextSelection(2)
    insertFieldChars(editor)
    editor.commands.insertContent('AUTHOR')
    const valueOf = vi.fn((instr: string) => (instr === 'AUTHOR' ? 'Ada' : ''))
    expect(closeFieldCodes(editor, valueOf)).toBe(1)
    expect(paragraphText()).toBe('aAdab')
    const [field] = collectFields(editor.state.doc)
    expect([field.instr, field.mark.attrs.codeEdit, field.mark.attrs.dirty]).toEqual([
      'AUTHOR',
      false,
      false,
    ])
    expect(closeFieldCodes(editor, valueOf)).toBe(0)
  })

  it('leaves a code this editor cannot compute for Word, and drops empty braces', () => {
    make([para(text('ab')), para(text('cd'))])
    editor.commands.setTextSelection(2)
    insertFieldChars(editor)
    // AutoCorrect curls typed quotes; a switch needs them straight
    editor.commands.insertContent('QUOTE \u201cx\u201d')
    editor.commands.setTextSelection(editor.state.doc.content.size - 2)
    insertFieldChars(editor)
    expect(closeFieldCodes(editor, () => '')).toBe(2)
    expect([paragraphText(0), paragraphText(1)]).toEqual(['a b', 'cd'])
    const [field] = collectFields(editor.state.doc)
    expect([field.instr, field.mark.attrs.dirty]).toEqual(['QUOTE "x"', true])
  })

  it('saves an open code as the field it spells and empty braces as nothing', () => {
    make([para(text('ab')), para(text('cd'))])
    editor.commands.setTextSelection(2)
    insertFieldChars(editor)
    editor.commands.insertContent('NUM')
    // formatting part of the code splits its text: it is still one field
    editor.commands.setTextSelection({ from: 3, to: 5 })
    editor.commands.setMark('bold')
    editor.commands.insertContentAt(6, 'PAGES')
    const first = runsOf(0)
    expect(first.filter((run) => run.instrField !== undefined)).toHaveLength(1)
    const saved = first.find((run) => run.instrField !== undefined)!
    expect([saved.instrField, saved.text, saved.fldDirty]).toEqual(['NUMPAGES', ' ', true])
    expect(first.map((run) => run.text).join('')).toBe('a b')
    editor.commands.setTextSelection(editor.state.doc.content.size - 2)
    insertFieldChars(editor)
    expect(runsOf(1).some((run) => run.instrField !== undefined)).toBe(false)
    expect(
      runsOf(1)
        .map((run) => run.text)
        .join(''),
    ).toBe('cd')
  })

  it('saves a code that ends its paragraph', () => {
    make([para(text('ab'))])
    editor.commands.setTextSelection(3)
    insertFieldChars(editor)
    editor.commands.insertContent('TIME')
    editor.commands.deleteRange({ from: 8, to: 9 })
    expect(paragraphText()).toBe('ab TIME')
    expect(runsOf(0).at(-1)?.instrField).toBe('TIME')
  })

  it('is not shown a second time by Toggle Field Codes', () => {
    make([para(text('ab'))])
    editor.commands.setTextSelection(2)
    insertFieldChars(editor)
    toggleAllFieldCodes(editor.view)
    expect(editor.view.dom.querySelectorAll('.doc-field-code')).toHaveLength(1)
    expect(editor.view.dom.querySelector('.doc-field-code-open')).not.toBeNull()
  })
})

describe('Outline view levels', () => {
  const doc = [
    heading(1, 'A'),
    heading(2, 'A1'),
    heading(3, 'A1a'),
    para(text('body')),
    heading(1, 'B'),
    heading(3, 'B-deep'),
    heading(1, 'C'),
  ]
  const levels = (titles: string[]) => titles.map((t) => t.slice(2))

  it('folds what lies below the level and nothing above it', () => {
    const refs = [
      { level: 1, text: 'A' },
      { level: 2, text: 'A1' },
      { level: 3, text: 'A1a' },
      { level: 1, text: 'B' },
      { level: 3, text: 'B-deep' },
      { level: 1, text: 'C' },
    ]
    expect(levels(foldsForLevel(refs, 1))).toEqual(['A', 'B'])
    expect(levels(foldsForLevel(refs, 2))).toEqual(['A1', 'B'])
    expect(foldsForLevel(refs, 3)).toEqual([])
  })

  it('shows a level or every heading in Outline view only', () => {
    make(doc)
    expect(showOutlineLevel(editor.view, 1)).toBe(false)
    editor.commands.setOutlineViewActive(true)
    const folded = () => outlineViewKey.getState(editor.state)!.collapsed.size
    const run = (payload: string) =>
      MENU_COMMAND_HANDLERS['outline-show']({ editor, hasDoc: true } as CommandHost, payload)
    run('1')
    expect(folded()).toBe(2)
    run('2')
    expect(folded()).toBe(2)
    run('all')
    expect(folded()).toBe(0)
    run('0')
    run('x')
    expect(folded()).toBe(0)
  })
})

describe('Browse, sections, windows and the split pane', () => {
  it('steps through the last Go To target, and through the Find after a Find', () => {
    make([para(text('one')), para(text('two')), para(text('three'))])
    const view = { switchWindow: vi.fn(), closePane: vi.fn() }
    const host = {
      editor,
      hasDoc: true,
      findStep: vi.fn(),
      pageLayout: () => null,
      view,
    } as unknown as CommandHost
    expect(currentBrowseObject()).toBe('page')
    // no layout yet: a page move has nowhere to go and the caret stays
    editor.commands.setTextSelection(1)
    MENU_COMMAND_HANDLERS.browse(host, 'next')
    MENU_COMMAND_HANDLERS['section-go'](host, 'next')
    expect(editor.state.selection.from).toBe(1)
    setBrowseObject('find')
    MENU_COMMAND_HANDLERS.browse(host, 'prev')
    expect(host.findStep).toHaveBeenLastCalledWith(-1)
    MENU_COMMAND_HANDLERS.browse(host, 'sideways')
    expect(host.findStep).toHaveBeenCalledTimes(1)
    MENU_COMMAND_HANDLERS['window-go'](host, 'next')
    MENU_COMMAND_HANDLERS['window-go'](host, 'prev')
    expect(view.switchWindow.mock.calls).toEqual([[1], [-1]])
    MENU_COMMAND_HANDLERS['close-pane'](host, undefined)
    expect(view.closePane).toHaveBeenCalledTimes(1)
    MENU_COMMAND_HANDLERS['window-go']({ ...host, hasDoc: false } as CommandHost, 'next')
    expect(view.switchWindow).toHaveBeenCalledTimes(2)
  })

  it('finds the next and previous structural target from the caret', () => {
    make([
      para(text('x', [{ type: 'instrField', attrs: { instr: 'DATE' } }])),
      para(text('plain')),
      para(text('y', [{ type: 'instrField', attrs: { instr: 'TIME' } }])),
    ])
    editor.commands.setTextSelection(5)
    const next = goToPosition(editor.view, null, 'field', 1)
    const prev = goToPosition(editor.view, null, 'field', -1)
    expect(next).toBeGreaterThan(5)
    expect(prev).toBeLessThan(5)
    expect(goToPosition(editor.view, null, 'page', 1)).toBe(null)
    expect(goToPosition(editor.view, null, 'bookmark', 1, 'missing')).toBe(null)
  })

  it('goes round the windows from the focused one', () => {
    const tabs = ['a', 'b', 'c'].map((id) => ({ id, title: id, focused: id === 'c' }))
    expect(neighbourWindow(tabs, 1)).toBe('a')
    expect(neighbourWindow(tabs, -1)).toBe('b')
    expect(neighbourWindow(tabs.slice(2), 1)).toBe(null)
    expect(neighbourWindow(tabs.slice(0, 2), 1)).toBe(null)
  })

  it('reads the window moves as caret moves', () => {
    expect(parseCaretMove('screen-down')).toEqual({ to: 'screen-down', extend: false })
    expect(parseCaretMove('window-start-extend')).toEqual({ to: 'window-start', extend: true })
  })
})
