/**
 * The editing commands behind chords of Word for Mac that had none: the
 * selection mode of F8, field locking and unlinking, outline levels, the ends
 * of a table row or column, the Spike and Go Back.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { CellSelection } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  extendModeKey,
  extendModeOn,
  growSelection,
  leaveExtendMode,
  shrinkSelection,
  unitRange,
} from '../src/renderer/editor/extend-selection'
import {
  cutToSpike,
  forgetEditPlaces,
  goBack,
  goToCell,
  goToField,
  insertSpike,
  parseCellMove,
  resultFields,
  setFieldsLocked,
  shiftOutlineLevel,
  spikeSize,
  unlinkFields,
} from '../src/renderer/editor/word-keys'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { refreshPageFields, renumberSeqFields } from '../src/renderer/editor/field-results'
import { refreshTableFormulas } from '../src/renderer/editor/table-formula'
import { foldOutlineAtCaret, outlineViewKey } from '../src/renderer/editor/outline-view'
import { MENU_COMMAND_HANDLERS, type CommandHost } from '../src/renderer/commands/menu-commands'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'

type Json = Record<string, unknown>
let editor: Editor

const para = (...content: Json[]): Json => ({ type: 'docParagraph', content })
const text = (value: string, marks?: Json[]): Json => ({
  type: 'text',
  text: value,
  ...(marks ? { marks } : {}),
})
const field = (value: string, instr: string, attrs: Json = {}): Json =>
  text(value, [{ type: 'instrField', attrs: { instr, ...attrs } }])
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
  leaveExtendMode()
  editor?.destroy()
})

/** the position of `needle` in the document's text */
function posOf(needle: string, offset = 0): number {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    const at = node.isText ? node.text!.indexOf(needle) : -1
    if (at >= 0 && found < 0) found = pos + at
  })
  return found + offset
}

const selected = (): string => {
  const { from, to } = editor.state.selection
  return editor.state.doc.textBetween(from, to, '|')
}

const key = (name: string, init: KeyboardEventInit = {}): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { key: name, ...init })
  Object.defineProperty(event, 'target', { value: editor.view.dom })
  return event
}

describe('extend selection (F8)', () => {
  it('turns on, then takes the word, the sentence, the paragraph and the document', () => {
    make([para(text('One two three. Four five six.')), para(text('Second paragraph.'))])
    editor.commands.setTextSelection(posOf('two', 1))
    growSelection(editor)
    expect(extendModeOn(editor)).toBe(true)
    expect(selected()).toBe('')
    growSelection(editor)
    expect(selected()).toBe('two')
    growSelection(editor)
    expect(selected()).toBe('One two three. ')
    growSelection(editor)
    expect(selected()).toBe('One two three. Four five six.')
    growSelection(editor)
    expect(selected()).toBe('One two three. Four five six.|Second paragraph.')
    growSelection(editor)
    expect(selected()).toBe('One two three. Four five six.|Second paragraph.')
  })

  it('shrinks back unit by unit to the caret it began at and stays on', () => {
    make([para(text('One two three. Four five six.'))])
    const origin = posOf('five', 2)
    editor.commands.setTextSelection(origin)
    for (let i = 0; i < 4; i++) growSelection(editor)
    expect(selected()).toBe('One two three. Four five six.')
    shrinkSelection(editor)
    expect(selected()).toBe('Four five six.')
    shrinkSelection(editor)
    expect(selected()).toBe('five')
    shrinkSelection(editor)
    expect(editor.state.selection.empty).toBe(true)
    expect(editor.state.selection.head).toBe(origin)
    shrinkSelection(editor)
    expect(extendModeOn(editor)).toBe(true)
  })

  it('does nothing on Shift+F8 outside the mode', () => {
    make([para(text('One two'))])
    editor.commands.setTextSelection({ from: 1, to: 4 })
    shrinkSelection(editor)
    expect(selected()).toBe('One')
    expect(extendModeOn()).toBe(false)
  })

  it('leaves on Escape with the selection kept, and on an edit', () => {
    make([para(text('One two three'))])
    editor.commands.setTextSelection(posOf('two'))
    growSelection(editor)
    growSelection(editor)
    expect(extendModeKey(key('Escape'), null, () => undefined)).toBe(true)
    expect(extendModeOn()).toBe(false)
    expect(selected()).toBe('two')
    growSelection(editor)
    editor.commands.insertContent('x')
    expect(extendModeOn()).toBe(false)
  })

  it('extends by the caret move a chord runs and to the next typed character', () => {
    make([para(text('One two three. Four'))])
    editor.commands.setTextSelection(posOf('two'))
    growSelection(editor)
    const moves: string[] = []
    expect(
      extendModeKey(key('ArrowDown', { metaKey: true }), 'para-down', (p) => moves.push(p)),
    ).toBe(true)
    expect(
      extendModeKey(key('End', { shiftKey: true }), 'line-end-extend', (p) => moves.push(p)),
    ).toBe(true)
    expect(moves).toEqual(['para-down-extend', 'line-end-extend'])
    expect(extendModeKey(key('.'), null, () => undefined)).toBe(true)
    expect(selected()).toBe('two three.')
    // a character the rest of the document does not have leaves the selection alone
    expect(extendModeKey(key('#'), null, () => undefined)).toBe(true)
    expect(selected()).toBe('two three.')
    // an editing key is the document's
    expect(extendModeKey(key('Backspace'), null, () => undefined)).toBe(false)
    expect(extendModeKey(key('b', { metaKey: true }), null, () => undefined)).toBe(false)
  })

  it('takes no key outside the mode or outside its editor', () => {
    make([para(text('One'))])
    expect(extendModeKey(key('Escape'), null, () => undefined)).toBe(false)
    growSelection(editor)
    const elsewhere = new KeyboardEvent('keydown', { key: 'Escape' })
    Object.defineProperty(elsewhere, 'target', { value: document.body })
    expect(extendModeKey(elsewhere, null, () => undefined)).toBe(false)
    expect(extendModeOn()).toBe(true)
  })

  it('reads a unit in a table cell and the whole story as the document', () => {
    make([para(text('Before'))])
    insertTableAt(editor, 2, 2)
    editor.commands.insertContent('cell words here')
    const at = posOf('words', 1)
    expect(unitRange(editor.state.doc, at, 1)).toEqual([posOf('words'), posOf('words') + 5])
    expect(unitRange(editor.state.doc, at, 3)).toEqual([posOf('cell'), posOf('here') + 4])
  })
})

describe('fields', () => {
  const doc = (): Json[] => [
    para(text('Date: '), field('1 May', 'DATE'), text(' and file '), field('a.docx', 'FILENAME')),
    para(
      text('See '),
      text('Figure 1', [{ type: 'refField', attrs: { name: '_Ref1', instr: ' REF _Ref1 \\h ' } }]),
      text(' end'),
    ),
  ]
  const locks = () => resultFields(editor.state.doc).map((f) => f.mark.attrs.locked === true)

  it('locks and unlocks the fields the selection touches, in one undo step each', () => {
    make(doc())
    editor.commands.setTextSelection(posOf('1 May', 2))
    expect(setFieldsLocked(editor, true)).toBe(1)
    expect(locks()).toEqual([true, false, false])
    expect(setFieldsLocked(editor, true)).toBe(0)
    editor.commands.selectAll()
    expect(setFieldsLocked(editor, true)).toBe(2)
    expect(locks()).toEqual([true, true, true])
    editor.commands.undo()
    expect(locks()).toEqual([true, false, false])
    editor.commands.setTextSelection(posOf('1 May'))
    expect(setFieldsLocked(editor, false)).toBe(1)
    expect(locks()).toEqual([false, false, false])
    // the result and the instruction are untouched
    expect(resultFields(editor.state.doc).map((f) => f.instr)).toEqual([
      'DATE',
      'FILENAME',
      ' REF _Ref1 \\h ',
    ])
    expect(editor.state.doc.textContent).toBe('Date: 1 May and file a.docxSee Figure 1 end')
  })

  it('writes the lock to the saved runs', () => {
    make(doc())
    editor.commands.selectAll()
    setFieldsLocked(editor, true)
    const runs = (editor.getJSON().content ?? []).flatMap((block) =>
      inlineToRuns((block.content ?? []) as never),
    )
    expect(runs.filter((r) => r.fldLock).map((r) => r.text)).toEqual([
      '1 May',
      'a.docx',
      'Figure 1',
    ])
    // and reads it back from them
    const marks = runsToInline(runs).flatMap((node) => node.marks ?? [])
    expect(marks.filter((m) => m.attrs?.locked === true)).toHaveLength(3)
  })

  it('unlinks a field into its result and undoes it', () => {
    make(doc())
    editor.commands.setTextSelection(posOf('a.docx', 3))
    expect(unlinkFields(editor)).toBe(1)
    expect(resultFields(editor.state.doc).map((f) => f.instr)).toEqual(['DATE', ' REF _Ref1 \\h '])
    expect(editor.state.doc.textContent).toBe('Date: 1 May and file a.docxSee Figure 1 end')
    editor.commands.undo()
    expect(resultFields(editor.state.doc)).toHaveLength(3)
    editor.commands.selectAll()
    expect(unlinkFields(editor)).toBe(3)
    expect(resultFields(editor.state.doc)).toHaveLength(0)
    expect(unlinkFields(editor)).toBe(0)
  })

  it('leaves a form field and a field part alone', () => {
    make([
      para(
        field('☐', 'FORMCHECKBOX', { beginXml: '<w:r/>' }),
        text(' '),
        field('cite', 'ADDIN ZOTERO_ITEM', { fieldId: 1, fieldPart: 'begin' }),
      ),
    ])
    editor.commands.selectAll()
    expect(setFieldsLocked(editor, true)).toBe(0)
    expect(unlinkFields(editor)).toBe(0)
  })

  it('changes nothing on a read-only document', () => {
    make(doc())
    editor.commands.selectAll()
    editor.setEditable(false)
    expect(setFieldsLocked(editor, true)).toBe(0)
    expect(unlinkFields(editor)).toBe(0)
    expect(resultFields(editor.state.doc)).toHaveLength(3)
  })

  it('goes to the next and the previous field and stops at the ends', () => {
    make(doc())
    editor.commands.setTextSelection(1)
    expect(goToField(editor, 1)).toBe(true)
    expect(selected()).toBe('1 May')
    expect(goToField(editor, 1)).toBe(true)
    expect(selected()).toBe('a.docx')
    expect(goToField(editor, 1)).toBe(true)
    expect(selected()).toBe('Figure 1')
    expect(goToField(editor, 1)).toBe(false)
    expect(goToField(editor, -1)).toBe(true)
    expect(selected()).toBe('a.docx')
    expect(goToField(editor, -1)).toBe(true)
    expect(goToField(editor, -1)).toBe(false)
    expect(selected()).toBe('1 May')
  })

  it('keeps a locked page field and caption number as they are when fields are refreshed', () => {
    make([
      para(field('1', 'PAGE', { locked: true }), text(' '), field('1', 'PAGE')),
      para(field('9', 'SEQ Figure', { locked: true }), text(' '), field('9', 'SEQ Figure')),
    ])
    const page = { index: 4, number: 5, display: '5', chaptered: false, sectionPages: 9 }
    const tr = editor.state.tr
    refreshPageFields(tr, { pageAt: () => page, pageCount: 9, edited: true }, [], true)
    renumberSeqFields(tr, [])
    editor.view.dispatch(tr)
    expect(editor.state.doc.textContent).toBe('1 59 2')
    expect(resultFields(editor.state.doc).map((f) => f.mark.attrs.locked)).toEqual([
      true,
      false,
      true,
      false,
    ])
  })

  it('keeps a locked formula as it is when formulas are recomputed', () => {
    make([para(field('1', '= 2 + 3', { locked: true }), text(' '), field('1', '= 2 + 4'))])
    const tr = editor.state.tr
    refreshTableFormulas(tr)
    editor.view.dispatch(tr)
    expect(editor.state.doc.textContent).toBe('1 6')
  })
})

describe('outline levels', () => {
  const kinds = () => {
    const out: string[] = []
    editor.state.doc.forEach((node) =>
      out.push(node.type.name === 'docHeading' ? `h${node.attrs.level}` : 'p'),
    )
    return out
  }

  it('promotes and demotes a heading one level, within 1 to 9', () => {
    make([heading(1, 'Top'), heading(2, 'Sub'), para(text('Body'))])
    editor.commands.setTextSelection(posOf('Sub'))
    expect(shiftOutlineLevel(editor, 1)).toBe(true)
    expect(kinds()).toEqual(['h1', 'h3', 'p'])
    expect(shiftOutlineLevel(editor, -1)).toBe(true)
    expect(shiftOutlineLevel(editor, -1)).toBe(true)
    expect(kinds()).toEqual(['h1', 'h1', 'p'])
    expect(shiftOutlineLevel(editor, -1)).toBe(false)
    editor.commands.undo()
    expect(kinds()).toEqual(['h1', 'h2', 'p'])
    for (let i = 0; i < 12; i++) shiftOutlineLevel(editor, 1)
    expect(kinds()).toEqual(['h1', 'h9', 'p'])
  })

  it('makes body text a heading of the level before it, or one below', () => {
    make([heading(2, 'Sub'), para(text('Body')), para(text('More'))])
    editor.commands.setTextSelection(posOf('Body'))
    expect(shiftOutlineLevel(editor, -1)).toBe(true)
    expect(kinds()).toEqual(['h2', 'h2', 'p'])
    editor.commands.setTextSelection(posOf('More'))
    expect(shiftOutlineLevel(editor, 1)).toBe(true)
    expect(kinds()).toEqual(['h2', 'h2', 'h3'])
    expect(editor.state.doc.textContent).toBe('SubBodyMore')
  })

  it('starts at Heading 1 with no heading before, and moves every selected paragraph', () => {
    make([para(text('First')), para(text('Second'))])
    editor.commands.selectAll()
    expect(shiftOutlineLevel(editor, -1)).toBe(true)
    expect(kinds()).toEqual(['h1', 'h1'])
    expect(shiftOutlineLevel(editor, 1)).toBe(true)
    expect(kinds()).toEqual(['h2', 'h2'])
  })

  it("gives body text under a heading inside the selection that heading's level", () => {
    make([heading(1, 'Top'), para(text('A')), heading(3, 'Deep'), para(text('B'))])
    editor.commands.selectAll()
    expect(shiftOutlineLevel(editor, 1)).toBe(true)
    expect(kinds()).toEqual(['h2', 'h3', 'h4', 'h5'])
  })

  it('keeps a style that confers no level when the document has no heading style for the new one', () => {
    make([{ type: 'docHeading', attrs: { level: 1, styleId: 'Fancy' }, content: [text('Top')] }])
    const styles = new Map([
      ['Fancy', { styleId: 'Fancy', name: 'Fancy', type: 'paragraph' }],
    ]) as never
    editor.commands.setTextSelection(posOf('Top'))
    shiftOutlineLevel(editor, 1, styles)
    expect(editor.state.doc.firstChild!.attrs).toMatchObject({ level: 2, styleId: 'Fancy' })
  })

  it('takes the heading style of the new level from the document', () => {
    make([heading(1, 'Top')])
    const styles = new Map([
      ['Titre1', { styleId: 'Titre1', name: 'heading 1', type: 'paragraph', headingLevel: 1 }],
      ['Titre2', { styleId: 'Titre2', name: 'heading 2', type: 'paragraph', headingLevel: 2 }],
    ]) as never
    editor.commands.setTextSelection(posOf('Top'))
    shiftOutlineLevel(editor, 1, styles)
    expect(editor.state.doc.firstChild!.attrs).toMatchObject({ level: 2, styleId: 'Titre2' })
  })

  it('folds and unfolds the heading at the caret in Outline view only', () => {
    make([heading(1, 'Top'), para(text('Body')), heading(2, 'Sub'), para(text('Deep'))])
    editor.commands.setTextSelection(posOf('Body'))
    const collapsed = () => outlineViewKey.getState(editor.state)!.collapsed.size
    expect(foldOutlineAtCaret(editor.view, true)).toBe(false)
    editor.commands.setOutlineViewActive(true)
    expect(foldOutlineAtCaret(editor.view, false)).toBe(false)
    expect(foldOutlineAtCaret(editor.view, true)).toBe(true)
    expect(collapsed()).toBe(1)
    expect(foldOutlineAtCaret(editor.view, true)).toBe(false)
    expect(foldOutlineAtCaret(editor.view, false)).toBe(true)
    expect(collapsed()).toBe(0)
    // a heading with nothing under it has nothing to fold
    editor.commands.setTextSelection(posOf('Deep'))
    expect(foldOutlineAtCaret(editor.view, true)).toBe(false)
  })
})

describe('table row and column ends', () => {
  function table(): void {
    make([para(text('Before'))])
    insertTableAt(editor, 3, 3)
    const cells: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.spec.tableRole === 'cell') cells.push(pos)
    })
    const tr = editor.state.tr
    cells.reverse().forEach((pos, i) => tr.insertText(`c${cells.length - 1 - i}`, pos + 2))
    editor.view.dispatch(tr)
  }
  const cellText = () => editor.state.selection.$head.node(-1).textContent

  it('reads a move and its extending twin', () => {
    expect(parseCellMove('row-end')).toEqual({ to: 'row-end', extend: false })
    expect(parseCellMove('column-start-extend')).toEqual({ to: 'column-start', extend: true })
    expect(parseCellMove('diagonal')).toBe(null)
  })

  it('puts the caret in the first or last cell of the row or column', () => {
    table()
    editor.commands.setTextSelection(posOf('c4', 1))
    expect(goToCell(editor, 'row-start', false)).toBe(true)
    expect(cellText()).toBe('c3')
    expect(editor.state.selection.empty).toBe(true)
    goToCell(editor, 'row-end', false)
    expect(cellText()).toBe('c5')
    goToCell(editor, 'column-start', false)
    expect(cellText()).toBe('c2')
    goToCell(editor, 'column-end', false)
    expect(cellText()).toBe('c8')
  })

  it('selects the cells up to the end of the row or column', () => {
    table()
    editor.commands.setTextSelection(posOf('c4', 1))
    goToCell(editor, 'row-end', true)
    const cells = editor.state.selection as CellSelection
    expect(cells).toBeInstanceOf(CellSelection)
    const texts: string[] = []
    cells.forEachCell((node) => texts.push(node.textContent))
    expect(texts).toEqual(['c4', 'c5'])
    goToCell(editor, 'column-end', true)
    const more: string[] = []
    ;(editor.state.selection as CellSelection).forEachCell((node) => more.push(node.textContent))
    expect(more).toEqual(['c4', 'c5', 'c7', 'c8'])
  })

  it('does nothing outside a table', () => {
    table()
    editor.commands.setTextSelection(posOf('Before', 2))
    expect(goToCell(editor, 'row-end', false)).toBe(false)
    expect(editor.state.selection.head).toBe(posOf('Before', 2))
  })
})

describe('the Spike', () => {
  it('collects cuts in order and inserts them as paragraphs, emptying itself', () => {
    make([para(text('alpha beta gamma')), para(text('end'))])
    editor.commands.setTextSelection({ from: posOf('alpha'), to: posOf('alpha') + 6 })
    expect(cutToSpike(editor)).toBe(true)
    editor.commands.setTextSelection({ from: posOf('gamma'), to: posOf('gamma') + 5 })
    expect(cutToSpike(editor)).toBe(true)
    expect(spikeSize()).toBe(2)
    expect(editor.state.doc.firstChild!.textContent).toBe('beta ')
    editor.commands.setTextSelection(posOf('end') + 3)
    expect(insertSpike(editor)).toBe(true)
    const paragraphs: string[] = []
    editor.state.doc.forEach((node) => paragraphs.push(node.textContent))
    expect(paragraphs).toEqual(['beta ', 'endalpha ', 'gamma'])
    expect(spikeSize()).toBe(0)
    expect(insertSpike(editor)).toBe(false)
    // one undo step takes the whole insertion back
    editor.commands.undo()
    expect(editor.state.doc.childCount).toBe(2)
  })

  it('cuts nothing from a caret or a read-only document', () => {
    make([para(text('alpha'))])
    editor.commands.setTextSelection(2)
    expect(cutToSpike(editor)).toBe(false)
    editor.commands.selectAll()
    editor.setEditable(false)
    expect(cutToSpike(editor)).toBe(false)
    expect(spikeSize()).toBe(0)
  })

  it('keeps the formatting of what was cut', () => {
    make([para(text('plain '), text('bold', [{ type: 'bold' }])), para(text('x'))])
    editor.commands.setTextSelection({ from: posOf('bold'), to: posOf('bold') + 4 })
    cutToSpike(editor)
    editor.commands.setTextSelection(posOf('x'))
    insertSpike(editor)
    expect(editor.state.doc.nodeAt(posOf('bold'))!.marks.map((m) => m.type.name)).toEqual(['bold'])
  })
})

describe('Go Back', () => {
  const paragraphOfCaret = () => editor.state.selection.$head.index(0)

  it('returns to the last edits, the latest first, then to where it set out', () => {
    make(['one', 'two', 'three', 'four', 'five'].map((t) => para(text(t))))
    expect(goBack(editor)).toBe(false)
    for (const word of ['one', 'three', 'four']) {
      editor.commands.setTextSelection(posOf(word) + word.length)
      editor.commands.insertContent('!')
    }
    editor.commands.setTextSelection(posOf('five'))
    const visited: number[] = []
    for (let i = 0; i < 5; i++) {
      expect(goBack(editor)).toBe(true)
      visited.push(paragraphOfCaret())
    }
    expect(visited).toEqual([3, 2, 0, 4, 3])
  })

  it('counts typing on in one paragraph as one place and follows later edits', () => {
    make(['one', 'two', 'three'].map((t) => para(text(t))))
    editor.commands.setTextSelection(posOf('three') + 5)
    editor.commands.insertContent('a')
    editor.commands.insertContent('b')
    editor.commands.setTextSelection(posOf('one'))
    editor.commands.insertContent('zzzz')
    editor.commands.setTextSelection(posOf('two'))
    expect(goBack(editor)).toBe(true)
    expect(paragraphOfCaret()).toBe(0)
    expect(goBack(editor)).toBe(true)
    expect(paragraphOfCaret()).toBe(2)
    expect(editor.state.selection.head).toBe(posOf('threeab') + 7)
    expect(goBack(editor)).toBe(true)
    expect(paragraphOfCaret()).toBe(1)
  })

  it('has nowhere to go in a document just loaded', () => {
    make(['one', 'two'].map((t) => para(text(t))))
    editor.commands.setTextSelection(posOf('one') + 3)
    editor.commands.insertContent('!')
    forgetEditPlaces(editor)
    editor.commands.setTextSelection(posOf('two'))
    expect(goBack(editor)).toBe(false)
    expect(editor.can().undo()).toBe(true)
  })

  it('goes to the edit before when the caret is still at the last one', () => {
    make(['one', 'two'].map((t) => para(text(t))))
    editor.commands.setTextSelection(posOf('one') + 3)
    editor.commands.insertContent('!')
    editor.commands.setTextSelection(posOf('two') + 3)
    editor.commands.insertContent('!')
    expect(goBack(editor)).toBe(true)
    expect(paragraphOfCaret()).toBe(0)
  })
})

describe('the commands', () => {
  const host = (over: Partial<CommandHost> = {}): CommandHost =>
    ({ editor, hasDoc: true, ...over }) as CommandHost

  it('run on the document, and not without one', () => {
    make([heading(1, 'Top'), para(text('Date: '), field('1 May', 'DATE'))])
    editor.commands.setTextSelection(posOf('1 May', 1))
    MENU_COMMAND_HANDLERS['field-lock'](host({ hasDoc: false }), 'lock')
    expect(resultFields(editor.state.doc)[0].mark.attrs.locked).toBe(false)
    MENU_COMMAND_HANDLERS['field-lock'](host(), 'lock')
    expect(resultFields(editor.state.doc)[0].mark.attrs.locked).toBe(true)
    MENU_COMMAND_HANDLERS['field-lock'](host(), 'unlock')
    expect(resultFields(editor.state.doc)[0].mark.attrs.locked).toBe(false)
    MENU_COMMAND_HANDLERS['field-go'](host(), 'prev')
    MENU_COMMAND_HANDLERS['field-go'](host(), 'next')
    expect(selected()).toBe('1 May')
    MENU_COMMAND_HANDLERS['field-unlink'](host())
    expect(resultFields(editor.state.doc)).toHaveLength(0)
    MENU_COMMAND_HANDLERS.spike(host(), 'cut')
    expect(editor.state.doc.textContent).toBe('TopDate: ')
    MENU_COMMAND_HANDLERS.spike(host(), 'insert')
    expect(editor.state.doc.textContent).toBe('TopDate: 1 May')
    editor.commands.setTextSelection(posOf('Top'))
    MENU_COMMAND_HANDLERS['outline-level'](host(), 'demote')
    expect(editor.state.doc.firstChild!.attrs.level).toBe(2)
    MENU_COMMAND_HANDLERS['outline-level'](host(), 'sideways')
    expect(editor.state.doc.firstChild!.attrs.level).toBe(2)
    MENU_COMMAND_HANDLERS['go-back'](host())
    expect(editor.state.selection.$head.index(0)).toBe(1)
    MENU_COMMAND_HANDLERS['extend-selection'](host(), 'grow')
    expect(extendModeOn(editor)).toBe(true)
    MENU_COMMAND_HANDLERS['table-cell-go'](host(), 'row-end')
    expect(editor.state.selection.$head.index(0)).toBe(1)
  })
})
