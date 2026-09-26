import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx } from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { AUTOCORRECT_KEY, setAutocorrectPref } from '../src/renderer/autocorrect-pref'
import { autocorrectKey, parseListPrefix } from '../src/renderer/editor/autocorrect'
import { pruneUnreferencedNumbering } from '../src/renderer/numbering-actions'
import { EMPTY_PENDING_NUMBERING } from '../src/renderer/doc-state'
import type { PmNode } from '../src/renderer/editor/convert'

const editors: Editor[] = []

async function createEditor() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.view.dom.setAttribute('contenteditable', 'true')
  editor.storage.autocorrect.hooks = {
    bulletNumId: () => 'B9',
    numberedNumId: (numFmt, pattern, start) => `N:${numFmt}:${pattern}:${start}`,
  }
  editor.commands.focus('end')
  editors.push(editor)
  return editor
}

/** feed characters the way the browser does: through handleTextInput, plain insert when no rule fires */
function type(editor: Editor, text: string) {
  for (const ch of text) {
    const { view } = editor
    const { from, to } = view.state.selection
    const deflt = () => view.state.tr.insertText(ch, from, to)
    const handled = view.someProp('handleTextInput', (f) => f(view, from, to, ch, deflt))
    if (!handled) view.dispatch(deflt())
  }
}

function backspace(editor: Editor) {
  const { view } = editor
  const event = new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true })
  return view.someProp('handleKeyDown', (f) => f(view, event))
}

const firstBlock = (editor: Editor) => editor.state.doc.firstChild!
const text = (editor: Editor) => editor.state.doc.textContent

beforeEach(() => localStorage.removeItem(AUTOCORRECT_KEY))
// destroying the view clears ProseMirror's DOMObserver timer, which otherwise fires after jsdom is gone
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

describe('autocorrect: smart quotes', () => {
  it('opens after a space or paragraph start and closes after a word', async () => {
    const editor = await createEditor()
    type(editor, "\"hi\" it's 'x'")
    expect(text(editor)).toBe('\u201chi\u201d it\u2019s \u2018x\u2019')
  })
  it('leaves CJK full-width quotes and text alone', async () => {
    const editor = await createEditor()
    type(editor, '\u4ed6\u8bf4\u201c\u4f60\u597d\u201d\u3002')
    expect(text(editor)).toBe('\u4ed6\u8bf4\u201c\u4f60\u597d\u201d\u3002')
  })
  it('is off when the switch is off', async () => {
    setAutocorrectPref('smartQuotes', false)
    const editor = await createEditor()
    type(editor, '"hi"')
    expect(text(editor)).toBe('"hi"')
  })
})

describe('autocorrect: dashes', () => {
  it('turns word--word into an em dash and word - word into an en dash on the following space', async () => {
    const editor = await createEditor()
    type(editor, 'red--blue ')
    expect(text(editor)).toBe('red\u2014blue ')
    type(editor, 'a - b ')
    expect(text(editor)).toBe('red\u2014blue a \u2013 b ')
  })
  it('does not touch a bare hyphenated word or a leading dash', async () => {
    const editor = await createEditor()
    type(editor, 'well-known -- ')
    expect(text(editor)).toBe('well-known -- ')
  })
})

describe('autocorrect: automatic lists', () => {
  it('"- " at the start of an empty paragraph becomes a bullet item without the prefix', async () => {
    const editor = await createEditor()
    type(editor, '- ')
    const node = firstBlock(editor)
    expect(node.type.name).toBe('docListItem')
    expect(node.attrs).toMatchObject({ kind: 'bullet', numId: 'B9', ilvl: 0 })
    expect(node.textContent).toBe('')
    type(editor, 'milk')
    expect(firstBlock(editor).textContent).toBe('milk')
  })
  it('"1. " / "a) " / "(1) " / CJK numerals start numbered lists in the typed format', async () => {
    const cases: Array<[string, string]> = [
      ['1. ', 'N:decimal:%1.:1'],
      ['3) ', 'N:decimal:%1):3'],
      ['(1) ', 'N:decimal:(%1):1'],
      ['b. ', 'N:lowerLetter:%1.:2'],
      ['i. ', 'N:lowerRoman:%1.:1'],
      ['A) ', 'N:upperLetter:%1):1'],
      ['\u4e00\u3001', 'N:chineseCountingThousand:%1\u3001:1'],
    ]
    for (const [typed, numId] of cases) {
      const editor = await createEditor()
      // the CJK comma is the delimiter for the numeral form; a space finishes the others
      type(editor, typed.endsWith(' ') ? typed : typed + ' ')
      const node = firstBlock(editor)
      expect(node.type.name, typed).toBe('docListItem')
      expect(node.attrs.numId, typed).toBe(numId)
    }
  })
  it('does not fire inside an existing list item or after other text', async () => {
    const editor = await createEditor()
    type(editor, '- ')
    type(editor, '1. ')
    expect(firstBlock(editor).type.name).toBe('docListItem')
    expect(firstBlock(editor).textContent).toBe('1. ')
    const plain = await createEditor()
    type(plain, 'see 1. ')
    expect(firstBlock(plain).type.name).toBe('docParagraph')
  })
  it('parseListPrefix rejects non-prefixes', () => {
    expect(parseListPrefix('0.')).toBeNull()
    expect(parseListPrefix('12345.')).toBeNull()
    expect(parseListPrefix('ab.')).toBeNull()
    expect(parseListPrefix('-x')).toBeNull()
  })
})

describe('autocorrect: symbols and ordinals', () => {
  it('replaces (c) (r) (tm) ... --> <-- and fractions', async () => {
    const editor = await createEditor()
    type(editor, '(c) (r) (tm) wait... a-->b c<--d 1/2 ')
    expect(text(editor)).toBe('\u00a9 \u00ae \u2122 wait\u2026 a\u2192b c\u2190d \u00bd ')
  })
  it('superscripts a correct ordinal suffix only', async () => {
    const editor = await createEditor()
    type(editor, '1st 2th ')
    const marks: string[] = []
    editor.state.doc.descendants((n) => {
      if (n.isText) {
        const sup = n.marks.some((m) => m.attrs.vertAlign === 'superscript')
        marks.push(`${n.text}${sup ? '^' : ''}`)
      }
    })
    expect(marks).toEqual(['1', 'st^', ' 2th '])
  })
})

describe('autocorrect: sentence capitalization (opt-in)', () => {
  it('is off by default and capitalizes sentence starts when enabled', async () => {
    const off = await createEditor()
    type(off, 'hello world. next ')
    expect(text(off)).toBe('hello world. next ')

    setAutocorrectPref('capitalize', true)
    const on = await createEditor()
    type(on, 'hello world. next one? yes e.g. apples 3. items http://x.y/z ')
    expect(text(on)).toBe('Hello world. Next one? Yes e.g. apples 3. items http://x.y/z ')
  })
})

describe('autocorrect: Backspace reverts the last correction', () => {
  it('restores the typed straight quote and a converted list paragraph', async () => {
    const editor = await createEditor()
    type(editor, 'x "')
    expect(text(editor)).toBe('x \u201c')
    expect(backspace(editor)).toBe(true)
    expect(text(editor)).toBe('x "')
    // the record is consumed: the next Backspace is an ordinary delete
    expect(autocorrectKey.getState(editor.state)).toBeNull()

    const list = await createEditor()
    type(list, '1. ')
    expect(firstBlock(list).type.name).toBe('docListItem')
    expect(backspace(list)).toBe(true)
    expect(firstBlock(list).type.name).toBe('docParagraph')
    expect(text(list)).toBe('1. ')
    // the allocated definition stays pending; whether it is saved follows the document
    const pending = {
      ...EMPTY_PENDING_NUMBERING,
      newDefs: [{ numId: 'N:decimal:%1.:1', kind: 'ordered' as const }],
    }
    const json = () => list.getJSON() as PmNode
    expect(pruneUnreferencedNumbering(pending, json()).newDefs).toEqual([])
    list.commands.undo()
    expect(firstBlock(list).type.name).toBe('docListItem')
    expect(firstBlock(list).attrs.numId).toBe('N:decimal:%1.:1')
    expect(pruneUnreferencedNumbering(pending, json()).newDefs).toHaveLength(1)
  })
  it('does not revert once the caret moved on', async () => {
    const editor = await createEditor()
    type(editor, '"a')
    expect(autocorrectKey.getState(editor.state)).toBeNull()
    backspace(editor)
    expect(text(editor).startsWith('\u201c')).toBe(true)
  })
})

describe('pruneUnreferencedNumbering', () => {
  it('keeps definitions referenced anywhere in the document (nested too) and drops the rest', () => {
    const pending = {
      ...EMPTY_PENDING_NUMBERING,
      newDefs: [
        { numId: '7', kind: 'ordered' as const },
        { numId: '8', kind: 'bullet' as const },
      ],
      restartNums: [
        { numId: '9', abstractNumId: 'a', startOverrides: {} },
        { numId: '10', abstractNumId: 'a', startOverrides: {} },
      ],
    }
    const cell = { type: 'docTableCell', content: [{ type: 'docListItem', attrs: { numId: 9 } }] }
    const doc = {
      type: 'doc',
      content: [
        { type: 'docListItem', attrs: { numId: '7' } },
        { type: 'docTable', content: [{ type: 'docTableRow', content: [cell] }] },
        { type: 'docParagraph' },
      ],
    }
    const out = pruneUnreferencedNumbering(pending, doc)
    expect(out.newDefs.map((d) => d.numId)).toEqual(['7'])
    expect(out.restartNums.map((r) => r.numId)).toEqual(['9'])
  })

  it('multi-step undo and redo: the definition is saved exactly when the list item exists', async () => {
    const editor = await createEditor()
    type(editor, '- ')
    type(editor, 'milk')
    const pending = {
      ...EMPTY_PENDING_NUMBERING,
      newDefs: [{ numId: 'B9', kind: 'bullet' as const }],
    }
    const kept = () =>
      pruneUnreferencedNumbering(pending, editor.getJSON() as PmNode).newDefs.length
    expect(kept()).toBe(1)
    editor.commands.undo()
    editor.commands.undo()
    expect(firstBlock(editor).type.name).toBe('docParagraph')
    expect(kept()).toBe(0)
    editor.commands.redo()
    expect(firstBlock(editor).type.name).toBe('docListItem')
    expect(kept()).toBe(1)
  })
})
