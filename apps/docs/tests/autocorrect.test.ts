import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx, type StyleUpsert } from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import {
  AUTOCORRECT_KEY,
  DEFAULT_REPLACE_ENTRIES,
  autocorrectLists,
  autocorrectPrefs,
  resetAutocorrectPrefs,
  setAutocorrectList,
  setAutocorrectPref,
  upsertEntry,
} from '../src/renderer/autocorrect-pref'
import {
  autoCompleteKey,
  autoCompleteMatch,
  autocorrectKey,
  changeBackFix,
  fixesKey,
  matchReplaceEntry,
  matchingCloser,
  parseListPrefix,
  stopCorrecting,
} from '../src/renderer/editor/autocorrect'
import { isDateLine } from '../src/renderer/editor/autoformat-styles'
import { pruneUnreferencedNumbering } from '../src/renderer/numbering-actions'
import { SPELLCHECK_KEY } from '../src/renderer/spellcheck-pref'
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

function press(editor: Editor, key: string) {
  const { view } = editor
  const event = new KeyboardEvent('keydown', { key, bubbles: true })
  return view.someProp('handleKeyDown', (f) => f(view, event))
}

const enter = (editor: Editor) => press(editor, 'Enter')

/** Enter the way the editor takes it: the AutoFormat rules first, then the paragraph split */
function enterAndSplit(editor: Editor) {
  const { view } = editor
  const event = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })
  const plugin = autocorrectKey.get(view.state)!
  if (!plugin.props.handleKeyDown!.call(plugin, view, event)) editor.commands.splitBlock()
}

const paragraphs = (editor: Editor) => {
  const out: string[] = []
  editor.state.doc.forEach((node) => out.push(node.textContent))
  return out
}

function backspace(editor: Editor) {
  const { view } = editor
  const event = new KeyboardEvent('keydown', { key: 'Backspace', bubbles: true })
  return view.someProp('handleKeyDown', (f) => f(view, event))
}

const firstBlock = (editor: Editor) => editor.state.doc.firstChild!
const text = (editor: Editor) => editor.state.doc.textContent

beforeEach(() => resetAutocorrectPrefs())
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

describe('autocorrect: suggestions from the spelling checker', () => {
  const dictionary: Record<string, string[]> = {
    recieve: ['receive'],
    Recieve: ['Receive'],
    wierd: ['weird', 'wired'],
    NASSA: ['NASA'],
  }
  const desktop = {
    spellIsMisspelled: (word: string) => word in dictionary,
    spellSuggestions: (word: string) => dictionary[word] ?? [],
  }
  beforeEach(() => {
    ;(window as unknown as { desktop: unknown }).desktop = desktop
    localStorage.removeItem(SPELLCHECK_KEY)
  })
  afterEach(() => {
    delete (window as unknown as { desktop?: unknown }).desktop
  })

  it('is off until the option is switched on', async () => {
    const editor = await createEditor()
    type(editor, 'recieve ')
    expect(text(editor)).toBe('recieve ')
  })
  it('replaces a word the checker has exactly one suggestion for, and Backspace brings it back', async () => {
    setAutocorrectPref('spellSuggest', true)
    const editor = await createEditor()
    type(editor, 'I recieve ')
    expect(text(editor)).toBe('I receive ')
    expect(fixesKey.getState(editor.state)!.at(-1)).toMatchObject({
      original: 'recieve',
      corrected: 'receive',
      rule: 'spellSuggest',
    })
    expect(backspace(editor)).toBe(true)
    expect(text(editor)).toBe('I recieve ')
    type(editor, 'Recieve.')
    expect(text(editor)).toBe('I recieve Receive.')
  })
  it('leaves words with several suggestions, capitals, parts of a longer token and unchecked text', async () => {
    setAutocorrectPref('spellSuggest', true)
    const editor = await createEditor()
    type(editor, 'wierd NASSA 2recieve ')
    expect(text(editor)).toBe('wierd NASSA 2recieve ')
    localStorage.setItem(SPELLCHECK_KEY, '0')
    type(editor, 'recieve ')
    expect(text(editor)).toBe('wierd NASSA 2recieve recieve ')
  })
  it('the table of entries wins over the checker', async () => {
    setAutocorrectPref('spellSuggest', true)
    setAutocorrectList('replace', [['recieve', 'got']])
    const editor = await createEditor()
    type(editor, 'recieve ')
    expect(text(editor)).toBe('got ')
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

describe('autocorrect: the dialog switches', () => {
  it('the master switch stops every rule', async () => {
    setAutocorrectPref('enabled', false)
    const editor = await createEditor()
    type(editor, '"a" (c) - ')
    expect(text(editor)).toBe('"a" (c) - ')
    expect(firstBlock(editor).type.name).toBe('docParagraph')
  })
  it('switches saved before the tabs existed still apply to the rules they were split into', () => {
    localStorage.setItem(AUTOCORRECT_KEY, JSON.stringify({ autoLists: false, symbols: false }))
    expect(autocorrectPrefs()).toMatchObject({
      autoBullets: false,
      autoNumbers: false,
      replaceText: false,
      fractions: false,
      smartQuotes: true,
    })
  })
  it('bulleted and numbered lists answer to their own switch', async () => {
    setAutocorrectPref('autoBullets', false)
    const editor = await createEditor()
    type(editor, '- ')
    expect(firstBlock(editor).type.name).toBe('docParagraph')
    editor.commands.undo()
    editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
    editor.commands.focus('end')
    type(editor, '1. ')
    expect(firstBlock(editor).type.name).toBe('docListItem')
  })
})

describe('autocorrect: replace text as you type', () => {
  it('finds the longest entry, word entries only at a delimiter after a whole word', () => {
    const entries = upsertEntry(DEFAULT_REPLACE_ENTRIES, 'teh', 'the')
    expect(matchReplaceEntry(entries, '(c', ')')).toMatchObject({ to: '\u00a9', word: false })
    expect(matchReplaceEntry(entries, '(C', ')')).toMatchObject({ to: '\u00a9' })
    expect(matchReplaceEntry(entries, '...', '.')).toBeNull()
    expect(matchReplaceEntry(entries, 'a teh', ' ')).toMatchObject({ to: 'the', word: true })
    expect(matchReplaceEntry(entries, 'Teh', ',')).toMatchObject({ to: 'The' })
    expect(matchReplaceEntry(entries, 'ateh', ' ')).toBeNull()
    expect(matchReplaceEntry(entries, 'te', 'h')).toBeNull()
  })
  it('built-in arrows, and an entry added in the dialog', async () => {
    setAutocorrectList('replace', upsertEntry(DEFAULT_REPLACE_ENTRIES, 'teh', 'the'))
    const editor = await createEditor()
    type(editor, 'a ==> b <=> c, teh end')
    expect(text(editor)).toBe('a \u21d2 b \u21d4 c, the end')
  })
  it('a deleted entry no longer fires', async () => {
    setAutocorrectList(
      'replace',
      DEFAULT_REPLACE_ENTRIES.filter(([from]) => from !== '(c)'),
    )
    const editor = await createEditor()
    type(editor, '(c) (r)')
    expect(text(editor)).toBe('(c) \u00ae')
  })
  it('Backspace puts back what was typed', async () => {
    setAutocorrectList('replace', upsertEntry(DEFAULT_REPLACE_ENTRIES, 'teh', 'the'))
    const editor = await createEditor()
    type(editor, 'teh ')
    expect(text(editor)).toBe('the ')
    backspace(editor)
    expect(text(editor)).toBe('teh ')
  })
})

describe('autocorrect: capitals', () => {
  it('two initial capitals, except the listed words', async () => {
    const editor = await createEditor()
    type(editor, 'THe PCs and IDs of OAuth at 5 MHz, USA, HEllo. ')
    expect(text(editor)).toBe('The PCs and IDs of OAuth at 5 MHz, USA, Hello. ')
  })
  it('an exception added in the dialog is left alone', async () => {
    setAutocorrectList('initialCaps', ['MYsql'])
    const editor = await createEditor()
    type(editor, 'MYsql OAuth ')
    expect(text(editor)).toBe('MYsql Oauth ')
  })
  it('names of days', async () => {
    const editor = await createEditor()
    type(editor, 'on monday, not mondays ')
    expect(text(editor)).toBe('on Monday, not mondays ')
  })
  it('a first-letter exception added in the dialog does not end a sentence', async () => {
    setAutocorrectPref('capitalize', true)
    setAutocorrectList('firstLetter', ['viz.'])
    const editor = await createEditor()
    type(editor, 'see viz. apples. pears ')
    expect(text(editor)).toBe('See viz. apples. Pears ')
  })
  it('the first word of a table cell follows its own switch, not the sentence one', async () => {
    setAutocorrectPref('capitalize', true)
    const cell = { type: 'docTableCell', content: [{ type: 'docParagraph' }] }
    const table = { type: 'docTable', content: [{ type: 'docTableRow', content: [cell] }] }
    const editor = await createEditor()
    editor.commands.setContent({ type: 'doc', content: [table, { type: 'docParagraph' }] })
    editor.commands.setTextSelection(4)
    type(editor, 'total ')
    expect(text(editor)).toBe('total ')
    setAutocorrectPref('capitalizeCells', true)
    type(editor, '. ')
    editor.commands.setContent({ type: 'doc', content: [table, { type: 'docParagraph' }] })
    editor.commands.setTextSelection(4)
    type(editor, 'total ')
    expect(text(editor)).toBe('Total ')
  })
})

describe('autocorrect: math names outside equations', () => {
  it('is off until switched on, then replaces a whole token at the space', async () => {
    const editor = await createEditor()
    type(editor, '\\alpha ')
    expect(text(editor)).toBe('\\alpha ')
    setAutocorrectPref('mathOutside', true)
    type(editor, '\\beta x<=y a <= b ')
    expect(text(editor)).toBe('\\alpha \u03b2 x<=y a \u2264 b ')
  })
})

describe('autocorrect: *bold* and _italic_', () => {
  it('is opt-in; the markers go and the words take the format', async () => {
    const editor = await createEditor()
    type(editor, '*a* ')
    expect(text(editor)).toBe('*a* ')
    setAutocorrectPref('boldItalic', true)
    type(editor, '*very bold* and _slanted_ x')
    expect(text(editor)).toBe('*a* very bold and slanted x')
    const marksOf = (word: string) => {
      const at = text(editor).indexOf(word) + 1
      return editor.state.doc
        .resolve(at + 1)
        .marks()
        .map((m) => m.type.name)
    }
    expect(marksOf('very')).toContain('bold')
    expect(marksOf('slanted')).toContain('italic')
    expect(marksOf('and')).not.toContain('bold')
    expect(marksOf('x')).not.toContain('italic')
  })
})

describe('autocorrect: on Enter', () => {
  it('--- becomes a bottom border and the caret goes to a new paragraph below', async () => {
    const editor = await createEditor()
    type(editor, '---')
    expect(enter(editor)).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    expect(text(editor)).toBe('')
    expect(firstBlock(editor).attrs.borders).toBe('b')
    expect(JSON.parse(firstBlock(editor).attrs.borderLines as string)).toMatchObject({
      b: { val: 'single', szPt: 0.75 },
    })
    expect(editor.state.doc.child(1).attrs.borders).toBeNull()
    expect(editor.state.selection.$from.parent).toBe(editor.state.doc.child(1))
    backspace(editor)
    expect(editor.state.doc.childCount).toBe(1)
    expect(text(editor)).toBe('---')
    expect(firstBlock(editor).attrs.borders).toBeNull()
  })
  it('=== is a double line; two characters or mixed characters are text', async () => {
    const editor = await createEditor()
    type(editor, '====')
    enter(editor)
    expect(JSON.parse(firstBlock(editor).attrs.borderLines as string).b.val).toBe('double')
    editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
    editor.commands.focus('end')
    type(editor, '--')
    enter(editor)
    type(editor, '--=')
    enter(editor)
    expect(text(editor)).toBe('----=')
    editor.state.doc.forEach((node) => expect(node.attrs.borders).toBeNull())
  })
  it('is off when Border lines is off', async () => {
    setAutocorrectPref('borderLines', false)
    const editor = await createEditor()
    type(editor, '---')
    enter(editor)
    expect(text(editor)).toBe('---')
    expect(firstBlock(editor).attrs.borders).toBeNull()
  })
  it('+--+--+ becomes a table of that many columns, and one undo brings the line back', async () => {
    const { insertTableAt } = await import('../src/renderer/components/ribbon-tabs')
    const editor = await createEditor()
    editor.storage.autocorrect.hooks!.insertTable = (cols) => insertTableAt(editor, 1, cols, true)
    type(editor, '+--+---+-+')
    expect(enter(editor)).toBe(true)
    expect(text(editor)).toBe('')
    expect(firstBlock(editor).type.name).toBe('docTable')
    expect(firstBlock(editor).firstChild!.childCount).toBe(3)
    editor.commands.undo()
    expect(firstBlock(editor).type.name).toBe('docParagraph')
    expect(text(editor)).toBe('+--+---+-+')
  })
  it('is off when Tables is off', async () => {
    setAutocorrectPref('tables', false)
    const asked: number[] = []
    const editor = await createEditor()
    editor.storage.autocorrect.hooks!.insertTable = (cols) => asked.push(cols)
    type(editor, '+--+')
    enter(editor)
    expect(asked).toEqual([])
    expect(text(editor)).toBe('+--+')
  })
})

describe('autocorrect: East Asian typing', () => {
  it('a closing bracket takes the width of the one it closes', async () => {
    expect(matchingCloser('\uff08a', ')')).toBe('\uff09')
    expect(matchingCloser('(a', '\uff09')).toBe(')')
    expect(matchingCloser('(a', ')')).toBeNull()
    expect(matchingCloser('\uff08a(b)', ')')).toBe('\uff09')
    expect(matchingCloser('a', ')')).toBeNull()
    const editor = await createEditor()
    type(editor, '\uff08note) [x]')
    expect(text(editor)).toBe('\uff08note\uff09 [x]')
    setAutocorrectPref('parens', false)
    type(editor, ' \uff08y)')
    expect(text(editor)).toBe('\uff08note\uff09 [x] \uff08y)')
  })
  it('a space between Asian and Western text is dropped, once switched on', async () => {
    const editor = await createEditor()
    type(editor, '\u4e2d\u6587 a')
    expect(text(editor)).toBe('\u4e2d\u6587 a')
    setAutocorrectPref('asianSpaces', true)
    type(editor, 'b \u6587 c  d')
    expect(text(editor)).toBe('\u4e2d\u6587 ab\u6587c  d')
  })
  it('a salutation gets its closing, right-aligned, below', async () => {
    const editor = await createEditor()
    type(editor, '\u62dd\u5553 ')
    expect(paragraphs(editor)).toEqual(['\u62dd\u5553 ', '\u656c\u5177'])
    expect(editor.state.doc.child(1).attrs.align).toBe('right')
    expect(editor.state.selection.$from.parent).toBe(firstBlock(editor))
    type(editor, ' ')
    expect(editor.state.doc.childCount).toBe(2)
  })
  it('Enter after a salutation does the same and leaves a line between', async () => {
    const editor = await createEditor()
    type(editor, '\u524d\u7565')
    enterAndSplit(editor)
    expect(paragraphs(editor)).toEqual(['\u524d\u7565', '', '\u8349\u3005'])
    setAutocorrectPref('closingPhrase', false)
    editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
    editor.commands.focus('end')
    type(editor, '\u62dd\u5553 ')
    expect(editor.state.doc.childCount).toBe(1)
  })
})

describe('autocorrect: indents from the keyboard', () => {
  it('leading spaces become a first-line indent on Enter, which the next paragraph keeps', async () => {
    const editor = await createEditor()
    type(editor, '\u3000\u3000body')
    enterAndSplit(editor)
    expect(paragraphs(editor)).toEqual(['body', ''])
    expect(firstBlock(editor).attrs).toMatchObject({ indentFirstLineChars: 200 })
    expect(Number(firstBlock(editor).attrs.indentFirstLine)).toBeGreaterThan(0)
    expect(editor.state.doc.child(1).attrs.indentFirstLineChars).toBe(200)
  })
  it('one space is not an indent, and the switch turns it off', async () => {
    const editor = await createEditor()
    type(editor, ' a')
    enterAndSplit(editor)
    expect(paragraphs(editor)).toEqual([' a', ''])
    setAutocorrectPref('indentSpaces', false)
    type(editor, '    b')
    enterAndSplit(editor)
    expect(paragraphs(editor)).toEqual([' a', '    b', ''])
  })
  it('Tab at the start of a paragraph indents the first line, then the paragraph; Backspace right after takes the last one back', async () => {
    const editor = await createEditor()
    type(editor, 'body')
    editor.commands.setTextSelection(1)
    expect(press(editor, 'Tab')).toBe(true)
    expect(firstBlock(editor).attrs).toMatchObject({ indentFirstLine: 720, indentLeft: null })
    expect(text(editor)).toBe('body')
    press(editor, 'Tab')
    expect(firstBlock(editor).attrs).toMatchObject({ indentFirstLine: 720, indentLeft: 720 })
    backspace(editor)
    expect(firstBlock(editor).attrs).toMatchObject({ indentFirstLine: 720, indentLeft: null })
    expect(text(editor)).toBe('body')
  })
  it('Tab inside the text, in an empty paragraph or with the switch off is not an indent', async () => {
    const editor = await createEditor()
    const plugin = autocorrectKey.get(editor.state)!
    const tab = () =>
      plugin.props.handleKeyDown!.call(
        plugin,
        editor.view,
        new KeyboardEvent('keydown', { key: 'Tab' }),
      )
    expect(tab()).toBe(false)
    type(editor, 'body')
    expect(tab()).toBe(false)
    setAutocorrectPref('tabIndents', false)
    editor.commands.setTextSelection(1)
    expect(tab()).toBe(false)
    expect(firstBlock(editor).attrs.indentFirstLine).toBeNull()
  })
})

describe('autocorrect: built-in heading styles', () => {
  async function headingEditor() {
    const editor = await createEditor()
    const parsed = await parseDocx(await buildBlankDocx())
    editor.storage.listNumbering.styles = parsed.styles
    return editor
  }
  it('is opt-in: a short line with a blank above, Enter twice, becomes Heading 1', async () => {
    const editor = await headingEditor()
    type(editor, 'Summary')
    enterAndSplit(editor)
    enterAndSplit(editor)
    expect(firstBlock(editor).type.name).toBe('docParagraph')
    setAutocorrectPref('headingStyles', true)
    editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
    editor.commands.focus('end')
    type(editor, 'Summary')
    enterAndSplit(editor)
    enterAndSplit(editor)
    expect(firstBlock(editor).type.name).toBe('docHeading')
    expect(firstBlock(editor).attrs.level).toBe(1)
    expect(editor.state.doc.childCount).toBe(2)
    backspace(editor)
    expect(firstBlock(editor).type.name).toBe('docParagraph')
  })
  it('a leading tab is a level down; a sentence or a line under body text stays a paragraph', async () => {
    setAutocorrectPref('headingStyles', true)
    setAutocorrectPref('tabIndents', false)
    const editor = await headingEditor()
    type(editor, '\tDetails')
    enterAndSplit(editor)
    enterAndSplit(editor)
    expect(firstBlock(editor).type.name).toBe('docHeading')
    expect(firstBlock(editor).attrs.level).toBe(2)
    expect(firstBlock(editor).textContent).toBe('Details')
    type(editor, 'It ends with a period.')
    enterAndSplit(editor)
    enterAndSplit(editor)
    type(editor, 'A period and a space. ')
    enterAndSplit(editor)
    enterAndSplit(editor)
    type(editor, 'Body text')
    enterAndSplit(editor)
    type(editor, 'Not a heading')
    enterAndSplit(editor)
    enterAndSplit(editor)
    const kinds: string[] = []
    editor.state.doc.forEach((node) => kinds.push(node.type.name))
    expect(kinds.filter((k) => k === 'docHeading')).toHaveLength(1)
  })
})

describe('autocorrect: a list item starts like the one before', () => {
  it('takes the bold of the term up to its delimiter', async () => {
    const editor = await createEditor()
    const bold = [{ type: 'bold' }]
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docListItem',
          attrs: { kind: 'bullet', numId: '1', ilvl: 0 },
          content: [
            { type: 'text', text: 'Term:', marks: bold },
            { type: 'text', text: ' meaning' },
          ],
        },
        { type: 'docListItem', attrs: { kind: 'bullet', numId: '1', ilvl: 0 } },
      ],
    })
    editor.commands.focus('end')
    type(editor, 'Next: more')
    const runs: string[] = []
    editor.state.doc.child(1).forEach((n) => {
      runs.push(`${n.text}${n.marks.some((m) => m.type.name === 'bold') ? '*' : ''}`)
    })
    expect(runs).toEqual(['Next:*', ' more'])
  })
  it('does nothing after a plain item or with the switch off', async () => {
    setAutocorrectPref('listStart', false)
    const editor = await createEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docListItem',
          attrs: { kind: 'bullet', numId: '1', ilvl: 0 },
          content: [
            { type: 'text', text: 'Term:', marks: [{ type: 'bold' }] },
            { type: 'text', text: ' meaning' },
          ],
        },
        { type: 'docListItem', attrs: { kind: 'bullet', numId: '1', ilvl: 0 } },
      ],
    })
    editor.commands.focus('end')
    type(editor, 'Next')
    expect(editor.state.doc.child(1).firstChild!.marks).toHaveLength(0)
  })
})

describe('autocorrect: formatted replacements', () => {
  it('an entry kept as formatted text goes in with its formatting', async () => {
    setAutocorrectList(
      'replace',
      upsertEntry(DEFAULT_REPLACE_ENTRIES, 'sig', 'Jo Doe', '<p><strong>Jo</strong> Doe</p>'),
    )
    const editor = await createEditor()
    type(editor, 'by sig.')
    expect(text(editor)).toBe('by Jo Doe.')
    const runs: string[] = []
    firstBlock(editor).forEach((n) => {
      runs.push(`${n.text}${n.marks.some((m) => m.type.name === 'bold') ? '*' : ''}`)
    })
    expect(runs).toEqual(['by ', 'Jo*', ' Doe.'])
    backspace(editor)
    expect(text(editor)).toBe('by sig.')
  })
})

describe('autocorrect: AutoComplete tip', () => {
  it('matches four typed characters that start a longer name, at a word boundary', () => {
    const names = ['Yours truly,', 'November']
    expect(autoCompleteMatch('Dear all, your', names)).toEqual({ length: 4, name: 'Yours truly,' })
    expect(autoCompleteMatch('yours tr', names)).toEqual({ length: 8, name: 'Yours truly,' })
    expect(autoCompleteMatch('you', names)).toBeNull()
    expect(autoCompleteMatch('xyours', names)).toBeNull()
    expect(autoCompleteMatch('November', names)).toBeNull()
  })
  it('Enter puts the entry in for what was typed; Escape or the switch dismisses the tip', async () => {
    const editor = await createEditor()
    editor.storage.autocorrect.hooks!.autoText = () => [
      { name: 'Yours truly,', text: 'Yours truly, Jo', html: '<p>Yours truly, <em>Jo</em></p>' },
    ]
    type(editor, 'So your')
    expect(autoCompleteKey.getState(editor.state)).toMatchObject({ text: 'Yours truly, Jo' })
    press(editor, 'Escape')
    expect(autoCompleteKey.getState(editor.state)).toBeNull()
    type(editor, 's')
    expect(enter(editor)).toBe(true)
    expect(text(editor)).toBe('So Yours truly, Jo')
    expect(editor.state.doc.childCount).toBe(1)
    setAutocorrectPref('autoComplete', false)
    type(editor, ' your')
    expect(autoCompleteKey.getState(editor.state)).toBeNull()
  })
  it('offers a month name', async () => {
    const editor = await createEditor()
    type(editor, 'in Nove')
    expect(autoCompleteKey.getState(editor.state)).toMatchObject({ text: 'November' })
    enter(editor)
    expect(text(editor)).toBe('in November')
  })
})

describe('autocorrect: the smart button keeps what was corrected', () => {
  const fixes = (editor: Editor) =>
    fixesKey.getState(editor.state)!.map((f) => [f.original, f.corrected, f.rule])
  it('remembers each correction as whole words, through later typing', async () => {
    setAutocorrectList('replace', upsertEntry(DEFAULT_REPLACE_ENTRIES, 'teh', 'the'))
    const editor = await createEditor()
    type(editor, 'THe teh (c) ok')
    expect(fixes(editor)).toEqual([
      ['THe', 'The', 'twoInitialCaps'],
      ['teh', 'the', 'replaceText'],
      ['(c)', '\u00a9', 'replaceText'],
    ])
    const marked = fixesKey.getState(editor.state)!
    expect(marked.map((f) => editor.state.doc.textBetween(f.from, f.to))).toEqual([
      'The',
      'the',
      '\u00a9',
    ])
  })
  it('forgets a correction that is typed into', async () => {
    const editor = await createEditor()
    type(editor, 'THe ')
    editor.commands.setTextSelection(3)
    type(editor, 'x')
    expect(fixes(editor)).toEqual([])
  })
  it('Change Back puts the typed text in again and drops the mark', async () => {
    const editor = await createEditor()
    type(editor, 'THe end')
    const [fix] = fixesKey.getState(editor.state)!
    editor.view.dispatch(changeBackFix(editor.state, fix))
    expect(text(editor)).toBe('THe end')
    expect(fixes(editor)).toEqual([])
  })
  it('Stop Automatically Correcting removes the entry, adds the exception, or switches the rule off', async () => {
    setAutocorrectList('replace', upsertEntry(DEFAULT_REPLACE_ENTRIES, 'teh', 'the'))
    const editor = await createEditor()
    type(editor, 'THe teh "x ')
    const [caps, entry, quote] = fixesKey.getState(editor.state)!
    stopCorrecting(caps)
    stopCorrecting(entry)
    stopCorrecting(quote)
    expect(autocorrectLists().initialCaps).toContain('THe')
    expect(autocorrectLists().replace.some(([from]) => from === 'teh')).toBe(false)
    expect(autocorrectPrefs().smartQuotes).toBe(false)
    type(editor, 'THe teh "y ')
    expect(text(editor)).toBe('The the \u201cx THe teh "y ')
  })
  it('two corrections of the same word are told apart', async () => {
    setAutocorrectList('replace', upsertEntry(DEFAULT_REPLACE_ENTRIES, 'teh', 'the'))
    const editor = await createEditor()
    type(editor, 'teh cat, teh dog ')
    const [first, second] = fixesKey.getState(editor.state)!
    expect(first.id).not.toBe(second.id)
    editor.view.dispatch(changeBackFix(editor.state, second))
    expect(text(editor)).toBe('the cat, teh dog ')
    expect(fixesKey.getState(editor.state)!.map((f) => f.id)).toEqual([first.id])
  })
  it('stopping a math name keeps its capitalized twin', async () => {
    setAutocorrectPref('mathOutside', true)
    const editor = await createEditor()
    type(editor, '\\delta ')
    stopCorrecting(fixesKey.getState(editor.state)![0])
    const names = autocorrectLists().math.map(([from]) => from)
    expect(names).not.toContain('\\delta')
    expect(names).toContain('\\Delta')
  })
})

describe('autocorrect: Date style and styles defined from formatting', () => {
  async function styledEditor() {
    const editor = await createEditor()
    const parsed = await parseDocx(await buildBlankDocx())
    editor.storage.listNumbering.styles = parsed.styles
    editor.storage.listNumbering.docDefaults = parsed.docDefaults
    const defined: StyleUpsert[] = []
    editor.storage.autocorrect.hooks = {
      ...editor.storage.autocorrect.hooks!,
      defineStyle: (up) => defined.push(up),
    }
    return { editor, defined }
  }
  const blocks = (editor: Editor) => {
    const out: Array<{ type: string; styleId: unknown; text: string }> = []
    editor.state.doc.forEach((node) =>
      out.push({ type: node.type.name, styleId: node.attrs.styleId, text: node.textContent }),
    )
    return out
  }

  it('reads a line that is a date and nothing else', () => {
    for (const date of [
      'November 23, 2026',
      'Monday, November 23, 2026',
      '23 November 2026',
      '2026-11-23',
      '11/23/2026',
      '2026\u5e7411\u670823\u65e5',
      '\u4ee4\u548c8\u5e7411\u670823\u65e5',
      '\u4ee4\u548c\u5143\u5e745\u67081\u65e5',
      '\uff12\uff10\uff12\uff16\u5e74\uff11\uff11\u6708\uff12\uff13\u65e5',
      '2026\ub144 11\uc6d4 23\uc77c',
    ])
      expect(isDateLine(date), date).toBe(true)
    for (const other of ['', 'Due November 23, 2026', 'November 2026', '1.2.3', '23/11', 'May'])
      expect(isDateLine(other), other).toBe(false)
  })

  it('Date style is opt-in: Enter after a date gives the line the Date style and body text below', async () => {
    const { editor, defined } = await styledEditor()
    type(editor, 'November 23, 2026')
    enterAndSplit(editor)
    expect(blocks(editor).map((b) => b.styleId)).toEqual([null, null])
    expect(defined).toEqual([])

    setAutocorrectPref('dateStyle', true)
    editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
    editor.commands.focus('end')
    type(editor, 'November 23, 2026')
    enterAndSplit(editor)
    expect(blocks(editor)).toEqual([
      { type: 'docParagraph', styleId: 'Date', text: 'November 23, 2026' },
      { type: 'docParagraph', styleId: null, text: '' },
    ])
    expect(defined).toEqual([
      expect.objectContaining({ styleId: 'Date', name: 'Date', builtIn: true, basedOn: 'Normal' }),
    ])
    expect(editor.state.selection.$from.parent).toBe(editor.state.doc.child(1))
    // one undo takes the new paragraph away, the next the style
    editor.commands.undo()
    expect(blocks(editor)).toEqual([
      { type: 'docParagraph', styleId: 'Date', text: 'November 23, 2026' },
    ])
    editor.commands.undo()
    expect(blocks(editor)).toEqual([
      { type: 'docParagraph', styleId: null, text: 'November 23, 2026' },
    ])
    editor.commands.focus('end')
    enterAndSplit(editor)
    type(editor, 'Dear all, see you on November 23, 2026')
    enterAndSplit(editor)
    expect(blocks(editor).map((b) => b.styleId)).toEqual(['Date', null, null])
  })

  it('a document that has a Date style gets no second one', async () => {
    setAutocorrectPref('dateStyle', true)
    const { editor, defined } = await styledEditor()
    const styles = new Map(editor.storage.listNumbering.styles)
    styles.set('a7', { styleId: 'a7', name: 'Date', type: 'paragraph' })
    editor.storage.listNumbering.styles = styles
    type(editor, '2026-11-23')
    enterAndSplit(editor)
    expect(blocks(editor).map((b) => b.styleId)).toEqual(['a7', null])
    expect(defined).toEqual([])
  })

  it('Define styles is opt-in: a bold short line becomes a heading whose style is made from it', async () => {
    const { editor, defined } = await styledEditor()
    const bold = [{ type: 'bold' }, { type: 'docTextStyle', attrs: { sizeHalfPoints: 36 } }]
    const line = (text: string, marks = bold) => ({
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text, marks }] }],
    })
    editor.commands.setContent(line('Quarterly results'))
    editor.commands.focus('end')
    enterAndSplit(editor)
    expect(blocks(editor).map((b) => b.type)).toEqual(['docParagraph', 'docParagraph'])

    setAutocorrectPref('defineStyles', true)
    editor.commands.setContent(line('Quarterly results'))
    editor.commands.focus('end')
    enterAndSplit(editor)
    expect(blocks(editor)).toEqual([
      { type: 'docHeading', styleId: 'Heading1', text: 'Quarterly results' },
      { type: 'docParagraph', styleId: null, text: '' },
    ])
    expect(editor.state.doc.firstChild!.attrs.level).toBe(1)
    // the formatting is the style's now, not the text's
    expect(editor.state.doc.firstChild!.firstChild!.marks).toEqual([])
    expect(defined).toHaveLength(1)
    expect(defined[0]).toMatchObject({
      styleId: 'Heading1',
      rPr: { bold: true, italic: false, sizeHalfPoints: 36 },
      pPr: { align: null },
    })
    // what is typed below is plain body text
    type(editor, 'Body')
    expect(editor.state.doc.child(1).firstChild!.marks).toEqual([])

    // a second line that looks the same takes the style just made, before the style table knows it
    editor.commands.setContent({
      type: 'doc',
      content: [
        ...(editor.getJSON().content ?? []),
        { type: 'docParagraph', content: [{ type: 'text', text: 'Outlook', marks: bold }] },
      ],
    })
    editor.commands.focus('end')
    enterAndSplit(editor)
    expect(blocks(editor).at(-2)).toEqual({
      type: 'docHeading',
      styleId: 'Heading1',
      text: 'Outlook',
    })
    expect(defined).toHaveLength(1)

    // another style table (another document): what was made before is not remembered
    editor.storage.listNumbering.styles = new Map(editor.storage.listNumbering.styles)
    editor.commands.setContent(line('Fresh start'))
    editor.commands.focus('end')
    enterAndSplit(editor)
    expect(blocks(editor)[0]).toMatchObject({ type: 'docHeading', text: 'Fresh start' })
    expect(defined).toHaveLength(2)
  })

  it('acts while Track Changes records: the heading is a tracked formatting change of the typed line', async () => {
    setAutocorrectPref('defineStyles', true)
    const { editor, defined } = await styledEditor()
    editor.commands.setContent({ type: 'doc', content: [{ type: 'docParagraph' }] })
    const storage = editor.storage.trackChanges as { enabled: boolean; author: string }
    storage.enabled = true
    storage.author = 'Tester'
    editor.commands.focus('end')
    editor.commands.setMark('bold')
    type(editor, 'Quarterly results')
    const typed = editor.state.doc.firstChild!.firstChild!
    expect(typed.marks.map((m) => m.type.name).sort()).toEqual(['bold', 'ins'])
    enterAndSplit(editor)
    // bold at the size of body text: the heading style that already looks like it
    expect(blocks(editor)).toEqual([
      { type: 'docHeading', styleId: 'Heading5', text: 'Quarterly results' },
      { type: 'docParagraph', styleId: null, text: '' },
    ])
    expect(defined).toEqual([])
    const head = editor.state.doc.firstChild!
    // the text stays a tracked insertion; the bold went into the style
    expect(head.firstChild!.marks.some((m) => m.type.name === 'ins')).toBe(true)
    expect(head.firstChild!.marks.some((m) => m.type.name === 'bold')).toBe(false)
    // the paragraph remembers what it was before it took the style
    const change = JSON.parse(String(head.attrs.pPrChange)) as { author: string }
    expect(change.author).toBe('Tester')
    storage.enabled = false
  })

  it('a second look takes the next free level; a sentence or plain text is left alone', async () => {
    setAutocorrectPref('defineStyles', true)
    const { editor, defined } = await styledEditor()
    const text = (value: string, marks: Array<{ type: string }>) => ({
      type: 'docParagraph',
      content: [{ type: 'text', text: value, marks }],
    })
    editor.commands.setContent({
      type: 'doc',
      content: [
        { type: 'docHeading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Title' }] },
        text('Italic and bold', [{ type: 'bold' }, { type: 'italic' }]),
      ],
    })
    editor.commands.focus('end')
    enterAndSplit(editor)
    expect(blocks(editor)[1]).toMatchObject({ type: 'docHeading', styleId: 'Heading2' })
    expect(defined.map((up) => up.styleId)).toEqual(['Heading2'])

    for (const [value, marks] of [
      ['It ends with a period.', [{ type: 'bold' }]],
      ['Plain line', []],
      ['Small italic', [{ type: 'italic' }]],
    ] as const) {
      editor.commands.setContent({ type: 'doc', content: [text(value, [...marks])] })
      editor.commands.focus('end')
      enterAndSplit(editor)
      expect(blocks(editor).map((b) => b.type)).toEqual(['docParagraph', 'docParagraph'])
    }
    // half of the line bold is not a line formatted as a whole
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'Half ', marks: [{ type: 'bold' }] },
            { type: 'text', text: 'bold' },
          ],
        },
      ],
    })
    editor.commands.focus('end')
    enterAndSplit(editor)
    expect(blocks(editor).map((b) => b.type)).toEqual(['docParagraph', 'docParagraph'])
    expect(defined).toHaveLength(1)
  })
})
