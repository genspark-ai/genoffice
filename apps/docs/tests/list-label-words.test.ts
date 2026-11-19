/**
 * Word counts list numbers as words, and a bullet written as text (Courier New
 * "o"), but not a symbol-font bullet. Word for Mac 16 probe 2026-11-19: three
 * items "alpha", "beta", "gamma" are six words numbered, three under a Symbol
 * bullet, six under "o". A grant form with 8 numbered and 5 "o" items came out
 * 13 words short of Word's 3919 without them.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { NumberingDef, NumberingLevel } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocumentContext } from '../src/renderer/ai/protocol'
import { countWords, listLabelText } from '../src/renderer/word-count'

const level = (l: Partial<NumberingLevel>): NumberingLevel =>
  ({ numFmt: 'decimal', lvlText: '%1.', start: 1, ...l }) as NumberingLevel
const def = (numId: string, lvl: NumberingLevel): NumberingDef => ({
  numId,
  abstractNumId: `a${numId}`,
  levels: { 0: lvl },
  startOverrides: {},
})
const defs = new Map<string, NumberingDef>([
  ['1', def('1', level({}))],
  ['2', def('2', level({ numFmt: 'bullet', lvlText: 'o', font: 'Courier New' }))],
  ['3', def('3', level({ numFmt: 'bullet', lvlText: '', font: 'Symbol' }))],
])

const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

function listOf(numId: string, paraMarkDel = false) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: ['alpha', 'beta', 'gamma'].map((t, i) => ({
        type: 'docListItem',
        attrs: {
          docxIndex: null,
          kind: 'bullet',
          numId,
          ilvl: 0,
          ...(paraMarkDel && i === 2 ? { paraMarkDel: true } : {}),
        },
        content: [{ type: 'text', text: t }],
      })),
    },
  })
  liveEditors.push(editor)
  return editor
}

describe('list labels in the word count', () => {
  it('counts numbers and text bullets, not symbol bullets', () => {
    expect(countWords(listLabelText(listOf('1').state.doc, defs))).toBe(3)
    expect(countWords(listLabelText(listOf('2').state.doc, defs))).toBe(3)
    expect(countWords(listLabelText(listOf('3').state.doc, defs))).toBe(0)
  })

  it('counts nothing without numbering definitions', () => {
    expect(listLabelText(listOf('1').state.doc, undefined)).toBe('')
    expect(listLabelText(listOf('1').state.doc, new Map())).toBe('')
  })

  it('leaves out the label of an item whose paragraph mark is deleted', () => {
    expect(countWords(listLabelText(listOf('1', true).state.doc, defs))).toBe(2)
  })

  it('reaches the AI stats line like the status bar', () => {
    const editor = listOf('1')
    ;(editor.storage as unknown as { listNumbering: { defs: typeof defs } }).listNumbering.defs =
      defs
    expect(buildDocumentContext(editor)).toContain('Full-text stats: words 6')
  })
})
