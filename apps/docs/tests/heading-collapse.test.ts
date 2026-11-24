/**
 * Word's Expand / Collapse of the body under a heading: the hidden range, the
 * view state through edits, and "Collapsed by default" from file to file.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { generateParagraphXml, parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmNodeToGeneratedBlock } from '../src/renderer/editor/convert'
import { collectHeadings } from '../src/renderer/editor/headings'
import {
  HEADING_BODY_KEEP,
  HEADING_COLLAPSE_KEEP,
  collapsedByDefaultAt,
  collapsedHeadingRanges,
  foldHeading,
  headingBodyRange,
  headingCollapseKey,
  headingFoldAt,
  hiddenRanges,
  sectionEndTest,
  setCollapsedByDefault,
} from '../src/renderer/editor/heading-collapse'
import { cutToSpike, insertSpike } from '../src/renderer/editor/word-keys'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string, attrs: Record<string, unknown> = {}): JSONContent => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: text ? [{ type: 'text', text }] : [],
})
const heading = (
  level: number,
  text: string,
  attrs: Record<string, unknown> = {},
): JSONContent => ({
  type: 'docHeading',
  attrs: { level, ...attrs },
  content: text ? [{ type: 'text', text }] : [],
})
const cell = (...content: JSONContent[]): JSONContent => ({ type: 'docTableCell', content })
const table = (...cells: JSONContent[]): JSONContent => ({
  type: 'docTable',
  content: [{ type: 'docTableRow', content: cells }],
})
const sectionBreak: JSONContent = {
  type: 'docProtected',
  attrs: { label: 'Section break paragraph' },
}

function createEditor(content: JSONContent[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

/** texts of the top-level blocks the folds leave visible */
function visible(editor: Editor): string[] {
  const hidden = collapsedHeadingRanges(editor.state)
  const out: string[] = []
  editor.state.doc.forEach((node, pos) => {
    if (!hidden.some((r) => pos >= r.from && pos < r.to)) out.push(node.textContent)
  })
  return out
}
const posOf = (editor: Editor, text: string): number => {
  let found = -1
  editor.state.doc.forEach((node, pos) => {
    if (found < 0 && node.textContent === text) found = pos
  })
  return found
}
const caretAtEnd = (editor: Editor, text: string) => {
  const pos = posOf(editor, text)
  const end = pos + editor.state.doc.nodeAt(pos)!.nodeSize - 1
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)))
}

const NESTED = [
  para('intro'),
  heading(1, 'Alpha'),
  para('a1'),
  heading(2, 'Alpha sub'),
  para('a2'),
  heading(3, 'Alpha deep'),
  para('a3'),
  heading(1, 'Beta'),
  para('b1'),
  heading(2, 'Beta sub 1'),
  para('b2'),
  heading(2, 'Beta sub 2'),
  para('b3'),
]

describe('headingBodyRange', () => {
  const rangeTexts = (content: JSONContent[], title: string, ends?: Set<number>) => {
    const { doc } = createEditor(content).state
    const headings = collectHeadings(doc, undefined, true)
    const index = headings.findIndex((h) => h.text === title)
    const range = headingBodyRange(doc, headings, index, sectionEndTest(ends))
    const out: string[] = []
    if (range) doc.nodesBetween(range.from, range.to, (n) => (out.push(n.textContent), false))
    return out
  }

  it('runs to the next heading of the same or a higher level', () => {
    expect(rangeTexts(NESTED, 'Alpha')).toEqual(['a1', 'Alpha sub', 'a2', 'Alpha deep', 'a3'])
    expect(rangeTexts(NESTED, 'Alpha sub')).toEqual(['a2', 'Alpha deep', 'a3'])
    expect(rangeTexts(NESTED, 'Alpha deep')).toEqual(['a3'])
    expect(rangeTexts(NESTED, 'Beta sub 1')).toEqual(['b2'])
  })

  it('the last heading takes everything to the end of the document', () => {
    expect(rangeTexts(NESTED, 'Beta')).toEqual(['b1', 'Beta sub 1', 'b2', 'Beta sub 2', 'b3'])
    expect(rangeTexts(NESTED, 'Beta sub 2')).toEqual(['b3'])
  })

  it('takes a table whole: a heading in a cell neither folds nor ends the fold', () => {
    const doc = [
      heading(2, 'Alpha'),
      para('a1'),
      table(cell(heading(1, 'In cell'), para('cell body'))),
      para('a2'),
      heading(1, 'Beta'),
    ]
    expect(rangeTexts(doc, 'Alpha')).toEqual(['a1', 'In cellcell body', 'a2'])
    const editor = createEditor(doc)
    const cellHeading = posOf(editor, 'In cellcell body') + 4
    expect(headingFoldAt(editor, cellHeading)).toBeNull()
  })

  it('stops before the paragraph that ends the section', () => {
    const doc = [
      heading(1, 'Alpha'),
      para('a1'),
      sectionBreak,
      para('a2'),
      heading(1, 'Beta'),
      para('b1'),
      para('carries the break', { docxIndex: 7 }),
      para('b2'),
    ]
    expect(rangeTexts(doc, 'Alpha')).toEqual(['a1'])
    expect(rangeTexts(doc, 'Beta', new Set([7]))).toEqual(['b1'])
    expect(rangeTexts(doc, 'Beta')).toEqual(['b1', 'carries the break', 'b2'])
  })

  it('has no range for a heading with nothing under it', () => {
    expect(rangeTexts([heading(1, 'Alpha'), heading(1, 'Beta'), para('b')], 'Alpha')).toEqual([])
    expect(rangeTexts([heading(1, 'Alpha'), sectionBreak, para('a')], 'Alpha')).toEqual([])
  })

  it('an outer fold swallows the folds inside it', () => {
    const { doc } = createEditor(NESTED).state
    const headings = collectHeadings(doc, undefined, true)
    const at = (title: string) => headings.find((h) => h.text === title)!.pos
    const ranges = hiddenRanges(doc, headings, [at('Alpha sub'), at('Alpha')], sectionEndTest())
    expect(ranges.map((r) => r.heading)).toEqual([at('Alpha')])
  })
})

describe('collapsing in the editor', () => {
  it('collapse, expand and toggle act on the heading at a position', () => {
    const editor = createEditor(NESTED)
    expect(foldHeading(editor, 'collapse', posOf(editor, 'Alpha'))).toBe(true)
    expect(visible(editor)).toEqual([
      'intro',
      'Alpha',
      'Beta',
      'b1',
      'Beta sub 1',
      'b2',
      'Beta sub 2',
      'b3',
    ])
    expect(foldHeading(editor, 'collapse', posOf(editor, 'Alpha'))).toBe(false)
    expect(foldHeading(editor, 'toggle', posOf(editor, 'Alpha'))).toBe(true)
    expect(visible(editor)).toHaveLength(NESTED.length)
    expect(editor.can().undo()).toBe(false)
  })

  it('without a position it folds the heading the caret is under, and parks the caret on it', () => {
    const editor = createEditor(NESTED)
    caretAtEnd(editor, 'a1')
    expect(foldHeading(editor, 'collapse')).toBe(true)
    expect(visible(editor)).not.toContain('a1')
    expect(editor.state.selection.$head.parent.textContent).toBe('Alpha')
  })

  it('Collapse All folds every heading with a body; Expand All opens them', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse-all')
    expect(visible(editor)).toEqual(['intro', 'Alpha', 'Beta'])
    // the inner folds are kept: opening the outer one shows them still folded
    foldHeading(editor, 'expand', posOf(editor, 'Alpha'))
    expect(visible(editor)).toEqual(['intro', 'Alpha', 'a1', 'Alpha sub', 'Beta'])
    foldHeading(editor, 'expand-all')
    expect(visible(editor)).toHaveLength(NESTED.length)
  })

  it('marks the heading and hides the body blocks with decorations only', () => {
    const editor = createEditor(NESTED)
    const before = editor.state.doc
    foldHeading(editor, 'collapse', posOf(editor, 'Beta sub 1'))
    expect(editor.state.doc.eq(before)).toBe(true)
    const dom = editor.view.dom
    expect(dom.querySelectorAll('.doc-heading-collapsed')).toHaveLength(1)
    expect(dom.querySelector('.doc-heading-collapsed')?.textContent).toBe('Beta sub 1')
    expect([...dom.querySelectorAll('.doc-collapsed-body')].map((el) => el.textContent)).toEqual([
      'b2',
    ])
    expect(dom.querySelector('.heading-toggle.is-collapsed')).not.toBeNull()
  })

  it('a document with nothing collapsed renders no fold markup at all', () => {
    const editor = createEditor(NESTED)
    expect(headingCollapseKey.getState(editor.state)?.collapsed).toEqual([])
    expect(
      editor.view.dom.querySelector('.heading-toggle, .doc-fold-gap, .doc-collapsed-body'),
    ).toBeNull()
  })

  it('the triangle of a collapsed heading expands it', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Alpha'))
    const toggle = editor.view.dom.querySelector('.heading-toggle.is-collapsed')!
    toggle.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
    expect(visible(editor)).toHaveLength(NESTED.length)
  })

  it('a selection landing in a hidden body opens it (Find, Go To, the navigation pane)', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse-all')
    const target = posOf(editor, 'a3') + 1
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, target, target + 2)),
    )
    expect(visible(editor)).toContain('a3')
    // only the folds on the way to it open
    expect(visible(editor)).not.toContain('b1')
  })

  it('folds follow their heading through edits above them', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Beta'))
    editor.commands.insertContentAt(1, 'more ')
    expect(visible(editor)).toEqual([
      'more intro',
      'Alpha',
      'a1',
      'Alpha sub',
      'a2',
      'Alpha deep',
      'a3',
      'Beta',
    ])
  })

  it('Enter at the end of a collapsed heading opens the new paragraph below the hidden body', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Alpha'))
    caretAtEnd(editor, 'Alpha')
    editor.commands.keyboardShortcut('Enter')
    const texts: string[] = []
    editor.state.doc.forEach((n) => texts.push(n.textContent))
    expect(texts.slice(0, 9)).toEqual([
      'intro',
      'Alpha',
      'a1',
      'Alpha sub',
      'a2',
      'Alpha deep',
      'a3',
      '',
      'Beta',
    ])
    expect(visible(editor).slice(0, 4)).toEqual(['intro', 'Alpha', '', 'Beta'])
    const $head = editor.state.selection.$head
    expect($head.parent.content.size).toBe(0)
    expect($head.before(1)).toBe(
      posOf(editor, 'a3') + editor.state.doc.nodeAt(posOf(editor, 'a3'))!.nodeSize,
    )
    editor.commands.undo()
    expect(editor.state.doc.childCount).toBe(NESTED.length)
  })

  it('deleting a selected collapsed heading takes its hidden body; typing over it renames it', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Alpha'))
    const select = () => {
      const pos = posOf(editor, 'Alpha')
      editor.commands.setTextSelection({ from: pos + 1, to: pos + 1 + 'Alpha'.length })
    }
    select()
    editor.commands.insertContent('Renamed')
    expect(editor.state.doc.childCount).toBe(NESTED.length)
    expect(visible(editor).slice(0, 3)).toEqual(['intro', 'Renamed', 'Beta'])
    editor.commands.undo()
    select()
    editor.commands.deleteSelection()
    const texts: string[] = []
    editor.state.doc.forEach((n) => texts.push(n.textContent))
    expect(texts).toEqual(['intro', '', 'Beta', 'b1', 'Beta sub 1', 'b2', 'Beta sub 2', 'b3'])
    editor.commands.undo()
    expect(editor.state.doc.childCount).toBe(NESTED.length)
  })

  it('the Spike takes a selected collapsed heading with its hidden body, and gives both back', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Alpha sub'))
    const pos = posOf(editor, 'Alpha sub')
    editor.commands.setTextSelection({ from: pos + 1, to: pos + 1 + 'Alpha sub'.length })
    expect(cutToSpike(editor)).toBe(true)
    const texts = () => {
      const out: string[] = []
      editor.state.doc.forEach((n) => out.push(n.textContent))
      return out
    }
    expect(texts()).not.toContain('a2')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(insertSpike(editor)).toBe(true)
    expect(texts().join('|')).toContain('Alpha sub|a2|Alpha deep|a3')
  })

  it('Replace over a selected collapsed heading leaves its hidden body', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Alpha'))
    const pos = posOf(editor, 'Alpha')
    const from = pos + 1
    const to = from + 'Alpha'.length
    editor.commands.setTextSelection({ from, to })
    editor.view.dispatch(editor.state.tr.insertText('', from, to).setMeta(HEADING_BODY_KEEP, true))
    expect(editor.state.doc.childCount).toBe(NESTED.length)
    expect(editor.state.doc.textContent).toContain('a3')
  })

  it('copying a selected collapsed heading copies its hidden body', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Alpha sub'))
    const pos = posOf(editor, 'Alpha sub')
    editor.commands.setTextSelection({ from: pos + 1, to: pos + 1 + 'Alpha sub'.length })
    const { state } = editor
    let slice = state.selection.content()
    editor.view.someProp('transformCopied', (f) => {
      slice = f(slice, editor.view)
    })
    const texts: string[] = []
    slice.content.forEach((n) => texts.push(n.textContent))
    expect(texts).toEqual(['Alpha sub', 'a2', 'Alpha deep', 'a3'])
  })
})

describe('"Collapsed by default"', () => {
  const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
  const p = (pPr: string, text: string) =>
    `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`
  const FOLD_STYLE =
    `<w:style w:type="paragraph" w:styleId="Fold" ${W15}><w:name w:val="Fold"/><w:basedOn w:val="Normal"/>` +
    '<w:pPr><w:outlineLvl w:val="0"/><w15:collapsed/></w:pPr></w:style>'

  async function openDoc(bodyXml: string) {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml, extraStylesXml: FOLD_STYLE, docRootExtraAttrs: W15 }),
    )
    const editor = createEditor([para('placeholder')])
    editor.storage.listNumbering.styles = parsed.styles
    editor
      .chain()
      .setMeta('addToHistory', false)
      .setContent(blocksToPmDoc(parsed.blocks) as never)
      .run()
    return { editor, parsed }
  }
  const CTX = { headingStyleIds: new Map<number, string>(), allocateHyperlinkRel: () => 'rId1' }
  const xmlOf = (editor: Editor, text: string) =>
    generateParagraphXml(
      pmNodeToGeneratedBlock(editor.state.doc.nodeAt(posOf(editor, text))!.toJSON()),
      CTX,
    )

  it('a document opens with its flagged headings collapsed, the style flag and w:val="0" honoured', async () => {
    const { editor } = await openDoc(
      p('<w:pStyle w:val="Heading1"/><w15:collapsed/>', 'Alpha') +
        p('', 'a1') +
        p('<w:pStyle w:val="Fold"/>', 'Beta') +
        p('', 'b1') +
        p('<w:pStyle w:val="Fold"/><w15:collapsed w:val="0"/>', 'Gamma') +
        p('', 'g1') +
        p('<w:pStyle w:val="Heading1"/>', 'Delta') +
        p('', 'd1'),
    )
    expect(visible(editor).filter(Boolean)).toEqual(['Alpha', 'Beta', 'Gamma', 'g1', 'Delta', 'd1'])
    expect(editor.state.doc.nodeAt(posOf(editor, 'Alpha'))!.attrs.collapsed).toBe(true)
    expect(editor.state.doc.nodeAt(posOf(editor, 'Beta'))!.attrs.collapsed).toBeNull()
  })

  it('folding by the triangle changes no paragraph and writes nothing', async () => {
    const { editor } = await openDoc(p('<w:pStyle w:val="Heading1"/>', 'Alpha') + p('', 'a1'))
    const before = xmlOf(editor, 'Alpha')
    foldHeading(editor, 'collapse', posOf(editor, 'Alpha'))
    expect(xmlOf(editor, 'Alpha')).toBe(before)
    expect(before).not.toContain('collapsed')
  })

  it('the dialog flag is written as w15:collapsed, folds the heading, and one undo takes it back', async () => {
    const { editor } = await openDoc(
      p('<w:pStyle w:val="Heading1"/>', 'Alpha') +
        p('', 'a1') +
        p('<w:pStyle w:val="Fold"/>', 'Beta') +
        p('', 'b1'),
    )
    editor.commands.setTextSelection(posOf(editor, 'a1') + 1)
    expect(collapsedByDefaultAt(editor)).toEqual({ enabled: false, checked: false })
    expect(setCollapsedByDefault(editor, true)).toBe(false)

    editor.commands.setTextSelection(posOf(editor, 'Alpha') + 1)
    expect(collapsedByDefaultAt(editor)).toEqual({ enabled: true, checked: false })
    setCollapsedByDefault(editor, true)
    expect(xmlOf(editor, 'Alpha')).toContain('<w15:collapsed/>')
    expect(visible(editor)).not.toContain('a1')
    editor.commands.undo()
    expect(xmlOf(editor, 'Alpha')).not.toContain('collapsed')
    expect(visible(editor)).toContain('a1')
    editor.commands.redo()
    expect(visible(editor)).not.toContain('a1')
    editor.commands.undo()

    // a style that folds: unchecking writes the explicit off, checking again clears it
    editor.commands.setTextSelection(posOf(editor, 'Beta') + 1)
    expect(collapsedByDefaultAt(editor)).toEqual({ enabled: true, checked: true })
    setCollapsedByDefault(editor, false)
    expect(xmlOf(editor, 'Beta')).toContain('<w15:collapsed w:val="0"/>')
    expect(visible(editor)).toContain('b1')
    setCollapsedByDefault(editor, true)
    expect(xmlOf(editor, 'Beta')).not.toContain('collapsed')
  })

  it('the reload after a save keeps the folds of the view; opening another document does not', () => {
    const editor = createEditor(NESTED)
    foldHeading(editor, 'collapse', posOf(editor, 'Beta'))
    const json = editor.getJSON()
    editor.chain().setMeta(HEADING_COLLAPSE_KEEP, true).setContent(json).run()
    expect(visible(editor)).toEqual([
      'intro',
      'Alpha',
      'a1',
      'Alpha sub',
      'a2',
      'Alpha deep',
      'a3',
      'Beta',
    ])
    // a paragraph opened below the hidden body stays in view through the reload
    caretAtEnd(editor, 'Beta')
    editor.commands.keyboardShortcut('Enter')
    editor.commands.insertContent('typed below')
    editor.chain().setMeta(HEADING_COLLAPSE_KEEP, true).setContent(editor.getJSON()).run()
    expect(visible(editor).slice(-2)).toEqual(['Beta', 'typed below'])
    editor.commands.setContent(json)
    expect(visible(editor)).toHaveLength(NESTED.length)
  })
})
