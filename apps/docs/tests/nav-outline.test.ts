import { afterEach, describe, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectHeadings, type HeadingRef } from '../src/renderer/editor/headings'
import { compileFind, findInDoc, type FindMatch } from '../src/renderer/editor/find'
import {
  buildHeadingTree,
  deleteHeadingSubtree,
  headingAtPos,
  headingKey,
  hitCounts,
  insertHeadingSibling,
  moveHeadingSubtree,
  resultContext,
  setHeadingLevel,
  subtreeRange,
  visibleRows,
} from '../src/renderer/editor/nav-outline'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string): JSONContent => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: text ? [{ type: 'text', text }] : [],
})
const heading = (level: number, text: string): JSONContent => ({
  type: 'docHeading',
  attrs: { level },
  content: [{ type: 'text', text }],
})

function createEditor(content: JSONContent[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

const outline = (editor: Editor) =>
  editor.state.doc.content.content.map((n) =>
    n.type.name === 'docHeading' ? `h${n.attrs.level}:${n.textContent}` : `p:${n.textContent}`,
  )

const refs = (levels: number[]): HeadingRef[] =>
  levels.map((level, i) => ({ level, text: `H${i}`, pos: i * 10 }))

const SAMPLE = [
  heading(1, 'Intro'),
  para('intro body'),
  heading(2, 'Scope'),
  para('scope body'),
  heading(3, 'Detail'),
  para('detail body'),
  heading(1, 'Methods'),
  para('methods body'),
  heading(2, 'Sampling'),
  para('sampling body'),
]

describe('buildHeadingTree', () => {
  it('nests each heading under the nearest earlier heading with a smaller level', () => {
    const tree = buildHeadingTree(refs([1, 2, 3, 1, 2]))
    expect(tree.map((n) => n.ref.text)).toEqual(['H0', 'H3'])
    expect(tree[0].children.map((n) => n.ref.text)).toEqual(['H1'])
    expect(tree[0].children[0].children.map((n) => n.ref.text)).toEqual(['H2'])
    expect(tree[1].children.map((n) => n.ref.text)).toEqual(['H4'])
    expect(tree[1].children[0].parent).toBe(3)
  })

  it('a skipped level nests under the closer ancestor; a deeper first heading is a root', () => {
    const tree = buildHeadingTree(refs([3, 1, 3]))
    expect(tree.map((n) => n.ref.text)).toEqual(['H0', 'H1'])
    expect(tree[1].children.map((n) => n.ref.text)).toEqual(['H2'])
  })

  it('visible rows hide collapsed subtrees and levels past the limit', () => {
    const tree = buildHeadingTree(refs([1, 2, 3, 1, 2]))
    const all = visibleRows(tree, new Set(), 9)
    expect(all.map((r) => [r.node.ref.text, r.depth, r.hasChildren])).toEqual([
      ['H0', 0, true],
      ['H1', 1, true],
      ['H2', 2, false],
      ['H3', 0, true],
      ['H4', 1, false],
    ])
    const collapsed = visibleRows(tree, new Set([headingKey({ level: 1, text: 'H0' })]), 9)
    expect(collapsed.map((r) => r.node.ref.text)).toEqual(['H0', 'H3', 'H4'])
    expect(collapsed[0].collapsed).toBe(true)
    const twoLevels = visibleRows(tree, new Set(), 2)
    expect(twoLevels.map((r) => r.node.ref.text)).toEqual(['H0', 'H1', 'H3', 'H4'])
    // a node whose only children are hidden by the level limit shows no disclosure triangle
    expect(twoLevels[1].hasChildren).toBe(false)
  })
})

describe('subtreeRange', () => {
  it('runs from the heading to the next heading of the same or a higher level', () => {
    const editor = createEditor(SAMPLE)
    const doc = editor.state.doc
    const hs = collectHeadings(doc)
    const size = doc.content.size
    const text = (r: { from: number; to: number }) => doc.textBetween(r.from, r.to, '|')
    expect(text(subtreeRange(hs, 0, size))).toBe(
      'Intro|intro body|Scope|scope body|Detail|detail body',
    )
    expect(text(subtreeRange(hs, 1, size))).toBe('Scope|scope body|Detail|detail body')
    expect(text(subtreeRange(hs, 2, size))).toBe('Detail|detail body')
    expect(subtreeRange(hs, 4, size).to).toBe(size)
  })

  it('headingAtPos picks the heading whose section holds the caret', () => {
    const hs = refs([1, 2, 1])
    expect(headingAtPos(hs, 0)).toBe(0)
    expect(headingAtPos(hs, 15)).toBe(1)
    expect(headingAtPos(hs, 20)).toBe(2)
    expect(headingAtPos(hs, 99)).toBe(2)
    expect(headingAtPos([{ level: 1, text: 'x', pos: 5 }], 2)).toBe(-1)
  })
})

describe('hitCounts', () => {
  it('counts own hits per heading and rolls them up the subtree', () => {
    const editor = createEditor(SAMPLE)
    const doc = editor.state.doc
    const hs = collectHeadings(doc)
    const compiled = compileFind('body', { matchCase: false, wholeWord: false, mode: 'literal' })
    if (!compiled.ok) throw new Error(compiled.error)
    const hits = findInDoc(doc, compiled.query, { matchCase: false, wholeWord: false })
    const counts = hitCounts(hs, hits, doc.content.size)
    expect(counts.own).toEqual([1, 1, 1, 1, 1])
    expect(counts.subtree).toEqual([3, 2, 1, 2, 1])
  })
})

describe('resultContext', () => {
  const match = (blockText: string, offset: number, text: string): FindMatch => ({
    from: 0,
    to: 0,
    text,
    groups: null,
    blockText,
    offset,
  })
  it('keeps 40 characters on either side with ellipses only where text was cut', () => {
    const left = 'L'.repeat(50)
    const right = 'R'.repeat(50)
    const ctx = resultContext(match(`${left}hit${right}`, 50, 'hit'))
    expect(ctx.before).toBe('…' + 'L'.repeat(40))
    expect(ctx.hit).toBe('hit')
    expect(ctx.after).toBe('R'.repeat(40) + '…')
  })
  it('short blocks carry no ellipses', () => {
    const ctx = resultContext(match('say hit now', 4, 'hit'))
    expect(ctx).toEqual({ before: 'say ', hit: 'hit', after: ' now' })
  })
})

describe('subtree edits', () => {
  it('moves a heading with its content before another heading in one transaction', () => {
    const editor = createEditor(SAMPLE)
    const hs = collectHeadings(editor.state.doc)
    const tr = editor.state.tr
    expect(moveHeadingSubtree(tr, hs, 3, 0, 'before')).toBe(true)
    editor.view.dispatch(tr)
    expect(outline(editor)).toEqual([
      'h1:Methods',
      'p:methods body',
      'h2:Sampling',
      'p:sampling body',
      'h1:Intro',
      'p:intro body',
      'h2:Scope',
      'p:scope body',
      'h3:Detail',
      'p:detail body',
    ])
    expect(editor.can().undo()).toBe(true)
    editor.commands.undo()
    expect(outline(editor)).toEqual(
      SAMPLE.map(
        (n) => `${n.type === 'docHeading' ? `h${n.attrs?.level}` : 'p'}:${n.content?.[0]?.text}`,
      ),
    )
  })

  it('moves a nested heading after a later subtree, keeping its own content', () => {
    const editor = createEditor(SAMPLE)
    const hs = collectHeadings(editor.state.doc)
    const tr = editor.state.tr
    expect(moveHeadingSubtree(tr, hs, 1, 4, 'after')).toBe(true)
    editor.view.dispatch(tr)
    expect(outline(editor)).toEqual([
      'h1:Intro',
      'p:intro body',
      'h1:Methods',
      'p:methods body',
      'h2:Sampling',
      'p:sampling body',
      'h2:Scope',
      'p:scope body',
      'h3:Detail',
      'p:detail body',
    ])
  })

  it('refuses drops inside the dragged subtree and onto itself', () => {
    const editor = createEditor(SAMPLE)
    const hs = collectHeadings(editor.state.doc)
    expect(moveHeadingSubtree(editor.state.tr, hs, 0, 2, 'before')).toBe(false)
    expect(moveHeadingSubtree(editor.state.tr, hs, 0, 2, 'after')).toBe(false)
    expect(moveHeadingSubtree(editor.state.tr, hs, 0, 0, 'before')).toBe(false)
    // before the heading that already follows the subtree is a no-op position
    expect(moveHeadingSubtree(editor.state.tr, hs, 0, 3, 'before')).toBe(false)
  })

  it('deletes heading and content, leaving a paragraph when nothing remains', () => {
    const editor = createEditor(SAMPLE)
    let hs = collectHeadings(editor.state.doc)
    let tr = editor.state.tr
    expect(deleteHeadingSubtree(tr, editor.schema, hs, 1)).toBe(true)
    editor.view.dispatch(tr)
    expect(outline(editor)).toEqual([
      'h1:Intro',
      'p:intro body',
      'h1:Methods',
      'p:methods body',
      'h2:Sampling',
      'p:sampling body',
    ])
    const solo = createEditor([heading(1, 'Only'), para('body')])
    hs = collectHeadings(solo.state.doc)
    tr = solo.state.tr
    expect(deleteHeadingSubtree(tr, solo.schema, hs, 0)).toBe(true)
    solo.view.dispatch(tr)
    expect(outline(solo)).toEqual(['p:'])
  })

  it('promotes, demotes and clamps heading levels', () => {
    const editor = createEditor(SAMPLE)
    const hs = collectHeadings(editor.state.doc)
    let tr = editor.state.tr
    expect(setHeadingLevel(tr, editor.schema, hs[2].pos, 2)).toBe(true)
    editor.view.dispatch(tr)
    expect(outline(editor)[4]).toBe('h2:Detail')
    tr = editor.state.tr
    expect(setHeadingLevel(tr, editor.schema, hs[0].pos, 0)).toBe(false)
    tr = editor.state.tr
    expect(setHeadingLevel(tr, editor.schema, hs[0].pos, 12)).toBe(true)
    editor.view.dispatch(tr)
    expect(outline(editor)[0]).toBe('h9:Intro')
  })

  it('turns a style-based heading paragraph into the style of the new level', () => {
    const editor = createEditor([
      { ...para('Styled'), attrs: { docxIndex: null, styleId: 'Heading2' } },
      para('body'),
    ])
    const styles = new Map([
      ['Heading1', { headingLevel: 1 }],
      ['Heading2', { headingLevel: 2 }],
    ])
    const tr = editor.state.tr
    expect(setHeadingLevel(tr, editor.schema, 0, 1, styles)).toBe(true)
    editor.view.dispatch(tr)
    const first = editor.state.doc.firstChild!
    expect(first.type.name).toBe('docParagraph')
    expect(first.attrs.styleId).toBe('Heading1')
  })

  it('inserts an empty sibling heading before or after the subtree with the caret inside', () => {
    const editor = createEditor(SAMPLE)
    const hs = collectHeadings(editor.state.doc)
    const tr = editor.state.tr
    expect(insertHeadingSibling(tr, editor.schema, hs, 1, 'after')).toBe(true)
    editor.view.dispatch(tr)
    expect(outline(editor)).toEqual([
      'h1:Intro',
      'p:intro body',
      'h2:Scope',
      'p:scope body',
      'h3:Detail',
      'p:detail body',
      'h2:',
      'h1:Methods',
      'p:methods body',
      'h2:Sampling',
      'p:sampling body',
    ])
    expect(editor.state.selection.$from.parent.type.name).toBe('docHeading')
    expect(editor.state.selection.$from.parent.textContent).toBe('')
    // the pane lists the blank heading right away; the TOC collector still skips it
    expect(collectHeadings(editor.state.doc, undefined, true).map((h) => h.text)).toContain('')
    expect(collectHeadings(editor.state.doc).map((h) => h.text)).not.toContain('')
  })
})
