import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { DecorationSet } from '@tiptap/pm/view'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setPageGaps } from '../src/renderer/editor/pagination-gaps'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: t ? [{ type: 'text', text: t }] : [],
})

const m = { marginTop: 96, marginBottom: 96, marginLeft: 90, marginRight: 90 }

const liveEditors: Editor[] = []
afterEach(() => {
  for (const e of liveEditors.splice(0)) e.destroy()
})

function createEditor(content: JsonNode[]): Editor {
  const element = document.createElement('div')
  document.body.appendChild(element)
  const editor = new Editor({
    element,
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

/** the page gap widgets: position, kind and whether they render mid-paragraph */
function gapsOf(editor: Editor): { pos: number; kind: string; inline: boolean }[] {
  const plugin = editor.state.plugins.find((p) =>
    (p as unknown as { key: string }).key.startsWith('paginationGaps$'),
  )!
  const set = plugin.getState(editor.state) as DecorationSet
  return set
    .find(undefined, undefined, (spec) => spec.gapKind != null)
    .map((d) => ({
      pos: d.from,
      kind: (d.spec as { gapKind: string }).gapKind,
      inline: editor.state.doc.resolve(d.from).parent.inlineContent,
    }))
}

const pageDomGaps = (editor: Editor) => editor.view.dom.querySelectorAll('.page-gap').length

describe('page gaps across edits next to them', () => {
  // "Hello world " ends page 1, "máxima de 25MB" opens page 2 (one paragraph)
  const text = 'Hello world máxima de 25MB'
  const breakAt = 1 + text.indexOf('máxima')

  it('a backspace at the start of a page keeps the mid-paragraph break', () => {
    const editor = createEditor([para(text), para('after')])
    setPageGaps(editor.view, [{ pos: breakAt, kind: 'inline', metrics: m }])
    expect(gapsOf(editor)).toEqual([{ pos: breakAt, kind: 'inline', inline: true }])
    // the space before "máxima": the character on the widget's side goes
    editor.view.dispatch(editor.state.tr.delete(breakAt - 1, breakAt))
    expect(gapsOf(editor)).toEqual([{ pos: breakAt - 1, kind: 'inline', inline: true }])
    expect(pageDomGaps(editor)).toBe(1)
    // a forward delete at the end of page 1 is the same edit
    editor.view.dispatch(editor.state.tr.delete(breakAt - 2, breakAt - 1))
    expect(gapsOf(editor)).toEqual([{ pos: breakAt - 2, kind: 'inline', inline: true }])
  })

  it('a selection cut across the break leaves it at the cut', () => {
    const editor = createEditor([para(text)])
    setPageGaps(editor.view, [{ pos: breakAt, kind: 'inline', metrics: m }])
    editor.view.dispatch(editor.state.tr.delete(breakAt - 6, breakAt + 3))
    expect(gapsOf(editor)).toEqual([{ pos: breakAt - 6, kind: 'inline', inline: true }])
  })

  it('a cut of whole pages drops the break inside it, as before', () => {
    const long = `${'x'.repeat(150)} ${'y'.repeat(150)}`
    const editor = createEditor([para(long)])
    setPageGaps(editor.view, [{ pos: 152, kind: 'inline', metrics: m }])
    editor.view.dispatch(editor.state.tr.delete(2, 300))
    expect(gapsOf(editor)).toEqual([])
    expect(pageDomGaps(editor)).toBe(0)
  })

  it('typing next to the break leaves it alone', () => {
    const editor = createEditor([para(text)])
    setPageGaps(editor.view, [{ pos: breakAt, kind: 'inline', metrics: m }])
    editor.view.dispatch(editor.state.tr.insertText('ab', breakAt - 1))
    expect(gapsOf(editor)).toEqual([{ pos: breakAt + 2, kind: 'inline', inline: true }])
  })

  it('deleting the empty paragraph that ends a page keeps the block break before the next one', () => {
    const editor = createEditor([para('page one'), para(''), para('page two')])
    const lead = editor.view.dom.children[2] as HTMLElement
    setPageGaps(editor.view, [{ el: lead, metrics: m }])
    const at = gapsOf(editor)[0].pos
    expect(gapsOf(editor)).toEqual([{ pos: at, kind: 'block', inline: false }])
    // the empty paragraph is the two tokens right before the gap
    editor.view.dispatch(editor.state.tr.delete(at - 2, at))
    expect(gapsOf(editor)).toEqual([{ pos: at - 2, kind: 'block', inline: false }])
    expect(pageDomGaps(editor)).toBe(1)
  })

  it('joining the page-leading paragraph into the one before keeps the break at the join', () => {
    const editor = createEditor([para('end of page one'), para('start of page two')])
    const lead = editor.view.dom.children[1] as HTMLElement
    setPageGaps(editor.view, [{ el: lead, metrics: m }])
    const at = gapsOf(editor)[0].pos
    // backspace at the start of page two: the paragraph boundary goes
    editor.view.dispatch(editor.state.tr.delete(at - 1, at + 1))
    expect(gapsOf(editor)).toEqual([{ pos: at - 1, kind: 'inline', inline: true }])
    const gap = editor.view.dom.querySelector('.page-gap')!
    expect(gap.classList.contains('page-gap-inline')).toBe(true)
    expect(gap.parentElement?.textContent).toBe('end of page onestart of page two')
  })

  it('a later pass replaces the kept break', () => {
    const editor = createEditor([para(text)])
    setPageGaps(editor.view, [{ pos: breakAt, kind: 'inline', metrics: m }])
    editor.view.dispatch(editor.state.tr.delete(breakAt - 1, breakAt))
    setPageGaps(editor.view, [{ pos: 1 + 'Hello '.length, kind: 'inline', metrics: m }])
    expect(gapsOf(editor)).toEqual([{ pos: 7, kind: 'inline', inline: true }])
    expect(pageDomGaps(editor)).toBe(1)
  })
})
