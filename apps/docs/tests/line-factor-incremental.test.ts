import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Decoration, DecorationSet } from '@tiptap/pm/view'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

const para = (text: string, attrs: Record<string, unknown> = {}) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: text ? [{ type: 'text', text }] : [],
})

function editorWith(content: unknown[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content } as never,
  })
  editors.push(editor)
  return editor
}

/** the live plugin's decoration set, flattened to a comparable list */
function liveDecos(editor: Editor): string[] {
  const plugin = editor.state.plugins.find((p) =>
    String((p as unknown as { key: string }).key).startsWith('lineFactorLive$'),
  )!
  const set = plugin.getState(editor.state) as DecorationSet
  return set
    .find()
    .map((d: Decoration) => {
      const type = (d as unknown as { type: { attrs?: unknown; spec?: unknown } }).type
      return JSON.stringify([d.from, d.to, type.attrs])
    })
    .sort()
}

/** a fresh editor on the same document builds its set from scratch (plugin init) */
function rebuiltDecos(editor: Editor): string[] {
  return liveDecos(editorWith(editor.getJSON().content as unknown[]))
}

const CJK = '中文 English 混排，句号。' // public-hygiene: fixture

describe('lineFactorLive incremental decorations', () => {
  it('matches a full rebuild after edits, appends, splits and deletions', () => {
    const editor = editorWith([
      para('Plain ascii paragraph one.'),
      para(CJK, { align: 'justify' }),
      para('Another line with symbols: 100% + 5 = x', { align: 'justify' }),
      para(''),
    ])
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))
    expect(liveDecos(editor).length).toBeGreaterThan(0)

    // text typed into the CJK paragraph
    const p2 = editor.state.doc.child(0).nodeSize + 1
    editor.view.dispatch(editor.state.tr.insertText('新内容 new', p2 + 3)) // public-hygiene: fixture
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))

    // streamed-style append at the end
    const tailText = '数字 123 年' // public-hygiene: fixture
    const tail = [para(CJK + ' tail'), para('ascii tail'), para(tailText)].map((n) =>
      editor.schema.nodeFromJSON(n),
    )
    editor.view.dispatch(editor.state.tr.insert(editor.state.doc.content.size, tail))
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))

    // split the first paragraph
    editor.view.dispatch(editor.state.tr.split(8))
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))

    // delete the (now third) block outright
    const third = editor.state.doc.child(0).nodeSize + editor.state.doc.child(1).nodeSize
    editor.view.dispatch(editor.state.tr.delete(third, third + editor.state.doc.child(2).nodeSize))
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))

    // mark and attribute steps move no positions but change the block
    editor.commands.setTextSelection({ from: 3, to: 12 })
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 28 })
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))
    editor.commands.unsetMark('docTextStyle')
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...editor.state.doc.child(0).attrs,
        align: 'justify',
      }),
    )
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))

    // replace everything (falls back to the full rebuild path)
    editor.commands.setContent({ type: 'doc', content: [para(CJK), para('x')] } as never)
    expect(liveDecos(editor)).toEqual(rebuiltDecos(editor))
  })
})
