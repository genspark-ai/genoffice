import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { DecorationSet } from '@tiptap/pm/view'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

const EA_HINT_RPR = '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
const styledSpace = {
  type: 'text',
  text: '   ',
  marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 72 } }],
}
const hintedQuote = (text: string) => ({
  type: 'text',
  text,
  marks: [{ type: 'docTextStyle', attrs: { rawRPr: EA_HINT_RPR } }],
})
const para = (...content: unknown[]) => ({ type: 'docParagraph', content })

function decosOf(editor: Editor, pluginName: string): string[] {
  const plugin = editor.state.plugins.find((p) =>
    String((p as unknown as { key: string }).key).startsWith(pluginName),
  )!
  const set = (plugin.props.decorations as (s: unknown) => DecorationSet | null).call(
    plugin,
    editor.state,
  )
  return (set?.find() ?? [])
    .map(
      (d) =>
        `${d.from}-${d.to}:${(d.spec as { class?: string }).class ?? ''}${JSON.stringify((d as unknown as { type: { attrs: unknown } }).type.attrs)}`,
    )
    .sort()
}

/** the same decorations a plugin computes from scratch for the current document */
function rebuilt(editor: Editor, pluginName: string): string[] {
  const probe = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: editor.state.doc.toJSON(),
  })
  editors.push(probe)
  return decosOf(probe, pluginName)
}

describe('incremental text decorations', () => {
  it('whitespace-run and hinted-quote decorations follow edits exactly like a rebuild', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          para({ type: 'text', text: 'lead ' }, styledSpace, { type: 'text', text: 'tail' }),
          para(hintedQuote('“quoted” plain'), { type: 'text', text: ' after' }),
          para({ type: 'text', text: 'plain paragraph' }),
        ],
      } as never,
    })
    editors.push(editor)
    for (const name of ['wsRunLineHeight', 'eaHintQuotes']) {
      expect(decosOf(editor, name).length).toBeGreaterThan(0)
      expect(decosOf(editor, name)).toEqual(rebuilt(editor, name))
    }

    const check = () => {
      for (const name of ['wsRunLineHeight', 'eaHintQuotes'])
        expect(decosOf(editor, name)).toEqual(rebuilt(editor, name))
    }
    // typing before the decorated runs shifts them
    editor.view.dispatch(editor.state.tr.insertText('xx', 2))
    check()
    // typing into the whitespace run retires it
    const wsPos = editor.state.doc.child(0).firstChild!.nodeSize + 2
    editor.view.dispatch(editor.state.tr.insertText('a', wsPos))
    check()
    // streamed-style append at the end brings new decorated runs
    const tail = [para(hintedQuote('‘more’'), { type: 'text', text: ' x' }), para(styledSpace)].map(
      (n) => editor.schema.nodeFromJSON(n),
    )
    editor.view.dispatch(editor.state.tr.insert(editor.state.doc.content.size, tail))
    check()
    // delete a whole block ahead of the decorated ones
    editor.view.dispatch(editor.state.tr.delete(0, editor.state.doc.child(0).nodeSize))
    check()
    // mark change without position moves
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 28 })
    check()
    // full replacement
    editor.commands.setContent({
      type: 'doc',
      content: [para(styledSpace), para(hintedQuote('“q”'))],
    } as never)
    check()
  })
})
