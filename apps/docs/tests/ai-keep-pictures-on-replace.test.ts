import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocContext, serializeRangeToHtml } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'
import { setModuleLang } from '../src/renderer/i18n/locale'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
const picture = (marks?: Array<{ type: string; attrs?: Record<string, unknown> }>): JsonNode => ({
  type: 'docInlineImage',
  attrs: { dataUrl: PNG, widthPx: 120, heightPx: 80 },
  ...(marks ? { marks } : {}),
})
const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})
const NUM_IDS = { bullet: null, ordered: null }

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

function pictureCount(editor: Editor): number {
  let n = 0
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docInlineImage') n++
    return true
  })
  return n
}

const replace = (editor: Editor, start: number, end: number, html: string) =>
  executeTool(
    editor,
    {
      id: 't',
      name: 'replace_blocks',
      input: { startBlockIndex: start, endBlockIndex: end, html },
    },
    NUM_IDS,
  )

setModuleLang('en')

describe('pictures are visible to the model and survive a block rewrite when echoed back', () => {
  it('read_blocks shows a picture-only paragraph as an img ref, not as an empty paragraph', () => {
    const editor = createEditor([
      para([picture()]),
      para([text('Caption: '), picture(), text(' end')]),
    ])
    expect(serializeRangeToHtml(editor, 0, 1)).toBe(
      '<p><img src="doc:0.0"></p>\n<p>Caption: <img src="doc:1.0"> end</p>',
    )
  })

  it('the block list names a picture-only paragraph instead of leaving the preview blank', () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    expect(buildDocContext(editor)).toContain('(1 picture, no text)')
  })

  it('a rewrite that omits the picture deletes it and the result says so', async () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    const r = await replace(editor, 0, 1, '<p>New intro</p><p>New body</p>')
    expect(r.isError).toBeFalsy()
    expect(r.output).toContain(
      '1 picture(s) that were in the replaced range are not in the new content',
    )
    expect(pictureCount(editor)).toBe(0)
  })

  it('echoing the img ref keeps the picture, placed where the tag sits, with no warning', async () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    const r = await replace(
      editor,
      0,
      1,
      '<p>Figure 1 <img src="doc:0.0"> shows it</p><p>New body</p>',
    )
    expect(r.isError).toBeFalsy()
    expect(r.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Figure 1  shows it')
    expect(editor.state.doc.child(0).content.child(1).type.name).toBe('docInlineImage')
    expect(editor.state.doc.child(0).content.child(1).attrs.dataUrl).toBe(PNG)
  })

  it('an unknown ref and a foreign <img> are ignored rather than inserted', async () => {
    const editor = createEditor([para([text('Body')])])
    const r = await replace(
      editor,
      0,
      0,
      '<p>A <img src="doc:9.9"> B <img src="https://x/y.png"> C</p>',
    )
    expect(r.isError).toBeFalsy()
    expect(pictureCount(editor)).toBe(0)
    expect(editor.state.doc.child(0).textContent).toBe('A  B  C')
  })

  it('a ref from outside the replaced range copies that picture in without counting as a loss', async () => {
    const editor = createEditor([para([picture()]), para([text('Body')])])
    const r = await replace(editor, 1, 1, '<p>Body with <img src="doc:0.0"></p>')
    expect(r.isError).toBeFalsy()
    expect(r.output).not.toContain('picture(s)')
    expect(pictureCount(editor)).toBe(2)
  })

  it('a withheld picture gets no ref, and a guessed ref to it resolves to nothing', async () => {
    const editor = createEditor([
      para([picture([{ type: 'redaction', attrs: { label: 'private' } }]), text(' after')]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).not.toContain('doc:')
    const r = await replace(editor, 0, 0, '<p>x <img src="doc:0.0"></p>')
    expect(r.isError).toBeFalsy()
    expect(pictureCount(editor)).toBe(0)
  })

  it('a picture pending deletion is neither numbered nor counted as lost', async () => {
    const editor = createEditor([
      para([picture([{ type: 'del', attrs: { author: 'a' } }]), picture(), text(' tail')]),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).toBe('<p><img src="doc:0.0"> tail</p>')
    const r = await replace(editor, 0, 0, '<p>kept <img src="doc:0.0"></p>')
    expect(r.output).not.toContain('picture(s)')
    const live = editor.state.doc.child(0)
    expect(live.content.child(1).type.name).toBe('docInlineImage')
    expect(live.content.child(1).marks.some((m) => m.type.name === 'del')).toBe(false)
  })

  it('replace_selection accepts an echoed ref instead of throwing', async () => {
    const editor = createEditor([para([text('Hello '), picture(), text(' world')])])
    // select "Hello <pic>" (positions 1..8 inside the paragraph)
    editor.commands.setTextSelection({ from: 1, to: 8 })
    const r = await executeTool(
      editor,
      { id: 't', name: 'replace_selection', input: { html: 'Hi <img src="doc:0.0">' } },
      NUM_IDS,
    )
    expect(r.isError, r.output).toBeFalsy()
    expect(pictureCount(editor)).toBe(1)
    expect(editor.state.doc.child(0).textContent).toBe('Hi  world')
  })

  it('echoing one picture twice does not hide that another was dropped', async () => {
    const editor = createEditor([para([picture(), picture()])])
    const r = await replace(editor, 0, 0, '<p><img src="doc:0.0"><img src="doc:0.0"></p>')
    expect(r.output).toContain('1 picture(s) that were in the replaced range')
    expect(pictureCount(editor)).toBe(2)
  })

  it("a carried-over picture does not strip the selection's inherited run marks from the new text", async () => {
    const editor = createEditor([
      para([
        { type: 'text', text: 'Bold ', marks: [{ type: 'bold' }] },
        picture([{ type: 'docTextStyle', attrs: { color: 'FF0000' } }]),
        { type: 'text', text: ' tail', marks: [{ type: 'bold' }] },
      ]),
    ])
    editor.commands.setTextSelection({ from: 1, to: 7 })
    const r = await executeTool(
      editor,
      { id: 't', name: 'replace_selection', input: { html: 'Strong <img src="doc:0.0">' } },
      NUM_IDS,
    )
    expect(r.isError, r.output).toBeFalsy()
    const first = editor.state.doc.child(0).content.child(0)
    expect(first.text).toBe('Strong ')
    expect(first.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })
})
