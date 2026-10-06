import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { DocRedaction } from '../src/renderer/editor/redaction'
import { liveText } from '../src/renderer/ai/doc-utils'

/**
 * liveText is the one place the document's text is read on its way to a model,
 * so a span has to be redacted here. A caller left reading the raw text would
 * hand over the whole document, and nothing downstream would notice.
 */

const SECRET = '13800138000'
const extensions = [...editorExtensions, DocRedaction]
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

async function open(blocks: Array<{ runs: Array<{ text: string }> }>): Promise<Editor> {
  const bodyXml = blocks
    .map(
      (b) =>
        `<w:p>${b.runs.map((r) => `<w:r><w:t xml:space="preserve">${r.text}</w:t></w:r>`).join('')}</w:p>`,
    )
    .join('')
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions,
    content: { type: 'doc', content: blocksToPmDoc(parsed.blocks).content as JsonNode[] },
  })
  editors.add(editor)
  return editor
}

function selectText(editor: Editor, needle: string) {
  let start = -1
  editor.state.doc.descendants((node, pos) => {
    if (start >= 0) return
    const at = node.textContent.indexOf(needle)
    if (at >= 0) start = pos + 1 + at
  })
  expect(start).toBeGreaterThan(0)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, start, start + needle.length),
    ),
  )
}

const firstBlock = (editor: Editor) => editor.state.doc.child(0)

describe('liveText redacts for a model', () => {
  it('replaces the withheld words with the marker', async () => {
    const editor = await open([{ runs: [{ text: `Call ${SECRET} now` }] }])
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    expect(liveText(firstBlock(editor))).toBe('Call {{客户电话}} now')
  })

  it('never returns the secret to a model', async () => {
    const editor = await open([{ runs: [{ text: `Call ${SECRET} now` }] }])
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    expect(liveText(firstBlock(editor))).not.toContain(SECRET)
  })

  it('leaves an unmarked block alone', async () => {
    const editor = await open([{ runs: [{ text: 'nothing secret' }] }])
    expect(liveText(firstBlock(editor))).toBe('nothing secret')
  })

  it('falls back to a neutral marker for a blank label', async () => {
    const editor = await open([{ runs: [{ text: `Call ${SECRET} now` }] }])
    selectText(editor, SECRET)
    editor.commands.setRedaction('x')
    const marks = firstBlock(editor).child(1)?.marks ?? []
    expect(liveText(firstBlock(editor))).toContain('{{')
    expect(marks.some((m) => m.type.name === 'redaction')).toBe(true)
  })

  it('shows the real words when the reader is the audience', async () => {
    // a comment preview is for a person: redacting it would show the reader a
    // document that differs from what they can see on the page
    const editor = await open([{ runs: [{ text: `Call ${SECRET} now` }] }])
    selectText(editor, SECRET)
    editor.commands.setRedaction('客户电话')
    expect(liveText(firstBlock(editor), true)).toBe(`Call ${SECRET} now`)
  })
})
