import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { hasUnanchoredComments } from '../src/renderer/file-actions'
import { editorExtensions } from '../src/renderer/editor/extensions'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const COMMENTS_XML =
  XML_DECL +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:comment w:id="7" w:author="Alice"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment>' +
  '</w:comments>'

const COMMENTED_P =
  '<w:p><w:r><w:t>before </w:t></w:r>' +
  '<w:commentRangeStart w:id="7"/>' +
  '<w:r><w:t>marked</w:t></w:r>' +
  '<w:commentRangeEnd w:id="7"/>' +
  '<w:r><w:commentReference w:id="7"/></w:r>' +
  '<w:r><w:t> after</w:t></w:r></w:p>'

async function loadEditor() {
  const bytes = await buildDocx({
    bodyXml: COMMENTED_P,
    extraParts: [
      {
        path: 'word/comments.xml',
        xml: COMMENTS_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
      },
    ],
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { parsed, editor }
}

describe('comment marks in the editor', () => {
  it('renders the covered run as a comment mark and round-trips untouched as original', async () => {
    const { parsed, editor } = await loadEditor()
    expect(editor.view.dom.querySelector('span[data-comment-ids="7"]')?.textContent).toBe('marked')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(plan.saveBlocks[0]).toEqual({ kind: 'original', docxIndex: 0 })
    editor.destroy()
  })

  it('keeps commentIds on the regenerated block when the paragraph is edited', async () => {
    const { parsed, editor } = await loadEditor()
    editor.commands.insertContentAt(editor.state.doc.content.size - 1, '!')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const block = plan.saveBlocks[0]
    if (block.kind !== 'generated') throw new Error(`expected generated, got ${block.kind}`)
    expect(block.block.runs).toEqual([
      { text: 'before ' },
      { text: 'marked', commentIds: ['7'] },
      { text: ' after!' },
    ])
    editor.destroy()
  })
})

describe('unanchored comments open the panel on load', () => {
  const withComments = (bodyXml: string) =>
    buildDocx({
      bodyXml,
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: COMMENTS_XML,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
    })

  it('false for a bare reference in an empty paragraph (the margin bubble anchors it to the block)', async () => {
    const parsed = await parseDocx(
      await withComments('<w:p><w:r><w:commentReference w:id="7"/></w:r></w:p>'),
    )
    expect(hasUnanchoredComments(parsed.comments, parsed.blocks)).toBe(false)
  })

  it('true when the comment id never appears in the body', async () => {
    const parsed = await parseDocx(await withComments('<w:p><w:r><w:t>hi</w:t></w:r></w:p>'))
    expect(hasUnanchoredComments(parsed.comments, parsed.blocks)).toBe(true)
  })

  it('false when the reference anchors on a text run', async () => {
    const parsed = await parseDocx(
      await withComments(
        '<w:p><w:r><w:t>hi</w:t></w:r><w:r><w:commentReference w:id="7"/></w:r></w:p>',
      ),
    )
    expect(hasUnanchoredComments(parsed.comments, parsed.blocks)).toBe(false)
  })

  it('false for a ranged comment', async () => {
    const parsed = await parseDocx(await withComments(COMMENTED_P))
    expect(hasUnanchoredComments(parsed.comments, parsed.blocks)).toBe(false)
  })
})
