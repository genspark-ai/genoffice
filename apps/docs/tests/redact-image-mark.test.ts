import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { modelTextOf, redactionCount } from '../src/renderer/ai/redact-view'

/**
 * Why the right-click item is offered for text only.
 *
 * These are characterisation tests: they pin down what the mark does to a
 * picture, so the decision to leave pictures out of the menu has a reason
 * behind it rather than a shrug. They are the empirical result of asking
 * whether a picture can be withheld, and they must keep passing.
 *
 * A picture alone in a paragraph is a *block* node (docProtected), and marks do
 * not apply to block nodes. A picture beside text, or in a table cell, is an
 * inline atom, and there the mark does stick and the model view is right — but
 * the .docx save path builds an image run from the node's attributes and never
 * looks at the node's marks, so the mark is dropped on save and the load path
 * never rebuilds it. Text takes neither of these detours: the mark rides the
 * run's rPr in both directions, which is what makes it safe to offer.
 */

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

const DRAWING = IMAGE_PARAGRAPH_XML.replace(/^<w:p>|<\/w:p>$/g, '')
/** text and a picture in one paragraph: the picture becomes an inline node */
const INLINE_XML = `<w:p><w:r><w:t>before </w:t></w:r>${DRAWING}</w:p>`
const BLOCK_XML = IMAGE_PARAGRAPH_XML
/** the same picture, but its run's rPr carries the element the mark rides on */
const INLINE_REDACT_XML = INLINE_XML.replace(
  '<w:r><w:drawing>',
  '<w:r><w:rPr><w:bdr w:val="single" w:sz="6" w:color="6B4FC0"/>' +
    '<go:redact w:label="id-photo"/></w:rPr><w:drawing>',
)

async function open(bodyXml: string) {
  const source = await buildDocx({ bodyXml, withImage: true })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed }
}

function findNode(editor: Editor, type: string): { node: any; pos: number } {
  let node: unknown = null
  let pos = -1
  editor.state.doc.descendants((n, p) => {
    if (node === null && n.type.name === type) {
      node = n
      pos = p
    }
  })
  return { node: node as never, pos }
}

const markNames = (n: { marks: readonly { type: { name: string } }[] }) =>
  n.marks.map((m) => m.type.name)

async function roundTrip(editor: Editor, parsed: Awaited<ReturnType<typeof open>>['parsed']) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const saved = await saveDocx(parsed, plan.saveBlocks)
  const reparsed = await parseDocx(saved)
  const reloaded = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(reparsed.blocks) as never,
  })
  editors.add(reloaded)
  return reloaded
}

describe('a picture cannot be withheld, and the reason is the docx round trip', () => {
  it('a picture alone in a paragraph is a block node, and blocks take no marks', async () => {
    const { editor } = await open(BLOCK_XML)
    expect(findNode(editor, 'docInlineImage').node).toBeNull()
    const { node, pos } = findNode(editor, 'docProtected')
    expect(node.type.name).toBe('docProtected')
    expect(node.attrs.blockType).toBe('image')

    // addMark is accepted and does nothing at all, with no error to warn anyone
    const mark = editor.state.schema.marks.redaction.create({ label: 'id-photo' })
    editor.view.dispatch(editor.state.tr.addMark(pos, pos + 1, mark))
    expect(markNames(findNode(editor, 'docProtected').node)).toEqual([])
  })

  it('an inline picture does take the mark, and the model view is right', async () => {
    const { editor } = await open(INLINE_XML)
    const { pos } = findNode(editor, 'docInlineImage')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    expect(editor.commands.setRedaction('id-photo')).toBe(true)
    expect(markNames(findNode(editor, 'docInlineImage').node)).toContain('redaction')
    // no path or data reaches the model. The marker is the whole of what the
    // picture contributes; the text run's trailing space is not carried across
    // the atom, which is existing behaviour of the view, not of the mark.
    expect(modelTextOf(editor.getJSON() as never)).toBe('before{{id-photo}}')
    expect(JSON.stringify(modelTextOf(editor.getJSON() as never))).not.toContain('base64')
  })

  it('but the mark on an inline picture is dropped by the save', async () => {
    // This is the reason the menu offers the item for text only. The decision
    // looks right in the editor and is gone after one save, with no warning.
    const { editor, parsed } = await open(INLINE_XML)
    const { pos } = findNode(editor, 'docInlineImage')
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    editor.commands.setRedaction('id-photo')
    expect(redactionCount(editor.getJSON() as never)).toBe(1)

    const reloaded = await roundTrip(editor, parsed)
    expect(markNames(findNode(reloaded, 'docInlineImage').node)).toEqual([])
    expect(redactionCount(reloaded.getJSON() as never)).toBe(0)
  })

  it('and the load path does not rebuild the mark either', async () => {
    // the mirror image of the save gap: a file that really does carry the
    // element on a picture run comes back with the mark missing
    const { editor } = await open(INLINE_REDACT_XML)
    const { node } = findNode(editor, 'docInlineImage')
    // the raw rPr is kept on the node, so the bytes survive; the mark is not rebuilt
    expect(String(node.attrs.rawRPr ?? '')).toContain('go:redact')
    expect(markNames(node)).toEqual([])
    expect(redactionCount(editor.getJSON() as never)).toBe(0)
  })

  it('the same treatment on text round trips, which is what makes it safe to offer', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>secret</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)
    editor.commands.setTextSelection({ from: 1, to: 7 })
    editor.commands.setRedaction('client')
    expect(redactionCount(editor.getJSON() as never)).toBe(1)

    const reloaded = await roundTrip(editor, parsed)
    expect(redactionCount(reloaded.getJSON() as never)).toBe(1)
    // the words are still there; only the model's view changed
    expect(reloaded.state.doc.textContent).toBe('secret')
    expect(modelTextOf(reloaded.getJSON() as never)).toBe('{{client}}')
  })
})
