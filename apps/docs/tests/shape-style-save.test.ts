/**
 * Shape rotation and outline width persist through save: a loaded shape has its
 * original drawing bytes patched in place (a:xfrm@rot / a:ln@w), an inserted
 * shape has its generated anchor rewritten.
 */
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  buildShapeParagraphXml,
  parseDocx,
  saveDocx,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

/**
 * The gallery shape buildShapeParagraphXml writes is centered; an explicit
 * posOffset keeps the parsed offsets stable (a margin-aligned X re-resolves
 * per wrap mode in the parser), exactly like shape-wrap-save.test.ts.
 */
const SHAPE_XML = buildShapeParagraphXml({ prst: 'rect', withTextbox: true, id: 42 }).replace(
  '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>',
  '<wp:positionH relativeFrom="column"><wp:posOffset>500000</wp:posOffset></wp:positionH>',
)

/** Same shape authored with a 45deg rotation, to exercise clearing it. */
const ROTATED_SHAPE_XML = SHAPE_XML.replace('<a:xfrm>', '<a:xfrm rot="2700000">')

const BLANK_BODY = '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>'

type ParsedDoc = Awaited<ReturnType<typeof parseDocx>>

async function openShapeDoc(bodyXml = SHAPE_XML) {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

function selectShape(editor: Editor): void {
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (pos === -1 && node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    return true
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
}

/** Write the selected shape's first display box, as the ribbon / rotate handle does. */
function patchFirstBox(editor: Editor, patch: Partial<TextboxDisplay>): void {
  let boxes: TextboxDisplay[] | null = null
  editor.state.doc.descendants((node) => {
    if (boxes === null && node.type.name === 'docProtected' && Array.isArray(node.attrs.textboxes))
      boxes = node.attrs.textboxes as TextboxDisplay[]
    return true
  })
  expect(boxes).toBeTruthy()
  editor.commands.updateAttributes('docProtected', { textboxes: [{ ...boxes![0]!, ...patch }] })
}

function firstBox(parsed: ParsedDoc): TextboxDisplay {
  const block = parsed.blocks.find((b) => b.textboxes?.length)
  return block!.textboxes![0]
}

async function saveShape(editor: Editor, parsed: ParsedDoc): Promise<Uint8Array> {
  return saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks)
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

describe('shape rotation and outline width survive save', () => {
  it('leaves an untouched loaded shape byte-identical', async () => {
    const { editor, parsed, source } = await openShapeDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('persists rotation as a:xfrm@rot in 60000ths of a degree', async () => {
    const { editor, parsed } = await openShapeDoc()
    selectShape(editor)
    patchFirstBox(editor, { rotDeg: 90 })
    const saved = await saveShape(editor, parsed)
    expect(await documentXml(saved)).toContain('rot="5400000"')
    expect(firstBox(await parseDocx(saved)).rotDeg).toBe(90)
    editor.destroy()
  })

  it('persists the outline width as a:ln@w in EMU', async () => {
    const { editor, parsed } = await openShapeDoc()
    selectShape(editor)
    patchFirstBox(editor, { borderWidthPx: 2 })
    const saved = await saveShape(editor, parsed)
    expect(await documentXml(saved)).toMatch(/<a:ln w="19050"/)
    expect(firstBox(await parseDocx(saved)).borderWidthPx).toBe(2)
    editor.destroy()
  })

  it('patches only rotation without touching the shape fill or outline color', async () => {
    const { editor, parsed } = await openShapeDoc()
    selectShape(editor)
    patchFirstBox(editor, { rotDeg: 90 })
    const xml = await documentXml(await saveShape(editor, parsed))
    expect(xml).toContain('<a:srgbClr val="4472C4"/>')
    expect(xml).toContain('<a:srgbClr val="2F5496"/>')
    editor.destroy()
  })

  it('clears the rotation back to 0 and drops the attribute', async () => {
    const { editor, parsed } = await openShapeDoc(ROTATED_SHAPE_XML)
    expect(firstBox(parsed).rotDeg).toBe(45)
    selectShape(editor)
    patchFirstBox(editor, { rotDeg: 0 })
    const saved = await saveShape(editor, parsed)
    const xml = await documentXml(saved)
    expect(xml).not.toMatch(/<a:xfrm[^>]*\brot=/)
    expect(firstBox(await parseDocx(saved)).rotDeg).toBeUndefined()
    editor.destroy()
  })

  it("saves rotation into an inserted shape's generated XML", async () => {
    const { editor, parsed } = await openShapeDoc(BLANK_BODY)
    insertShapeAt(editor, 'rect')
    selectShape(editor)
    patchFirstBox(editor, { rotDeg: 90 })
    const saved = await saveShape(editor, parsed)
    expect(await documentXml(saved)).toContain('rot="5400000"')
    expect(firstBox(await parseDocx(saved)).rotDeg).toBe(90)
    editor.destroy()
  })

  it("saves an outline width into an inserted shape's generated XML", async () => {
    const { editor, parsed } = await openShapeDoc(BLANK_BODY)
    insertShapeAt(editor, 'rect')
    selectShape(editor)
    patchFirstBox(editor, { borderWidthPx: 3 })
    const saved = await saveShape(editor, parsed)
    expect(await documentXml(saved)).toMatch(/<a:ln w="28575"/)
    expect(firstBox(await parseDocx(saved)).borderWidthPx).toBe(3)
    editor.destroy()
  })
})
