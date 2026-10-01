/**
 * Shape Format wrap and stacking persist through save: a loaded shape has its
 * original drawing bytes patched in place (position preserved), an inserted
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
import { bringToFront, setFloatingWrap, shapeWrapOf } from '../src/renderer/editor/floating-z-order'

/**
 * The gallery shape buildShapeParagraphXml writes is centered; an explicit
 * posOffset mirrors a dragged shape and keeps the parsed offsets stable across
 * wrap modes (a margin-aligned X re-resolves per mode in the parser).
 */
const SHAPE_XML = buildShapeParagraphXml({ prst: 'rect', withTextbox: true, id: 42 }).replace(
  '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>',
  '<wp:positionH relativeFrom="column"><wp:posOffset>500000</wp:posOffset></wp:positionH>',
)

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

async function saveShape(editor: Editor, parsed: ParsedDoc): Promise<Uint8Array> {
  return saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks)
}

function firstBox(parsed: ParsedDoc): TextboxDisplay {
  const block = parsed.blocks.find((b) => b.textboxes?.length)
  return block!.textboxes![0]
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

describe('shape wrap and stacking survive save', () => {
  it('leaves an untouched loaded shape byte-identical', async () => {
    const { editor, parsed, source } = await openShapeDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('persists behind wrap on a loaded shape, keeping its position', async () => {
    const { editor, parsed } = await openShapeDoc()
    const original = firstBox(parsed)
    selectShape(editor)
    setFloatingWrap(editor, 'behind')
    const reparsed = await parseDocx(await saveShape(editor, parsed))
    const box = firstBox(reparsed)
    expect(shapeWrapOf(box)).toBe('behind')
    expect(box.behind).toBe(true)
    expect(box.offsetXEmu).toBe(original.offsetXEmu)
    expect(box.offsetYEmu).toBe(original.offsetYEmu)
    editor.destroy()
  })

  it('persists a stacking rank on a loaded shape', async () => {
    const { editor, parsed } = await openShapeDoc()
    selectShape(editor)
    bringToFront(editor)
    const reparsed = await parseDocx(await saveShape(editor, parsed))
    expect(firstBox(reparsed).z).toBe(1)
    editor.destroy()
  })

  it('persists square -> front on a loaded shape, keeping its position', async () => {
    const { editor, parsed } = await openShapeDoc()
    const original = firstBox(parsed)
    selectShape(editor)
    setFloatingWrap(editor, 'front')
    const reparsed = await parseDocx(await saveShape(editor, parsed))
    const box = firstBox(reparsed)
    expect(box.floating).toBe(true)
    expect(box.behind).toBeFalsy()
    expect(box.offsetXEmu).toBe(original.offsetXEmu)
    expect(box.offsetYEmu).toBe(original.offsetYEmu)
    editor.destroy()
  })

  it('converts a loaded shape back to inline', async () => {
    const { editor, parsed } = await openShapeDoc()
    selectShape(editor)
    setFloatingWrap(editor, null)
    const saved = await saveShape(editor, parsed)
    const box = firstBox(await parseDocx(saved))
    expect(box.floating).toBeFalsy()
    expect(shapeWrapOf(box)).toBeNull()
    expect(await documentXml(saved)).not.toContain('<wp:anchor')
    editor.destroy()
  })

  it('persists behind wrap and rank on an inserted shape', async () => {
    const { editor, parsed } = await openShapeDoc('<w:p><w:r><w:t>Body text</w:t></w:r></w:p>')
    insertShapeAt(editor, 'rect')
    selectShape(editor)
    setFloatingWrap(editor, 'behind')
    bringToFront(editor)
    const reparsed = await parseDocx(await saveShape(editor, parsed))
    const box = firstBox(reparsed)
    expect(box.behind).toBe(true)
    expect(box.z).toBe(1)
    editor.destroy()
  })
})
