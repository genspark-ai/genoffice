/**
 * A shape read from a file that holds no text keeps a new rotation, wrap,
 * stacking rank and size through save, alone or beside a shape with text.
 */
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { bringToFront, setFloatingWrap, shapeWrapOf } from '../src/renderer/editor/floating-z-order'

const EMU = 9525

/** one anchored rectangle as Word writes it: no wps:cNvPr, so only its place in the paragraph names it */
const drawing = (id: number, x: number, text = '') =>
  `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
  `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${x * EMU}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${20 * EMU}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${120 * EMU}" cy="${80 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapNone/><wp:docPr id="${id}" name="Rectangle ${id}"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${120 * EMU}" cy="${80 * EMU}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>` +
  (text
    ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>`
    : '') +
  `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`

const TEXTLESS = `<w:p>${drawing(7, 40)}</w:p>`
const MIXED = `<w:p>${drawing(7, 40)}${drawing(8, 260, 'Label')}</w:p>`

type ParsedDoc = Awaited<ReturnType<typeof parseDocx>>

async function open(bodyXml: string) {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (pos === -1 && node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    return true
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return { editor, parsed, source, pos }
}

const boxesOf = (parsed: ParsedDoc): TextboxDisplay[] =>
  parsed.blocks.find((b) => b.textboxes?.length)!.textboxes!

function patchBox(editor: Editor, pos: number, index: number, patch: Partial<TextboxDisplay>) {
  const boxes = editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
  editor.commands.updateAttributes('docProtected', {
    textboxes: boxes.map((box, i) => (i === index ? { ...box, ...patch } : box)),
  })
}

async function save(editor: Editor, parsed: ParsedDoc) {
  const bytes = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
  )
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return { xml, reparsed: await parseDocx(bytes) }
}

describe('a shape without text saves what is changed on it', () => {
  it('saves untouched as the bytes it was read from', async () => {
    const { editor, parsed, source } = await open(MIXED)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('keeps a rotation', async () => {
    const { editor, parsed, pos } = await open(TEXTLESS)
    expect(boxesOf(parsed)[0].paras).toEqual([])
    patchBox(editor, pos, 0, { rotDeg: 90 })
    const { xml, reparsed } = await save(editor, parsed)
    expect(xml).toContain('<a:xfrm rot="5400000">')
    expect(boxesOf(reparsed)[0].rotDeg).toBe(90)
    editor.destroy()
  })

  it('keeps a wrap and its place', async () => {
    const { editor, parsed } = await open(TEXTLESS)
    const original = boxesOf(parsed)[0]
    setFloatingWrap(editor, 'behind')
    const { xml, reparsed } = await save(editor, parsed)
    expect(xml).toContain('behindDoc="1"')
    const box = boxesOf(reparsed)[0]
    expect(shapeWrapOf(box)).toBe('behind')
    expect(box.offsetXEmu).toBe(original.offsetXEmu)
    expect(box.offsetYEmu).toBe(original.offsetYEmu)
    editor.destroy()
  })

  it('keeps a stacking rank', async () => {
    const { editor, parsed } = await open(TEXTLESS)
    bringToFront(editor)
    const { reparsed } = await save(editor, parsed)
    expect(boxesOf(reparsed)[0].z).toBe(boxesOf(parsed)[0].z! + 1)
    editor.destroy()
  })

  it('keeps a new size', async () => {
    const { editor, parsed, pos } = await open(TEXTLESS)
    patchBox(editor, pos, 0, { widthPx: 200, heightPx: 100 })
    const { xml, reparsed } = await save(editor, parsed)
    expect(xml).toContain(`<wp:extent cx="${200 * EMU}" cy="${100 * EMU}"/>`)
    expect(boxesOf(reparsed)[0]).toMatchObject({ widthPx: 200, heightPx: 100 })
    editor.destroy()
  })

  it('turns each of two shapes in a paragraph, the one without text and the one with', async () => {
    const { editor, parsed, pos } = await open(MIXED)
    patchBox(editor, pos, 0, { rotDeg: 90 })
    patchBox(editor, pos, 1, { rotDeg: 180 })
    const { reparsed } = await save(editor, parsed)
    const [bare, labelled] = boxesOf(reparsed)
    expect(bare.rotDeg).toBe(90)
    expect(bare.paras).toEqual([])
    expect(labelled.rotDeg).toBe(180)
    expect(labelled.paras[0].runs[0].text).toBe('Label')
    editor.destroy()
  })
})
