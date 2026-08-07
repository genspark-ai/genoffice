/**
 * Shape insertion tests - Item 2: insert shape → save → re-parse
 *
 * Verifies:
 *  1. buildShapeParagraphXml produces XML with the correct prstGeom
 *  2. Word structure assertions after saving each shape type (wps:wsp + a:prstGeom)
 *  3. The prst field is correct after re-parsing
 *  4. Double-click enters text editing (has wps:txbx)
 *  5. Text content can be patched in
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
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

const WIDTH_EMU = 1800000
const HEIGHT_EMU = 1080000
const DEFAULT_FILL = '4472C4'
const DEFAULT_BORDER = '2F5496'

async function openBlankDoc() {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

function makeShapeTextbox(prst: string): TextboxDisplay {
  return {
    fill: DEFAULT_FILL,
    borderColor: DEFAULT_BORDER,
    widthPx: Math.round(WIDTH_EMU / 9525),
    heightPx: Math.round(HEIGHT_EMU / 9525),
    prst,
    paras: [{ runs: [{ text: '' }] }],
  }
}

/** Insert shape as docProtected genXml node */
function insertShape(editor: Editor, prst: string) {
  const xml = buildShapeParagraphXml({
    prst,
    widthEmu: WIDTH_EMU,
    heightEmu: HEIGHT_EMU,
    id: 1,
    fillHex: DEFAULT_FILL,
    borderHex: DEFAULT_BORDER,
    withTextbox: true,
  })
  editor
    .chain()
    .insertContentAt(editor.state.doc.content.size, {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: `Shape(${prst})`,
        genXml: xml,
        textboxes: [makeShapeTextbox(prst)],
      },
    })
    .run()
}

describe('shape insertion', () => {
  it('buildShapeParagraphXml generates paragraph XML with the correct prstGeom', () => {
    const prst = 'triangle'
    const xml = buildShapeParagraphXml({
      prst,
      widthEmu: WIDTH_EMU,
      heightEmu: HEIGHT_EMU,
      id: 1,
      withTextbox: true,
      fillHex: DEFAULT_FILL,
      borderHex: DEFAULT_BORDER,
    })
    expect(xml).toContain('wps:wsp')
    expect(xml).toContain(`prstGeom prst="${prst}"`)
    expect(xml).toContain('w:txbxContent')
    expect(xml).toContain('wp:anchor')
    expect(xml).toContain(DEFAULT_FILL)
    expect(xml).toContain(DEFAULT_BORDER)
    expect(xml).toContain('mc:AlternateContent')
    expect(xml).toContain(
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    )
    // mc:Fallback should NOT have xmlns:mc (to allow stripping by parse.ts)
    expect(xml).not.toContain('mc:Fallback xmlns:mc=')
  })

  it('shape without a text box does not generate wps:txbx', () => {
    const xml = buildShapeParagraphXml({
      prst: 'ellipse',
      widthEmu: WIDTH_EMU,
      heightEmu: HEIGHT_EMU,
      id: 1,
      withTextbox: false,
    })
    expect(xml).toContain('wps:wsp')
    expect(xml).toContain('prstGeom prst="ellipse"')
    expect(xml).not.toContain('wps:txbx')
  })

  const SHAPE_TYPES = [
    'rect',
    'roundRect',
    'ellipse',
    'triangle',
    'diamond',
    'pentagon',
    'hexagon',
    'star5',
    'rightArrow',
  ]

  SHAPE_TYPES.forEach((prst) => {
    it(`shape ${prst}: insert→save→Word structure assertions`, async () => {
      const { editor, parsed } = await openBlankDoc()
      insertShape(editor, prst)

      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const xmlBlock = plan.saveBlocks.find((b) => b.kind === 'xml') as
        { kind: 'xml'; xml: string } | undefined
      expect(xmlBlock).toBeDefined()
      expect(xmlBlock?.xml).toContain('wps:wsp')
      expect(xmlBlock?.xml).toContain(`prstGeom prst="${prst}"`)
      expect(xmlBlock?.xml).toContain('w:txbxContent')
      expect(xmlBlock?.xml).toContain('wp:anchor')
      editor.destroy()
    })
  })

  it('shape keeps the prst field after save and reparse', async () => {
    const { editor, parsed } = await openBlankDoc()
    insertShape(editor, 'diamond')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.textboxes)
    expect(block).toBeDefined()
    expect(block?.textboxes?.[0].prst).toBe('diamond')
    expect(block?.textboxes?.[0].fill).toBe(DEFAULT_FILL)
    editor.destroy()
  })

  it('shape text content can be written and reparsed', async () => {
    const { editor, parsed } = await openBlankDoc()
    const xml = buildShapeParagraphXml({
      prst: 'roundRect',
      widthEmu: WIDTH_EMU,
      heightEmu: HEIGHT_EMU,
      id: 2,
      fillHex: DEFAULT_FILL,
      borderHex: DEFAULT_BORDER,
      withTextbox: true,
    })
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Shape(roundRect)',
          genXml: xml,
          textboxes: [
            { ...makeShapeTextbox('roundRect'), paras: [{ runs: [{ text: 'Shape text' }] }] },
          ],
        },
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.textboxes)
    expect(block?.textboxes?.[0].prst).toBe('roundRect')
    expect(block?.textboxes?.[0].paras[0].runs[0].text).toBe('Shape text')
    editor.destroy()
  })

  it('rect shape prst is not stored in TextboxDisplay (only non-rect shapes store prst)', async () => {
    const { editor, parsed } = await openBlankDoc()
    const xml = buildShapeParagraphXml({
      prst: 'rect',
      widthEmu: WIDTH_EMU,
      heightEmu: HEIGHT_EMU,
      id: 3,
      fillHex: 'FFFFFF',
      borderHex: '000000',
      withTextbox: true,
    })
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Shape(rect)',
          genXml: xml,
          textboxes: [
            {
              fill: 'FFFFFF',
              borderColor: '000000',
              widthPx: Math.round(WIDTH_EMU / 9525),
              heightPx: Math.round(HEIGHT_EMU / 9525),
              paras: [{ runs: [{ text: '' }] }],
            },
          ],
        },
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.textboxes)
    // rect is the default prstGeom, should not have prst set in the display model
    expect(block?.textboxes?.[0].prst).toBeUndefined()
    editor.destroy()
  })

  it('inserts and saves a shape at top level when the cursor is inside a table cell', async () => {
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const source = await buildDocx({ bodyXml: table })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    let cellPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (cellPos < 0 && node.type.name === 'docTableCell') cellPos = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, cellPos + 2)),
    )

    insertShapeAt(editor, 'diamond')

    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(1).attrs.textboxes?.[0]?.prst).toBe('diamond')
    const saved = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks.some((block) => block.textboxes?.[0]?.prst === 'diamond')).toBe(true)
    editor.destroy()
  })

  it('inserts a line arrow as a stroke-only connector and round-trips it', async () => {
    const { editor, parsed } = await openBlankDoc()
    insertShapeAt(editor, 'lineArrow')

    const box = editor.state.doc.lastChild?.attrs.textboxes?.[0] as TextboxDisplay
    expect(box.prst).toBe('lineArrow')
    expect(box.readOnly).toBe(true)
    expect(box.fill).toBeUndefined()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xmlBlock = plan.saveBlocks.find((b) => b.kind === 'xml') as
      { kind: 'xml'; xml: string } | undefined
    expect(xmlBlock?.xml).toContain('prst="straightConnector1"')
    expect(xmlBlock?.xml).toContain('<a:tailEnd type="triangle"/>')
    expect(xmlBlock?.xml).toContain('<a:noFill/>')

    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.textboxes?.length)
    expect(block?.textboxes?.[0].prst).toBe('lineArrow')
    expect(block?.textboxes?.[0].readOnly).toBe(true)
    editor.destroy()
  })

  it('straight lines ignore the drawn height; bent connectors keep it', async () => {
    const { editor, parsed } = await openBlankDoc()
    insertShapeAt(editor, 'line', { widthEmu: 2700000, heightEmu: 1800000 })
    insertShapeAt(editor, 'lineBent', { widthEmu: 1800000, heightEmu: 1350000 })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const boxes = reparsed.blocks.flatMap((b) => b.textboxes ?? [])
    const straight = boxes.find((b) => b.prst === 'line')
    const bent = boxes.find((b) => b.prst === 'lineBent')
    // 114300 EMU = 12 px grab band
    expect(straight?.heightPx).toBe(12)
    expect(straight?.widthPx).toBe(Math.round(2700000 / 9525))
    expect(bent?.heightPx).toBe(Math.round(1350000 / 9525))
    editor.destroy()
  })

  it('moves a shape with its handle and persists the floating position', async () => {
    const { editor, parsed } = await openBlankDoc()
    insertShapeAt(editor, 'ellipse')
    const wrapper = editor.view.dom.querySelector('.doc-protected-textboxes') as HTMLElement
    const handle = wrapper.querySelector('.doc-move-handle') as HTMLElement

    handle.dispatchEvent(
      new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }),
    )
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 40, clientY: 30 }))
    window.dispatchEvent(new MouseEvent('mouseup', { clientX: 40, clientY: 30 }))

    const moved = editor.state.doc.lastChild
    expect(moved?.attrs.imageOffsetXEmu).toBe(30 * 9525)
    expect(moved?.attrs.imageOffsetYEmu).toBe(20 * 9525)

    const saved = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const reparsed = await parseDocx(saved)
    const shape = reparsed.blocks.find((block) => block.textboxes?.[0]?.prst === 'ellipse')
    expect(shape?.imageOffsetXEmu).toBe(30 * 9525)
    expect(shape?.imageOffsetYEmu).toBe(20 * 9525)
    editor.destroy()
  })
})
