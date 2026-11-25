import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { equationBlockJson, inlineEquationNodeJson } from '../src/renderer/editor/equation'
import { editorExtensions } from '../src/renderer/editor/extensions'

const V5 = [5, 1, 0, 7, 0, 68, 83, 77, 84, 55, 0, 1]
const EOC = 0xfffffffe
const FREE = 0xffffffff
const word = (value: number) => [value & 255, value >> 8]
const ch = (glyph: string, face = 3) => [2, 0, 128 + face, ...word(glyph.charCodeAt(0))]
const line = (...records: number[][]) => [1, 0, ...records.flat(), 0]
const template = (selector: number, variation: number, ...records: number[][]) => [
  3,
  0,
  selector,
  variation,
  0,
  ...records.flat(),
  0,
]

function buildMathTypeOle(mtef: Uint8Array): string {
  const sectorSize = 512
  const headerSize = 28
  const nativeSize = headerSize + mtef.length
  const miniStreamSize = Math.ceil(nativeSize / 64) * 64
  const dataSectors = Math.ceil(miniStreamSize / sectorSize)
  const dataId = 2
  const fatId = dataId + dataSectors
  const bytes = new Uint8Array((fatId + 2) * sectorSize)
  const data = new DataView(bytes.buffer)
  const u16 = (offset: number, value: number) => data.setUint16(offset, value, true)
  const u32 = (offset: number, value: number) => data.setUint32(offset, value, true)
  const sector = (id: number) => (id + 1) * sectorSize
  const directory = sector(0)
  const miniFat = sector(1)
  const fat = sector(fatId)

  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
  u16(24, 0x3e)
  u16(26, 3)
  u16(28, 0xfffe)
  u16(30, 9)
  u16(32, 6)
  u32(44, 1)
  u32(48, 0)
  u32(56, 4096)
  u32(60, 1)
  u32(64, 1)
  u32(68, EOC)
  u32(72, 0)
  for (let i = 0; i < 109; i++) u32(76 + 4 * i, i === 0 ? fatId : FREE)

  bytes.fill(255, fat, fat + sectorSize)
  u32(fat, EOC)
  u32(fat + 4, EOC)
  u32(fat + fatId * 4, 0xfffffffd)
  for (let i = 0; i < dataSectors; i++) {
    u32(fat + (dataId + i) * 4, i + 1 === dataSectors ? EOC : dataId + i + 1)
  }

  const writeDir = (
    offset: number,
    name: string,
    type: number,
    child: number,
    start: number,
    size: number,
  ) => {
    for (let i = 0; i < name.length; i++) u16(offset + i * 2, name.charCodeAt(i))
    u16(offset + 64, (name.length + 1) * 2)
    bytes.set([type, 1], offset + 66)
    u32(offset + 68, FREE)
    u32(offset + 72, FREE)
    u32(offset + 76, child)
    u32(offset + 116, start)
    u32(offset + 120, size)
  }
  const miniCount = miniStreamSize / 64
  writeDir(directory, 'Root Entry', 5, 1, dataId, miniStreamSize)
  writeDir(directory + 128, 'Equation Native', 2, FREE, 0, nativeSize)
  bytes.fill(255, miniFat, miniFat + sectorSize)
  for (let i = 0; i < miniCount; i++) {
    u32(miniFat + i * 4, i + 1 === miniCount ? EOC : i + 1)
  }

  const nativeOffset = sector(dataId)
  u16(nativeOffset, headerSize)
  u32(nativeOffset + 2, 0x20000)
  u32(nativeOffset + 8, mtef.length)
  bytes.set(mtef, nativeOffset + headerSize)
  return Buffer.from(bytes).toString('base64')
}

const SAMPLE_OLE_B64 = buildMathTypeOle(
  Uint8Array.from([
    ...V5,
    ...line(ch('x'), ch('=', 6), template(11, 0, line(ch('a')), line(ch('b')))),
    0,
  ]),
)
const PNG_1X1_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

const OBJECT_XML =
  '<w:object w:dxaOrig="1440" w:dyaOrig="720">' +
  '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:36pt" o:ole="">' +
  '<v:imagedata r:id="rId30" o:title=""/>' +
  '</v:shape>' +
  '<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" ' +
  'DrawAspect="Content" ObjectID="_1700000002" r:id="rId31"/>' +
  '</w:object>'

const RELS =
  '<Relationship Id="rId30" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
  '<Relationship Id="rId31" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>'

async function openMathTypeDoc(bodyXml: string, oleBase64 = SAMPLE_OLE_B64) {
  const source = await buildDocx({
    bodyXml,
    extraRels: RELS,
    binaryParts: [
      {
        path: 'word/media/image1.png',
        base64: PNG_1X1_B64,
        extension: 'png',
        contentType: 'image/png',
      },
      {
        path: 'word/embeddings/oleObject1.bin',
        base64: oleBase64,
        extension: 'bin',
        contentType: 'application/vnd.openxmlformats-officedocument.oleObject',
      },
    ],
  })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('MathType editor rendering and conversion', () => {
  it('renders inline MathType as display-style MathML without negative position overlap and keeps untouched save clean', async () => {
    const { editor, parsed } = await openMathTypeDoc(
      '<w:p>' +
        '<w:r><w:tab/></w:r>' +
        `<w:r><w:rPr><w:position w:val="-50"/><w:sz w:val="24"/></w:rPr>${OBJECT_XML}</w:r>` +
        '<w:r><w:tab/><w:t>(5)</w:t></w:r>' +
        '</w:p>',
    )

    const mathEl = editor.view.dom.querySelector('.doc-inline-math')
    expect(mathEl).not.toBeNull()
    expect(mathEl!.innerHTML).toContain('<mfrac>')
    expect(mathEl!.innerHTML).toContain('displaystyle="true"')
    // WMF negative w:position must not shift the native MathML down into the next line
    const styledParent = mathEl!.closest('[style*="top:25pt"]')
    expect(styledParent).toBeNull()

    // Untouched document reports zero dirty blocks
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    editor.destroy()
  })

  it('marks an inline MathType equation dirty when converted to native OMML even with identical LaTeX', async () => {
    const { editor, parsed } = await openMathTypeDoc(
      '<w:p>' +
        '<w:r><w:t xml:space="preserve">Eq: </w:t></w:r>' +
        `<w:r><w:rPr><w:position w:val="-20"/></w:rPr>${OBJECT_XML}</w:r>` +
        '</w:p>',
    )

    let mathPos = -1
    let mathLatex = ''
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docInlineMath') {
        mathPos = pos
        mathLatex = String(node.attrs.latex ?? '')
      }
    })
    expect(mathPos).toBeGreaterThanOrEqual(0)
    expect(mathLatex).toBe('x=\\frac{a}{b}')

    // Convert in-place using the same attrs produced by EquationModal
    const next = inlineEquationNodeJson(mathLatex)
    editor.view.dispatch(editor.state.tr.setNodeMarkup(mathPos, undefined, next.attrs))

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const run = reparsed.blocks[0].runs?.find((r) => r.math)
    expect(run?.math?.oleXml).toBeUndefined()
    expect(run?.math?.omml).toContain('<m:f>')
    editor.destroy()
  })

  it('renders standalone MathType blocks as vector MathML with an edit/convert button and converts cleanly', async () => {
    const { editor, parsed } = await openMathTypeDoc(`<w:p><w:r>${OBJECT_XML}</w:r></w:p>`)

    const host = editor.view.dom.querySelector('.doc-formula-math')
    expect(host).not.toBeNull()
    expect(host!.innerHTML).toContain('<mfrac>')
    expect(editor.view.dom.querySelector('.doc-formula-edit')).not.toBeNull()

    const planBefore = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(planBefore.changedCount).toBe(0)

    // Replace with native OMML block (same path as EquationModal)
    const next = editor.schema.nodeFromJSON(equationBlockJson('x=\\frac{a}{b}'))
    const first = editor.state.doc.child(0)
    editor.view.dispatch(editor.state.tr.replaceWith(0, first.nodeSize, next))

    const planAfter = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(planAfter.changedCount).toBe(1)
    const saved = await saveDocx(parsed, planAfter.saveBlocks)
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].formulaDisplay?.oleMathType).toBeUndefined()
    expect(reparsed.blocks[0].formulaDisplay?.mathml).toContain('<mfrac>')
    editor.destroy()
  })

  it('uses vertical-align instead of relative top offset for lowered fallback inline OLE images', async () => {
    const { editor, parsed } = await openMathTypeDoc(
      '<w:p>' +
        '<w:r><w:tab/></w:r>' +
        `<w:r><w:rPr><w:position w:val="-50"/></w:rPr>${OBJECT_XML}</w:r>` +
        '<w:r><w:tab/><w:t>(5)</w:t></w:r>' +
        '</w:p>',
      'AAECAwQ=',
    )

    const img = editor.view.dom.querySelector('img.doc-inline-img') as HTMLImageElement | null
    expect(img).not.toBeNull()
    expect(img!.style.verticalAlign).toBe('-25pt')
    expect(img!.closest('[style*="top:25pt"]')).toBeNull()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    editor.destroy()
  })

  it('renders and converts a MathType equation inside a table cell while keeping untouched tables clean', async () => {
    const { editor, parsed } = await openMathTypeDoc(
      '<w:tbl><w:tr><w:tc><w:p>' +
        '<w:r><w:tab/></w:r>' +
        `<w:r><w:rPr><w:position w:val="-50"/><w:sz w:val="24"/></w:rPr>${OBJECT_XML}</w:r>` +
        '<w:r><w:tab/><w:t>(5)</w:t></w:r>' +
        '</w:p></w:tc></w:tr></w:tbl>',
    )

    const mathEl = editor.view.dom.querySelector('td .doc-inline-math')
    expect(mathEl).not.toBeNull()
    expect(mathEl!.innerHTML).toContain('<mfrac>')

    const planBefore = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(planBefore.changedCount).toBe(0)

    let mathPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docInlineMath') mathPos = pos
    })
    expect(mathPos).toBeGreaterThanOrEqual(0)
    const next = inlineEquationNodeJson('x=\\frac{a}{b}')
    editor.view.dispatch(editor.state.tr.setNodeMarkup(mathPos, undefined, next.attrs))

    const planAfter = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(planAfter.changedCount).toBe(1)
    const saved = await saveDocx(parsed, planAfter.saveBlocks)
    const reparsed = await parseDocx(saved)
    const cellRuns = reparsed.blocks[0].table?.rows[0]?.[0]?.richParas?.[0]?.runs
    const mathRun = cellRuns?.find((r) => r.math)
    expect(mathRun?.math?.oleXml).toBeUndefined()
    expect(mathRun?.math?.omml).toContain('<m:f>')
    editor.destroy()
  })
})
