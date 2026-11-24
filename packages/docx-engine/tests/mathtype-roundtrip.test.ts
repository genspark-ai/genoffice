import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  latexToOmml,
  mathParagraphXml,
  ommlToLatex,
  ommlToMathML,
  parseDocx,
  saveDocx,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

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

/** Build a minimal valid OLE2 container carrying an Equation Native MTEF v5 stream */
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

/** MTEF v5 for: a = b / c */
const SAMPLE_MTEF = Uint8Array.from([
  ...V5,
  ...line(ch('a'), ch('=', 6), template(11, 0, line(ch('b')), line(ch('c')))),
  0,
])
const SAMPLE_OLE_B64 = buildMathTypeOle(SAMPLE_MTEF)

const OBJECT_XML =
  '<w:object w:dxaOrig="1440" w:dyaOrig="720">' +
  '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:36pt" o:ole="">' +
  '<v:imagedata r:id="rId30" o:title=""/>' +
  '</v:shape>' +
  '<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" ' +
  'DrawAspect="Content" ObjectID="_1700000001" r:id="rId31"/>' +
  '</w:object>'

const RELS =
  '<Relationship Id="rId30" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
  '<Relationship Id="rId31" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>'

// 1x1 transparent PNG placeholder as fallback preview
const PNG_1X1_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

async function buildMathTypeDocx(bodyXml: string, oleBase64 = SAMPLE_OLE_B64): Promise<Uint8Array> {
  return buildDocx({
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
}

describe('MathType OLE decoding and round-trip', () => {
  it('decodes a standalone MathType paragraph to vector MathML and preserves OLE bytes on untouched save', async () => {
    const src = await buildMathTypeDocx(
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr>${OBJECT_XML}</w:r></w:p>`,
    )
    const parsed = await parseDocx(src)
    const block = parsed.blocks[0]

    expect(block.type).toBe('passthrough')
    expect(block.oleProgId).toBe('Equation.DSMT4')
    expect(block.formulaDisplay?.oleMathType).toBe(true)
    expect(block.formulaDisplay?.mathml).toContain('<mfrac>')
    expect(block.formulaDisplay?.latex).toBe('a=\\frac{b}{c}')
    expect(block.formulaDisplay?.sizeHalfPoints).toBe(24)
    expect(block.formulaDisplay?.align).toBe('center')

    // Untouched save preserves original <w:object> and oleObject1.bin byte-for-byte
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: block.docxIndex! }])
    const outZip = await JSZip.loadAsync(saved)
    const docXml = await outZip.file('word/document.xml')!.async('string')
    expect(docXml).toContain('ProgID="Equation.DSMT4"')
    expect(docXml).not.toContain('<m:oMath>')
    const binOut = await outZip.file('word/embeddings/oleObject1.bin')!.async('base64')
    expect(binOut).toBe(SAMPLE_OLE_B64)
  })

  it('decodes an inline MathType equation with equation number and preserves OLE when sibling text is edited', async () => {
    const src = await buildMathTypeDocx(
      '<w:p>' +
        '<w:r><w:tab/></w:r>' +
        `<w:r><w:rPr><w:position w:val="-50"/><w:sz w:val="24"/></w:rPr>${OBJECT_XML}</w:r>` +
        '<w:r><w:tab/><w:t>(5)</w:t></w:r>' +
        '</w:p>',
    )
    const parsed = await parseDocx(src)
    const block = parsed.blocks[0]
    expect(block.type).toBe('paragraph')

    const mathRun = block.runs?.find((r) => r.math)
    expect(mathRun).toBeDefined()
    expect(mathRun?.math?.oleProgId).toBe('Equation.DSMT4')
    expect(mathRun?.math?.oleXml).toContain('<w:object')
    expect(mathRun?.math?.oleRawRPr).toContain('w:position w:val="-50"')
    expect(mathRun?.sizeHalfPoints).toBe(24)
    expect(mathRun?.positionHalfPoints).toBe(-50)
    expect(ommlToLatex(mathRun!.math!.omml)).toBe('a=\\frac{b}{c}')

    // Edit sibling "(5)" -> "(6)" while keeping the MathType run untouched
    const editedRuns = block.runs!.map((r) =>
      r.text.includes('(5)') ? { ...r, text: '\t(6)' } : r,
    )
    const saved = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs: editedRuns } },
    ])
    const outZip = await JSZip.loadAsync(saved)
    const docXml = await outZip.file('word/document.xml')!.async('string')
    expect(docXml).toContain('(6)')
    expect(docXml).toContain('ProgID="Equation.DSMT4"')
    expect(docXml).toContain('w:position w:val="-50"')
    expect(docXml).not.toContain('<m:oMath>')
  })

  it('writes native OMML when a MathType equation is explicitly converted', async () => {
    const src = await buildMathTypeDocx(
      '<w:p>' +
        '<w:r><w:t xml:space="preserve">Eq: </w:t></w:r>' +
        `<w:r><w:rPr><w:position w:val="-20"/></w:rPr>${OBJECT_XML}</w:r>` +
        '</w:p>',
    )
    const parsed = await parseDocx(src)
    const block = parsed.blocks[0]
    const mathRun = block.runs!.find((r) => r.math)!

    // Clearing oleXml simulates the user converting the MathType equation to native OMML
    const convertedRuns = block.runs!.map((r) =>
      r === mathRun ? { text: r.text, math: { omml: r.math!.omml } } : r,
    )
    const saved = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs: convertedRuns } },
    ])
    const reparsed = await parseDocx(saved)
    const reparsedMath = reparsed.blocks[0].runs?.find((r) => r.math)
    expect(reparsedMath).toBeDefined()
    expect(reparsedMath?.math?.oleXml).toBeUndefined()
    expect(ommlToLatex(reparsedMath!.math!.omml)).toBe('a=\\frac{b}{c}')
  })

  it('decodes MathType OLE equations inside table cells', async () => {
    const src = await buildMathTypeDocx(
      `<w:tbl><w:tr><w:tc><w:p><w:r>${OBJECT_XML}</w:r></w:p></w:tc></w:tr></w:tbl>`,
    )
    const parsed = await parseDocx(src)
    const cellRuns = parsed.blocks[0].table?.rows[0]?.[0]?.richParas?.[0]?.runs
    const mathRun = cellRuns?.find((r) => r.math)
    expect(mathRun?.math?.oleXml).toContain('Equation.DSMT4')
    expect(ommlToLatex(mathRun!.math!.omml)).toBe('a=\\frac{b}{c}')
  })

  it('supports converting a standalone MathType block into a native oMathPara paragraph', async () => {
    const src = await buildMathTypeDocx(`<w:p><w:r>${OBJECT_XML}</w:r></w:p>`)
    const parsed = await parseDocx(src)
    const latex = parsed.blocks[0].formulaDisplay!.latex!
    const saved = await saveDocx(parsed, [
      { kind: 'xml', xml: mathParagraphXml(latexToOmml(latex)) },
    ])
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].formulaDisplay?.oleMathType).toBeUndefined()
    expect(reparsed.blocks[0].formulaDisplay?.mathml).toContain('<mfrac>')
  })

  it('renders and round-trips bold and script math alphabets in MathML and LaTeX', () => {
    const latex =
      '\\mathbf{w}_c^{\\text{cov}} = \\sum_{k \\in \\mathcal{K}_c} \\omega_{k,c} \\mathbf{w}_c^k'
    const omml = `<m:oMath>${latexToOmml(latex)}</m:oMath>`
    const mathml = ommlToMathML(omml)
    // Unicode Mathematical Bold Small W (U+1D430) and Mathematical Script Capital K (U+1D4A6)
    expect(mathml).toContain('\u{1D430}')
    expect(mathml).toContain('\u{1D4A6}')
    const decompiled = ommlToLatex(omml)
    expect(decompiled).toContain('\\mathbf{w}')
    expect(decompiled).toContain('\\mathcal{K}')
    expect(ommlToMathML(`<m:oMath>${latexToOmml(decompiled!)}</m:oMath>`)).toBe(mathml)
  })
})
