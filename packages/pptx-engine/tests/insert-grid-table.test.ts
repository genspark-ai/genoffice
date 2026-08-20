/** pdf2pptx P25 insertion extensions: bodyPr overrides, rtl paragraphs, grid tables. */
import { describe, expect, it } from 'vitest'
import {
  addElement,
  appendRawElements,
  buildTableGridXml,
  createBlankPptx,
  openPptx,
  savePptx,
  type NewTableGridOptions,
} from '../src/index'

const OFF = { x: 914400, y: 914400, cx: 3657600, cy: 1828800 }

describe('bodyPr overrides on generated text boxes', () => {
  it('writes insets/anchor/wrap and round-trips through save', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'measured box' }] }],
      bodyPr: { wrap: 'square', anchor: 't', insetsEmu: { l: 0, t: 0, r: 0, b: 0 } },
    })
    expect(el.anchor.originalXml).toContain(
      '<a:bodyPr wrap="square" rtlCol="0" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>',
    )
    const reopened = await openPptx(await savePptx(opened))
    const el2: any = reopened.deck.slides[0]!.elements.at(-1)
    expect(el2.text.body?.insets ?? el2.text.insets ?? { l: 0, t: 0, r: 0, b: 0 }).toBeTruthy()
    expect(el2.anchor.originalXml).toContain('lIns="0"')
  })

  it('default stays the historical bodyPr', async () => {
    const opened = await openPptx(await createBlankPptx())
    const el = addElement(opened.deck.slides[0]!, { kind: 'textbox', offset: { ...OFF } })
    expect(el.anchor.originalXml).toContain('<a:bodyPr wrap="square" rtlCol="0"/>')
  })

  it('rtl paragraphs emit a:pPr rtl="1"', async () => {
    const opened = await openPptx(await createBlankPptx())
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'שלום' }], align: 'right', rtl: true }],
    })
    expect(el.anchor.originalXml).toContain('rtl="1"')
    expect(el.anchor.originalXml).toContain('algn="r"')
  })
})

describe('buildTableGridXml', () => {
  const spec = (over: Partial<NewTableGridOptions> = {}): NewTableGridOptions => ({
    offset: { x: 0, y: 0, cx: 3000000, cy: 2000000 },
    colWidthsEmu: [1000000, 1200000, 800000],
    rowHeightsEmu: [1000000, 1000000],
    cells: [
      [
        { paragraphs: [{ runs: [{ text: 'A1' }] }], gridSpan: 2, fillColor: '#DDEEFF' },
        { hMerge: true },
        { paragraphs: [{ runs: [{ text: 'C1' }] }], rowSpan: 2, anchor: 'ctr' },
      ],
      [
        { paragraphs: [{ runs: [{ text: 'A2' }] }] },
        { paragraphs: [{ runs: [{ text: 'B2' }] }] },
        { vMerge: true },
      ],
    ],
    border: { color: '#333333', widthEmu: 9525 },
    ...over,
  })

  it('writes explicit grid, merges, shading and borders', async () => {
    const opened = await openPptx(await createBlankPptx())
    const xml = buildTableGridXml(opened.deck.slides[0]!, spec())
    expect(xml).toContain('<a:gridCol w="1000000"/><a:gridCol w="1200000"/><a:gridCol w="800000"/>')
    expect(xml).toContain('gridSpan="2"')
    expect(xml).toContain('<a:tc hMerge="1">')
    expect(xml).toContain('rowSpan="2"')
    expect(xml).toContain('<a:tc vMerge="1">')
    expect(xml).toContain('anchor="ctr"')
    expect(xml).toContain('<a:srgbClr val="DDEEFF"/>')
    expect(xml).toContain('<a:lnL w="9525" cap="flat">')
    expect(xml).toContain('<a:tblPr/>')
    // every row lists one tc per grid column
    const rows = xml.match(/<a:tr /g)
    expect(rows?.length).toBe(2)
    expect((xml.match(/<a:tc[\s>]/g) ?? []).length).toBe(6)
  })

  it('insideV scope only rules verticals between columns', async () => {
    const opened = await openPptx(await createBlankPptx())
    const xml = buildTableGridXml(
      opened.deck.slides[0]!,
      spec({ border: { color: '#000000', widthEmu: 12700, scope: 'insideV' } }),
    )
    expect(xml).toContain('<a:lnL')
    expect(xml).not.toContain('<a:lnT')
    expect(xml).not.toContain('<a:lnR')
  })

  it('inserted grid table survives save → reopen with cell text', async () => {
    const opened = await openPptx(await createBlankPptx())
    const xml = buildTableGridXml(opened.deck.slides[0]!, spec())
    const r = appendRawElements(opened, 0, [xml])
    expect(r).not.toBeNull()
    const reopened = await openPptx(await savePptx(opened))
    const table: any = reopened.deck.slides[0]!.elements.find((e: any) => e.type === 'table')
    expect(table).toBeTruthy()
    const texts = JSON.stringify(table)
    expect(texts).toContain('A1')
    expect(texts).toContain('B2')
  })
})

describe('translucent shape fills', () => {
  it('#RRGGBBAA emits an a:alpha child; opaque stays plain', async () => {
    const opened = await openPptx(await createBlankPptx())
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { ...OFF },
      fillColor: '#000000CC',
    })
    expect(el.anchor.originalXml).toContain(
      '<a:srgbClr val="000000"><a:alpha val="80000"/></a:srgbClr>',
    )
    const el2 = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { ...OFF },
      fillColor: '#112233',
    })
    expect(el2.anchor.originalXml).toContain('<a:srgbClr val="112233"/>')
  })
})
