import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openPptx,
  savePptx,
  addElement,
  addTable,
  deleteElement,
  createBlankPptx,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 3657600, cy: 914400 }

describe('add/delete element', () => {
  it('added textbox survives save → reopen with text and geometry', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const before = slide.elements.length
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'Newly inserted textbox', bold: true }] }],
    })
    expect(el.type).toBe('text')
    expect(slide.elements.length).toBe(before + 1)

    const reopened = await openPptx(await savePptx(opened))
    const slide2 = reopened.deck.slides[0]!
    expect(slide2.elements.length).toBe(before + 1)
    const el2: any = slide2.elements[slide2.elements.length - 1]
    expect(el2.type).toBe('text')
    expect(el2.transform.offset).toEqual(OFF)
    const text = el2.text.paragraphs
      .flatMap((p: any) => p.runs)
      .map((r: any) => r.text)
      .join('')
    expect(text).toBe('Newly inserted textbox')
    expect(el2.text.paragraphs[0].runs[0].bold).toBe(true)
  })

  it('added shape round-trips preset geometry and solid fill', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    addElement(slide, { kind: 'ellipse', offset: { ...OFF }, fillColor: '#C43E1C' })

    const reopened = await openPptx(await savePptx(opened))
    const el2: any = reopened.deck.slides[0]!.elements.at(-1)
    expect(el2.type).toBe('shape')
    expect(el2.presetGeometry).toBe('ellipse')
    expect(el2.fill).toEqual({ type: 'solid', color: '#C43E1C' })
  })

  it('run fontFamily writes both latin and ea slots and round-trips', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: '中文 Latin', fontFamily: '微软雅黑' }] }],
    })
    expect(el.anchor.originalXml).toContain('<a:latin typeface="微软雅黑"/>')
    expect(el.anchor.originalXml).toContain('<a:ea typeface="微软雅黑"/>')

    const reopened = await openPptx(await savePptx(opened))
    const el2: any = reopened.deck.slides[0]!.elements.at(-1)
    expect(el2.text.paragraphs[0].runs[0].fontFamily).toBe('微软雅黑')
  })

  it('cNvPr ids stay unique after two inserts', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const a = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const b = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const idOf = (xml: string) => /<p:cNvPr\s[^>]*\bid="(\d+)"/.exec(xml)![1]
    expect(idOf(a.anchor.originalXml)).not.toBe(idOf(b.anchor.originalXml))
  })

  it('delete element persists through save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const before = slide.elements.length
    expect(before).toBeGreaterThan(1)
    const victim = slide.elements[0]!
    expect(deleteElement(opened, slide, victim.id)).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides[0]!.elements.length).toBe(before - 1)
  })

  it('delete-only edit still marks the deck dirty for save', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    deleteElement(opened, slide, slide.elements[0]!.id)
    expect(slide.structureDirty).toBe(true)
  })
})

describe('insert-time shape options (genpptx parity)', () => {
  it('adjustments write prstGeom avLst guides and round-trip', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'roundRect',
      offset: { ...OFF },
      adjustments: { adj: 25000 },
    })
    expect(el.anchor.originalXml).toContain(
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>',
    )
    // The in-memory model must carry the guides too, or shape handles reset them
    expect(el.adjust).toEqual({ adj: 25000 })
    const reopened = await openPptx(await savePptx(opened))
    const el2: any = reopened.deck.slides[0]!.elements.at(-1)
    expect(el2.presetGeometry).toBe('roundRect')
    expect(el2.anchor.originalXml).toContain('fmla="val 25000"')
    expect(el2.adjust).toEqual({ adj: 25000 })
  })

  it('connector kinds carry adjustments into avLst and the model', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'lineBent',
      offset: { ...OFF },
      adjustments: { adj1: 30000 },
    })
    expect(el.anchor.originalXml).toContain(
      '<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 30000"/></a:avLst></a:prstGeom>',
    )
    expect(el.adjust).toEqual({ adj1: 30000 })
  })

  it('bodyPr autoFit "shrink"/"resize" emit normAutofit/spAutoFit children', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const shrink = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: 'x' }] }],
      bodyPr: { autoFit: 'shrink' },
    })
    expect(shrink.anchor.originalXml).toContain('<a:normAutofit/></a:bodyPr>')
    expect(shrink.text!.autofit).toBe('shrink')
    const grow = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      bodyPr: { anchor: 'ctr', autoFit: 'resize' },
    })
    expect(grow.anchor.originalXml).toContain('anchor="ctr"><a:spAutoFit/></a:bodyPr>')
    expect(grow.text!.autofit).toBe('resize')
  })

  it('click-to-type text box: wrap="none" + spAutoFit round-trips at the PowerPoint 14.5x29pt size', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    // PowerPoint for Mac drops a 14.5 x 29 pt empty box on click (1 pt = 12700 EMU)
    const offset = { x: 914400, y: 914400, cx: 14.5 * 12700, cy: 29 * 12700 }
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...offset },
      bodyPr: { wrap: 'none', autoFit: 'resize' },
    })
    expect(el.anchor.originalXml).toContain(
      '<a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr>',
    )
    expect(el.text!.wrap).toBe(false)
    expect(el.text!.autofit).toBe('resize')

    const reopened = await openPptx(await savePptx(opened))
    const slide2 = reopened.deck.slides[0]!
    const el2: any = slide2.elements[slide2.elements.length - 1]
    expect(el2.type).toBe('text')
    expect(el2.transform.offset).toEqual(offset)
    expect(el2.text.wrap).toBe(false)
    expect(el2.text.autofit).toBe('resize')
  })
})

describe('addTable explicit grid options (genpptx parity)', () => {
  it('length-matched colWidthsEmu/rowHeightsEmu override the equal split', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 3000000, cy: 1000000 },
      colWidthsEmu: [1000000, 2000000],
      rowHeightsEmu: [400000, 600000],
    })!
    const el: any = opened.deck.slides[0]!.elements.find((e) => e.id === r.elementId)
    expect(el.anchor.originalXml).toContain('<a:gridCol w="1000000"/><a:gridCol w="2000000"/>')
    expect(el.anchor.originalXml).toContain('<a:tr h="400000">')
    expect(el.anchor.originalXml).toContain('<a:tr h="600000">')
  })

  it('cellProps write merges and vertical anchors per cell', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 2,
      offset: { x: 0, y: 0, cx: 2000000, cy: 1000000 },
      cellProps: [
        [{ gridSpan: 2, anchor: 'ctr' }, { hMerge: true }],
        [undefined, { anchor: 'b' }],
      ],
    })!
    const xml = (opened.deck.slides[0]!.elements.find((e) => e.id === r.elementId) as any).anchor
      .originalXml as string
    expect(xml).toContain('<a:tc gridSpan="2">')
    expect(xml).toContain('<a:tc hMerge="1">')
    expect(xml).toContain('<a:tcPr anchor="ctr"/>')
    expect(xml).toContain('<a:tcPr anchor="b"/>')
    const reopened = await openPptx(await savePptx(opened))
    const tbl: any = reopened.deck.slides[0]!.elements.at(-1)
    expect(tbl.type).toBe('table')
  })

  it('length-mismatched lists fall back to the equal split', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 3,
      offset: { x: 0, y: 0, cx: 3000000, cy: 1000000 },
      colWidthsEmu: [1, 2],
    })!
    const el: any = opened.deck.slides[0]!.elements.find((e) => e.id === r.elementId)
    expect(el.anchor.originalXml).toContain('<a:gridCol w="1000000"/>'.repeat(3))
  })

  it('clamps hostile dims and spans instead of throwing or emitting NaN', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const start = Date.now()
    const r = addTable(opened, 0, {
      rows: NaN,
      cols: Infinity,
      offset: { x: 0, y: 0, cx: 3000000, cy: 1000000 },
      cellProps: [[{ gridSpan: Infinity }]],
    })!
    expect(Date.now() - start).toBeLessThan(10000)
    const xml = (opened.deck.slides[0]!.elements.find((e) => e.id === r.elementId) as any).anchor
      .originalXml as string
    expect(xml).not.toContain('NaN')
    expect(xml).not.toContain('Infinity')
    expect(xml.match(/<a:tr /g)?.length).toBeLessThanOrEqual(75)
  })

  it('clamps spans to the cells remaining right of / below the cell', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addTable(opened, 0, {
      rows: 2,
      cols: 3,
      offset: { x: 0, y: 0, cx: 3000000, cy: 1000000 },
      cellProps: [[undefined, { gridSpan: 5 }, { gridSpan: 3, rowSpan: 4 }], [{ rowSpan: 2 }]],
    })!
    const xml = (opened.deck.slides[0]!.elements.find((e) => e.id === r.elementId) as any).anchor
      .originalXml as string
    expect(xml.match(/gridSpan="\d+"/g)).toEqual(['gridSpan="2"'])
    expect(xml.match(/rowSpan="\d+"/g)).toEqual(['rowSpan="2"'])
  })
})

describe('shape outline color', () => {
  it('keeps an #RRGGBBAA stroke translucent like the fill', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#11223380',
      stroke: { color: '#ff000080', widthEmu: 12700 },
    })
    const xml = el.anchor.originalXml
    expect(xml).toContain(
      '<a:solidFill><a:srgbClr val="112233"><a:alpha val="50196"/></a:srgbClr></a:solidFill>',
    )
    expect(xml).toContain(
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50196"/></a:srgbClr></a:solidFill></a:ln>',
    )
  })
})
