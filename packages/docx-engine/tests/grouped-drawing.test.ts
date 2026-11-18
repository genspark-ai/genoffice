import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { saveDocx } from '../src/patch'
import { buildDocx } from './helpers/build-docx'
import { GROUPED_DRAWING } from './helpers/grouped-drawing'
import { patchDrawingShapes } from '../src/drawing-edit'

describe('grouped DrawingML child coordinates', () => {
  it('maps small child extents before converting to pixels and keeps every line', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: GROUPED_DRAWING }))
    const boxes = doc.blocks[0].textboxes!
    expect(boxes).toHaveLength(5)
    expect(boxes[0]).toMatchObject({
      widthPx: 200,
      heightPx: 40,
      offsetXEmu: 190500,
      offsetYEmu: 190500,
    })
    expect(boxes[1]).toMatchObject({ widthPx: 200, heightPx: 40, offsetXEmu: 3048000 })
    expect(boxes[2]).toMatchObject({ shapeId: '12', widthPx: 100, heightPx: 0, prst: 'lineArrow' })
    expect(boxes[3]).toMatchObject({ shapeId: '13', widthPx: 0, heightPx: 100 })
    expect(boxes[4]).toMatchObject({ shapeId: '14', widthPx: 100, heightPx: 60 })
    expect(boxes[2].lineStyle?.tailEnd).toEqual({ type: 'triangle', width: 'med', length: 'med' })
    expect(boxes[3].lineStyle?.dash).toBe('dash')
    expect(boxes[4].lineStyle?.compound).toBe('thickThin')
  })

  it('does not change authored bytes when the group is untouched', async () => {
    const bytes = await buildDocx({ bodyXml: GROUPED_DRAWING })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(
      doc,
      doc.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(saved).toEqual(bytes)
  })

  it('uses outer shape ordinals when a preceding textbox contains a nested drawing', async () => {
    const nested =
      '<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="000000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    const xml = GROUPED_DRAWING.replace(
      '<w:t>Source</w:t></w:r>',
      `<w:t>Source</w:t></w:r>${nested}`,
    ).replace('<wps:cNvPr id="11" name="Box 11"/>', '')
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const before = doc.blocks[0].textboxes!
    const index = before.findIndex((b) =>
      b.paras.some((p) => p.runs.some((r) => r.text === 'Destination')),
    )
    const after = before.map((b, i) => (i === index ? { ...b, widthPx: 240 } : b))
    const changed = patchDrawingShapes(xml, before, after)
    const reparsed = await parseDocx(await buildDocx({ bodyXml: changed }))
    expect(
      reparsed.blocks[0].textboxes!.find((b) =>
        b.paras.some((p) => p.runs.some((r) => r.text === 'Destination')),
      )?.widthPx,
    ).toBe(240)
  })

  it('round-trips geometry written with paired offset and extent tags', async () => {
    const xml = GROUPED_DRAWING.replace(/<a:(off|ext)\b([^>]*)\/>/g, '<a:$1$2></a:$1>')
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const before = doc.blocks[0].textboxes!
    const after = before.map((b, i) =>
      i === 1 ? { ...b, widthPx: 240, offsetXEmu: b.offsetXEmu! + 95250 } : b,
    )
    const changed = patchDrawingShapes(xml, before, after)
    const saved = await parseDocx(await buildDocx({ bodyXml: changed }))
    expect(saved.blocks[0].textboxes![1]).toMatchObject({ widthPx: 240, offsetXEmu: 3143250 })
  })

  it('excludes fallback drawings from source drawing ordinals', async () => {
    const run = GROUPED_DRAWING.slice('<w:p><w:r>'.length, -'</w:r></w:p>'.length)
    const alt = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><mc:Choice Requires="wpg">${run}</mc:Choice><mc:Fallback>${run}</mc:Fallback></mc:AlternateContent>`
    const second = run.replaceAll('id="10"', 'id="20"').replaceAll('id="11"', 'id="21"')
    const xml = `<w:p><w:r>${alt}${second}</w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    const before = doc.blocks[0].textboxes!
    const after = before.map((b) => (b.shapeId === '21' ? { ...b, widthPx: 240 } : b))
    const changed = patchDrawingShapes(xml, before, after)
    const saved = await parseDocx(await buildDocx({ bodyXml: changed }))
    expect(saved.blocks[0].textboxes!.find((b) => b.shapeId === '21')?.widthPx).toBe(240)
    expect(saved.blocks[0].textboxes!.find((b) => b.shapeId === '11')?.widthPx).toBe(200)
  })

  it('does not paint a grouped line whose stroke is explicitly disabled', async () => {
    const xml = GROUPED_DRAWING.replace(
      /(<wps:cNvPr id="13"[\s\S]*?<a:ln[^>]*>)<a:solidFill>[\s\S]*?<\/a:solidFill>/,
      '$1<a:noFill/>',
    )
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].textboxes!.some((b) => b.shapeId === '13')).toBe(false)
    expect(doc.blocks[0].textboxes!.find((b) => b.shapeId === '14')?.widthPx).toBe(100)
  })
})
