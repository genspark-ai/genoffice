import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type Run } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const documentXml = async (bytes: Uint8Array): Promise<string> =>
  (await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) ?? ''

async function saved(runs: Run[], bodyXml = '<w:p><w:r><w:t>plain</w:t></w:r></w:p>') {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const bytes = await saveDocx(parsed, [{ kind: 'generated', block: { type: 'paragraph', runs } }])
  return { xml: await documentXml(bytes), runs: (await parseDocx(bytes)).blocks[0].runs ?? [] }
}

const fill = (color: string) =>
  `<w14:textFill><w14:solidFill>${color}</w14:solidFill></w14:textFill>`
const srgb = (hex: string) => fill(`<w14:srgbClr w14:val="${hex}"/>`)
const fileRun = (rPr: string) => `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>kept</w:t></w:r></w:p>`
const firstRun = async (rPr: string) =>
  (await parseDocx(await buildDocx({ bodyXml: fileRun(rPr) }))).blocks[0].runs![0]

describe('w14:textFill (solid)', () => {
  it('is written after the outline in the colour of the run, with w:color beside it', async () => {
    const { xml, runs } = await saved([
      {
        text: 'filled',
        color: 'ED7D31',
        fxOwn: { outline: { color: 'ED7D31', widthPt: 0.75 }, fill: 'ED7D31' },
      },
    ])
    expect(xml).toMatch(/<w:document[^>]*xmlns:w14=/)
    expect(xml).toContain('<w:color w:val="ED7D31"/>')
    expect(xml).toMatch(/<\/w14:textOutline><w14:textFill>/)
    expect(xml).toContain(srgb('ED7D31'))
    expect(runs[0].color).toBe('ED7D31')
  })

  it('follows the colour the run has at the save, and goes with it', async () => {
    const recoloured = await saved([{ text: 'x', color: '00B050', fxOwn: { fill: 'ED7D31' } }])
    expect(recoloured.xml).toContain(srgb('00B050'))
    expect(recoloured.xml).not.toContain('ED7D31')
    const automatic = await saved([{ text: 'x', fxOwn: { fill: 'ED7D31' } }])
    expect(automatic.xml).not.toContain('w14:textFill')
  })

  it('is the colour of a run that has no w:color', async () => {
    expect((await firstRun(srgb('C00000'))).color).toBe('C00000')
  })

  it('stays byte for byte when the run is saved untouched', async () => {
    for (const rPr of [
      `<w:color w:val="C00000"/>${srgb('C00000')}`,
      `<w:color w:val="000000"/>${srgb('C00000')}`,
      srgb('C00000'),
      '<w:color w:val="000000" w:themeColor="text1"/>' + fill('<w14:schemeClr w14:val="tx1"/>'),
      fill('<w14:srgbClr w14:val="C00000"><w14:alpha w14:val="25000"/></w14:srgbClr>'),
    ]) {
      const run = await firstRun(rPr)
      const { xml } = await saved([{ ...run, text: 'kept still' }], fileRun(rPr))
      expect(xml).toContain(`<w:rPr>${rPr}</w:rPr>`)
    }
  })

  it('takes the new colour when the colour of a run from a file changes', async () => {
    const rPr = `<w:b/><w:color w:val="C00000"/>${srgb('C00000')}<w14:ligatures w14:val="none"/>`
    const run = await firstRun(rPr)
    const { xml } = await saved([{ ...run, color: '0070C0' }], fileRun(rPr))
    expect(xml).toContain(
      `<w:b/><w:color w:val="0070C0"/>${srgb('0070C0')}<w14:ligatures w14:val="none"/>`,
    )
    const { color: _was, ...plain } = run
    const cleared = await saved([plain], fileRun(rPr))
    expect(cleared.xml).not.toContain('w14:textFill')
    expect(cleared.xml).not.toContain('w:color')
  })

  it('replaces a gradient when a colour is picked over it or a fill is set in the editor', async () => {
    const gradient =
      '<w14:textFill><w14:gradFill><w14:gsLst>' +
      '<w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>' +
      '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>' +
      '</w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>'
    const rPr = `<w:color w:val="FF0000"/>${gradient}`
    const run = await firstRun(rPr)
    const untouched = await saved([{ ...run, text: 'kept still' }], fileRun(rPr))
    expect(untouched.xml).toContain(gradient)
    const recoloured = await saved([{ ...run, color: '0070C0' }], fileRun(rPr))
    expect(recoloured.xml).not.toContain('w14:gradFill')
    expect(recoloured.xml).toContain(srgb('0070C0'))
    const filled = await saved(
      [{ ...run, color: '0070C0', fxOwn: { fill: '0070C0' } }],
      fileRun(rPr),
    )
    expect(filled.xml).not.toContain('w14:gradFill')
    expect(filled.xml).toContain(srgb('0070C0'))
  })

  it('is taken off a run when the editor cleared it', async () => {
    const rPr = `<w:color w:val="C00000"/>${srgb('C00000')}`
    const run = await firstRun(rPr)
    const { xml } = await saved([{ ...run, fxOwn: { fill: null } }], fileRun(rPr))
    expect(xml).not.toContain('w14:textFill')
    expect(xml).toContain('<w:color w:val="C00000"/>')
  })
})

const gradFill = (stops: string, spread = '<w14:lin w14:ang="5400000" w14:scaled="0"/>') =>
  `<w14:textFill><w14:gradFill><w14:gsLst>${stops}</w14:gsLst>${spread}</w14:gradFill></w14:textFill>`
const gs = (pos: number, hex: string, inner = '') =>
  `<w14:gs w14:pos="${pos}"><w14:srgbClr w14:val="${hex}"${inner ? `>${inner}</w14:srgbClr>` : '/>'}</w14:gs>`

describe('w14:textFill (what the page draws)', () => {
  it('reads a gradient with its stops, its angle and the transparency of a stop', async () => {
    const run = await firstRun(
      '<w:color w:val="4472C4"/>' +
        gradFill(gs(100000, '4472C4') + gs(0, 'FFFFFF', '<w14:alpha w14:val="25000"/>')),
    )
    expect(run.color).toBe('4472C4')
    expect(run.textFill).toEqual({
      kind: 'gradient',
      stops: [
        { pos: 0, color: 'FFFFFF', alpha: 0.75 },
        { pos: 1, color: '4472C4' },
      ],
      angleDeg: 90,
    })
  })

  it('reads a path gradient, an empty fill and a transparent colour', async () => {
    const path = '<w14:path w14:path="circle"><w14:fillToRect w14:l="50000"/></w14:path>'
    expect(
      (await firstRun(gradFill(gs(0, 'FF0000') + gs(100000, '0000FF'), path))).textFill,
    ).toMatchObject({ kind: 'gradient', path: 'circle' })
    expect((await firstRun('<w14:textFill><w14:noFill/></w14:textFill>')).textFill).toEqual({
      kind: 'none',
    })
    const faint = fill('<w14:srgbClr w14:val="C00000"><w14:alpha w14:val="60000"/></w14:srgbClr>')
    expect((await firstRun(`<w:color w:val="C00000"/>${faint}`)).textFill).toEqual({
      kind: 'solid',
      color: 'C00000',
      alpha: 0.4,
    })
  })

  it('says nothing beside a colour that already says it, and speaks up when the two differ', async () => {
    expect((await firstRun(`<w:color w:val="C00000"/>${srgb('C00000')}`)).textFill).toBeUndefined()
    expect((await firstRun(srgb('C00000'))).textFill).toBeUndefined()
    expect((await firstRun(`<w:color w:val="000000"/>${srgb('C00000')}`)).textFill).toEqual({
      kind: 'solid',
      color: 'C00000',
    })
  })

  it('writes a gradient set in the editor and reads the same one back', async () => {
    const gradient = {
      kind: 'gradient' as const,
      stops: [
        { pos: 0, color: 'B4C7E7' },
        { pos: 1, color: '4472C4' },
      ],
      angleDeg: 45,
    }
    const { xml, runs } = await saved([
      { text: 'g', color: '4472C4', fxOwn: { fill: '4472C4', textFill: gradient } },
    ])
    expect(xml).toContain('<w:color w:val="4472C4"/>')
    expect(xml).toContain(
      gradFill(
        gs(0, 'B4C7E7') + gs(100000, '4472C4'),
        '<w14:lin w14:ang="2700000" w14:scaled="0"/>',
      ),
    )
    expect(runs[0].textFill).toEqual(gradient)
  })

  it('keeps a gradient of the file byte for byte, and gives it up for a colour picked over it', async () => {
    const rPr = `<w:color w:val="4472C4"/>${gradFill(gs(0, 'FFFFFF') + gs(100000, '4472C4'))}`
    const run = await firstRun(rPr)
    const kept = await saved([{ ...run, text: 'kept still' }], fileRun(rPr))
    expect(kept.xml).toContain(`<w:rPr>${rPr}</w:rPr>`)
    const { textFill: _gone, ...plain } = run
    const recoloured = await saved([{ ...plain, color: 'C00000' }], fileRun(rPr))
    expect(recoloured.xml).not.toContain('w14:gradFill')
    expect(recoloured.xml).toContain(`<w:color w:val="C00000"/>${srgb('C00000')}`)
  })

  it('hands the fill back to the colour when a gradient set in the editor is taken off', async () => {
    const { xml } = await saved([
      { text: 'g', color: '4472C4', fxOwn: { fill: '4472C4', textFill: null } },
    ])
    expect(xml).toContain(srgb('4472C4'))
    expect(xml).not.toContain('w14:gradFill')
  })
})
