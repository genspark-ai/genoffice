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

  it('replaces a gradient only when a fill is set in the editor', async () => {
    const gradient =
      '<w14:textFill><w14:gradFill><w14:gsLst>' +
      '<w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>' +
      '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>' +
      '</w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>'
    const rPr = `<w:color w:val="FF0000"/>${gradient}`
    const run = await firstRun(rPr)
    const recoloured = await saved([{ ...run, color: '0070C0' }], fileRun(rPr))
    expect(recoloured.xml).toContain(gradient)
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
