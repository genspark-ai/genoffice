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

const OUTLINE = { color: 'FF0000', widthPt: 1.5 }
const SHADOW = { color: '000000', alpha: 0.6, blurPt: 4, distPt: 3, dirDeg: 45 }
const GLOW = { color: '4472C4', radiusPt: 8, alpha: 0.6 }

describe('text effects set in the editor (w14 outline, shadow, glow)', () => {
  it('writes the three effects on a new run and reads them back', async () => {
    const { xml, runs } = await saved([
      { text: 'a ' },
      { text: 'fancy', bold: true, fxOwn: { outline: OUTLINE, shadow: SHADOW, glow: GLOW } },
    ])
    // Word 2010 markup needs its namespace on the root
    expect(xml).toMatch(
      /<w:document[^>]*xmlns:w14="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordml"/,
    )
    // in the schema's order: glow, shadow, outline
    expect(xml).toMatch(
      /<w14:glow [^>]*>.*<\/w14:glow><w14:shadow [^>]*>.*<\/w14:shadow><w14:textOutline /,
    )
    expect(xml).toContain('<w14:glow w14:rad="101600">')
    expect(xml).toContain('w14:blurRad="50800" w14:dist="38100" w14:dir="2700000"')
    expect(xml).toContain('<w14:textOutline w14:w="19050"')
    expect(runs[1]).toMatchObject({
      text: 'fancy',
      bold: true,
      textOutline: OUTLINE,
      textShadow: SHADOW,
      glow: GLOW,
    })
    expect(runs[0].glow).toBeUndefined()
  })

  it('reads a shadow Word wrote', async () => {
    const body =
      '<w:p><w:r><w:rPr><w14:shadow w14:blurRad="38100" w14:dist="25400" w14:dir="5400000" ' +
      'w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="t">' +
      '<w14:srgbClr w14:val="112233"><w14:alpha w14:val="40000"/></w14:srgbClr></w14:shadow>' +
      '</w:rPr><w:t>shaded</w:t></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(parsed.blocks[0].runs?.[0].textShadow).toEqual({
      color: '112233',
      alpha: 0.4,
      blurPt: 3,
      distPt: 2,
      dirDeg: 90,
    })
  })

  it('changes and clears effects on a run from a file, and leaves untouched ones as they are', async () => {
    // a run with a glow, a bevelled outline the model does not describe in full, and a ligature setting
    const outline =
      '<w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr">' +
      '<w14:solidFill><w14:srgbClr w14:val="00AA00"/></w14:solidFill>' +
      '<w14:prstDash w14:val="dash"/><w14:round/></w14:textOutline>'
    const body =
      '<w:p><w:r><w:rPr><w:b/><w14:glow w14:rad="63500"><w14:srgbClr w14:val="FFC000"/></w14:glow>' +
      `${outline}<w14:ligatures w14:val="none"/></w:rPr><w:t>kept</w:t></w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const run = parsed.blocks[0].runs![0]
    expect(run.glow).toMatchObject({ color: 'FFC000', radiusPt: 5 })

    // the glow goes, a shadow comes, the outline is not touched
    const { xml, runs } = await saved([{ ...run, fxOwn: { glow: null, shadow: SHADOW } }], body)
    expect(xml).not.toContain('w14:glow')
    expect(xml).toContain(outline)
    expect(xml).toContain('<w14:ligatures w14:val="none"/>')
    expect(xml).toMatch(/<w14:shadow [^>]*>.*<\/w14:shadow><w14:textOutline /)
    expect(runs[0]).toMatchObject({ bold: true, textShadow: SHADOW })
    expect(runs[0].glow).toBeUndefined()

    // a run the editor did not change keeps its bytes
    const same = await saved([run], body)
    expect(same.xml).toContain(
      '<w14:glow w14:rad="63500"><w14:srgbClr w14:val="FFC000"/></w14:glow>',
    )
  })
})
