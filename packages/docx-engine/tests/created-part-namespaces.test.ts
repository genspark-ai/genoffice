import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  inlinePictureDrawingXml,
  NEW_PICTURE_SVG_SUFFIX,
  NEW_PICTURE_TOKEN_PREFIX,
  parseDocx,
  saveDocx,
} from '../src/index'
import type { HeaderFooter, Run } from '../src/types'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'
import { packageXmlProblems, xmlProblem } from './helpers/strict-xml'

const SVG_BASE64 = Buffer.from(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 4 4"><rect width="4" height="4"/></svg>',
).toString('base64')
const TOKEN = `${NEW_PICTURE_TOKEN_PREFIX}1`
const ICON = inlinePictureDrawingXml({
  rId: TOKEN,
  svgRId: TOKEN + NEW_PICTURE_SVG_SUFFIX,
  docPrId: TOKEN,
  widthPx: 96,
  heightPx: 96,
})
const NEW_PICTURES = {
  [TOKEN]: { base64: TINY_PNG_BASE64, mime: 'image/png' as const, svgBase64: SVG_BASE64 },
}
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
/** a text box as a body run carries it: every prefix relies on the root */
const TEXT_BOX =
  '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="0A1B2C3D">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="9" name="Text Box 9"/>' +
  `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="${WPS}">` +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p w14:paraId="11223344"><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
  '<mc:Fallback><w:pict><v:rect style="width:72pt;height:36pt" o:gfxdata="AA=="><w10:wrap type="none"/></v:rect></w:pict></mc:Fallback>' +
  '</mc:AlternateContent>'

const picture = (xml: string): Run =>
  ({ text: '', image: { dataUrl: `data:image/png;base64,${TINY_PNG_BASE64}`, xml } }) as Run
const part = async (bytes: Uint8Array, name: string) =>
  (await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? ''
const plainBody = () => buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' })
const KEEP = [{ kind: 'original' as const, docxIndex: 0 }]

describe('a part the save creates declares every prefix written into it', () => {
  it.each(['header', 'footer'] as const)('an icon in a new %s', async (kind) => {
    const hf: HeaderFooter = {
      text: 'Logo',
      paras: [{ runs: [{ text: 'Logo ' } as Run, picture(ICON)] }],
      newPictures: NEW_PICTURES,
    }
    const saved = await saveDocx(await parseDocx(await plainBody()), KEEP, { [kind]: hf })
    const xml = await part(saved, `word/${kind}1.xml`)
    expect(xml).toContain('<wp:inline ')
    expect(xml).toContain('<asvg:svgBlip ')
    expect(xmlProblem(xml)).toBeNull()
    const root = /<w:(?:hdr|ftr)\b[^>]*>/.exec(xml)![0]
    expect(root).toContain('xmlns:wp=')
    // declared where it is used: the root does not repeat it
    expect(root).not.toContain('xmlns:asvg=')
    expect(await packageXmlProblems(saved)).toEqual([])
  })

  it('a text box, its fallback and Word 2010 attributes in a new first-page header', async () => {
    const hf: HeaderFooter = { text: '', paras: [{ runs: [picture(TEXT_BOX)] }] }
    const saved = await saveDocx(await parseDocx(await plainBody()), KEEP, {
      headerFirst: hf,
      titlePg: true,
    })
    const xml = await part(saved, 'word/header1.xml')
    expect(xml).toContain('<wps:txbx>')
    expect(xmlProblem(xml)).toBeNull()
    const root = /<w:hdr\b[^>]*>/.exec(xml)![0]
    for (const prefix of ['mc', 'wp', 'wp14', 'wps', 'w14', 'v', 'o', 'w10'])
      expect(root).toContain(`xmlns:${prefix}=`)
    expect(root.match(/xmlns:v=/g)).toHaveLength(1)
  })

  it('content from the body brings a prefix only the body root declares', async () => {
    const source = await JSZip.loadAsync(await plainBody())
    const body = await source.file('word/document.xml')!.async('string')
    source.file(
      'word/document.xml',
      body.replace('<w:document ', '<w:document xmlns:vendor="urn:example:vendor" '),
    )
    const parsed = await parseDocx(await source.generateAsync({ type: 'uint8array' }))
    const moved = '<w:drawing vendor:mark="1"><vendor:note/></w:drawing>'
    const saved = await saveDocx(parsed, KEEP, {
      footer: { text: '', paras: [{ runs: [picture(moved)] }] },
    })
    const xml = await part(saved, 'word/footer1.xml')
    expect(xml).toContain('<vendor:note/>')
    expect(/<w:ftr\b[^>]*>/.exec(xml)![0]).toContain('xmlns:vendor="urn:example:vendor"')
    expect(xmlProblem(xml)).toBeNull()
  })

  it('an existing header with a short root gets them too', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
        headerXml: '<w:p><w:r><w:t>Old</w:t></w:r></w:p>',
      }),
    )
    const saved = await saveDocx(parsed, KEEP, {
      header: {
        text: 'New',
        paras: [{ runs: [{ text: 'New ' } as Run, picture(ICON)] }],
        newPictures: NEW_PICTURES,
      },
    })
    const xml = await part(saved, 'word/header1.xml')
    expect(xml).toContain('<wp:inline ')
    expect(xmlProblem(xml)).toBeNull()
  })
})

describe('the strict reader the save tests run under', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'

  it('rejects an unbound prefix and a prefix named but not declared', () => {
    expect(xmlProblem(`<w:hdr ${W}><wp:inline/></w:hdr>`)).toMatch(/unbound namespace prefix/)
    expect(xmlProblem(`<w:hdr ${W} ${MC} mc:Ignorable="w14"/>`)).toMatch(/undeclared prefix w14/)
    expect(
      xmlProblem(
        `<w:hdr ${W} ${MC}><mc:AlternateContent><mc:Choice Requires="wps"/></mc:AlternateContent></w:hdr>`,
      ),
    ).toMatch(/undeclared prefix wps/)
    expect(xmlProblem(`<w:hdr ${W} ${MC} xmlns:w14="urn:w14" mc:Ignorable="w14"/>`)).toBeNull()
  })

  it('reports a part the save broke, not one that arrived broken', async () => {
    const zip = await JSZip.loadAsync(await plainBody())
    zip.file('word/header9.xml', `<w:hdr ${W}><wp:inline/></w:hdr>`)
    const source = await zip.generateAsync({ type: 'uint8array' })
    expect(await packageXmlProblems(source)).toHaveLength(1)
    expect(await packageXmlProblems(source, source)).toEqual([])
  })
})
