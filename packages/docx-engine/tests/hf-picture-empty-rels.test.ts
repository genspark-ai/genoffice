import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  inlinePictureDrawingXml,
  NEW_PICTURE_TOKEN_PREFIX,
  parseDocx,
  saveDocx,
  type HeaderFooter,
} from '../src'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'
import { packageXmlProblems } from './helpers/strict-xml'

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const EXISTING_REL =
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'
const roots = [
  ['self-closing', `<Relationships xmlns="${REL_NS}"/>`],
  ['self-closing with whitespace', `<?xml version="1.0"?>\n<Relationships xmlns="${REL_NS}" />`],
  ['paired', `<Relationships xmlns="${REL_NS}"></Relationships>`],
  ['absent', undefined],
  ['existing relationship', `<Relationships xmlns="${REL_NS}">${EXISTING_REL}</Relationships>`],
] as const

describe.each(['header', 'footer'] as const)('%s picture relationship roots', (kind) => {
  it.each(roots)('round-trips an inserted picture with a %s part', async (name, root) => {
    const other = kind === 'header' ? 'footer' : 'header'
    const zip = await JSZip.loadAsync(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body stays unchanged</w:t></w:r></w:p>',
        headerXml: '<w:p><w:r><w:t>Logo</w:t></w:r></w:p>',
        footerXml: '<w:p><w:r><w:t>Footer</w:t></w:r></w:p>',
      }),
    )
    const relsPath = `word/_rels/${kind}1.xml.rels`
    if (root !== undefined) zip.file(relsPath, root)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const token = NEW_PICTURE_TOKEN_PREFIX + '1'
    const text = kind === 'header' ? 'Logo' : 'Footer'
    const hf: HeaderFooter = {
      text,
      paras: [
        {
          runs: [
            { text: text + ' ' },
            {
              text: '',
              image: {
                dataUrl: 'data:image/png;base64,' + TINY_PNG_BASE64,
                xml: inlinePictureDrawingXml({
                  rId: token,
                  docPrId: token,
                  widthPx: 32,
                  heightPx: 32,
                }),
              },
            },
          ],
        },
      ],
      newPictures: { [token]: { base64: TINY_PNG_BASE64, mime: 'image/png' } },
    }
    const saved = await saveDocx(await parseDocx(bytes), [{ kind: 'original', docxIndex: 0 }], {
      [kind]: hf,
    })
    expect(await packageXmlProblems(saved)).toEqual([])
    const after = await JSZip.loadAsync(saved)
    const xml = await after.file(`word/${kind}1.xml`)!.async('string')
    const rels = await after.file(relsPath)!.async('string')
    const embed = /r:embed="([^"]+)"/.exec(xml)![1]
    expect(embed).not.toContain(NEW_PICTURE_TOKEN_PREFIX)
    const relationships = rels.match(/<Relationship\b[^>]*\/>/g) ?? []
    expect(relationships).toHaveLength(name === 'existing relationship' ? 2 : 1)
    const imageRel = relationships.find((rel) => rel.includes(`Id="${embed}"`))
    expect(imageRel).toContain(`Type="${IMAGE_REL}"`)
    const target = /Target="([^"]+)"/.exec(imageRel!)![1]
    expect(await after.file(`word/${target}`)!.async('base64')).toBe(TINY_PNG_BASE64)
    if (name === 'existing relationship') expect(rels).toContain(EXISTING_REL)
    for (const path of ['word/document.xml', 'word/styles.xml', `word/${other}1.xml`]) {
      expect(await after.file(path)!.async('uint8array')).toEqual(
        await zip.file(path)!.async('uint8array'),
      )
    }
    const reopened = await parseDocx(saved)
    const paras = kind === 'header' ? reopened.headerParas : reopened.footerParas
    expect(paras?.flatMap((p) => p.runs).find((run) => run.image)?.image?.dataUrl).toBe(
      'data:image/png;base64,' + TINY_PNG_BASE64,
    )
  })
})
