/**
 * Run formatting that reached the screen wrong or not at all: legacy text
 * effects, w14:glow, w:position, pattern shading, the Hyperlink character
 * style, style-level w:bidi and a field result cut by a paragraph mark.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W14_NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"'

const para = (runs: string, pPr = '') => `<w:p>${pPr}${runs}</w:p>`
const run = (rPr: string, text: string) => `<w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r>`

describe('legacy text effects', () => {
  it('reads imprint, outline, shadow and the baseline shift', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: para(
          run('<w:b/><w:imprint/><w:dstrike/><w:color w:val="FFFFFF"/>', 'engraved') +
            run('<w:outline/>', 'hollow') +
            run('<w:shadow/>', 'cast') +
            run('<w:position w:val="-8"/>', 'lowered') +
            run('<w:position w:val="0"/>', 'flat'),
        ),
      }),
    )
    const runs = doc.blocks[0].runs ?? []
    expect(runs.map((r) => r.textEffect)).toEqual([
      'imprint',
      'outline',
      'shadow',
      undefined,
      undefined,
    ])
    expect(runs[0].color).toBe('FFFFFF')
    expect(runs[0].dstrike).toBe(true)
    expect(runs[1].dstrike).toBeUndefined()
    expect(runs[3].positionHalfPoints).toBe(-8)
    expect(runs[4].positionHalfPoints).toBeUndefined()
  })

  it('reads a w14:glow halo with radius, theme colour and alpha', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: para(
          run(
            '<w14:glow w14:rad="101600"><w14:schemeClr w14:val="accent3"><w14:alpha w14:val="60000"/></w14:schemeClr></w14:glow>',
            'glowing',
          ),
        ),
        docRootExtraAttrs: W14_NS,
      }),
    )
    const glow = doc.blocks[0].runs?.[0].glow
    expect(glow?.radiusPt).toBe(8)
    expect(glow?.alpha).toBe(0.6)
    expect(glow?.color).toMatch(/^[0-9A-F]{6}$/)
  })
})

describe('run shading patterns', () => {
  it('blends a pct15 pattern over a white fill instead of painting white', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: para(run('<w:shd w:val="pct15" w:color="auto" w:fill="FFFFFF"/>', 'boxed')),
      }),
    )
    expect(doc.blocks[0].runs?.[0]).toMatchObject({ shading: 'FFFFFF', shadingDisplay: 'D9D9D9' })
  })
})

describe('Hyperlink character style', () => {
  it('stays on the run so the document style paints the link', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: para(
          '<w:hyperlink w:anchor="_top"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>top</w:t></w:r></w:hyperlink>',
        ),
      }),
    )
    const r = doc.blocks[0].runs?.[0]
    expect(r?.link?.href).toBe('#_top')
    expect(r?.styleId).toBe('Hyperlink')
  })

  it('records the leading result run style of a TOC entry', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc1">' +
          '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>H1</w:t></w:r>' +
          '<w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r></w:hyperlink></w:p>',
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>',
      }),
    )
    expect(doc.blocks[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'H1',
      runStyleId: 'Hyperlink',
    })
  })
})

describe('style-level w:bidi', () => {
  const RTL_NORMAL =
    '<w:style w:type="paragraph" w:styleId="RtlBody"><w:name w:val="Rtl Body"/>' +
    '<w:pPr><w:bidi/><w:jc w:val="distribute"/></w:pPr></w:style>'

  it('swaps an explicit jc left/right like a direct w:bidi does', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          para(
            run('', 'heading'),
            '<w:pPr><w:pStyle w:val="RtlBody"/><w:jc w:val="right"/></w:pPr>',
          ) +
          para(run('', 'plain'), '<w:pPr><w:jc w:val="right"/></w:pPr>') +
          para(run('', 'own'), '<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>'),
        extraStylesXml: RTL_NORMAL,
      }),
    )
    expect(doc.styles.get('RtlBody')?.display).toMatchObject({ bidi: true, align: 'distribute' })
    expect(doc.blocks[0].format?.align).toBe('left')
    expect(doc.blocks[0].format?.bidi).toBeUndefined()
    expect(doc.blocks[1].format?.align).toBe('right')
    expect(doc.blocks[2].format).toMatchObject({ bidi: true, align: 'left' })
  })
})

describe('field result split by a paragraph mark', () => {
  it('keeps the formatted result runs of the paragraph that lacks the field end', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:t xml:space="preserve">Author. (2008). </w:t></w:r>' +
          run('<w:i/>', 'Title.') +
          '</w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      }),
    )
    const field = doc.blocks[0].fieldDisplay
    expect(field?.kind).toBe('text')
    expect(field?.runs?.map((r) => [r.text, r.italic ?? false])).toEqual([
      ['Author. (2008). ', false],
      ['Title.', true],
    ])
  })
})
