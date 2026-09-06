import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES =
  '<w:style w:type="paragraph" w:styleId="Rule"><w:name w:val="Rule"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="48" w:space="4" w:color="1F4E79"/></w:pBdr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="RuleOff"><w:name w:val="RuleOff"/><w:basedOn w:val="Rule"/>' +
  '<w:pPr><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="RuleBoxed"><w:name w:val="RuleBoxed"/><w:basedOn w:val="Rule"/>' +
  '<w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="4F81BD" w:themeColor="accent1"/></w:pBdr></w:pPr></w:style>'

describe('style-level paragraph borders (w:style/w:pPr/w:pBdr)', () => {
  it('reads the sides into StyleDisplay and merges them per side along basedOn', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:pPr><w:pStyle w:val="Rule"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml: STYLES,
      }),
    )
    expect(doc.styles.get('Rule')!.display?.borderSides).toEqual({
      b: { color: '1F4E79', szPt: 6 },
    })
    // explicit none on the child cancels the inherited bottom (kept as null, not dropped)
    expect(doc.styles.get('RuleOff')!.display?.borderSides).toEqual({ b: null })
    // a child adding a top keeps the parent's bottom; the theme reference wins over the literal
    const boxed = doc.styles.get('RuleBoxed')!.display?.borderSides
    expect(boxed?.b).toEqual({ color: '1F4E79', szPt: 6 })
    expect(boxed?.t?.szPt).toBe(1)
    expect(boxed?.t?.color).toMatch(/^[0-9A-F]{6}$/i)
    expect(boxed?.t?.color).toBe(doc.themeColors?.accent1)
  })

  it('leaves the paragraph model untouched: style borders are display-only', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="Rule"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>' +
          '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="FF0000" w:themeColor="accent2"/></w:pBdr></w:pPr><w:r><w:t>direct</w:t></w:r></w:p>',
        extraStylesXml: STYLES,
      }),
    )
    const [styled, direct] = doc.blocks.filter((b) => !b.hidden)
    expect(styled.format?.borders).toBeUndefined()
    expect(styled.format?.borderReset).toBeUndefined()
    // direct borders keep the literal color so the raw/model save comparison still matches
    expect(direct.format?.borders).toBe('t')
    expect(direct.format?.borderLines).toEqual({ t: { color: 'FF0000', szPt: 0.5 } })
  })

  it('records direct none/nil sides as a reset that cancels the style side', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="Rule"/><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
        extraStylesXml: STYLES,
      }),
    )
    const [p] = doc.blocks.filter((b) => !b.hidden)
    expect(p.format?.borders).toBeUndefined()
    expect(p.format?.borderReset).toBe('b')
  })

  it('header paragraphs take the style border unless the direct pBdr resets that side', async () => {
    const header =
      '<w:p><w:pPr><w:pStyle w:val="Rule"/></w:pPr><w:r><w:t>ruled</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Rule"/><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr></w:pPr><w:r><w:t>plain</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Rule"/><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="FF0000"/></w:pBdr></w:pPr><w:r><w:t>boxed</w:t></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
        extraStylesXml: STYLES,
        sectPrExtra: '<w:headerReference w:type="default" r:id="rIdHdr"/>',
        extraRels:
          '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
        extraParts: [
          {
            path: 'word/header1.xml',
            xml:
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
              header +
              '</w:hdr>',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
          },
        ],
      }),
    )
    const paras = doc.headerParas!
    expect(paras[0]).toMatchObject({
      borders: 'b',
      borderLines: { b: { color: '1F4E79', szPt: 6 } },
    })
    expect(paras[1].borders).toBeUndefined()
    // a direct side only replaces that side; the style's other sides stay
    expect(paras[2]).toMatchObject({
      borders: 'tb',
      borderLines: { t: { color: 'FF0000', szPt: 0.5 }, b: { color: '1F4E79', szPt: 6 } },
    })
  })
})
