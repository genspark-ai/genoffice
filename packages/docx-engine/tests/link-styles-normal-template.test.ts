import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const STYLES =
  `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>` +
  '<w:pPrDefault><w:pPr><w:spacing w:before="120" w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="400"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>' +
  '</w:styles>'
const BODY =
  '<w:p><w:r><w:rPr><w:szCs w:val="24"/></w:rPr><w:t>Bibliography entry sized only through szCs</w:t></w:r></w:p>'

function settings(inner: string) {
  return {
    path: 'word/settings.xml',
    xml: `<w:settings ${W}>${inner}</w:settings>`,
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  }
}

describe('settings.xml w:linkStyles', () => {
  it('without an attached template swaps docDefaults and Normal for Normal.dotm', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: STYLES,
        extraParts: [settings('<w:linkStyles/>')],
      }),
    )
    expect(doc.docDefaults).toMatchObject({
      sizeHalfPoints: 24,
      kernHalfPoints: 2,
      spaceAfterTwips: 160,
      lineRawTwips: 278,
      lineRule: 'auto',
    })
    expect(doc.docDefaults?.spaceBeforeTwips).toBeUndefined()
    // after=200/line=276 differ from the template: Normal carries the template values
    expect(doc.styles.get('Normal')?.display).toEqual({
      spaceAfterTwips: 160,
      lineRawTwips: 278,
      lineRule: 'auto',
      lineSpacing: 278 / 240,
    })
    const body = doc.styles.get('Body')?.display
    expect(body?.sizeHalfPoints).toBeUndefined()
    expect(body?.indentLeftTwips).toBe(720)
    expect(body?.lineRawTwips).toBe(278)
  })

  it('writes onto Normal only the pPrDefault spacing values the package declared differently', async () => {
    const stylesWith = (spacing: string) =>
      `<w:styles ${W}><w:docDefaults><w:pPrDefault><w:pPr><w:spacing ${spacing}/></w:pPr></w:pPrDefault></w:docDefaults>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
      '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>'
    const parse = async (spacing: string) =>
      (
        await parseDocx(
          await buildDocx({
            bodyXml: BODY,
            stylesXml: stylesWith(spacing),
            extraParts: [settings('<w:linkStyles/>')],
          }),
        )
      ).styles.get('Normal')?.display
    // equal to Normal.dotm: Normal's own line=240 is dropped, nothing written back
    expect(await parse('w:after="160" w:line="278" w:lineRule="auto"')).toBeUndefined()
    expect(await parse('w:after="200" w:line="278" w:lineRule="auto"')).toEqual({
      spaceAfterTwips: 160,
    })
    expect(await parse('w:after="160" w:line="276" w:lineRule="auto"')).toEqual({
      lineRawTwips: 278,
      lineRule: 'auto',
      lineSpacing: 278 / 240,
    })
  })

  it('clears Normal even when it is based on another style', async () => {
    const styles =
      `<w:styles ${W}><w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/>` +
      '<w:pPr><w:spacing w:after="400"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:basedOn w:val="Base"/>' +
      '<w:rPr><w:b/></w:rPr></w:style></w:styles>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: styles,
        extraParts: [settings('<w:linkStyles/>')],
      }),
    )
    const normal = doc.styles.get('Normal')
    expect(normal?.display).toBeUndefined()
    expect(normal?.numPr).toBeUndefined()
    expect(normal?.basedOn).toBeUndefined()
    expect(doc.styles.get('Base')?.display?.sizeHalfPoints).toBe(20)
  })

  it('targets the Normal style rather than the declared default', async () => {
    const styles =
      `<w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Body"><w:name w:val="Body Text"/>` +
      '<w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/>' +
      '<w:pPr><w:spacing w:after="400"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:style></w:styles>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        stylesXml: styles,
        extraParts: [settings('<w:linkStyles/>')],
      }),
    )
    expect(doc.styles.get('Normal')?.display).toBeUndefined()
    expect(doc.styles.get('Body')?.display?.sizeHalfPoints).toBe(20)
    expect(doc.styles.get('Body')?.isDefault).toBe(true)
  })

  it('keeps the package styles when the flag is off or a template is attached', async () => {
    for (const inner of [
      '',
      '<w:linkStyles w:val="0"/>',
      '<w:attachedTemplate r:id="rId9"/><w:linkStyles/>',
    ]) {
      const doc = await parseDocx(
        await buildDocx({ bodyXml: BODY, stylesXml: STYLES, extraParts: [settings(inner)] }),
      )
      expect(doc.docDefaults?.sizeHalfPoints).toBe(22)
      expect(doc.styles.get('Normal')?.display?.sizeHalfPoints).toBe(20)
    }
  })
})
