import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

describe('heading style without w:sz', () => {
  it('is body-sized, not the built-in 16pt heading size', async () => {
    const css = docStyleCss(
      await parseDocx(
        await buildDocx({
          bodyXml: '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>T</w:t></w:r></w:p>',
          extraStylesXml:
            '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
            '<w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>' +
            '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>' +
            '<w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="28"/></w:rPr></w:style>',
        }),
      ),
    )
    const rule = (id: string) => css.split('\n').find((l) => l.includes(`[data-style="${id}"]`))
    expect(rule('Heading1')).toContain('font-size:inherit')
    expect(rule('Heading2')).toContain('font-size:14pt')
  })
})

describe('paragraph style with a Korean East Asian face', () => {
  it('re-anchors the hangul line factor for its paragraphs', async () => {
    const css = docStyleCss(
      await parseDocx(
        await buildDocx({
          bodyXml:
            '<w:p><w:pPr><w:pStyle w:val="NormalWeb"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>\uD55C\uAE00</w:t></w:r></w:p>',
          stylesXml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic"/></w:rPr></w:rPrDefault></w:docDefaults>' +
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
            '<w:style w:type="paragraph" w:styleId="NormalWeb"><w:name w:val="Normal (Web)"/><w:basedOn w:val="Normal"/>' +
            '<w:rPr><w:rFonts w:ascii="\uAD74\uB9BC" w:eastAsia="\uAD74\uB9BC" w:hAnsi="\uAD74\uB9BC"/></w:rPr></w:style>' +
            '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/>' +
            '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="SimSun" w:hAnsi="Arial"/></w:rPr></w:style></w:styles>',
        }),
      ),
    )
    const rule = (id: string) => css.split('\n').find((l) => l.includes(`[data-style="${id}"]`))
    expect(css).toContain('--doc-line-factor-kr:1.7371')
    expect(rule('NormalWeb')).toContain('--doc-line-factor-kr:1.3029')
    expect(rule('Body')).not.toContain('--doc-line-factor-kr')
  })
})

describe('empty EA theme slot resolved to a Japanese face', () => {
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  const scripts =
    '<a:font script="Jpan" typeface="ＭＳ ゴシック"/>' +
    '<a:font script="Hang" typeface="맑은 고딕"/>'
  const parts = (eaLang: string) => [
    {
      path: 'word/theme/theme1.xml',
      xml:
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme ${NS} name="T"><a:themeElements><a:fontScheme name="T">` +
        `<a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>${scripts}</a:majorFont>` +
        `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>${scripts}</a:minorFont>` +
        '</a:fontScheme></a:themeElements></a:theme>',
      contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
    },
    {
      path: 'word/settings.xml',
      xml:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        `<w:themeFontLang w:val="en-US" w:eastAsia="${eaLang}"/></w:settings>`,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    },
  ]
  const cssFor = async (eaLang: string) =>
    docStyleCss(
      await parseDocx(
        await buildDocx({
          bodyXml: '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>한글</w:t></w:r></w:p>',
          extraStylesXml:
            '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
            '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:ascii="NanumGothic" w:hAnsi="NanumGothic" w:eastAsia="NanumGothic"/></w:rPr></w:style>',
          extraParts: parts(eaLang),
        }),
      ),
    )
  const rule = (css: string) => css.split('\n').find((l) => l.includes('[data-style="Title"]'))

  it('sizes hangul lines by the Malgun fallback under themeFontLang ja-JP', async () => {
    expect(rule(await cssFor('ja-JP'))).toContain('--doc-line-factor-kr:1.7371')
  })

  it('leaves a Korean face resolved under ko-KR to the document factor', async () => {
    expect(rule(await cssFor('ko-KR'))).not.toContain('--doc-line-factor-kr')
  })
})
