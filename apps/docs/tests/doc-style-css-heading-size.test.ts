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
