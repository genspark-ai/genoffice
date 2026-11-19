import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="28"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:rPr><w:i/><w:sz w:val="32"/><w:bCs/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Em"><w:name w:val="Em"/>' +
  '<w:rPr><w:b/><w:szCs w:val="20"/></w:rPr></w:style>'

describe('.doc-cs gate in the style CSS', () => {
  it('rtl runs take the Cs size of their chain in place of the base w:sz', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>', extraStylesXml: STYLES }),
    )
    const css = docStyleCss(parsed as never)
    // Normal: szCs 14pt; Heading1 inherits it; the character style's own szCs wins the tie
    expect(css).toContain('.doc-page span.doc-cs { font-size:14pt }')
    expect(css).toContain('.doc-page [data-style="Heading1"] .doc-cs { font-size:14pt }')
    expect(css).toContain('.doc-page span[data-style="Em"].doc-cs { font-size:10pt }')
    expect(css.indexOf('span[data-style="Em"].doc-cs')).toBeGreaterThan(
      css.indexOf('[data-style="Heading1"] .doc-cs'),
    )
    // weight / style come resolved from the engine (marks.ts doc-cs-(n)b / -(n)i), never
    // from a per-style counter that cannot see the table-style layer
    expect(css).not.toMatch(/\.doc-cs[^{]*\{[^}]*font-(weight|style)/)
  })
})
