import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx } from '../src/index'
import type { GenerateContext } from '../src/generate'
import { buildDocx } from './helpers/build-docx'

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

const CUSTOM_STYLES =
  '<w:style w:type="paragraph" w:styleId="Body1"><w:name w:val="Body 1"/><w:basedOn w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="HDR"><w:name w:val="HDR"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'

describe('heading by direct w:outlineLvl only', () => {
  it('marks outline-only headings and leaves styled headings alone', async () => {
    const bytes = await buildDocx({
      extraStylesXml: CUSTOM_STYLES,
      bodyXml: [
        p('<w:outlineLvl w:val="1"/>', 'unstyled body paragraph'),
        p('<w:pStyle w:val="Body1"/><w:outlineLvl w:val="1"/>', 'custom body style'),
        p('<w:pStyle w:val="Heading2"/>', 'real heading'),
        p('<w:pStyle w:val="Heading2"/><w:outlineLvl w:val="0"/>', 'heading with direct level'),
        p('<w:pStyle w:val="HDR"/>', 'style-level outlineLvl'),
      ].join(''),
    })
    const blocks = (await parseDocx(bytes)).blocks.filter((b) => !b.hidden)
    expect(blocks.map((b) => [b.type, b.level, b.outlineOnly])).toEqual([
      ['heading', 2, true],
      ['heading', 2, true],
      ['heading', 2, undefined],
      ['heading', 1, undefined],
      ['heading', 1, undefined],
    ])
  })
})

describe('generateParagraphXml for outline-only headings', () => {
  const ctx = {
    headingStyleIds: new Map([[2, 'Heading2']]),
  } as unknown as GenerateContext

  it('writes the outline level instead of a Heading style', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 2, outlineOnly: true, runs: [{ text: 'x' }] },
      ctx,
    )
    expect(xml).toContain('<w:outlineLvl w:val="1"/>')
    expect(xml).not.toContain('w:pStyle')
  })

  it('keeps the Heading style fallback for ordinary headings', () => {
    const xml = generateParagraphXml({ type: 'heading', level: 2, runs: [{ text: 'x' }] }, ctx)
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>')
    expect(xml).not.toContain('w:outlineLvl')
  })
})
