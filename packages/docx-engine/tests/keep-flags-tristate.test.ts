import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const HEADING_STYLE =
  '<w:style w:type="paragraph" w:styleId="KeepHead"><w:name w:val="Keep Head"/>' +
  '<w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="1"/></w:pPr></w:style>'

const para = (pPr: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="KeepHead"/>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

describe('keepNext / keepLines tri-state', () => {
  it('an explicit w:val="0" overrides a keep-on style, absent stays undefined', async () => {
    const docx = await buildDocx({
      bodyXml:
        para('<w:keepNext w:val="0"/><w:keepLines w:val="0"/>', 'off') +
        para('', 'inherit') +
        para('<w:keepNext/><w:keepLines/>', 'on'),
      extraStylesXml: HEADING_STYLE,
    })
    const parsed = await parseDocx(docx)
    const [off, inherit, on] = parsed.blocks.filter((b) => !b.hidden)
    expect(off.format?.keepNext).toBe(false)
    expect(off.format?.keepLines).toBe(false)
    expect(inherit.format?.keepNext).toBeUndefined()
    expect(inherit.format?.keepLines).toBeUndefined()
    expect(on.format?.keepNext).toBe(true)
    expect(on.format?.keepLines).toBe(true)
    expect(parsed.styles.get('KeepHead')?.display?.keepNext).toBe(true)
  })
})
