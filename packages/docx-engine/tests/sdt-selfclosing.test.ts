import { describe, expect, it } from 'vitest'
import { parseSdtBlock } from '../src/parse-sdt'

const sdt = (content: string) =>
  '<w:sdt><w:sdtPr><w:alias w:val="A"/><w:text/></w:sdtPr><w:sdtContent>' +
  content +
  '</w:sdtContent></w:sdt>'

describe('parseSdtBlock with self-closing paragraphs', () => {
  it('takes a self-closing <w:p/> as the whole first paragraph', () => {
    // Word writes empty paragraphs inside rich-text content controls as
    // <w:p w14:paraId="…"/>; the extracted slice must not swallow the next
    // paragraph into the same block
    const r = parseSdtBlock(sdt('<w:p w14:paraId="1A2B3C4D"/><w:p><w:r><w:t>x</w:t></w:r></w:p>'))
    expect(r).not.toBeNull()
    expect(r!.pXml).toBe('<w:p w14:paraId="1A2B3C4D"/>')
  })

  it('takes a bare <w:p/> the same way', () => {
    const r = parseSdtBlock(sdt('<w:p/><w:p><w:r><w:t>x</w:t></w:r></w:p>'))
    expect(r).not.toBeNull()
    expect(r!.pXml).toBe('<w:p/>')
  })

  it('still matches a regular open/close pair', () => {
    const r = parseSdtBlock(sdt('<w:p><w:r><w:t>Jane Doe</w:t></w:r></w:p>'))
    expect(r).not.toBeNull()
    expect(r!.pXml).toBe('<w:p><w:r><w:t>Jane Doe</w:t></w:r></w:p>')
  })
})
