import { describe, expect, it } from 'vitest'
import { inlineLazyMediaInHtml } from '../src/main/lazy-media-inline'

const HASH = 'a'.repeat(64)
const url = (part: string) => `genoffice-docx-media://${HASH}/word/media/${part}`

describe('inlineLazyMediaInHtml', () => {
  it('replaces every lazily served picture with its bytes, reading each URL once', async () => {
    const reads: string[] = []
    const html = `<p><img src="${url('image1.png')}" width="10"><img src="${url('image1.png')}"><img src="${url("photo's%202.jpg")}"></p>`
    const out = await inlineLazyMediaInHtml(html, async (u) => {
      reads.push(u)
      return u.endsWith('.png')
        ? { body: new Uint8Array([137, 80, 78, 71]), mime: 'image/png' }
        : { body: new Uint8Array([255, 216]), mime: 'image/jpeg' }
    })
    expect(reads).toEqual([url('image1.png'), url("photo's%202.jpg")])
    expect(out).toBe(
      '<p><img src="data:image/png;base64,iVBORw==" width="10"><img src="data:image/png;base64,iVBORw=="><img src="data:image/jpeg;base64,/9g="></p>',
    )
  })

  it('keeps a reference the source cannot serve and leaves other URLs alone', async () => {
    const html = `<img src="${url('gone.png')}"><img src="data:image/png;base64,AAAA"><a href="https://example.com/x">x</a>`
    expect(await inlineLazyMediaInHtml(html, async () => null)).toBe(html)
  })
})
