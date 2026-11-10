import { describe, expect, it } from 'vitest'
import { sniffImageMime } from '../src/renderer/ai/tools'

// the shared ai:fetch-image handler labels every non-png/gif content-type as
// jpeg — the sniff must catch mislabeled webp/svg before they enter the docx
describe('sniffImageMime', () => {
  const b64 = (bytes: string) => Buffer.from(bytes, 'binary').toString('base64')

  it('identifies png, gif and jpeg by magic bytes', () => {
    expect(sniffImageMime(b64('\x89PNG\r\n\x1a\n....'))).toBe('image/png')
    expect(sniffImageMime(b64('GIF89a......'))).toBe('image/gif')
    expect(sniffImageMime(b64('\xff\xd8\xff\xe0....JFIF'))).toBe('image/jpeg')
  })

  it('rejects webp and svg regardless of the reported content-type', () => {
    expect(sniffImageMime(b64('RIFF\0\0\0\0WEBPVP8 '))).toBeNull()
    expect(sniffImageMime(b64('<svg xmlns="ht'))).toBeNull()
    expect(sniffImageMime('%%%not-base64%%%')).toBeNull()
  })
})
