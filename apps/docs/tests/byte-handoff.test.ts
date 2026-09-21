import { describe, expect, it } from 'vitest'
import { handOffBytes, pendingHandoffCount, takeHandoff } from '../src/main/byte-handoff'

describe('byte handoff', () => {
  it('serves the bytes exactly once behind a media-scheme URL', () => {
    const bytes = Buffer.from('PK\u0003\u0004 document')
    const url = handOffBytes(bytes)
    expect(url.startsWith('genoffice-docx-media://handoff/')).toBe(true)
    expect(pendingHandoffCount()).toBe(1)
    expect(takeHandoff(url)).toBe(bytes)
    expect(takeHandoff(url)).toBeNull()
    expect(pendingHandoffCount()).toBe(0)
  })

  it('ignores lazy-media and foreign URLs', () => {
    handOffBytes(Buffer.from('x'))
    expect(takeHandoff(`genoffice-docx-media://${'a'.repeat(64)}/word/media/image1.png`)).toBeNull()
    expect(takeHandoff('genoffice-docx-media://handoff/not-a-token')).toBeNull()
    expect(takeHandoff('https://example.com/handoff/x')).toBeNull()
    expect(pendingHandoffCount()).toBe(1)
  })
})
