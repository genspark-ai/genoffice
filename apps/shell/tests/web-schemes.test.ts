import { describe, expect, it } from 'vitest'
import { rewriteIn, rewriteOut } from '../src/web/schemes'

describe('web bridge scheme rewriting', () => {
  it('rewrites scheme literals to same-origin paths for docs payloads', () => {
    const reply = {
      dataUrl: 'genoffice-docx-media://handoff/abc-uuid',
      nested: { list: ['genoffice-docx-media://hash/media/image1.png'] },
    }
    expect(rewriteOut('docs', reply)).toEqual({
      dataUrl: '/__media://handoff/abc-uuid',
      nested: { list: ['/__media://hash/media/image1.png'] },
    })
  })

  it('maps paths back to scheme literals on the way in', () => {
    expect(rewriteIn('markdown', { src: '/__mdasset:///notes/pic.png' })).toEqual({
      src: 'md-asset:///notes/pic.png',
    })
  })

  it('rewrites every html scheme in one payload', () => {
    const out = rewriteOut('html', {
      frame: 'html-preview://view-9/',
      img: 'html-asset://local/notes/a/',
    })
    expect(out).toEqual({ frame: '/__preview://view-9/', img: '/__htmlasset://local/notes/a/' })
  })

  it('leaves other apps untouched', () => {
    const value = { src: 'genoffice-docx-media://hash/p', md: 'md-asset://x' }
    expect(rewriteOut('sheets', value)).toBe(value)
  })

  it('passes typed arrays through the rewrite unharmed', () => {
    // document bytes ride next to strings in the same payload; index-key
    // reconstruction would turn them into {0:1, 1:2, …} and handlers'
    // Buffer.from(data) would throw
    const bytes = new Uint8Array([1, 2, 3])
    const out = rewriteIn('docs', { data: bytes, url: '/__media://h/p' }) as {
      data: Uint8Array
      url: string
    }
    expect(out.data).toBe(bytes)
    expect(out.url).toBe('genoffice-docx-media://h/p')
  })

  it('keeps classed instances (Date) intact', () => {
    const date = new Date(0)
    const out = rewriteOut('docs', { when: date })
    expect((out as { when: Date }).when).toBe(date)
  })
})
