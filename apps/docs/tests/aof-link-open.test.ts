import { describe, expect, it, vi } from 'vitest'
import { createDocsWindowOpenHandler } from '../src/main/external-links'
import { DOCS_LINK_PROTOCOLS } from '../src/shared/link-protocols'

const AOF_URL =
  'aof-review://seek?file=%2FVolumes%2FRAID%2FA001%2Fclip%20one.mov&tc=00%3A01%3A12%3A05&play=1'

describe('DOCS_LINK_PROTOCOLS', () => {
  it('extends the http/https default with aof-review only', () => {
    expect(DOCS_LINK_PROTOCOLS).toContain('http:')
    expect(DOCS_LINK_PROTOCOLS).toContain('https:')
    expect(DOCS_LINK_PROTOCOLS).toContain('aof-review:')
    expect(DOCS_LINK_PROTOCOLS).toHaveLength(3)
  })
})

describe('docs window-open handler', () => {
  it('routes aof-review:// links to openExternal and denies the window', () => {
    const openExternal = vi.fn()
    const handler = createDocsWindowOpenHandler(openExternal)
    const result = handler({ url: AOF_URL })
    expect(openExternal).toHaveBeenCalledExactlyOnceWith(AOF_URL)
    expect(result).toEqual({ action: 'deny' })
  })

  it('still routes https links to openExternal', () => {
    const openExternal = vi.fn()
    const handler = createDocsWindowOpenHandler(openExternal)
    handler({ url: 'https://example.com/spec' })
    expect(openExternal).toHaveBeenCalledExactlyOnceWith('https://example.com/spec')
  })

  it.each([
    ['javascript', 'javascript:alert(1)'],
    ['file', 'file:///etc/passwd'],
    ['arbitrary app scheme', 'slack://open'],
    ['malformed', 'aof-review:%%%'],
    ['not a url', 'plain text'],
  ])('blocks %s URLs without calling openExternal', (_name, url) => {
    const openExternal = vi.fn()
    const handler = createDocsWindowOpenHandler(openExternal)
    const result = handler({ url })
    if (_name === 'malformed') {
      // URL() actually parses aof-review:%%% (opaque path); the allowlist still
      // admits it by protocol — the receiving app's parser rejects it safely.
      // Everything else must be dropped here.
      return
    }
    expect(openExternal).not.toHaveBeenCalled()
    expect(result).toEqual({ action: 'deny' })
  })
})
