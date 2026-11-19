import { describe, expect, it } from 'vitest'

import { safeExternalUrl } from '../src/index'

describe('safeExternalUrl', () => {
  it('accepts http and https URLs', () => {
    expect(safeExternalUrl('https://example.com/a?b=c')).toBe('https://example.com/a?b=c')
    expect(safeExternalUrl('http://example.com')).toBe('http://example.com')
  })

  it('rejects non-string input', () => {
    expect(safeExternalUrl(undefined)).toBeNull()
    expect(safeExternalUrl(null)).toBeNull()
    expect(safeExternalUrl(42)).toBeNull()
    expect(safeExternalUrl({ toString: () => 'https://example.com' })).toBeNull()
  })

  it('rejects malformed URLs', () => {
    expect(safeExternalUrl('not a url')).toBeNull()
    expect(safeExternalUrl('')).toBeNull()
    expect(safeExternalUrl('http//missing-colon.com')).toBeNull()
  })

  it('rejects dangerous protocols by default', () => {
    expect(safeExternalUrl('file:///etc/passwd')).toBeNull()
    expect(safeExternalUrl('javascript:alert(1)')).toBeNull()
    expect(safeExternalUrl('smb://server/share')).toBeNull()
    expect(safeExternalUrl('mailto:a@b.com')).toBeNull()
    // prefix tricks that pass a bare startsWith('http') check
    expect(safeExternalUrl('httpx://evil.example')).toBeNull()
  })

  it('supports a custom protocol allowlist', () => {
    const opts = { allowedProtocols: ['https:', 'mailto:'] }
    expect(safeExternalUrl('mailto:a@b.com', opts)).toBe('mailto:a@b.com')
    expect(safeExternalUrl('http://example.com', opts)).toBeNull()
  })

  it('trims surrounding whitespace instead of returning it verbatim', () => {
    // The WHATWG parser strips surrounding whitespace while parsing, so the
    // untrimmed input validates but would fail in shell.openExternal.
    expect(safeExternalUrl('  https://example.com/a?b=c  ')).toBe('https://example.com/a?b=c')
    expect(safeExternalUrl('\nhttp://example.com\n')).toBe('http://example.com')
    expect(safeExternalUrl('   ')).toBeNull()
    expect(safeExternalUrl('  file:///etc/passwd  ')).toBeNull()
  })
})
