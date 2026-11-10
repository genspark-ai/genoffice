import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { rendererUrl, resolveRendererFile } from '../src/renderer-scheme'

describe('rendererUrl', () => {
  it('builds the scheme URL with the query when no dev server is configured', () => {
    expect(rendererUrl(undefined, 'sheets', { mode: 'tab' })).toBe(
      'genoffice-app://sheets/index.html?mode=tab',
    )
    expect(rendererUrl(undefined, 'docs')).toBe('genoffice-app://docs/index.html')
  })

  it('appends the query to a dev URL that already carries params', () => {
    expect(rendererUrl('http://localhost:5174/?x=1', 'sheets', { mode: 'tab' })).toBe(
      'http://localhost:5174/?x=1&mode=tab',
    )
  })

  it('throws a descriptive error for a malformed dev URL', () => {
    expect(() => rendererUrl(':::', 'sheets')).toThrow(
      'Invalid dev URL for renderer "sheets": ":::"',
    )
    expect(() => rendererUrl('', 'docs')).toThrow('Invalid dev URL for renderer "docs": ""')
  })

  it('rejects oversized query params', () => {
    const many = Object.fromEntries(Array.from({ length: 25 }, (_, i) => [`k${i}`, 'v']))
    expect(() => rendererUrl(undefined, 'docs', many)).toThrow(/Too many/)
    expect(() => rendererUrl(undefined, 'docs', { k: 'x'.repeat(5000) })).toThrow(/too long/)
  })
})

describe('resolveRendererFile', () => {
  // Resolve through node:path so expectations match on Windows and POSIX.
  const rootDir = resolve('/out/sheets/renderer')
  const roots = new Map([['sheets', rootDir]])

  it('maps the path under the host root', () => {
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/index.html?mode=tab')).toBe(
      join(rootDir, 'index.html'),
    )
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/assets/a%20b.js')).toBe(
      join(rootDir, 'assets', 'a b.js'),
    )
  })

  it('keeps dot segments inside the root and rejects unknown hosts and unparsable URLs', () => {
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/../../etc/passwd')).toBe(
      join(rootDir, 'etc', 'passwd'),
    )
    expect(resolveRendererFile(roots, 'genoffice-app://docs/index.html')).toBeNull()
    expect(resolveRendererFile(roots, 'not a url')).toBeNull()
  })

  it('blocks encoded traversal sequences', () => {
    // Percent-encoded dots are normalized by the URL parser and stay inside.
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/%2e%2e/%2e%2e/etc/passwd')).toBe(
      join(rootDir, 'etc', 'passwd'),
    )
    // Encoded separators decode to a traversal that escapes the root.
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/%2e%2e%2fetc%2fpasswd')).toBeNull()
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/..%5c..%5cetc%5cpasswd')).toBeNull()
  })

  it('ignores query and fragment suffixes when mapping files', () => {
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/index.html?mode=tab#section')).toBe(
      join(rootDir, 'index.html'),
    )
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/index.html#frag')).toBe(
      join(rootDir, 'index.html'),
    )
  })

  it('rejects the root directory itself, which is not a file to serve', () => {
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/')).toBeNull()
    expect(resolveRendererFile(roots, 'genoffice-app://sheets')).toBeNull()
    expect(resolveRendererFile(roots, 'genoffice-app://sheets/.')).toBeNull()
  })

  it('rejects overlong and NUL paths', () => {
    expect(resolveRendererFile(roots, `genoffice-app://sheets/${'a'.repeat(5000)}.js`)).toBeNull()
  })
})
