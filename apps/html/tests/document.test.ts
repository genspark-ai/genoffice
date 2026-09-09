import { describe, expect, it } from 'vitest'
import { parseDocText, serializeDocText } from '../src/renderer/document/envelope'
import { applyPatches, invertPatches, validatePatchSet } from '../src/renderer/document/patch'
import { buildParseMap, elementCovering } from '../src/renderer/document/parse-map'

describe('envelope', () => {
  it.each([
    ['<p>a</p>\n', 'lf + trailing'],
    ['<p>a</p>', 'lf no trailing'],
    ['﻿<p>a</p>\r\n<p>b</p>\r\n', 'bom + crlf'],
    ['<p>a</p>\r\n<p>b</p>', 'crlf no trailing'],
    ['', 'empty'],
  ])('round-trips %s (%s)', (raw) => {
    expect(serializeDocText(parseDocText(raw))).toBe(raw)
  })

  it('normalizes the body to LF and strips the BOM', () => {
    const doc = parseDocText('﻿<a>\r\n</a>\r\n')
    expect(doc.text).toBe('<a>\n</a>')
    expect(doc.envelope).toEqual({ bom: true, eol: '\r\n', trailingNewline: true })
  })

  it('picks the majority EOL when a file mixes them', () => {
    const doc = parseDocText('a\r\nb\r\nc\nd\r\n')
    expect(doc.envelope.eol).toBe('\r\n')
    expect(doc.text).toBe('a\nb\nc\nd')
  })
})

describe('patches', () => {
  const text = '<p>hello</p><p>world</p>'

  it('applies multiple patches from the end so offsets stay valid', () => {
    const patches = [
      { from: 3, to: 8, text: 'HELLO' },
      { from: 15, to: 20, text: 'there' },
    ]
    expect(applyPatches(text, patches)).toBe('<p>HELLO</p><p>there</p>')
  })

  it('inverts patches back to the original', () => {
    const patches = [
      { from: 3, to: 8, text: 'HI' },
      { from: 15, to: 20, text: 'everyone' },
    ]
    const next = applyPatches(text, patches)
    expect(applyPatches(next, invertPatches(text, patches))).toBe(text)
  })

  it('rejects stale, out-of-bounds and overlapping sets', () => {
    const base = { origin: 'ai' as const, label: 't' }
    expect(validatePatchSet({ ...base, baseVersion: 1, patches: [] }, 2, 10)).toEqual({
      kind: 'stale',
      baseVersion: 1,
      currentVersion: 2,
    })
    expect(
      validatePatchSet(
        { ...base, baseVersion: 2, patches: [{ from: 5, to: 11, text: '' }] },
        2,
        10,
      ),
    ).toEqual({ kind: 'bounds', index: 0 })
    expect(
      validatePatchSet(
        {
          ...base,
          baseVersion: 2,
          patches: [
            { from: 0, to: 5, text: '' },
            { from: 4, to: 6, text: '' },
          ],
        },
        2,
        10,
      ),
    ).toEqual({ kind: 'overlap', index: 1 })
    expect(
      validatePatchSet(
        {
          ...base,
          baseVersion: 2,
          patches: [
            { from: 0, to: 5, text: '' },
            { from: 5, to: 6, text: '' },
          ],
        },
        2,
        10,
      ),
    ).toBeNull()
  })
})

describe('parse map', () => {
  it('maps start/end tags and inner ranges for a plain document', () => {
    const text = '<!doctype html>\n<html><body><h1 class="t">Hi</h1><img src="a.png"></body></html>'
    const map = buildParseMap(text, 1)
    const h1 = map.elements.find((e) => e.tag === 'h1')!
    expect(text.slice(...h1.startTag)).toBe('<h1 class="t">')
    expect(text.slice(...h1.inner)).toBe('Hi')
    expect(text.slice(...h1.endTag!)).toBe('</h1>')
    expect(text.slice(...h1.range)).toBe('<h1 class="t">Hi</h1>')
    expect(h1.path).toBe('html > body > h1:nth-of-type(1)')
    const img = map.elements.find((e) => e.tag === 'img')!
    expect(img.endTag).toBeNull()
    expect(text.slice(...img.range)).toBe('<img src="a.png">')
    expect(img.inner[0]).toBe(img.inner[1])
  })

  it('skips implied elements and handles implied end tags without overlapping siblings', () => {
    const text = '<div id=a>\n<p>one\n<p>two <b>bold\n<ul><li>x<li>y</ul>\n</div>'
    const map = buildParseMap(text, 1)
    expect(map.elements.map((e) => e.tag)).toEqual(['div', 'p', 'p', 'b', 'ul', 'li', 'li'])
    const [p1, p2] = map.elements.filter((e) => e.tag === 'p')
    expect(text.slice(...p1!.range)).toBe('<p>one\n')
    expect(p1!.endTag).toBeNull()
    expect(p1!.range[1]).toBeLessThanOrEqual(p2!.range[0])
    const [li1, li2] = map.elements.filter((e) => e.tag === 'li')
    expect(text.slice(...li1!.range)).toBe('<li>x')
    expect(text.slice(...li2!.range)).toBe('<li>y')
    // the reconstructed <b> (adoption agency) appears once, keyed by its start tag
    expect(map.elements.filter((e) => e.tag === 'b')).toHaveLength(1)
  })

  it('keeps sids stable across an edit and assigns fresh ones to new elements', () => {
    const v1 = '<body><section><h2>A</h2><p>x</p></section><section><h2>B</h2></section></body>'
    const m1 = buildParseMap(v1, 1)
    const secondH2 = m1.elements.filter((e) => e.tag === 'h2')[1]!
    const v2 =
      '<body><section><h2>A</h2><p>x</p><p>NEW</p></section><section><h2>B</h2></section></body>'
    const m2 = buildParseMap(v2, 2, m1)
    const secondH2After = m2.elements.filter((e) => e.tag === 'h2')[1]!
    expect(secondH2After.sid).toBe(secondH2.sid)
    const newP = m2.elements.filter((e) => e.tag === 'p')[1]!
    expect(m1.bySid.has(newP.sid)).toBe(false)
  })

  it('finds the smallest element covering a range and counts recovery errors', () => {
    const text = '<div><p>hello <em>world</em></p></div><span>'
    const map = buildParseMap(text, 1)
    const idx = text.indexOf('world')
    expect(elementCovering(map, idx, idx + 5)!.tag).toBe('em')
    expect(elementCovering(map, 5, idx + 5)!.tag).toBe('p')
    expect(map.errorCount).toBeGreaterThan(0)
  })

  it('does not choke on script content that looks like markup', () => {
    const text =
      '<html><head><script>var s = "</div>"; if (a < b) {}</script></head><body></body></html>'
    const map = buildParseMap(text, 1)
    expect(map.elements.map((e) => e.tag)).toEqual(['html', 'head', 'script', 'body'])
  })
})

describe('preview document', () => {
  it('injects a base tag into head only when the author has none', async () => {
    const { buildPreviewDocument, assetBaseHref } = await import('../src/main/preview-document')
    const base = assetBaseHref('/Users/h/My Docs')
    expect(base).toBe('html-asset://local/Users/h/My%20Docs/')
    expect(buildPreviewDocument('<html><head><title>x</title></head></html>', base)).toBe(
      `<html><head><base href="${base}"><title>x</title></head></html>`,
    )
    expect(buildPreviewDocument('<p>frag</p>', base)).toBe(`<base href="${base}"><p>frag</p>`)
    const authored = '<html><head><base href="https://x/"></head></html>'
    expect(buildPreviewDocument(authored, base)).toBe(authored)
    expect(buildPreviewDocument('<p>a</p>', null)).toBe('<p>a</p>')
    expect(assetBaseHref('C:\\Users\\h\\docs\\')).toBe('html-asset://local/C:/Users/h/docs/')
  })
})
