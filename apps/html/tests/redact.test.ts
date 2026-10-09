import { describe, expect, it } from 'vitest'
import {
  MARK_ATTR,
  MARK_ATTR_PREFIX,
  buildProjection,
  checkPlaceholders,
  collectPlaceholders,
  collectWithheld,
  isWholePlaceholder,
  placeholderInstruction,
  placeholderSource,
  readPlaceholderLabel,
  sanitizeLabel,
} from '../src/renderer/ai/redact'
import { buildParseMap } from '../src/renderer/document/parse-map'

/**
 * The projection is the whole feature for an HTML page: the reader's file is
 * untouched, and this is the only text a model ever sees. Every test here is
 * about one of three things — what gets withheld, what does NOT, and whether
 * offsets survive the substitution (they must, or `read_source`'s line numbering
 * silently starts returning the wrong lines).
 */

const LABEL = 'API key'
const SECRET = 'sk-ABC123'

const project = (html: string) => buildProjection(html, buildParseMap(html, 1, null))

describe('a page with nothing withheld', () => {
  it('is handed over verbatim', () => {
    const html = '<p>Hello</p><img src="logo.png">'
    const p = project(html)
    expect(p.empty).toBe(true)
    expect(p.view).toBe(html)
  })

  it('translates offsets as themselves', () => {
    const p = project('<p>Hello</p>')
    for (const at of [0, 3, 7, 12]) expect(p.toViewOffset(at)).toBe(at)
    for (const at of [0, 3, 7, 12]) expect(p.toRawOffset(at)).toBe(at)
  })
})

describe('withholding visible text', () => {
  const html = `<p>Call <span ${MARK_ATTR}="${LABEL}">13800138000</span> now</p>`

  it('swaps the span content for the marker', () => {
    const p = project(html)
    expect(p.view).toBe(
      `<p>Call <span ${MARK_ATTR}="${LABEL}">${placeholderSource(LABEL)}</span> now</p>`,
    )
    expect(p.view).not.toContain('13800138000')
  })

  it('keeps the wrapper so the outline and the source describe one document', () => {
    // get_outline reads the parse map, which still has the span. Dropping it
    // from the view would leave the model holding two different descriptions of
    // the same page — and the span is what the guard addresses by sid.
    const p = project(html)
    expect(p.view).toContain('<span')
    expect(collectWithheld(p.raw, buildParseMap(p.raw, 1, null))).toHaveLength(1)
  })

  it('leaves the page itself byte-for-byte identical', () => {
    expect(project(html).raw).toBe(html)
  })

  it('tells the model the withheld text was a value, not a sentence', () => {
    // the label is shown in the prompt instruction anyway, so what matters here
    // is only that the words themselves are gone
    const p = project(html)
    expect(p.view).not.toContain('13800138000')
    expect(p.labels()).toEqual([LABEL])
  })

  it('reports the label for the instruction', () => {
    expect(project(html).labels()).toEqual([LABEL])
  })
})

describe('withholding an attribute value', () => {
  it('withholds only the marked attribute, not its neighbours', () => {
    const marked = `<img src="https://cdn.example.com/logo.png" alt="company logo" ${MARK_ATTR_PREFIX}src="${LABEL}">`
    const p = project(marked)
    expect(p.view).toContain('alt="company logo"')
    expect(p.view).not.toContain('cdn.example.com')
    expect(p.view).toContain(placeholderSource(LABEL))
  })

  it('names the attribute it is withholding, for the refusal message', () => {
    const marked = `<img src="secret.png" ${MARK_ATTR_PREFIX}src="${LABEL}">`
    const spans = collectWithheld(marked, buildParseMap(marked, 1, null))
    expect(spans[0]!.kind).toBe('attr')
    expect(spans[0]!.where).toContain('src')
  })

  it('does not withhold an unmarked attribute whose value names a marked one', () => {
    // `alt` holds the text "src=…" — a scanner matching on the name rather than
    // on the parsed attribute list would swallow the wrong value
    const marked = `<img src="secret.png" alt="src= not this" ${MARK_ATTR_PREFIX}src="logo">`
    const p = project(marked)
    expect(p.view).toContain('src= not this')
    expect(p.view).toContain(placeholderSource('logo'))
    expect(p.view).not.toContain('secret.png')
  })

  it('reads a single-quoted attribute value', () => {
    const marked = `<img src='secret.png' alt='logo' ${MARK_ATTR_PREFIX}src='brand'>`
    const p = project(marked)
    expect(p.view).toContain(`alt='logo'`)
    expect(p.view).toContain(placeholderSource('brand'))
    expect(p.view).not.toContain('secret.png')
  })

  it('reads an unquoted attribute value', () => {
    const marked = `<img src=secret.png alt=logo ${MARK_ATTR_PREFIX}src=brand>`
    const p = project(marked)
    expect(p.view).toContain('alt=logo')
    expect(p.view).toContain(placeholderSource('brand'))
    expect(p.view).not.toContain('secret.png')
  })

  it('ignores a mark for an attribute the element does not have', () => {
    const marked = `<img src="a.png" ${MARK_ATTR_PREFIX}href="${LABEL}">`
    expect(project(marked).empty).toBe(true)
  })
})

describe('withholding a value inside a script', () => {
  const html = `<script>const KEY = /*gx:redact:${LABEL}*/ "${SECRET}"; render(KEY);</script>`

  it('withholds the string literal the comment stands in front of', () => {
    const p = project(html)
    expect(p.view).not.toContain(SECRET)
    expect(p.view).toContain(placeholderSource(LABEL))
    // the rest of the script is untouched, so the model can still read the logic
    expect(p.view).toContain('render(KEY)')
  })

  it('leaves the comment in the file — the page needs it inert, not gone', () => {
    const p = project(html)
    expect(p.raw).toBe(html)
    expect(p.raw).toContain('/*gx:redact:')
  })

  it('honours the mark only inside a real script', () => {
    const inText = `<p>literal /*gx:redact:${LABEL}*/ "${SECRET}" in a paragraph</p>`
    expect(project(inText).empty).toBe(true)
  })

  it('handles a single-quoted literal', () => {
    const single = `<script>const K = /*gx:redact:${LABEL}*/ '${SECRET}';</script>`
    expect(project(single).view).not.toContain(SECRET)
  })

  it('leaves the quotes alone so the script still parses', () => {
    const p = project(html)
    // the quotes are part of the source, only the literal's content was withheld
    expect(p.raw.slice(p.raw.indexOf('"'), p.raw.indexOf('"') + 1)).toBe('"')
  })

  it('a mark with no literal after it is not a span', () => {
    const dangling = `<script>/*gx:redact:${LABEL}*/ render();</script>`
    expect(project(dangling).empty).toBe(true)
  })
})

describe('offsets survive the substitution', () => {
  const html = [
    '<p>line one</p>',
    `<p><span ${MARK_ATTR}="${LABEL}">${SECRET}</span></p>`,
    '<p>line three</p>',
  ].join('\n')

  it('a raw offset before the span is unchanged', () => {
    const p = project(html)
    const at = html.indexOf('line one')
    expect(p.toViewOffset(at)).toBe(at)
  })

  it('an offset after the span shifts by exactly the length difference', () => {
    const p = project(html)
    const at = html.indexOf('line three')
    const expected = at - (SECRET.length - placeholderSource(LABEL).length)
    expect(p.toViewOffset(at)).toBe(expected)
  })

  it('round-trips an offset that is outside every span', () => {
    const p = project(html)
    for (const needle of ['line one', 'line three']) {
      const raw = html.indexOf(needle) + 2
      expect(p.toRawOffset(p.toViewOffset(raw))).toBe(raw)
    }
  })

  it('an offset inside the span lands on the marker, never on the words', () => {
    const p = project(html)
    const inside = html.indexOf(SECRET) + 3
    const view = p.toViewOffset(inside)
    expect(p.view.slice(view, view + SECRET.length)).not.toContain(SECRET)
  })

  it('projects a raw range the way read_source slices one', () => {
    const p = project(html)
    const from = html.indexOf('<p>line three')
    expect(p.projectRange(from, html.length)).toBe('<p>line three</p>')
    // and a range that starts inside the last line, mid-text
    const mid = html.indexOf('line three') + 4
    expect(p.projectRange(mid, html.length)).toBe(' three</p>')
  })

  it('a withheld line is still a line: numbering does not shift the count', () => {
    const p = project(html)
    expect(p.view.split('\n')).toHaveLength(html.split('\n').length)
  })

  it('survives two spans in the same file', () => {
    const two = `<p><span ${MARK_ATTR}="one">AAAA</span> mid <span ${MARK_ATTR}="two">BBBB</span></p>`
    const p = project(two)
    expect(p.view).toBe(
      `<p><span ${MARK_ATTR}="one">${placeholderSource('one')}</span> mid ` +
        `<span ${MARK_ATTR}="two">${placeholderSource('two')}</span></p>`,
    )
    expect(p.labels()).toEqual(['one', 'two'])
    // the tail after the LAST span's content is what has to land in the right
    // place: both substitutions happened, so its shift is the sum of both. The
    // tail runs to the end of the file, wrapper and all.
    const afterLast = p.toViewOffset(p.spans[1]!.rawTo)
    expect(p.view.slice(afterLast)).toBe(`${placeholderSource('two')}</span></p>`)
    // and the end of the document maps onto the end of the view
    expect(p.toViewOffset(two.length)).toBe(p.view.length)
  })
})

describe('the pure helpers, ported', () => {
  it('sanitises a label', () => {
    expect(sanitizeLabel('  a<b>{"c"}  ')).toBe('abc')
    expect(sanitizeLabel('x'.repeat(80)).length).toBeLessThanOrEqual(40)
  })

  it('builds and reads a marker', () => {
    expect(placeholderSource('client phone')).toBe('{{client phone}}')
    expect(isWholePlaceholder('{{a}}')).toBe(true)
    expect(isWholePlaceholder('{{a')).toBe(false)
    expect(readPlaceholderLabel('{{a}}')).toBe('a')
    expect(readPlaceholderLabel('plain')).toBeNull()
  })

  it('collects markers', () => {
    expect(collectPlaceholders('a {{x}} b {{y}}')).toEqual(['{{x}}', '{{y}}'])
  })

  it('reports a mangled marker', () => {
    // a marker cut in half is both a malformed leftover and a marker gone
    expect(
      checkPlaceholders('{{a}}', '{{a')
        .map((i) => i.reason)
        .sort(),
    ).toEqual(['missing', 'split'])
    expect(checkPlaceholders('{{a}}', '{{b}}')[0]!.reason).toMatch(/missing|unknown/)
    expect(checkPlaceholders('{{a}}', '{{a}}')).toHaveLength(0)
  })

  it('the instruction names the placeholders and says what a marker may stand for', () => {
    const text = placeholderInstruction(['a', 'b', 'a'])
    expect(text).toContain('{{a}}')
    expect(text).toContain('{{b}}')
    expect(text.match(/\{\{a\}\}/g)).toHaveLength(1)
    expect(text).toMatch(/attribute/)
    expect(text).toMatch(/script/i)
  })
})
