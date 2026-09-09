import { describe, expect, it } from 'vitest'
import { findSnippet } from '../src/renderer/document/match'
import { buildParseMap } from '../src/renderer/document/parse-map'
import { compileOps, type HtmlOp } from '../src/renderer/document/ops'
import { applyPatches } from '../src/renderer/document/patch'

const DOC = `<!doctype html>
<html>
<head><title>T</title><style>.card { color: red; }</style></head>
<body>
  <section class="hero">
    <h1 id="title">Hello &amp; welcome</h1>
    <p class="lead">First paragraph.</p>
  </section>
  <ul>
    <li>one</li>
    <li>two</li>
    <li>three</li>
  </ul>
  <img src="a.png">
</body>
</html>
`

function run(ops: HtmlOp[], text = DOC) {
  const map = buildParseMap(text, 1)
  const compiled = compileOps(text, map, ops)
  return {
    compiled,
    next: compiled.errors.length ? text : applyPatches(text, compiled.patches),
    map,
  }
}

function sidOf(text: string, tag: string, nth = 0): number {
  const map = buildParseMap(text, 1)
  return map.elements.filter((e) => e.tag === tag)[nth]!.sid
}

describe('findSnippet ladder', () => {
  it('exact and unique', () => {
    const r = findSnippet(DOC, '<p class="lead">First paragraph.</p>')
    expect(r.ok && r.rung).toBe(0)
  })
  it('refuses ambiguous matches instead of taking the first', () => {
    const r = findSnippet(DOC, '<li>')
    expect(r.ok).toBe(false)
    if (!r.ok) expect(r.kind === 'ambiguous' && r.candidates.length).toBe(3)
  })
  it('tolerates trailing whitespace and indentation drift', () => {
    const r = findSnippet(
      DOC,
      '<h1 id="title">Hello &amp; welcome</h1>   \n<p class="lead">First paragraph.</p>',
    )
    expect(r.ok).toBe(true)
    if (r.ok)
      expect(DOC.slice(r.from, r.to)).toBe(
        '<h1 id="title">Hello &amp; welcome</h1>\n    <p class="lead">First paragraph.</p>',
      )
  })
  it('folds entities and curly quotes on the canonical rung', () => {
    const r = findSnippet(DOC, 'Hello & welcome')
    expect(r.ok).toBe(true)
    if (r.ok) expect(DOC.slice(r.from, r.to)).toBe('Hello &amp; welcome')
  })
  it('block-anchors multi-line snippets with a slightly different interior', () => {
    const r = findSnippet(DOC, '<ul>\n<li>one</li>\n<li>2</li>\n<li>three</li>\n</ul>')
    expect(r.ok).toBe(true)
    if (r.ok) expect(DOC.slice(r.from, r.to).startsWith('<ul>')).toBe(true)
  })
  it('reports the nearest snippet when nothing matches', () => {
    const r = findSnippet(DOC, '<p class="lead">Second paragraph!!</p>')
    expect(r.ok).toBe(false)
    if (!r.ok && r.kind === 'not_found') expect(r.nearest?.similarity).toBeGreaterThan(0.6)
  })
})

describe('compileOps', () => {
  it('str_replace edits only the matched bytes', () => {
    const { next, compiled } = run([
      { op: 'str_replace', old: 'First paragraph.', new: 'Opening line.' },
    ])
    expect(compiled.errors).toEqual([])
    expect(next).toBe(DOC.replace('First paragraph.', 'Opening line.'))
  })
  it('str_replace scoped to a sid disambiguates repeated markup', () => {
    const li = sidOf(DOC, 'li', 1)
    const { next, compiled } = run([
      { op: 'str_replace', old: '<li>', new: '<li class="x">', sid: li },
    ])
    expect(compiled.errors).toEqual([])
    expect(next).toBe(DOC.replace('<li>two', '<li class="x">two'))
  })
  it('replace_all rewrites every exact occurrence and never succeeds silently', () => {
    const { next, compiled } = run([
      { op: 'str_replace', old: '<li>', new: '<li class="i">', replace_all: true },
    ])
    expect(compiled.errors).toEqual([])
    expect((next.match(/<li class="i">/g) ?? []).length).toBe(3)
    // loosely matching text (entity folded) is not good enough for replace_all
    const loose = run([{ op: 'str_replace', old: 'Hello & welcome', new: 'Hi', replace_all: true }])
    expect(loose.compiled.patches).toEqual([])
    expect(loose.compiled.errors[0]?.kind).toBe('not_found')
    expect(loose.compiled.errors[0]?.message).toMatch(/exact source text/)
  })

  it('rejects an empty old string (with and without replace_all) instead of looping', () => {
    expect(
      run([{ op: 'str_replace', old: '', new: 'x', replace_all: true }]).compiled.errors[0]?.kind,
    ).toBe('bad_args')
    expect(run([{ op: 'str_replace', old: '', new: 'x' }]).compiled.errors[0]?.kind).toBe(
      'bad_args',
    )
  })

  it('set_text escapes and replaces only the inner content', () => {
    const h1 = sidOf(DOC, 'h1')
    const { next } = run([{ op: 'set_text', sid: h1, text: 'A <b> & c' }])
    expect(next).toContain('<h1 id="title">A &lt;b&gt; &amp; c</h1>')
  })
  it('replace_element / insert_html / remove keep the rest byte-identical', () => {
    const p = sidOf(DOC, 'p')
    const ul = sidOf(DOC, 'ul')
    const img = sidOf(DOC, 'img')
    const { next, compiled } = run([
      { op: 'replace_element', sid: p, html: '<p class="lead">Rewritten.</p>' },
      { op: 'insert_html', sid: ul, position: 'append', html: '<li>four</li>' },
      { op: 'remove', sid: img },
    ])
    expect(compiled.errors).toEqual([])
    expect(next).toContain('<p class="lead">Rewritten.</p>')
    expect(next).toContain('<li>three</li>\n  <li>four</li></ul>')
    expect(next).not.toContain('<img')
    expect(next).toContain('</ul>\n</body>')
  })
  it('set_attr and set_style edit the start tag in place, preserving other attributes', () => {
    const h1 = sidOf(DOC, 'h1')
    const { next } = run([
      { op: 'set_attr', sid: h1, name: 'id', value: 'main-title' },
      { op: 'set_attr', sid: h1, name: 'data-x', value: 'y "q"' },
    ])
    expect(next).toContain('<h1 id="main-title" data-x="y &quot;q&quot;">')
    const sec = sidOf(next, 'section')
    const styled = run(
      [{ op: 'set_style', sid: sec, styles: { color: 'blue', padding: '4px' } }],
      next,
    ).next
    expect(styled).toContain('<section class="hero" style="color: blue; padding: 4px">')
    const sec2 = sidOf(styled, 'section')
    const unstyled = run(
      [{ op: 'set_style', sid: sec2, styles: { color: null, padding: null } }],
      styled,
    ).next
    expect(unstyled).toContain('<section class="hero">')
    const removed = run(
      [{ op: 'set_attr', sid: sidOf(unstyled, 'section'), name: 'class', value: null }],
      unstyled,
    ).next
    expect(removed).toContain('<section>')
  })
  it('move relocates an element and rejects moving into itself', () => {
    const one = sidOf(DOC, 'li', 0)
    const three = sidOf(DOC, 'li', 2)
    const { next, compiled } = run([{ op: 'move', sid: one, position: 'after', ref_sid: three }])
    expect(compiled.errors).toEqual([])
    expect(next.indexOf('<li>one</li>')).toBeGreaterThan(next.indexOf('<li>three</li>'))
    // whole lines travel: indentation and line breaks stay tidy
    expect(next).toContain('<ul>\n    <li>two</li>\n    <li>three</li>\n    <li>one</li>\n  </ul>')
    const two = sidOf(DOC, 'li', 1)
    const up = run([{ op: 'move', sid: two, position: 'before', ref_sid: one }]).next
    expect(up).toContain('<ul>\n    <li>two</li>\n    <li>one</li>\n    <li>three</li>\n  </ul>')
    const ul = sidOf(DOC, 'ul')
    expect(
      run([{ op: 'move', sid: ul, position: 'after', ref_sid: one }]).compiled.errors[0]?.kind,
    ).toBe('bad_args')
  })
  it('rejects void-element content ops, unknown sids and overlapping batches as a whole', () => {
    const img = sidOf(DOC, 'img')
    const p = sidOf(DOC, 'p')
    const sec = sidOf(DOC, 'section')
    const { compiled } = run([
      { op: 'set_text', sid: img, text: 'x' },
      { op: 'set_text', sid: 9999, text: 'x' },
      { op: 'replace_element', sid: sec, html: '<section></section>' },
      { op: 'set_text', sid: p, text: 'inner' },
    ])
    expect(compiled.errors.map((e) => e.kind)).toEqual(['void_element', 'no_such_sid', 'overlap'])
  })
})

describe('structural ops', () => {
  const T = '<div>\n  <p class="x">Hello <b>big</b> world</p>\n  <img src="a.png">\n</div>'
  it('set_tag renames both tags and keeps attributes', () => {
    const p = sidOf(T, 'p')
    const { next, compiled } = run([{ op: 'set_tag', sid: p, tag: 'h2' }], T)
    expect(compiled.errors).toEqual([])
    expect(next).toContain('<h2 class="x">Hello <b>big</b> world</h2>')
    expect(
      run([{ op: 'set_tag', sid: sidOf(T, 'img'), tag: 'p' }], T).compiled.errors[0]?.kind,
    ).toBe('bad_args')
  })
  it('set_text_node and wrap_text address direct text nodes', () => {
    const p = sidOf(T, 'p')
    const edited = run([{ op: 'set_text_node', sid: p, index: 1, text: ' & universe' }], T).next
    expect(edited).toContain('<b>big</b> &amp; universe</p>')
    const wrapped = run(
      [{ op: 'wrap_text', sid: p, index: 0, start: 0, end: 5, tag: 'em' }],
      T,
    ).next
    expect(wrapped).toContain('<p class="x"><em>Hello</em> <b>big</b> world</p>')
    const linked = run(
      [
        {
          op: 'wrap_text',
          sid: p,
          index: 1,
          start: 1,
          end: 6,
          tag: 'a',
          attrs: { href: 'https://x' },
        },
      ],
      T,
    ).next
    expect(linked).toContain('<b>big</b> <a href="https://x">world</a></p>')
    expect(
      run([{ op: 'wrap_text', sid: p, index: 5, start: 0, end: 1, tag: 'em' }], T).compiled
        .errors[0]?.kind,
    ).toBe('bad_args')
  })
  it('wrap_text maps decoded (DOM) offsets across entities', async () => {
    const { decodedToRaw } = await import('../src/renderer/document/ops')
    const raw = 'A &amp; B&nbsp;C'
    expect(decodedToRaw(raw, 0)).toBe(0)
    expect(decodedToRaw(raw, 2)).toBe(2) // before &amp;
    expect(decodedToRaw(raw, 3)).toBe(7) // after &amp;
    expect(decodedToRaw(raw, 4)).toBe(8) // "B"
    expect(decodedToRaw(raw, 5)).toBe(9) // before &nbsp;
    expect(decodedToRaw(raw, 6)).toBe(15) // "C"
    expect(decodedToRaw(raw, 99)).toBe(raw.length)
    const E = '<p>A &amp; B&nbsp;C</p>'
    const p = sidOf(E, 'p')
    // DOM selection "B" = decoded offsets [4, 5)
    expect(
      run([{ op: 'wrap_text', sid: p, index: 0, start: 4, end: 5, tag: 'strong' }], E).next,
    ).toBe('<p>A &amp; <strong>B</strong>&nbsp;C</p>')
  })

  it('unwrap removes only the tags', () => {
    const b = sidOf(T, 'b')
    expect(run([{ op: 'unwrap', sid: b }], T).next).toContain('<p class="x">Hello big world</p>')
    expect(run([{ op: 'unwrap', sid: sidOf(T, 'img') }], T).compiled.errors[0]?.kind).toBe(
      'void_element',
    )
  })
})

describe('preview instrumentation', () => {
  it('adds data-sid to every mapped start tag and appends the inspector before </body>', async () => {
    const { instrumentForPreview } = await import('../src/renderer/preview/instrument')
    const src = '<html><body><p class="a">x</p><img src="i.png"/><br></body></html>'
    const map = buildParseMap(src, 1)
    const out = instrumentForPreview(src, map, 'console.log(1)')
    expect(instrumentForPreview(src, map, "Number('__GX_VERSION__')")).toContain("Number('1')")
    const p = map.elements.find((e) => e.tag === 'p')!
    const img = map.elements.find((e) => e.tag === 'img')!
    const br = map.elements.find((e) => e.tag === 'br')!
    expect(out).toContain(`<p class="a" data-sid="${p.sid}">`)
    expect(out).toContain(`<img src="i.png" data-sid="${img.sid}"/>`)
    expect(out).toContain(`<br data-sid="${br.sid}">`)
    expect(out).toMatch(/<script data-gx-inspector>console\.log\(1\)<\/script><\/body>/)
    // a fragment without </body> still gets the script
    expect(instrumentForPreview('<p>x</p>', buildParseMap('<p>x</p>', 1), 's')).toMatch(
      /<\/p><script data-gx-inspector>s<\/script>$/,
    )
  })
})
