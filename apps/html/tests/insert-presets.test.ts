import { describe, expect, it } from 'vitest'
import {
  INSERT_KINDS,
  insertOp,
  insertPresetHtml,
  type InsertLabels,
} from '../src/renderer/document/insert-presets'
import { compileOps } from '../src/renderer/document/ops'
import { buildParseMap } from '../src/renderer/document/parse-map'
import { applyPatches } from '../src/renderer/document/patch'

const LABELS: InsertLabels = {
  heading: 'New heading',
  paragraph: 'Type <here> & "there".',
  listItem: 'Item',
  button: 'Button',
  sectionTitle: 'Section',
  sectionBody: 'Body',
  imageAlt: '',
  tableHeaders: ['A', 'B', 'C'],
  tableCell: 'Cell',
}

const DOC = `<!doctype html>
<html>
<head><title>T</title></head>
<body>
  <section class="hero">
    <h1 id="title">Hello</h1>
    <p class="lead">First paragraph.</p>
  </section>
  <ul>
    <li>one</li>
    <li>two</li>
  </ul>
  <table><tbody><tr><td>cell</td></tr></tbody></table>
</body>
</html>
`

/** `tag` or `tag#n` (nth element with that tag, 1-based) */
function apply(text: string, pick: string | null, html: string) {
  const map = buildParseMap(text, 1)
  const [tag, nth = '1'] = (pick ?? '').split('#')
  const selected = pick ? map.elements.filter((e) => e.tag === tag)[Number(nth) - 1] : undefined
  if (pick) expect(selected).toBeDefined()
  const op = insertOp(map, selected, html)
  expect(op).not.toBeNull()
  const compiled = compileOps(text, map, [op!])
  expect(compiled.errors).toEqual([])
  return applyPatches(text, compiled.patches)
}

describe('insert presets', () => {
  it('every kind produces markup that parses back to exactly one new top-level element', () => {
    for (const kind of INSERT_KINDS) {
      const html = insertPresetHtml(kind, LABELS, { imageSrc: 'assets/a.png' })
      const next = apply(DOC, null, html)
      const before = buildParseMap(DOC, 1)
      const after = buildParseMap(next, 2)
      const body = after.elements.find((e) => e.tag === 'body')!
      const topLevelBefore = before.elements.filter((e) => e.depth === body.depth + 1).length
      const topLevelAfter = after.elements.filter((e) => e.depth === body.depth + 1).length
      expect(topLevelAfter, kind).toBe(topLevelBefore + 1)
      expect(after.errorCount, kind).toBe(before.errorCount)
    }
  })

  it('escapes placeholder text and the image source', () => {
    expect(insertPresetHtml('paragraph', LABELS)).toBe(
      '<p>Type &lt;here> &amp; &quot;there&quot;.</p>',
    )
    expect(insertPresetHtml('image', LABELS, { imageSrc: 'a"b.png' })).toContain(
      'src="a&quot;b.png"',
    )
  })

  it('lands right after the selected element, on its own line', () => {
    const next = apply(DOC, 'h1', '<p>x</p>')
    expect(next).toContain('<h1 id="title">Hello</h1>\n<p>x</p>\n    <p class="lead">')
  })

  it('appends to the body when nothing is selected', () => {
    const next = apply(DOC, null, '<hr>')
    expect(next).toMatch(/<\/table><hr>\n<\/body>|<\/table>\n<hr>\n<\/body>/)
    const after = buildParseMap(next, 2)
    const hr = after.elements.find((e) => e.tag === 'hr')!
    expect(after.bySid.get(hr.parentSid!)!.tag).toBe('body')
  })

  it('never puts a block between list items or table cells: it goes after the whole list / table', () => {
    for (const path of ['li#2', 'td']) {
      const next = apply(DOC, path, '<p>x</p>')
      const after = buildParseMap(next, 2)
      const p = after.elements.find((e) => e.tag === 'p' && !e.path.includes('section'))!
      expect(after.bySid.get(p.parentSid!)!.tag, path).toBe('body')
    }
  })

  it('without a body in the map, lands after the last top-level content element', () => {
    const text = '<div>only</div>'
    const map = buildParseMap(text, 1)
    const stripped = { ...map, elements: map.elements.filter((e) => e.tag !== 'body') }
    const div = stripped.elements.find((e) => e.tag === 'div')!
    expect(insertOp(stripped, undefined, '<p>x</p>')).toEqual({
      op: 'insert_html',
      sid: div.sid,
      position: 'after',
      html: '\n<p>x</p>',
    })
  })

  it('the table preset has one header row and two body rows with a cell per header', () => {
    const next = apply(DOC, null, insertPresetHtml('table', LABELS))
    const after = buildParseMap(next, 2)
    const tables = after.elements.filter((e) => e.tag === 'table')
    expect(tables).toHaveLength(2)
    const fresh = tables[1]!
    const inside = (tag: string) =>
      after.elements.filter(
        (e) => e.tag === tag && e.range[0] > fresh.range[0] && e.range[1] < fresh.range[1],
      )
    expect(inside('th')).toHaveLength(3)
    expect(inside('td')).toHaveLength(6)
    expect(inside('tr')).toHaveLength(3)
  })

  it('the table preset sizes to the picked columns and body rows', () => {
    const html = insertPresetHtml(
      'table',
      { ...LABELS, tableHeaders: ['A', 'B'] },
      { tableBodyRows: 4 },
    )
    const next = apply(DOC, null, html)
    const after = buildParseMap(next, 2)
    const fresh = after.elements.filter((e) => e.tag === 'table')[1]!
    const inside = (tag: string) =>
      after.elements.filter(
        (e) => e.tag === tag && e.range[0] > fresh.range[0] && e.range[1] < fresh.range[1],
      )
    expect(inside('th')).toHaveLength(2)
    expect(inside('td')).toHaveLength(8)
    expect(inside('tr')).toHaveLength(5)
  })

  it('documents without an explicit body still get an insert point', () => {
    const at = (doc: string) => {
      const map = buildParseMap(doc, 1)
      const op = insertOp(map, undefined, '<p>x</p>')
      return op && op.op === 'insert_html'
        ? { tag: map.bySid.get(op.sid)!.tag, position: op.position }
        : op
    }
    expect(at('<h1>x</h1><p>y</p>')).toEqual({ tag: 'p', position: 'after' })
    expect(
      at('<!doctype html><html><head><title>t</title></head><h1>x</h1><p>y</p></html>'),
    ).toEqual({
      tag: 'p',
      position: 'after',
    })
    expect(at('<html><head><title>t</title></head></html>')).toEqual({
      tag: 'head',
      position: 'after',
    })
    expect(at('<html><body><main><h1>a</h1></main></body></html>')).toEqual({
      tag: 'main',
      position: 'append',
    })
    expect(at('just text')).toBeNull()
  })
})
