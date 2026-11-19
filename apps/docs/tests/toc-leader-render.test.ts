import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { FieldDisplay } from '@genoffice/docx-engine'
import { renderFieldSpec } from '../src/renderer/editor/protected-render'
import { tocLeaderGap } from '../src/renderer/editor/decoration-extensions'

type Span = [string, Record<string, string>, string]

const cells = (field: FieldDisplay) =>
  (renderFieldSpec(field) as unknown[]).filter(Array.isArray) as unknown[][]
const body = (field: FieldDisplay) =>
  cells(field).find((c) => (c[1] as Record<string, string>).class === 'doc-toc-body')!
const tail = (field: FieldDisplay) =>
  (body(field).filter(Array.isArray) as unknown[][]).find(
    (c) => (c[1] as Record<string, string>).class === 'doc-toc-tail',
  )!
const leaderSpan = (field: FieldDisplay) =>
  (tail(field).filter(Array.isArray) as unknown[][]).find((c) =>
    (c[1] as Record<string, string>).class?.startsWith('doc-toc-dots'),
  ) as Span

describe('TOC entry leader rendering', () => {
  const base: FieldDisplay = { kind: 'tocLine', left: 'Title', right: '3', level: 1 }

  it('draws the leader as repeated glyphs of the entry font, dots when the leader is unknown', () => {
    expect(leaderSpan(base)[2]).toBe('.'.repeat(220))
    expect(leaderSpan({ ...base, leader: 'dot' })[2]).toBe('.'.repeat(220))
    expect(leaderSpan({ ...base, leader: 'hyphen' })[2]).toBe('-'.repeat(220))
    expect(leaderSpan({ ...base, leader: 'underscore' })[2]).toBe('_'.repeat(220))
    expect(leaderSpan({ ...base, leader: 'middleDot' })[2]).toBe('\u00b7'.repeat(220))
  })

  it('a bare right tab leaves the gap empty, heavy is a bold underscore run', () => {
    const none = leaderSpan({ ...base, leader: 'none' })
    expect(none[2]).toBe('')
    expect(none[1].class).toBe('doc-toc-dots')
    const heavy = leaderSpan({ ...base, leader: 'heavy' })
    expect(heavy[2]).toBe('_'.repeat(220))
    expect(heavy[1].class).toBe('doc-toc-dots doc-toc-leader-heavy')
  })

  it('an unstyled entry indents its body; the number cell is as wide as its tab stop', () => {
    const attrs = (c: unknown[]) => c[1] as Record<string, string>
    expect(attrs(body({ ...base, indentLeftTwips: 240 })).style).toBe('--doc-toc-indent:12pt')
    expect(attrs(tail(base)).style).toBeUndefined()
    const numbered = body({
      ...base,
      num: '1.',
      indentLeftTwips: 480,
      numTabTwips: 400,
      hangTwips: -400,
    })
    const [num] = numbered.filter(Array.isArray) as Span[]
    expect(num[1].class).toBe('doc-toc-num')
    expect(num[1].style).toBe('min-width:20pt')
    expect(attrs(numbered).style).toBe('--doc-toc-indent:24pt;--doc-toc-hang:-20pt')
    expect(attrs(body({ ...base, rightTabTwips: 8647 })).style).toBe('--doc-toc-stop:432.35pt')
    const bare = body({ ...base, num: '1.' }).filter(Array.isArray) as Span[]
    expect(bare[0][1].style).toBeUndefined()
    expect(attrs(body(base)).style).toBeUndefined()
  })

  it('number, title, leader and page share one inline flow; leader and page are one unbreakable unit', () => {
    const [num, title, unit] = body({ ...base, num: '2' }).filter(Array.isArray) as unknown[][]
    expect((num[1] as Record<string, string>).class).toBe('doc-toc-num')
    expect((title[1] as Record<string, string>).class).toBe('doc-toc-title')
    expect((unit[1] as Record<string, string>).class).toBe('doc-toc-tail')
    const [, page] = unit.filter(Array.isArray) as Span[]
    expect(page[1].class).toBe('doc-toc-page')
    expect(page[2]).toBe('3')
  })
})

describe('TOC entry line strut', () => {
  const lineAttrs = (field: FieldDisplay) =>
    (renderFieldSpec(field) as unknown[])[1] as Record<string, string>
  const base: FieldDisplay = { kind: 'tocLine', left: 'Title', right: '3', level: 1 }

  it('an own face sets its metric factor; a style-only face follows the per-style script var', () => {
    expect(lineAttrs({ ...base, fontFamily: 'Arial' }).style).toMatch(/--doc-line-factor:1\.1\d+;/)
    expect(lineAttrs(base).style).toMatch(/--doc-line-factor:var\(--doc-line-factor-latin,1\.2\);/)
    expect(lineAttrs({ ...base, left: '\u7b2c\u4e00\u7ae0' }).style).toMatch(
      /--doc-line-factor:var\(--doc-line-factor-cjk,1\.7\);/,
    )
    for (const f of [base, { ...base, fontFamily: 'Arial' }])
      expect(lineAttrs(f).style).toContain(
        'line-height:var(--doc-line-max, calc(var(--doc-line-factor,1.2) * var(--doc-line-mult,1)))',
      )
  })
})

describe('TOC entry line layout', () => {
  const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
  const block = (sel: string) =>
    new RegExp(sel.replace(/\./g, '\\.') + '\\s*\\{([^}]*)\\}').exec(css)?.[1] ?? ''

  it('the body carries the first-line indent; the leader is sized by the measured gap', () => {
    expect(block('.doc-toc-body')).toMatch(/text-indent:\s*var\(--doc-toc-hang/)
    expect(block('.doc-toc-tail')).toMatch(/white-space:\s*nowrap/)
    expect(block('.doc-toc-dots')).toMatch(/width:\s*var\(--doc-toc-gap/)
  })
})

describe('tocLeaderGap', () => {
  it('fills the free width of the last title line when the page fits after it', () => {
    expect(tocLeaderGap(300, 0, 500, 20, 1)).toBe(180)
  })
  it('drops a bare page number when the title stops short of the stop', () => {
    expect(tocLeaderGap(490, 0, 500, 20, 1)).toBe(0)
    expect(tocLeaderGap(480, 0, 500, 20, 1)).toBe(0)
  })
  it('drops leader + page filling the next line when the title runs past the stop', () => {
    expect(tocLeaderGap(500, 40, 500, 20, 1)).toBe(440)
    expect(tocLeaderGap(520, 40, 500, 20, 1)).toBe(440)
  })
})
