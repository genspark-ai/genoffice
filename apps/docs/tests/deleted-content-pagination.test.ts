import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { paraDeletionClass } from '../src/renderer/editor/decoration-extensions'
import { measureBlocks } from '../src/renderer/pagination'

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

const rectOf = (top: number, height: number) =>
  ({ top, height, bottom: top + height, left: 0, right: 100, width: 100, x: 0, y: top }) as DOMRect

function flow(html: string, rootClass = '') {
  const root = document.createElement('div')
  root.className = rootClass
  const pm = document.createElement('div')
  root.appendChild(pm)
  const add = (top: number, height: number, inner: string) => {
    const el = document.createElement('p')
    el.innerHTML = inner
    el.getBoundingClientRect = () => rectOf(top, height)
    pm.appendChild(el)
  }
  add(0, 20, 'one')
  add(20, 20, html)
  add(40, 20, 'after')
  return pm
}

describe('tracked-deleted page breaks', () => {
  it('a deleted trailing page break does not push the next block', () => {
    const { blocks } = measureBlocks(
      flow('text<span class="doc-del"><br class="doc-page-br"></span>'),
      0,
      1,
    )
    expect(blocks[1].breakAfter).toBeUndefined()
    expect(blocks[1].innerBreaks ?? []).toHaveLength(0)
  })

  it('a live trailing page break still pushes the next block', () => {
    const { blocks } = measureBlocks(flow('text<br class="doc-page-br">'), 0, 1)
    expect(blocks[1].breakAfter).toBe(true)
  })

  it('the Original view restores the deleted break', () => {
    const { blocks } = measureBlocks(
      flow('text<span class="doc-del"><br class="doc-page-br"></span>', 'rev-display-original'),
      0,
      1,
    )
    expect(blocks[1].breakAfter).toBe(true)
  })
})

const del = { marks: [{ type: { name: 'del' } }] }
const live = { marks: [] as { type: { name: string } }[] }
const node = (children: (typeof del)[], attrs: Record<string, unknown> = {}) => ({
  attrs,
  childCount: children.length,
  forEach: (f: (c: typeof del) => void) => children.forEach(f),
})

describe('paraDeletionClass', () => {
  it('deleted mark + all content deleted → collapse', () => {
    expect(paraDeletionClass(node([del, del], { paraMarkDel: '{}' }))).toBe(
      'doc-para-mark-del doc-para-del-collapse',
    )
    expect(paraDeletionClass(node([], { paraMarkDel: '{}' }))).toBe(
      'doc-para-mark-del doc-para-del-collapse',
    )
  })

  it('deleted mark with live text keeps its line', () => {
    expect(paraDeletionClass(node([del, live], { paraMarkDel: '{}' }))).toBe('doc-para-mark-del')
  })

  it('live mark with every child deleted keeps one empty line', () => {
    expect(paraDeletionClass(node([del]))).toBe('doc-para-content-del')
    expect(paraDeletionClass(node([del, live]))).toBeNull()
    expect(paraDeletionClass(node([]))).toBeNull()
  })
})

describe('markup-view CSS', () => {
  const rule = (selector: string) => {
    const re = new RegExp(
      `(^|[,\\n])\\s*${selector.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')}\\s*[,{][^}]*}`,
    )
    return re.exec(css)?.[0] ?? ''
  }
  it('collapsed paragraphs are hidden in No Markup and Simple Markup too', () => {
    expect(rule('.rev-display-none .doc-para-del-collapse')).toMatch(/display:\s*none/)
    expect(rule('.rev-display-simple .doc-para-del-collapse')).toMatch(/display:\s*none/)
  })
  it('the strut outranks the same-specificity has-ppr-change::after content:none', () => {
    const strut = css.indexOf('.rev-balloon .doc-para-content-del::after')
    for (const mode of ['.rev-display-none', '.rev-display-simple', '.rev-balloon']) {
      expect(strut).toBeGreaterThan(css.indexOf(`${mode} .has-ppr-change::after`))
    }
  })

  it('the strut resets the pPrChange badge chrome it shares ::after with', () => {
    const r = rule('.rev-balloon .doc-para-content-del::after')
    expect(r).toMatch(/background:\s*none/)
    expect(r).toMatch(/border:\s*none/)
    expect(r).toMatch(/padding:\s*0/)
    expect(r).toMatch(/font-size:\s*inherit/)
    expect(r).toMatch(/margin-left:\s*0/)
  })

  it('content-deleted paragraphs get a strut where deletions are hidden', () => {
    for (const mode of ['.rev-display-none', '.rev-display-simple', '.rev-balloon']) {
      expect(rule(`${mode} .doc-para-content-del::after`)).toMatch(/content:\s*'\\200b'/)
      expect(rule(`${mode} .doc-para-content-del > br.ProseMirror-trailingBreak`)).toMatch(
        /display:\s*none/,
      )
    }
  })
})
