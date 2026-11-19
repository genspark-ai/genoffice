import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { paraDeletionClass } from '../src/renderer/editor/decoration-extensions'
import { measureBlocks } from '../src/renderer/pagination'
import { alignJoinedParagraphs } from '../src/renderer/pagination-measure'

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

const rule = (selector: string) => {
  const re = new RegExp(
    `(^|[,\\n])\\s*${selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*[,{][^}]*}`,
  )
  return re.exec(css)?.[0] ?? ''
}

describe('markup-view CSS', () => {
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

describe('pending paragraph-mark deletion joins paragraphs in final-text views', () => {
  const two = (paraMarkDel: string | null) => [
    {
      type: 'docParagraph',
      attrs: { docxIndex: null, paraMarkDel },
      content: [{ type: 'text', text: 'head' }],
    },
    { type: 'docParagraph', attrs: { docxIndex: null }, content: [{ type: 'text', text: 'tail' }] },
    { type: 'docParagraph', attrs: { docxIndex: null }, content: [{ type: 'text', text: 'rest' }] },
  ]

  it('decorates head and tail in No Markup / Simple Markup only', async () => {
    const { Editor } = await import('@tiptap/core')
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const { revisionDisplayState } = await import('../src/renderer/editor/marks')
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: two('{"author":"Ann"}') },
    })
    const classes = () =>
      Array.from(editor.view.dom.children).map((el) =>
        Array.from(el.classList)
          .filter((c) => c.startsWith('doc-para-'))
          .sort()
          .join(' '),
      )
    expect(classes()).toEqual(['doc-para-mark-del', '', ''])
    for (const mode of ['none', 'simple'] as const) {
      revisionDisplayState.mode = mode
      editor.view.dispatch(editor.state.tr)
      expect(classes()).toEqual(['doc-para-join-head doc-para-mark-del', 'doc-para-join-tail', ''])
    }
    revisionDisplayState.mode = 'original'
    editor.view.dispatch(editor.state.tr)
    expect(classes()).toEqual(['doc-para-mark-del', '', ''])
    revisionDisplayState.mode = 'all'
    editor.destroy()
  })

  it('chains: a tail whose own mark is deleted is also a head', async () => {
    const { paraJoinClasses } = await import('../src/renderer/editor/decoration-extensions')
    const { Editor } = await import('@tiptap/core')
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const content = two('{"author":"Ann"}')
    content[1].attrs = { ...content[1].attrs, paraMarkDel: '{"author":"Ann"}' }
    const editor = new Editor({ extensions: editorExtensions, content: { type: 'doc', content } })
    expect([...paraJoinClasses(editor.state.doc, 'none')]).toEqual([
      [0, 'doc-para-join-head'],
      [1, 'doc-para-join-tail doc-para-join-head'],
      [2, 'doc-para-join-tail'],
    ])
    expect(paraJoinClasses(editor.state.doc, 'all').size).toBe(0)
    editor.destroy()
  })

  it('skips a fully deleted paragraph between head and tail', async () => {
    const { paraJoinClasses } = await import('../src/renderer/editor/decoration-extensions')
    const { Editor } = await import('@tiptap/core')
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const content = two('{"author":"Ann"}')
    content[1] = {
      type: 'docParagraph',
      attrs: { docxIndex: null, paraMarkDel: '{"author":"Ann"}' },
      content: [{ type: 'text', text: 'gone', marks: [{ type: 'del', attrs: { author: 'Ann' } }] }],
    } as unknown as (typeof content)[number]
    const editor = new Editor({ extensions: editorExtensions, content: { type: 'doc', content } })
    expect([...paraJoinClasses(editor.state.doc, 'none')]).toEqual([
      [0, 'doc-para-join-head'],
      [2, 'doc-para-join-tail'],
    ])
    editor.destroy()
  })

  it('a tail carrying a page break stays its own block', () => {
    const pm = document.createElement('div')
    const add = (top: number, height: number, cls: string, inner: string) => {
      const el = document.createElement('p')
      el.className = cls
      el.innerHTML = inner
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    add(0, 20, 'doc-para-mark-del doc-para-join-head', 'head')
    add(0, 40, 'doc-para-join-tail', 'tail<br class="doc-page-br">')
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks).toHaveLength(2)
    expect(blocks[1].breakAfter).toBe(true)
  })

  it('a tracked-deleted break in the tail does not prevent folding', () => {
    const pm = document.createElement('div')
    const add = (top: number, height: number, cls: string, inner: string) => {
      const el = document.createElement('p')
      el.className = cls
      el.innerHTML = inner
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    add(0, 20, 'doc-para-mark-del doc-para-join-head', 'head')
    add(0, 40, 'doc-para-join-tail', 'tail<span class="doc-del"><br class="doc-page-br"></span>')
    expect(measureBlocks(pm, 0, 1).blocks.map((b) => b.height)).toEqual([40])
  })

  it('folds a chain of joined paragraphs into the first block', () => {
    const pm = document.createElement('div')
    const add = (top: number, height: number, cls: string) => {
      const el = document.createElement('p')
      el.className = cls
      el.textContent = 'x'
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    add(0, 20, 'doc-para-mark-del doc-para-join-head')
    add(0, 40, 'doc-para-mark-del doc-para-join-tail doc-para-join-head')
    add(20, 40, 'doc-para-join-tail')
    add(60, 20, 'after')
    expect(measureBlocks(pm, 0, 1).blocks.map((b) => [b.top, b.height])).toEqual([
      [0, 60],
      [60, 20],
    ])
  })

  it('folds the tail across a hidden fully deleted paragraph', () => {
    const pm = document.createElement('div')
    const add = (top: number, height: number, cls: string, hidden = false) => {
      const el = document.createElement('p')
      el.className = cls
      el.textContent = 'x'
      if (hidden) el.style.display = 'none'
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    add(0, 20, 'doc-para-mark-del doc-para-join-head')
    add(20, 0, 'doc-para-del-collapse', true)
    add(0, 40, 'doc-para-join-tail')
    const heights = measureBlocks(pm, 0, 1).blocks.map((b) => [b.top, b.height])
    expect(heights[0]).toEqual([0, 40])
    expect(heights.some(([, h]) => h === 40 && heights.length <= 2)).toBe(true)
  })

  it('the tail keeps its block and is lifted onto the head line by the measured vars', () => {
    expect(rule('.rev-display-none .doc-para-join-tail[data-join-on]')).toMatch(
      /margin-top:\s*calc\(-1 \* var\(--join-lift[^)]*\)\)\s*!important/,
    )
    expect(rule('.rev-display-simple .doc-para-join-tail[data-join-on]')).toMatch(
      /text-indent:\s*var\(--join-indent[^)]*\)\s*!important/,
    )
    expect(css).not.toMatch(/\.doc-para-join-(head|tail)\s*\{[^}]*display:\s*inline/)
  })

  it('alignJoinedParagraphs lifts the tail by one line and indents it to the head end', () => {
    const pm = document.createElement('div')
    const para = (
      cls: string,
      text: string,
      box: { x: number; w: number; top: number; h: number },
    ) => {
      const el = document.createElement('p')
      el.className = cls
      el.textContent = text
      el.style.lineHeight = '20px'
      el.getBoundingClientRect = () =>
        ({ ...rectOf(box.top, box.h), left: box.x, right: box.x + box.w, width: box.w }) as DOMRect
      pm.appendChild(el)
      return el
    }
    const head = para('doc-para-mark-del doc-para-join-head', 'head', {
      x: 30,
      w: 400,
      top: 0,
      h: 40,
    })
    const gone = para('doc-para-del-collapse', 'gone', { x: 30, w: 400, top: 40, h: 0 })
    gone.style.display = 'none'
    const tail = para('doc-para-join-tail', 'tail', { x: 10, w: 420, top: 40, h: 20 })
    const line = (top: number, left: number, right: number) =>
      ({ top, bottom: top + 20, left, right, width: right - left, height: 20 }) as DOMRect
    const rectsOf = new Map<Node, DOMRect[]>([
      [head.firstChild!, [line(0, 30, 430), line(20, 30, 150)]],
      [tail.firstChild!, [line(20, 150, 430)]],
    ])
    const proto = Range.prototype as unknown as {
      getClientRects?: () => DOMRect[]
      selectNodeContents: (n: Node) => void
      __node?: Node
    }
    const orig = { rects: proto.getClientRects, select: proto.selectNodeContents }
    proto.selectNodeContents = function (this: typeof proto, n: Node) {
      this.__node = n
    }
    proto.getClientRects = function (this: typeof proto) {
      return rectsOf.get(this.__node!) ?? []
    }
    try {
      alignJoinedParagraphs(pm, 1)
      expect(tail.dataset.joinOn).toBe('1')
      expect(tail.style.getPropertyValue('--join-lift')).toBe('20.00px')
      expect(tail.style.getPropertyValue('--join-indent')).toBe('140.00px')
      expect(head.dataset.joinOn).toBeUndefined()

      // the tail's first word wrapped below the shared line: start a new line unindented
      rectsOf.set(tail.firstChild!, [line(40, 10, 100)])
      alignJoinedParagraphs(pm, 1)
      expect(tail.style.getPropertyValue('--join-lift')).toBe('0.00px')
      expect(tail.style.getPropertyValue('--join-indent')).toBe('0.00px')

      // leaving the final-text views drops the classes and the override
      head.classList.remove('doc-para-join-head')
      alignJoinedParagraphs(pm, 1)
      expect(tail.dataset.joinOn).toBeUndefined()
    } finally {
      if (orig.rects) proto.getClientRects = orig.rects
      else delete proto.getClientRects
      proto.selectNodeContents = orig.select
    }
  })

  it('measureBlocks folds the tail into the head block instead of counting a break', () => {
    const pm = document.createElement('div')
    const add = (top: number, height: number, cls: string) => {
      const el = document.createElement('p')
      el.className = cls
      el.textContent = cls || 'x'
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    add(0, 20, 'before')
    add(20, 20, 'doc-para-mark-del doc-para-join-head')
    add(20, 40, 'doc-para-join-tail')
    add(60, 20, 'after')
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks.map((b) => [b.top, b.height])).toEqual([
      [0, 20],
      [20, 40],
      [60, 20],
    ])
  })
})
