import { afterEach, describe, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  RESULTS_LIST_CAP,
  compileFind,
  foldPatternWidth,
  expandReplacement,
  findInDoc,
  findInDocAsync,
  foldWidth,
  matchContext,
  wildcardToRegExp,
  type FindOptions,
} from '../src/renderer/editor/find'
import {
  lineAtOffset,
  lineIndexAtPos,
  listGoToItems,
  resolveGoToIndex,
  visiblePageStartBlocks,
  type PageLayout,
} from '../src/renderer/editor/goto'
import type { EditorView } from '@tiptap/pm/view'
import { findMatches } from '../src/renderer/components/FindPanel'

// non-ASCII fixtures are built from code points (public-hygiene: no CJK literals in tests)
const cp = (...codes: number[]) => String.fromCodePoint(...codes)
const FULL_ABC = cp(0xff21, 0xff22, 0xff23)
const CAT = cp(0x732b)

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string, attrs: Record<string, unknown> = {}) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: text ? [{ type: 'text', text }] : [],
})

function createEditor(content: JSONContent[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

const opts = (o: Partial<FindOptions> = {}): FindOptions => ({
  matchCase: false,
  wholeWord: false,
  mode: 'literal',
  matchWidth: false,
  ...o,
})

function find(editor: Editor, query: string, o: Partial<FindOptions> = {}) {
  const compiled = compileFind(query, opts(o))
  if (!compiled.ok) throw new Error(compiled.error)
  return findInDoc(editor.state.doc, compiled.query, opts(o))
}

const texts = (editor: Editor, query: string, o: Partial<FindOptions> = {}) =>
  find(editor, query, o).map((m) => editor.state.doc.textBetween(m.from, m.to))

describe('wildcardToRegExp', () => {
  const src = (p: string) => {
    const t = wildcardToRegExp(p)
    if ('error' in t) throw new Error(t.error)
    return t
  }
  it('translates each Word token', () => {
    const any = '[^\\u0000\\n]'
    expect(src('a?c').source).toBe(`a${any}c`)
    expect(src('a*c').source).toBe(`a${any}*?c`)
    expect(src('[a-z]').source).toBe('[a-z]')
    expect(src('[!0-9]').source).toBe('[^0-9\\u0000\\n]')
    expect(src('<the>').source).toContain('(?<![\\p{L}\\p{N}_])(?=[\\p{L}\\p{N}_])the(?<=')
    expect(src('o@').source).toBe('o+')
    expect(src('o{2,3}').source).toBe('o{2,3}')
    expect(src('(a)(b)').source).toBe('(a)(b)')
    expect(src('^t^#^$').source).toBe('\\t\\d\\p{L}')
    expect(src('^32').source).toBe('\\u{20}')
  })
  it('escapes regex metacharacters and honors backslash escapes', () => {
    expect(src('a.b+c').source).toBe('a\\.b\\+c')
    expect(src('\\?\\*').source).toBe('\\?\\*')
  })
  it('turns ^p into paragraph anchors only at the ends', () => {
    expect(src('^pabc')).toMatchObject({ anchorStart: true, anchorEnd: false })
    expect(src('abc^p')).toMatchObject({ anchorStart: false, anchorEnd: true })
    expect(wildcardToRegExp('a^pb')).toEqual({ error: 'paragraph-mark-position' })
  })
  it('reports malformed patterns instead of throwing', () => {
    expect(wildcardToRegExp('(ab')).toEqual({ error: 'unbalanced-paren' })
    expect(wildcardToRegExp('ab)')).toEqual({ error: 'unbalanced-paren' })
    expect(wildcardToRegExp('[ab')).toEqual({ error: 'unbalanced-bracket' })
    expect(wildcardToRegExp('@')).toEqual({ error: 'nothing-to-repeat' })
    expect(wildcardToRegExp('a\\')).toEqual({ error: 'trailing-backslash' })
    expect(wildcardToRegExp('^z')).toEqual({ error: 'bad-caret-code' })
    expect(compileFind('(', opts({ mode: 'regex' }))).toEqual({
      ok: false,
      error: 'invalid-pattern',
    })
  })
})

describe('foldWidth', () => {
  it('folds full-width ASCII and half-width katakana without changing length', () => {
    const s = FULL_ABC + cp(0x3000, 0xff76, 0xff9e)
    const folded = foldWidth(s)
    expect(folded).toBe('ABC ' + cp(0x30ab, 0x309b))
    expect(folded.length).toBe(s.length)
  })
})

describe('findInDoc', () => {
  it('matches across marks and folds case, keeping offsets aligned', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          { type: 'text', text: cp(0x130, 0x130) + ' Hel' },
          { type: 'text', text: 'lo world', marks: [{ type: 'bold' }] },
        ],
      },
    ])
    expect(texts(editor, 'HELLO')).toEqual(['Hello'])
  })

  it('whole word respects CJK-adjacent and underscore boundaries', () => {
    const editor = createEditor([para(`cat cats _cat ${CAT} cat${CAT} (cat)`)])
    expect(texts(editor, 'cat', { wholeWord: true })).toHaveLength(2)
  })

  it('treats full-width text as half-width unless Match width is on', () => {
    const editor = createEditor([para(`${FULL_ABC} and ABC`)])
    expect(texts(editor, 'abc')).toEqual([FULL_ABC, 'ABC'])
    expect(texts(editor, 'abc', { matchWidth: true })).toEqual(['ABC'])
  })

  it('wildcards find the shortest run like Word and keep groups', () => {
    const editor = createEditor([para('short test, sort tart')])
    expect(texts(editor, 's*t', { mode: 'wildcard' })).toEqual(['short', 'st', 'sort'])
    const m = find(editor, '(s)(?)rt', { mode: 'wildcard' })
    expect(m.map((x) => x.groups)).toEqual([['s', 'o']])
  })

  it('regex mode uses JS syntax with Unicode classes', () => {
    const editor = createEditor([para('a1 b22 c333')])
    expect(texts(editor, '\\p{L}\\d{2,}', { mode: 'regex' })).toEqual(['b22', 'c333'])
  })

  it('literal ^t and ^p codes', () => {
    const editor = createEditor([para('x\ty'), para('end here'), para('here end')])
    expect(texts(editor, 'x^ty')).toEqual(['x\ty'])
    expect(texts(editor, 'end^p')).toEqual(['end'])
    expect(find(editor, 'end^p').map((m) => m.from)).toEqual([find(editor, 'here end')[0].to - 3])
    expect(texts(editor, '^phere')).toEqual(['here'])
  })

  it('skips protected blocks and never loops on zero-width matches', () => {
    const editor = createEditor([para('aaa')])
    expect(texts(editor, 'x*', { mode: 'regex' })).toEqual([])
  })

  it('^l finds hard breaks; wildcards and dots never cross a break or a leaf', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          { type: 'text', text: 'ab' },
          { type: 'hardBreak' },
          { type: 'text', text: 'cd' },
          { type: 'docInlineImage' },
          { type: 'text', text: 'ef' },
        ],
      },
    ])
    expect(find(editor, 'b^lc').map((m) => [m.from, m.to])).toEqual([[2, 5]])
    expect(texts(editor, 'b?c', { mode: 'wildcard' })).toEqual([])
    expect(texts(editor, 'a*d', { mode: 'wildcard' })).toEqual([])
    expect(texts(editor, 'd.e', { mode: 'regex' })).toEqual([])
    expect(texts(editor, 'd[^x]e', { mode: 'regex' })).toEqual([])
    expect(texts(editor, 'c[!x]', { mode: 'wildcard' })).toEqual(['cd'])
  })

  it('a rejected leaf-spanning hit does not hide a later hit inside its span', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          { type: 'text', text: 'x' },
          { type: 'docInlineImage' },
          { type: 'text', text: 'xy xy' },
        ],
      },
    ])
    // the first exec swallows "x<img>xy"; the rescan must still surface both "xy"
    expect(texts(editor, 'x.*?y', { mode: 'regex' })).toEqual(['xy', 'xy'])
    expect(texts(editor, 'x.', { mode: 'regex', wholeWord: true })).toEqual(['xy', 'xy'])
  })

  it('folds the pattern too and reads groups from the original text', () => {
    const editor = createEditor([para(`${FULL_ABC}1 and ABC2`)])
    expect(texts(editor, `${FULL_ABC}?`, { mode: 'wildcard' })).toEqual([`${FULL_ABC}1`, 'ABC2'])
    expect(texts(editor, `(${FULL_ABC})(\\d)`, { mode: 'regex' })).toEqual([`${FULL_ABC}1`, 'ABC2'])
    const [m] = find(editor, '(abc)(\\d)', { mode: 'regex' })
    expect(m.groups).toEqual([FULL_ABC, '1'])
    expect(texts(editor, `${FULL_ABC}?`, { mode: 'wildcard', matchWidth: true })).toEqual([
      `${FULL_ABC}1`,
    ])
    // a full-width char whose twin is a regex metacharacter is kept as typed
    expect(foldPatternWidth(`${FULL_ABC}${cp(0xff0e)}`, foldWidth)).toBe(`ABC${cp(0xff0e)}`)
  })

  it('the async scan returns the same matches and honors abort', async () => {
    const editor = createEditor(Array.from({ length: 40 }, (_, i) => para(`row ${i} cat`)))
    const compiled = compileFind('cat', opts())
    if (!compiled.ok) throw new Error(compiled.error)
    const sync = findInDoc(editor.state.doc, compiled.query, opts())
    const async = await findInDocAsync(editor.state.doc, compiled.query, opts(), { aborted: false })
    expect(async?.map((m) => m.from)).toEqual(sync.map((m) => m.from))
    expect(await findInDocAsync(editor.state.doc, compiled.query, opts(), { aborted: true })).toBe(
      null,
    )
  })

  it('findMatches keeps the legacy range-only API', () => {
    const editor = createEditor([para('Cat cats CAT')])
    expect(findMatches(editor, 'cat', { matchCase: false, wholeWord: true })).toHaveLength(2)
  })
})

describe('expandReplacement', () => {
  const m = (text: string, groups: string[] | null) => ({
    from: 1,
    to: 1 + text.length,
    text,
    groups,
    blockText: text,
    offset: 0,
  })
  it('expands groups in regex and wildcard syntax', () => {
    expect(expandReplacement('$2-$1 $& $$', m('ab', ['a', 'b']), 'regex')).toEqual(['b-a ab $'])
    expect(expandReplacement('\\2\\1 ^&', m('ab', ['a', 'b']), 'wildcard')).toEqual(['ba ab'])
  })
  it('splits on ^p and translates ^t', () => {
    expect(expandReplacement('x^py^tz', m('q', null), 'literal')).toEqual(['x', 'y\tz'])
  })
  it('leaves unknown sequences alone', () => {
    expect(expandReplacement('$x ^q \\', m('q', null), 'regex')).toEqual(['$x ^q \\'])
  })
})

describe('replace through the editor', () => {
  it('regex replace-all with groups and a paragraph split', () => {
    const editor = createEditor([para('John Smith, Jane Doe')])
    const matches = find(editor, '(\\w+) (\\w+)', { mode: 'regex' })
    editor.commands.command(({ tr }) => {
      for (const mm of [...matches].reverse()) {
        const segs = expandReplacement('$2^p$1', mm, 'regex')
        const joined = segs.join('\u0000')
        tr.insertText(joined, mm.from, mm.to)
        const p = mm.from + segs[0].length
        tr.delete(p, p + 1)
        tr.split(p)
      }
      return true
    })
    expect(editor.state.doc.childCount).toBe(3)
    expect(editor.state.doc.textContent).toBe('SmithJohn, DoeJane')
    expect(editor.state.doc.child(0).textContent).toBe('Smith')
  })
})

describe('results list', () => {
  it('provides context around a hit and caps the visible list', () => {
    const editor = createEditor([para(`${'a'.repeat(50)} needle ${'b'.repeat(50)}`)])
    const [m] = find(editor, 'needle')
    const ctx = matchContext(m)
    expect(ctx.hit).toBe('needle')
    expect(ctx.before.startsWith('…')).toBe(true)
    expect(ctx.before.length).toBe(31)
    expect(ctx.after.endsWith('…')).toBe(true)
    const many = createEditor(Array.from({ length: 12 }, () => para('x '.repeat(60))))
    const hits = find(many, 'x')
    expect(hits.length).toBeGreaterThan(RESULTS_LIST_CAP)
    expect(hits.slice(0, RESULTS_LIST_CAP)).toHaveLength(RESULTS_LIST_CAP)
  })
})

describe('go to', () => {
  it('lists headings, tables and bookmarks in document order', () => {
    const editor = createEditor([
      {
        type: 'docHeading',
        attrs: { level: 1, docxIndex: null },
        content: [{ type: 'text', text: 'One' }],
      },
      para('body', { bookmarks: ['bm_a', 'bm_b'] }),
      {
        type: 'docHeading',
        attrs: { level: 2, docxIndex: null },
        content: [{ type: 'text', text: 'Two' }],
      },
    ])
    const doc = editor.state.doc
    expect(listGoToItems(doc, 'heading').map((i) => i.label)).toEqual(['One', 'Two'])
    expect(listGoToItems(doc, 'bookmark').map((i) => i.label)).toEqual(['bm_a', 'bm_b'])
    expect(listGoToItems(doc, 'table')).toEqual([])
  })

  it('resolves absolute, relative and empty inputs like Word', () => {
    expect(resolveGoToIndex('3', 0, 10)).toBe(2)
    expect(resolveGoToIndex('+2', 4, 10)).toBe(6)
    expect(resolveGoToIndex('-9', 4, 10)).toBe(0)
    expect(resolveGoToIndex('', 4, 10)).toBe(5)
    expect(resolveGoToIndex('', 4, 10, -1)).toBe(3)
    expect(resolveGoToIndex('99', 0, 10)).toBe(9)
    expect(resolveGoToIndex('abc', 0, 10)).toBe(null)
    expect(resolveGoToIndex('1', 0, 0)).toBe(null)
  })

  it('maps visible pages to their first blocks, skipping parity blanks', () => {
    const blocks = [
      { top: 0, height: 100 },
      { top: 100, height: 100 },
      { top: 200, height: 100 },
      { top: 300, height: 100 },
    ]
    const slices = [
      { start: 0, end: 200, section: 0 },
      { start: 200, end: 200, section: 0 },
      { start: 200, end: 400, section: 0 },
    ]
    expect(visiblePageStartBlocks({ blocks, slices, sections: [] })).toEqual([0, 2])
  })

  it('resolves the caret line inside its block, not the block start', () => {
    const boxes = [
      { offsetInBlock: 0, height: 20 },
      { offsetInBlock: 20, height: 20 },
      { offsetInBlock: 40, height: 20 },
    ]
    expect(lineAtOffset(boxes, 5)).toBe(0)
    expect(lineAtOffset(boxes, 35)).toBe(1)
    expect(lineAtOffset(boxes, 999)).toBe(2)
    expect(lineAtOffset(undefined, 35)).toBe(0)

    const first = document.createElement('p')
    const second = document.createElement('p')
    const root = document.createElement('div')
    root.append(first, second)
    // half zoom: 60 layout px render as 30 css px
    second.getBoundingClientRect = () => ({ top: 200, height: 30 }) as DOMRect
    const layout: PageLayout = {
      blocks: [
        { top: 0, height: 60, el: first, lineBoxes: boxes },
        { top: 60, height: 60, el: second, lineBoxes: boxes },
      ],
      slices: [],
      sections: [],
    }
    const view = {
      dom: root,
      domAtPos: () => ({ node: second, offset: 0 }),
      coordsAtPos: () => ({ top: 215, bottom: 220, left: 0, right: 0 }),
    } as unknown as EditorView
    // caret at css y 217.5 → layout offset 35 → second line of the second block → 3 + 1
    expect(lineIndexAtPos(view, layout, 1)).toBe(4)
  })
})
