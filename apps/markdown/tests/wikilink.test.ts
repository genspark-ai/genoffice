/**
 * `[[wiki links]]` in two halves. `wikilink.ts` decides what a link IS (syntax);
 * `wiki-resolve.ts` decides what it points AT (resolution). Both are covered
 * here because the interesting failures are the seams: a link that renders but
 * does not resolve, and a resolver that answers but is never consulted.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { collectWikiTargets, wikilinkStart } from '../src/renderer/editor/wikilink'
import { parseDocText } from '../src/renderer/markdown/docText'
import { buildSourceMap, spliceMarkdown } from '../src/renderer/markdown/sourceSplice'
import type { WikiLinkResolved } from '../src/shared/ipc'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (v: string) => v }

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((e) => e.destroy())
  vi.useRealTimers()
})

const ANSWERS: Record<string, WikiLinkResolved> = {
  '[[Alpha]]': { path: '/notes/Alpha.md', fragment: '', found: true },
  '[[Beta|Bee]]': { path: '/notes/Beta.md', fragment: '', found: true },
  '[[Gamma#Section]]': { path: '/notes/Gamma.md', fragment: 'Section', found: true },
  '[[Nowhere]]': { path: '', fragment: '', found: false },
}

function open(
  source: string,
  wiki: {
    onOpen?: (path: string, fragment: string) => void
    onUnresolved?: (target: string) => void
    answers?: Record<string, WikiLinkResolved>
    documentPath?: string
  } = {},
) {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
      wikiLink: {
        getDocumentPath: () => wiki.documentPath ?? '/notes/Index.md',
        onOpen: wiki.onOpen ?? (() => {}),
        onUnresolved: wiki.onUnresolved ?? (() => {}),
        resolve: async (_from, targets) => {
          const out: Record<string, WikiLinkResolved> = {}
          for (const t of targets) {
            out[t] = wiki.answers?.[t] ?? { path: '', fragment: '', found: false }
          }
          return out
        },
      },
    }),
    content: '',
  })
  editors.push(editor)
  const envelope = parseDocText(source)
  editor
    .chain()
    .setMeta('addToHistory', false)
    .setContent(envelope.body, { contentType: 'markdown' })
    .run()
  return { editor, envelope }
}

/** let the resolver's debounce fire and its promise settle */
const settled = () => new Promise((r) => setTimeout(r, 200))

/**
 * ProseMirror reaches `handleClick` through `posAtCoords`, which jsdom cannot
 * answer (no layout, every rect is zero). The handler is what carries the
 * behaviour, so it is invoked the way the editor would: on the element the
 * user actually pressed.
 */
function clickWikiLink(editor: Editor, raw: string): boolean {
  const el = editor.view.dom.querySelector(
    `[data-wiki-target="${CSS.escape(raw)}"]`,
  ) as HTMLElement | null
  if (!el) throw new Error(`no link rendered for ${raw}`)
  let handled = false
  editor.view.someProp('handleClick', (f) => {
    handled = f(editor.view, 0, { target: el } as unknown as MouseEvent) === true
    return handled
  })
  return handled
}

describe('syntax', () => {
  it('wikilinkStart points at the first unescaped [[', () => {
    expect(wikilinkStart('a [[Note]] b')).toBe(2)
    expect(wikilinkStart('no brackets here')).toBe(-1)
    expect(wikilinkStart('escaped \\[[x] tail')).toBe(-1)
  })

  it('collects each distinct target once, in document order', () => {
    const { editor } = open('[[A]] then [[A]] then [[B]]')
    expect(collectWikiTargets(editor.state.doc)).toEqual(['[[A]]', '[[B]]'])
  })

  it.each([
    ['a bare link', '[[Note]]'],
    ['an alias', 'See [[Note|Display text]] here.'],
    ['a heading ref', 'Jump to [[Note#Some Heading]].'],
    ['a block ref', 'Block [[Note#^a1b2c3]].'],
    ['a same-file ref', 'Same file [[#Section]].'],
    ['an embed', 'Embed ![[Note]] inline.'],
    ['two in one line', 'before [[A]] middle [[B]] after'],
  ])('parses %s and writes it back byte for byte', (_label, src) => {
    const { editor, envelope } = open(src)
    const map = buildSourceMap(editor, editor.state.doc, envelope.body)
    const out = map ? spliceMarkdown(editor, editor.state.doc, map) : editor.getMarkdown()
    expect(out).toBe(src)
  })

  it('leaves a standard markdown link to the link mark', () => {
    const { editor } = open('A [label](https://example.com) link.')
    const names: string[] = []
    editor.state.doc.descendants((n) => {
      if (!n.isText) return
      for (const m of n.marks) names.push(m.type.name)
    })
    expect(names).toContain('link')
    expect(names).not.toContain('wikiLink')
  })

  it.each([
    ['code span', 'Use `[[Foo]]` here.'],
    ['code block', '```\n[[NotALink]]\n```'],
    ['escaped brackets', 'Literal \\[\\[not a link\\]\\] here.'],
    ['an unclosed link', 'Broken [[Foo and more text.'],
  ])('does not link inside %s', (_label, src) => {
    const { editor } = open(src)
    expect(collectWikiTargets(editor.state.doc)).toEqual([])
  })
})

describe('resolution', () => {
  it('marks a resolved link and leaves an unresolved one dimmed', async () => {
    const { editor } = open('See [[Alpha]] and [[Nowhere]].', { answers: ANSWERS })
    await settled()
    const cls = [...editor.view.dom.querySelectorAll('[data-wiki-target]')].map(
      (el) => el.className,
    )
    expect(cls.some((c) => c.includes('wiki-link-resolved'))).toBe(true)
    expect(cls.some((c) => c.includes('wiki-link-unresolved'))).toBe(true)
  })

  it('clicking a resolved link opens the file it resolved to', async () => {
    const onOpen = vi.fn()
    const { editor } = open('[[Alpha]]', { onOpen, answers: ANSWERS })
    await settled()
    expect(clickWikiLink(editor, '[[Alpha]]')).toBe(true)
    expect(onOpen).toHaveBeenCalledWith('/notes/Alpha.md', '')
  })

  it('clicking a heading ref carries the fragment through', async () => {
    const onOpen = vi.fn()
    const { editor } = open('[[Gamma#Section]]', { onOpen, answers: ANSWERS })
    await settled()
    expect(clickWikiLink(editor, '[[Gamma#Section]]')).toBe(true)
    expect(onOpen).toHaveBeenCalledWith('/notes/Gamma.md', 'Section')
  })

  it('clicking an unresolved link reports it instead of opening anything', async () => {
    const onOpen = vi.fn()
    const onUnresolved = vi.fn()
    const { editor } = open('[[Nowhere]]', { onOpen, onUnresolved, answers: ANSWERS })
    await settled()
    expect(clickWikiLink(editor, '[[Nowhere]]')).toBe(true)
    expect(onOpen).not.toHaveBeenCalled()
    expect(onUnresolved).toHaveBeenCalledWith('[[Nowhere]]')
  })

  it('the resolution never becomes document state', async () => {
    const { editor } = open('[[Alpha]]', { answers: ANSWERS })
    const before = editor.state.doc.toJSON()
    await settled()
    expect(editor.state.doc.toJSON()).toEqual(before)
  })

  it('a resolver that throws leaves the editor working', async () => {
    const { editor } = open('[[Alpha]]')
    // replace the seam with a failing one
    const bad = new Editor({
      extensions: buildExtensions({
        slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
        slashItems: () => [],
        wikiLink: {
          getDocumentPath: () => '/notes/Index.md',
          resolve: async () => {
            throw new Error('resolver down')
          },
        },
      }),
      content: '',
    })
    editors.push(bad)
    bad.chain().setContent('[[Alpha]]', { contentType: 'markdown' }).run()
    await expect(settled()).resolves.toBeUndefined()
    expect(bad.state.doc.textContent).toBe('[[Alpha]]')
    void editor
  })
})
