/**
 * `[[wiki links]]` in two halves. `wikilink.ts` decides what a link IS (syntax);
 * `wiki-resolve.ts` decides what it points AT (resolution). Both are covered
 * here because the interesting failures are the seams: a link that renders but
 * does not resolve, and a resolver that answers but is never consulted.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  collectWikiTargets,
  wikiTargetOf,
  wikilinkStart,
  wikiFragmentPosition,
} from '../src/renderer/editor/wikilink'
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
  Alpha: { path: '/notes/Alpha.md', fragment: '', found: true },
  Beta: { path: '/notes/Beta.md', fragment: '', found: true },
  'Gamma#Section': { path: '/notes/Gamma.md', fragment: 'Section', found: true },
  Nowhere: { path: '', fragment: '', found: false },
}

function open(
  source: string,
  wiki: {
    onOpen?: (path: string, fragment: string) => void
    onUnresolved?: (target: string, embed: boolean) => void
    answers?: Record<string, WikiLinkResolved>
    documentPath?: string
    resolve?: (from: string, targets: string[]) => Promise<Record<string, WikiLinkResolved>>
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
        resolve:
          wiki.resolve ??
          (async (_from, targets) => {
            const out: Record<string, WikiLinkResolved> = {}
            for (const t of targets) {
              out[t] = wiki.answers?.[t] ?? { path: '', fragment: '', found: false }
            }
            return out
          }),
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
function clickWikiLink(editor: Editor, raw: string, modifier = true): boolean {
  const el = editor.view.dom.querySelector(
    `[data-wiki-target="${CSS.escape(raw)}"]`,
  ) as HTMLElement | null
  if (!el) throw new Error(`no link rendered for ${raw}`)
  const event = {
    target: el,
    metaKey: modifier,
    ctrlKey: false,
  } as unknown as MouseEvent
  let handled = false
  editor.view.someProp('handleClick', (f) => {
    handled = f(editor.view, 0, event) === true
    return handled
  })
  return handled
}

describe('syntax', () => {
  it('wikiTargetOf drops the alias or width hint from a link and from an embed inner', () => {
    expect(wikiTargetOf('[[Note|shown]]')).toBe('Note')
    expect(wikiTargetOf('img.png|200')).toBe('img.png')
    expect(wikiTargetOf('Note#Heading')).toBe('Note#Heading')
  })
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

  it('writes a literal * or _ inside a link back unescaped on a full re-serialize', () => {
    const src = 'See [[Note *em*|a_b]] now.'
    const { editor } = open(src)
    expect(editor.getMarkdown()).toBe(src)
    expect(editor.view.dom.querySelector('.wiki-link em')).toBeNull()
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

describe('editing', () => {
  /**
   * The bytes a save writes. The source map is built when the file is opened,
   * the way App builds it, so an edit goes through the real splice path rather
   * than a plain re-serialization of the whole document.
   */
  function saver(source: string) {
    const { editor, envelope } = open(source)
    const map = buildSourceMap(editor, editor.state.doc, envelope.body)
    expect(map, 'the splice save path needs a source map').not.toBeNull()
    return { editor, save: () => spliceMarkdown(editor, editor.state.doc, map!) }
  }

  /** where a keystroke lands when the cursor sits just after `needle` */
  function after(editor: Editor, needle: string): number {
    return editor.state.doc.textContent.indexOf(needle) + needle.length + 1
  }

  it('keeps a character typed inside a link', () => {
    const { editor, save } = saver('See [[Note]] here.')
    editor.view.dispatch(editor.state.tr.insertText('s', after(editor, 'Note')))
    expect(editor.state.doc.textContent).toBe('See [[Notes]] here.')
    expect(save()).toBe('See [[Notes]] here.')
  })

  it('keeps a character typed right after a link', () => {
    const { editor, save } = saver('See [[Note]] here.')
    editor.view.dispatch(editor.state.tr.insertText('X', after(editor, ']]')))
    expect(editor.state.doc.textContent).toBe('See [[Note]]X here.')
    expect(save()).toBe('See [[Note]]X here.')
  })

  it('keeps a link the edit broke open', () => {
    const { editor, save } = saver('See [[Note]] here.')
    // `e]]`: what is left is ordinary text, and that is what must be written
    editor.view.dispatch(editor.state.tr.delete(after(editor, 'Not'), after(editor, ']]')))
    expect(editor.state.doc.textContent).toBe('See [[Not here.')
    expect(save()).toBe('See [[Not here.')
  })

  it('keeps an alias the edit changed', () => {
    const { editor, save } = saver('See [[Note|Display]] here.')
    editor.view.dispatch(editor.state.tr.insertText('!', after(editor, 'Display')))
    expect(save()).toBe('See [[Note|Display!]] here.')
  })

  it('resolves the link as the text spells it now, not as it was parsed', async () => {
    const resolve = vi.fn(async (_from: string, targets: string[]) => {
      const out: Record<string, WikiLinkResolved> = {}
      for (const t of targets) out[t] = { path: '', fragment: '', found: false }
      return out
    })
    const { editor } = open('[[Alpha]]', { resolve })
    await settled()
    expect(resolve).toHaveBeenLastCalledWith('/notes/Index.md', ['Alpha'])
    editor.view.dispatch(editor.state.tr.insertText('b', after(editor, 'Alpha')))
    await settled()
    // the target is the text the user sees, so the edit follows the link
    expect(resolve).toHaveBeenLastCalledWith('/notes/Index.md', ['Alphab'])
  })

  it('asks the resolver for file names, not literals, and keys the answers back', async () => {
    const resolve = vi.fn(async (_from: string, targets: string[]) => {
      const out: Record<string, WikiLinkResolved> = {}
      for (const t of targets) out[t] = ANSWERS[t] ?? { path: '', fragment: '', found: false }
      return out
    })
    const { editor } = open('[[Alpha]] [[Beta|Bee]] [[Gamma#Section]] [[Alpha|Again]]', { resolve })
    await settled()
    expect(resolve).toHaveBeenLastCalledWith('/notes/Index.md', ['Alpha', 'Beta', 'Gamma#Section'])
    // the alias display pass splits an aliased link's DOM around the hidden
    // syntax, duplicating the resolver's span; the link is the visible reading
    const resolved = [
      ...editor.view.dom.querySelectorAll('.wiki-link-resolved:not(.wiki-alias-hidden)'),
    ].map((el) => el.getAttribute('data-wiki-target'))
    expect(resolved).toEqual(['[[Alpha]]', '[[Beta|Bee]]', '[[Gamma#Section]]', '[[Alpha|Again]]'])
  })

  it('resolves two adjacent links as two links, not one run', () => {
    const { editor } = open('[[Alpha]][[Beta]]')
    expect(collectWikiTargets(editor.state.doc)).toEqual(['[[Alpha]]', '[[Beta]]'])
  })

  it('resolves a link split by an inline mark inside its brackets', () => {
    const { editor } = open('[[Note *emphasis*]]')
    expect(collectWikiTargets(editor.state.doc)).toEqual(['[[Note *emphasis*]]'])
  })

  it('stops resolving a link once the edit takes its brackets away', async () => {
    const { editor } = open('See [[Note]] here.', { answers: ANSWERS })
    expect(collectWikiTargets(editor.state.doc)).toEqual(['[[Note]]'])
    editor.view.dispatch(editor.state.tr.delete(after(editor, 'Not'), after(editor, ']]')))
    expect(collectWikiTargets(editor.state.doc)).toEqual([])
    await settled()
    // nothing carries the old target any more, so there is nothing to activate
    expect(editor.view.dom.querySelector('[data-wiki-target]')).toBeNull()
  })
})

describe('activation', () => {
  it('a plain click only places the caret; the modifier is what opens', async () => {
    const onOpen = vi.fn()
    const { editor } = open('[[Alpha]]', { onOpen, answers: ANSWERS })
    await settled()
    expect(clickWikiLink(editor, '[[Alpha]]', false)).toBe(false)
    expect(onOpen).not.toHaveBeenCalled()
    expect(clickWikiLink(editor, '[[Alpha]]', true)).toBe(true)
    expect(onOpen).toHaveBeenCalledWith('/notes/Alpha.md', '')
  })

  it("a plain click outside a link is not the wiki handler's business", () => {
    const { editor } = open('See [[Alpha]].')
    let handled = false
    editor.view.someProp('handleClick', (f) => {
      handled = f(editor.view, 0, { target: editor.view.dom, metaKey: true } as MouseEvent) === true
      return handled
    })
    expect(handled).toBe(false)
  })
})

describe('fragments', () => {
  it('finds the heading a #fragment names', () => {
    const { editor } = open('# Title\n\nIntro.\n\n## Some Heading\n\nBody.\n')
    const pos = wikiFragmentPosition(editor.state.doc, 'Some Heading')
    expect(pos).not.toBeNull()
    expect(editor.state.doc.resolve(pos!).parent.textContent).toBe('Some Heading')
  })

  it('matches a heading whatever its case and spacing', () => {
    const { editor } = open('## Setup   Steps\n')
    expect(wikiFragmentPosition(editor.state.doc, 'setup steps')).not.toBeNull()
  })

  it('has no position for a heading the note does not have', () => {
    const { editor } = open('## Setup\n')
    expect(wikiFragmentPosition(editor.state.doc, 'Nowhere')).toBeNull()
    expect(wikiFragmentPosition(editor.state.doc, '')).toBeNull()
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
    expect(onUnresolved).toHaveBeenCalledWith('[[Nowhere]]', false)
  })

  it('clicking an unresolved embed reports it flagged as an embed', async () => {
    const onOpen = vi.fn()
    const onUnresolved = vi.fn()
    const { editor } = open('![[Nowhere]]', { onOpen, onUnresolved, answers: ANSWERS })
    await settled()
    // an embed opens on a plain click, no modifier needed
    expect(clickWikiLink(editor, 'Nowhere')).toBe(true)
    expect(onOpen).not.toHaveBeenCalled()
    expect(onUnresolved).toHaveBeenCalledWith('Nowhere', true)
  })

  it('the resolution never becomes document state', async () => {
    const { editor } = open('[[Alpha]]', { answers: ANSWERS })
    const before = editor.state.doc.toJSON()
    await settled()
    expect(editor.state.doc.toJSON()).toEqual(before)
  })

  it('does not walk the vault again while the target set is unchanged', async () => {
    const resolve = vi.fn(async () => ({}))
    const { editor } = open('See [[Nowhere]] and some prose.', { resolve })
    await settled()
    expect(resolve).toHaveBeenCalledTimes(1)
    // typing prose around an unresolved link asks the same question again; the
    // walk it triggers can read twenty thousand entries, so it is skipped
    editor.commands.insertContentAt(editor.state.doc.content.size - 1, ' and more')
    await settled()
    expect(resolve).toHaveBeenCalledTimes(1)
    // a link the edit changed is a new question
    const inside = editor.state.doc.textContent.indexOf('Nowhere') + 'Nowhere'.length + 1
    editor.view.dispatch(editor.state.tr.insertText('s', inside))
    await settled()
    expect(resolve).toHaveBeenLastCalledWith('/notes/Index.md', ['Nowheres'])
    expect(resolve).toHaveBeenCalledTimes(2)
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
