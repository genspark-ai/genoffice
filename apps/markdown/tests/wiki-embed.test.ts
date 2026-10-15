/**
 * `![[...]]` embeds: what the syntax is (`wikiEmbed.ts`), how it resolves
 * (the same batch as links), and that an existing vault round-trips untouched.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'

import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  collectWikiEmbedTargets,
  embedInner,
  embedWidth,
  isImagePath,
  wikiEmbedStart,
} from '../src/renderer/editor/wikiEmbed'
import { parseDocText } from '../src/renderer/markdown/docText'
import { buildSourceMap, spliceMarkdown } from '../src/renderer/markdown/sourceSplice'
import type { WikiLinkResolved } from '../src/shared/ipc'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (v: string) => v }

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((e) => e.destroy())
  vi.useRealTimers()
})

function open(
  source: string,
  wiki: {
    onOpen?: (path: string, fragment: string) => void
    onUnresolved?: (target: string) => void
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

describe('syntax', () => {
  it('wikiEmbedStart points at the first unescaped ![[', () => {
    expect(wikiEmbedStart('a ![[Note]] b')).toBe(2)
    expect(wikiEmbedStart('plain [[Note]] only')).toBe(-1)
    expect(wikiEmbedStart('escaped \\![[x]] tail')).toBe(-1)
  })

  it('folds the literal down to the resolver form and back out', () => {
    expect(embedInner('![[Note]]')).toBe('Note')
    expect(embedInner('![[Note#Heading]]')).toBe('Note#Heading')
    expect(embedInner('![[img.png|200]]')).toBe('img.png|200')
  })

  it('reads a numeric alias as a width hint and leaves words alone', () => {
    expect(embedWidth('![[img.png|200]]')).toBe(200)
    expect(embedWidth('![[img.png]]')).toBeNull()
    expect(embedWidth('![[Note|Alias]]')).toBeNull()
  })

  it('knows a picture from a note', () => {
    expect(isImagePath('/notes/img.PNG')).toBe(true)
    expect(isImagePath('/notes/Note.md')).toBe(false)
    expect(isImagePath('/notes/noext')).toBe(false)
  })

  it('collects each distinct embed target once, in document order', () => {
    const { editor } = open('![[A]] then ![[A]] then ![[B]]')
    expect(collectWikiEmbedTargets(editor.state.doc)).toEqual(['A', 'B'])
  })

  it('an embed is a node, a link is a mark, one line can carry both', () => {
    const { editor } = open('An ![[Alpha]] and a [[Beta]].')
    const kinds: string[] = []
    editor.state.doc.descendants((n) => {
      if (n.type.name === 'wikiEmbed') kinds.push('embed')
      if (n.isText && n.marks.some((m) => m.type.name === 'wikiLink')) kinds.push('link')
    })
    expect(kinds).toEqual(['embed', 'link'])
  })

  it.each([
    ['a bare embed', 'Embed ![[Note]] inline.'],
    ['a heading embed', 'See ![[Note#Some Heading]] here.'],
    ['a sized picture', '![[img.png|200]]'],
    ['a plain picture', '![[img.png]]'],
  ])('parses %s and writes it back byte for byte', (_label, src) => {
    const { editor, envelope } = open(src)
    const map = buildSourceMap(editor, editor.state.doc, envelope.body)
    const out = map ? spliceMarkdown(editor, editor.state.doc, map) : editor.getMarkdown()
    expect(out).toBe(src)
  })

  it('a bang that is not an embed stays plain text', () => {
    const { editor } = open('A [[Note]] with a lone ! mark.')
    expect(collectWikiEmbedTargets(editor.state.doc)).toEqual([])
  })
})

describe('resolution', () => {
  it('a picture the asset handler refuses falls back to its card, not an empty box', async () => {
    // the resolver walks the whole vault but md-asset only serves files under
    // an open document's own directory, so this one resolves and then 403s
    const { editor } = open('![[diagram.png]]', {
      answers: { 'diagram.png': { path: '/vault/assets/diagram.png', fragment: '', found: true } },
    })
    await settled()
    const node = editor.view.dom.querySelector('[data-wiki-embed]') as HTMLElement | null
    expect(node).not.toBeNull()
    expect(node!.className).toContain('wiki-embed-image')
    const img = editor.view.dom.querySelector('img.wiki-embed-img') as HTMLImageElement
    expect(img).not.toBeNull()

    img.dispatchEvent(new Event('error'))
    await settled()
    const after = editor.view.dom.querySelector('[data-wiki-embed]') as HTMLElement | null
    expect(after).not.toBeNull()
    // the card is the honest rendering: the text is back and the class is off
    expect(after!.className).not.toContain('wiki-embed-image')
    expect(editor.view.dom.querySelector('img.wiki-embed-img')).toBeNull()
  })

  it('a resolved note embed carries its answer and opens on a plain click', async () => {
    const onOpen = vi.fn()
    const { editor } = open('![[Gamma#Section]]', {
      onOpen,
      answers: { 'Gamma#Section': { path: '/notes/Gamma.md', fragment: 'Section', found: true } },
    })
    await settled()
    const el = editor.view.dom.querySelector('[data-wiki-embed]') as HTMLElement | null
    expect(el).not.toBeNull()
    expect(el!.getAttribute('data-wiki-target')).toBe('Gamma#Section')
    expect(el!.className).toContain('wiki-embed-resolved')
    // no modifier: an embed has no text to edit, so a plain click opens
    const event = { target: el, metaKey: false, ctrlKey: false } as unknown as MouseEvent
    let handled = false
    editor.view.someProp('handleClick', (f) => {
      handled = f(editor.view, 0, event) === true
      return handled
    })
    expect(handled).toBe(true)
    expect(onOpen).toHaveBeenCalledWith('/notes/Gamma.md', 'Section')
  })

  it('an unresolved embed stays dashed and reports the target as written', async () => {
    const onUnresolved = vi.fn()
    const { editor } = open('![[Nowhere]]', {
      onUnresolved,
      answers: {},
    })
    await settled()
    const el = editor.view.dom.querySelector('[data-wiki-embed]') as HTMLElement | null
    expect(el!.className).toContain('wiki-embed-unresolved')
    const event = { target: el, metaKey: false, ctrlKey: false } as unknown as MouseEvent
    let handled = false
    editor.view.someProp('handleClick', (f) => {
      handled = f(editor.view, 0, event) === true
      return handled
    })
    expect(handled).toBe(true)
    expect(onUnresolved).toHaveBeenCalledWith('Nowhere')
  })

  it('a resolved picture renders an image widget beside the hidden literal', async () => {
    const { editor } = open('![[pic.png]]', {
      answers: { 'pic.png': { path: '/notes/pic.png', fragment: '', found: true } },
    })
    await settled()
    const node = editor.view.dom.querySelector('.wiki-embed') as HTMLElement | null
    expect(node!.className).toContain('wiki-embed-image')
    const img = editor.view.dom.querySelector('img.wiki-embed-img') as HTMLImageElement | null
    expect(img).not.toBeNull()
    expect(img!.getAttribute('src')).toContain('md-asset://')
  })

  it('a picture embed with a width hint sizes the image', async () => {
    const { editor } = open('![[pic.png|120]]', {
      answers: { 'pic.png|120': { path: '/notes/pic.png', fragment: '', found: true } },
    })
    await settled()
    const img = editor.view.dom.querySelector('img.wiki-embed-img') as HTMLImageElement | null
    expect(img!.style.width).toBe('120px')
  })

  it('links and embeds resolve in one batch, each under its own key', async () => {
    const seen: string[][] = []
    const { editor } = open('A [[Alpha]] and an ![[Beta]].', {
      resolve: async (_from, targets) => {
        seen.push(targets)
        return {
          '[[Alpha]]': { path: '/a.md', fragment: '', found: true },
          Beta: { path: '/b.md', fragment: '', found: true },
        }
      },
    })
    await settled()
    expect(seen[0]).toEqual(['[[Alpha]]', 'Beta'])
    const linkEl = editor.view.dom.querySelector('.wiki-link-resolved')
    expect(linkEl).not.toBeNull()
    const embedEl = editor.view.dom.querySelector('.wiki-embed-resolved')
    expect(embedEl).not.toBeNull()
  })
})
