import { Mark, Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { WikiLinkResolved } from '../../shared/ipc'
import { collectWikiEmbedTargets, embedWidth, isImagePath, wikiEmbedSpans } from './wikiEmbed'
import { resolveImageSrc } from './localImage'

/**
 * `[[target]]`, `[[target|alias]]`, `[[target#heading]]`, `[[target#^blockid]]`.
 * The target may hold anything but a bracket, pipe or newline; the alias runs
 * to the closing brackets. `[[]]` carries no name and is not a link.
 */
const WIKI_LINK_RE = /^\[\[([^[\]|\n]+?)(?:\|([^[\]\n]+))?\]\]/

/**
 * Index of the first unescaped `[[`. The stock link tokenizer opens on a
 * single `[`, so both candidates start at the same offset and marked breaks
 * that tie by extension order — pointing `start` at the first bracket is what
 * puts a wiki link in the running at all.
 */
export function wikilinkStart(src: string): number {
  for (let i = 0; i < src.length - 1; i++) {
    if (src[i] !== '[' || src[i + 1] !== '[') continue
    if (i > 0 && src[i - 1] === '\\') continue // `\[` is an escaped literal
    return i
  }
  return -1
}

/**
 * A `[[...]]` literal and the text it covers.
 *
 * The literal is read from the text, not from the mark's attributes: the text
 * IS the link (that is what `parseMarkdown` puts there), so it is what an edit
 * changes, what the save writes and what has to be resolved. Attributes are
 * written once while parsing and keep answering for the pre-edit target.
 */
interface WikiSpan {
  /** the link as written, brackets included: `[[Note|Alias]]` */
  raw: string
  /** document positions the literal covers */
  from: number
  to: number
}

/** the `[[...]]` literals in `text`, in order; one run can spell several */
function wikiLiterals(text: string): { start: number; end: number; raw: string }[] {
  const out: { start: number; end: number; raw: string }[] = []
  for (let at = 0; at < text.length;) {
    const match = WIKI_LINK_RE.exec(text.slice(at))
    if (!match) break
    out.push({ start: at, end: at + match[0].length, raw: match[0] })
    at += match[0].length
  }
  return out
}

const isWikiText = (node: PmNode | null | undefined): boolean =>
  !!node?.isText && node.marks.some((mark) => mark.type.name === 'wikiLink')

/**
 * Every link in the document, with the positions it covers.
 *
 * The siblings under one mark are gathered before the run is read, because a
 * literal can be spread over several text nodes — an inline mark inside the
 * brackets (`[[Note *x*]]`) splits it, and an edit can leave the mark wider
 * than the link (`[[Note]]X`) or wider than any link at all (`[[Note` after the
 * closing brackets are deleted). What makes a link is the brackets, not the
 * extent of the mark.
 */
function wikiSpans(doc: PmNode): WikiSpan[] {
  const spans: WikiSpan[] = []
  const walked = new Set<PmNode>()
  doc.descendants((node, pos, parent, index) => {
    if (!parent || !isWikiText(node) || walked.has(node)) return
    const kids: PmNode[] = []
    parent.forEach((child) => kids.push(child))
    let first = index
    while (first > 0 && isWikiText(kids[first - 1])) first--
    let runStart = pos
    for (let i = index; i > first; i--) runStart -= kids[i]!.nodeSize
    let text = ''
    for (let i = first; i < kids.length && isWikiText(kids[i]); i++) {
      walked.add(kids[i]!)
      text += kids[i]!.text ?? ''
    }
    for (const link of wikiLiterals(text)) {
      // a literal split by an inline mark covers several nodes; one the edit
      // extended past its brackets covers part of a single node
      let cursor = runStart
      for (let i = first; i < kids.length; i++) {
        const end = cursor + kids[i]!.nodeSize
        const from = Math.max(cursor, runStart + link.start)
        const to = Math.min(end, runStart + link.end)
        if (to > from) spans.push({ raw: link.raw, from, to })
        cursor = end
        if (cursor >= runStart + link.end) break
      }
    }
  })
  return spans
}

/** every distinct wiki target in the document, in document order */
export function collectWikiTargets(doc: PmNode): string[] {
  const seen = new Set<string>()
  const out: string[] = []
  for (const span of wikiSpans(doc)) {
    if (seen.has(span.raw)) continue
    seen.add(span.raw)
    out.push(span.raw)
  }
  return out
}

/**
 * Where a `#fragment` lands in a note: inside the heading it names, or null
 * when the note has no such heading. A `#^block` ref needs block ids, which
 * this schema does not carry yet, so it opens the note without a position.
 */
export function wikiFragmentPosition(doc: PmNode, fragment: string): number | null {
  const want = fragment.trim().replace(/\s+/g, ' ').toLowerCase()
  if (!want) return null
  let found: number | null = null
  doc.descendants((node, pos) => {
    if (found !== null || node.type.name !== 'heading') return
    if (node.textContent.replace(/\s+/g, ' ').trim().toLowerCase() === want) found = pos + 1
  })
  return found
}

export interface WikiLinkOptions {
  /** absolute path of the open document; '' while it is untitled */
  getDocumentPath(): string
  /** a link whose target was found was activated */
  onOpen(path: string, fragment: string): void
  /** a link whose target was not found was activated, with the target as written */
  onUnresolved(target: string): void
  /** seam for tests; defaults to the real IPC */
  resolve(fromPath: string, targets: string[]): Promise<Record<string, WikiLinkResolved>>
}

interface WikiState {
  decorations: DecorationSet
  answers: Record<string, WikiLinkResolved>
}

const EMPTY: WikiState = { decorations: DecorationSet.empty, answers: {} }

const key = new PluginKey<WikiState>('wikiLinkResolve')

/**
 * The text under a wiki mark. The serializer asks for the opening and closing
 * wrapper by rendering the mark against a placeholder node — a plain object
 * carrying `content`, not a real ProseMirror node — so both shapes are read.
 */
function markedText(node: any): string {
  return node.textContent ?? node.content?.[0]?.text ?? ''
}

/**
 * Resolution lives in plugin state, not on the mark: a mark attribute is
 * document state, so writing the answer there would turn every resolution
 * into a document edit — and the block source map must not see those.
 */
export const WikiLinkResolve = Extension.create<WikiLinkOptions>({
  name: 'wikiLinkResolve',

  addOptions() {
    return {
      getDocumentPath: () => '',
      onOpen: () => {},
      onUnresolved: () => {},
      resolve: (fromPath, targets) => window.markdownApi.resolveWikiLinks({ fromPath, targets }),
    }
  },

  addProseMirrorPlugins() {
    const options = this.options
    let timer: ReturnType<typeof setTimeout> | undefined
    let generation = 0
    /** what the last resolution asked for; a pause that asks the same thing again is not a new question */
    let asked = ''

    const resolveNow = async (view: EditorView) => {
      // one batch answers links and embeds: an embed is the same lookup with
      // the `!` folded away, and both halves pay for a single traversal
      const targets = [
        ...collectWikiTargets(view.state.doc),
        ...collectWikiEmbedTargets(view.state.doc),
      ]
      if (targets.length === 0) {
        asked = ''
        if (view.isDestroyed) return
        view.dispatch(view.state.tr.setMeta(key, { decorations: DecorationSet.empty, answers: {} }))
        return
      }
      const fromPath = options.getDocumentPath()
      const question = `${fromPath}\u0000${targets.join('\u0000')}`
      // the walk it triggers can read twenty thousand entries; typing prose
      // around an unresolved link changes no target, so it asks nothing
      if (question === asked) return
      const mine = ++generation
      let answers: Record<string, WikiLinkResolved>
      try {
        answers = await options.resolve(fromPath, targets)
      } catch {
        return // an unavailable resolver must not break editing
      }
      // a later keystroke started a newer resolution; its answer is the current one
      if (mine !== generation || view.isDestroyed) return
      asked = question

      const linkDecorations = wikiSpans(view.state.doc).map((span) =>
        Decoration.inline(span.from, span.to, {
          class: answers[span.raw]?.found ? 'wiki-link-resolved' : 'wiki-link-unresolved',
          'data-wiki-target': span.raw,
        }),
      )
      const embedDecorations: Decoration[] = []
      for (const span of wikiEmbedSpans(view.state.doc)) {
        const answer = answers[span.inner]
        const found = answer?.found === true
        // a resolved picture swaps the literal for the image itself, so the
        // node also carries the class that hides its own text
        const picture = found && isImagePath(answer!.path)
        embedDecorations.push(
          Decoration.node(span.from, span.to, {
            class: `${found ? 'wiki-embed-resolved' : 'wiki-embed-unresolved'}${
              picture ? ' wiki-embed-image' : ''
            }`,
            'data-wiki-target': span.inner,
            'data-wiki-embed': '1',
          }),
        )
        if (picture) {
          const width = embedWidth(`![[${span.inner}]]`)
          embedDecorations.push(
            Decoration.widget(
              span.to,
              () => {
                const img = document.createElement('img')
                img.className = 'wiki-embed-img'
                img.src = resolveImageSrc(answer!.path)
                img.alt = span.inner
                if (width !== null) img.style.width = `${width}px`
                return img
              },
              { side: 10 },
            ),
          )
        }
      }
      view.dispatch(
        view.state.tr.setMeta(key, {
          decorations: DecorationSet.create(view.state.doc, [
            ...linkDecorations,
            ...embedDecorations,
          ]),
          answers,
        }),
      )
    }

    const schedule = (view: EditorView) => {
      if (timer) clearTimeout(timer)
      // one traversal per burst of typing, not one per keystroke
      timer = setTimeout(() => void resolveNow(view), 150)
    }

    return [
      new Plugin<WikiState>({
        key,
        state: {
          init: () => EMPTY,
          apply: (tr, old) =>
            tr.getMeta(key) ?? { ...old, decorations: old.decorations.map(tr.mapping, tr.doc) },
        },
        props: {
          decorations: (state) => (key.getState(state) ?? EMPTY).decorations,
          handleClick: (view, _pos, event) => {
            const el = (event.target as HTMLElement | null)?.closest<HTMLElement>(
              '[data-wiki-target]',
            )
            if (!el) return false
            // a link's text stays editable, so the modifier is what leaves the
            // document (as with `openOnClick: false` elsewhere); an embed has no
            // text to edit, so a plain click is what opens it
            const embed = el.getAttribute('data-wiki-embed') === '1'
            if (!embed && !event.metaKey && !event.ctrlKey) return false
            const raw = el.getAttribute('data-wiki-target') ?? ''
            const answer = (key.getState(view.state) ?? EMPTY).answers[raw]
            if (answer?.found) options.onOpen(answer.path, answer.fragment)
            else options.onUnresolved(raw)
            return true
          },
        },
        view: (view) => {
          schedule(view)
          return {
            update: (v, prevState) => {
              if (v.state.doc !== prevState.doc) schedule(v)
            },
            destroy: () => {
              if (timer) clearTimeout(timer)
            },
          }
        },
      }),
    ]
  },
})

export const WikiLink = Mark.create({
  name: 'wikiLink',
  // above Link (1000): a wiki link must open before a markdown link claims the
  // same bracket, and ProseMirror derives mark ranks from this order
  priority: 1100,

  markdownTokenizer: {
    name: 'wikiLink',
    level: 'inline',
    start: wikilinkStart,
    tokenize: (src: string) => {
      const match = WIKI_LINK_RE.exec(src)
      if (!match) return undefined
      // the literal is the token; there is nothing else to carry, because the
      // text it becomes is the link (see `wikiSpans`)
      return { type: 'wikiLink', raw: match[0] }
    },
  },

  parseMarkdown: (token: Record<string, unknown>, helpers: any) =>
    helpers.applyMark('wikiLink', helpers.parseInline([{ type: 'text', text: String(token.raw) }])),

  /**
   * The marked text verbatim. A mark is serialized as an opening and a closing
   * wrapper around the text it covers, both taken from this renderMarkdown
   * against a placeholder; the literal never appears in that placeholder, so
   * the wrapper is empty and the text is written out as it stands — which is
   * what makes an edit inside or after a link survive the save.
   */
  renderMarkdown: (node: any) => markedText(node),

  renderHTML() {
    // no target here: what a link points at is an answer from the resolver,
    // and it arrives as a decoration once the target set is resolved
    return ['a', { class: 'wiki-link' }, 0]
  },
})
