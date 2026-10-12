import { Mark, Extension, mergeAttributes } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { WikiLinkResolved } from '../../shared/ipc'

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

/** every distinct wiki target in the document, in document order */
export function collectWikiTargets(doc: PmNode): string[] {
  const seen = new Set<string>()
  const out: string[] = []
  doc.descendants((node) => {
    if (!node.isText) return
    for (const mark of node.marks) {
      if (mark.type.name !== 'wikiLink') continue
      const raw = mark.attrs.raw as string
      if (typeof raw !== 'string' || raw === '' || seen.has(raw)) continue
      seen.add(raw)
      out.push(raw)
    }
  })
  return out
}

const key = new PluginKey<WikiState>('wikiLinkResolve')

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

    const resolveNow = async (view: EditorView) => {
      const targets = collectWikiTargets(view.state.doc)
      if (targets.length === 0) {
        if (view.isDestroyed) return
        view.dispatch(view.state.tr.setMeta(key, { decorations: DecorationSet.empty, answers: {} }))
        return
      }
      const mine = ++generation
      let answers: Record<string, WikiLinkResolved>
      try {
        answers = await options.resolve(options.getDocumentPath(), targets)
      } catch {
        return // an unavailable resolver must not break editing
      }
      // a later keystroke started a newer resolution; its answer is the current one
      if (mine !== generation || view.isDestroyed) return

      const decorations: Decoration[] = []
      view.state.doc.descendants((node, pos) => {
        if (!node.isText) return
        for (const mark of node.marks) {
          if (mark.type.name !== 'wikiLink') continue
          const raw = mark.attrs.raw as string
          decorations.push(
            Decoration.inline(pos, pos + node.nodeSize, {
              class: answers[raw]?.found ? 'wiki-link-resolved' : 'wiki-link-unresolved',
              'data-wiki-target': raw,
            }),
          )
        }
      })
      view.dispatch(
        view.state.tr.setMeta(key, {
          decorations: DecorationSet.create(view.state.doc, decorations),
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

  addAttributes() {
    return {
      raw: { default: '' },
      target: { default: '' },
      alias: { default: null as string | null },
    }
  },

  markdownTokenizer: {
    name: 'wikiLink',
    level: 'inline',
    start: wikilinkStart,
    tokenize: (src: string) => {
      const match = WIKI_LINK_RE.exec(src)
      if (!match) return undefined
      return { type: 'wikiLink', raw: match[0], target: match[1].trim(), alias: match[2] ?? null }
    },
  },

  parseMarkdown: (token: Record<string, unknown>, helpers: any) =>
    helpers.applyMark(
      'wikiLink',
      helpers.parseInline([{ type: 'text', text: String(token.raw) }]),
      {
        raw: String(token.raw),
        target: String(token.target),
        alias: token.alias as string | null,
      },
    ),

  renderMarkdown: (node: any) => {
    const alias = node.attrs.alias as string | null
    return `[[${node.attrs.target}${alias ? `|${alias}` : ''}]]`
  },

  renderHTML({ HTMLAttributes }: { HTMLAttributes: Record<string, string> }) {
    return [
      'a',
      mergeAttributes(
        { class: 'wiki-link', 'data-wiki-target': HTMLAttributes.raw },
        HTMLAttributes,
      ),
      0,
    ]
  },
})
