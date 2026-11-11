/**
 * `![[Note]]`, `![[Note#Heading]]`, `![[image.png]]`, `![[image.png|200]]` —
 * Obsidian's embed form. A link is text you can edit; an embed is a THING on
 * the line (a picture, another note), so it is an atom node carrying its
 * literal, not a mark around text: there is nothing inside to place a caret
 * in, and an atom's attribute cannot drift from what the save writes the way
 * a mark's did.
 */
import { Node } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'

/** the literal a wiki embed node carries: `![[Note#Heading]]` */
export const WIKI_EMBED_RE = /^!\[\[([^[\]|\n]+?)(?:\|([^[\]\n]+))?\]\]/

/**
 * Index of the first `![[`. One position before a wiki link's own `[[`, so
 * the tokenizer whose start is earlier wins the same span — what makes
 * `![[Note]]` an embed while `[[Note]]` stays a link.
 */
export function wikiEmbedStart(src: string): number {
  for (let i = 0; i < src.length - 2; i++) {
    if (src[i] !== '!' || src[i + 1] !== '[' || src[i + 2] !== '[') continue
    if (i > 0 && src[i - 1] === '\\') continue // `\!` is an escaped literal
    return i
  }
  return -1
}

/** media extensions an embed renders as a picture rather than a note card */
const IMAGE_EXTS = new Set([
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
  '.webp',
  '.svg',
  '.bmp',
  '.avif',
  '.ico',
])

export function isImagePath(path: string): boolean {
  const dot = path.lastIndexOf('.')
  if (dot < 0) return false
  return IMAGE_EXTS.has(path.slice(dot).toLowerCase())
}

/** an embed's target without the brackets and bang: `Note#Heading|200` */
export function embedInner(raw: string): string {
  const match = WIKI_EMBED_RE.exec(raw)
  if (!match) return raw
  return `${match[1] ?? ''}${match[2] !== undefined ? `|${match[2]}` : ''}`
}

/** a trailing `|200` width hint, in px, when the embed carries one */
export function embedWidth(raw: string): number | null {
  const match = WIKI_EMBED_RE.exec(raw)
  const hint = match?.[2] ?? ''
  return /^\d{1,4}$/.test(hint) ? Number(hint) : null
}

export interface WikiEmbedSpan {
  /** the target as the resolver keys it: `Note#Heading`, no brackets or bang */
  inner: string
  /** document positions of the whole node */
  from: number
  to: number
}

/** every embed in the document, with the positions it covers */
export function wikiEmbedSpans(doc: PmNode): WikiEmbedSpan[] {
  const spans: WikiEmbedSpan[] = []
  doc.descendants((node, pos) => {
    if (node.type.name !== 'wikiEmbed') return
    spans.push({
      inner: embedInner(String(node.attrs.raw ?? '')),
      from: pos,
      to: pos + node.nodeSize,
    })
  })
  return spans
}

/** every distinct embed target in the document, in document order */
export function collectWikiEmbedTargets(doc: PmNode): string[] {
  const seen = new Set<string>()
  const out: string[] = []
  for (const span of wikiEmbedSpans(doc)) {
    if (seen.has(span.inner)) continue
    seen.add(span.inner)
    out.push(span.inner)
  }
  return out
}

export const WikiEmbed = Node.create({
  name: 'wikiEmbed',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      /** the literal as written, brackets and bang included — written out verbatim */
      raw: { default: '![[]]' },
    }
  },

  markdownTokenizer: {
    name: 'wikiEmbed',
    level: 'inline',
    start: wikiEmbedStart,
    tokenize: (src: string) => {
      const match = WIKI_EMBED_RE.exec(src)
      if (!match) return undefined
      return { type: 'wikiEmbed', raw: match[0] }
    },
  },

  parseMarkdown: (token: Record<string, unknown>, helpers: any) =>
    helpers.createNode('wikiEmbed', { raw: String(token.raw) }),

  /** the literal verbatim, so an existing vault round-trips untouched */
  renderMarkdown: (node: any) => String(node.attrs?.raw ?? ''),

  renderHTML({ node }: { node: PmNode }) {
    return [
      'span',
      { class: 'wiki-embed', 'data-wiki-raw': String(node.attrs.raw ?? '') },
      String(node.attrs.raw ?? ''),
    ]
  },
})
