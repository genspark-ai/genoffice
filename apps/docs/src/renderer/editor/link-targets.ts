/**
 * Targets of Word's Insert Hyperlink dialog beyond a web address: places in
 * this document (its top, headings, bookmarks) and e-mail addresses.
 */
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'

/** Word's built-in bookmark for the start of the document */
export const TOP_ANCHOR = '_top'

export interface HeadingTarget {
  pos: number
  level: number
  text: string
  /** bookmark names already on the heading (hidden first) */
  anchors: string[]
}

export interface BookmarkTarget {
  pos: number
  name: string
  preview: string
}

const names = (node: PmNode, attr: 'bookmarks' | 'hiddenBookmarks'): string[] => {
  const value = node.attrs?.[attr] as unknown
  return Array.isArray(value) ? (value as string[]) : []
}

export function collectLinkTargets(doc: PmNode): {
  headings: HeadingTarget[]
  bookmarks: BookmarkTarget[]
} {
  const headings: HeadingTarget[] = []
  const bookmarks: BookmarkTarget[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'docHeading' && node.textContent.trim()) {
      headings.push({
        pos,
        level: Math.min(Math.max(Number(node.attrs.level) || 1, 1), 9),
        text: node.textContent.trim(),
        anchors: [...names(node, 'hiddenBookmarks'), ...names(node, 'bookmarks')],
      })
    }
    for (const name of names(node, 'bookmarks'))
      bookmarks.push({ pos, name, preview: node.textContent.slice(0, 40) })
    return !node.isLeaf
  })
  return { headings, bookmarks }
}

/** position of the block carrying a bookmark (user or hidden), null when the document has none by that name */
export function anchorPos(doc: PmNode, name: string): number | null {
  let found: number | null = null
  doc.descendants((node, pos) => {
    if (found !== null) return false
    if (names(node, 'bookmarks').includes(name) || names(node, 'hiddenBookmarks').includes(name))
      found = pos
    return found === null && !node.isLeaf
  })
  return found
}

function allAnchorNames(doc: PmNode): Set<string> {
  const out = new Set<string>()
  doc.descendants((node) => {
    for (const n of names(node, 'bookmarks')) out.add(n.toLowerCase())
    for (const n of names(node, 'hiddenBookmarks')) out.add(n.toLowerCase())
    return !node.isLeaf
  })
  return out
}

/**
 * Name of the hidden bookmark a link to a heading points at: an underscore and
 * the heading text with everything but letters and digits turned into
 * underscores, 40 characters at most (the bookmark name limit), numbered when taken.
 */
export function headingAnchorName(text: string, taken: ReadonlySet<string>): string {
  const body = text
    .trim()
    .replace(/[^\p{L}\p{N}]+/gu, '_')
    .replace(/^_+|_+$/g, '')
  const base = `_${body || 'Heading'}`.slice(0, 36)
  let name = base
  for (let n = 1; taken.has(name.toLowerCase()); n++) name = `${base}_${n}`
  return name
}

const REUSABLE = /^_(?:Toc|Ref)\d/

/**
 * The anchor a link to the heading at `pos` uses. `current` (the anchor of the
 * link being edited) stays when the heading carries it, and so does a
 * table-of-contents or cross-reference bookmark; otherwise a hidden bookmark
 * is added to the heading in `tr`.
 */
export function ensureHeadingAnchor(tr: Transaction, pos: number, current?: string): string | null {
  const node = tr.doc.nodeAt(pos)
  if (!node || node.type.name !== 'docHeading') return null
  const hidden = names(node, 'hiddenBookmarks')
  if (current && (hidden.includes(current) || names(node, 'bookmarks').includes(current)))
    return current
  const wanted = headingAnchorName(node.textContent, new Set())
  // a second heading with the same text carries the numbered form of the name
  const own = (n: string) =>
    n === wanted || (n.startsWith(`${wanted}_`) && /^\d+$/.test(n.slice(wanted.length + 1)))
  const existing = hidden.find(own) ?? hidden.find((n) => REUSABLE.test(n))
  if (existing) return existing
  const name = headingAnchorName(node.textContent, allAnchorNames(tr.doc))
  tr.setNodeMarkup(pos, undefined, { ...node.attrs, hiddenBookmarks: [...hidden, name] })
  return name
}

export interface MailtoParts {
  address: string
  subject: string
}

export function parseMailto(href: string): MailtoParts | null {
  const m = /^mailto:([^?]*)(?:\?(.*))?$/i.exec(href.trim())
  if (!m) return null
  const decode = (s: string) => {
    try {
      return decodeURIComponent(s)
    } catch {
      return s
    }
  }
  const subject = (m[2] ?? '')
    .split('&')
    .map((pair) => pair.split('='))
    .find(([key]) => key.toLowerCase() === 'subject')
  return { address: decode(m[1]), subject: subject ? decode(subject.slice(1).join('=')) : '' }
}

export function buildMailto({ address, subject }: MailtoParts): string {
  const to = address.trim().replace(/^mailto:/i, '')
  if (!to) return ''
  const s = subject.trim()
  return `mailto:${to}${s ? `?subject=${encodeURIComponent(s)}` : ''}`
}

export type LinkKind = 'web' | 'doc' | 'email'

export function linkKindOf(href: string): LinkKind {
  if (href.startsWith('#')) return 'doc'
  return /^mailto:/i.test(href.trim()) ? 'email' : 'web'
}
