/**
 * Word's Review → Show Markup filters and Reviewing Pane model, layered over
 * collectRevisions without touching the recorder: which revisions the current
 * view shows, per-author ink assignment, and the pane's entry list.
 */
import type { Node as PmNode } from '@tiptap/pm/model'
import type { RevisionRange } from './revisions'

export const REVISION_INK_COUNT = 8

/** Accept / Reject menu targets (Word's dropdown): one change, every change, the filtered view, one author, the selection */
export type RevisionScope = 'next' | 'all' | 'shown' | 'selection' | { author: string }

export type RevisionGroup = 'insDel' | 'format'

export interface RevisionFilter {
  insDel: boolean
  format: boolean
  /** authors unticked under Specific People */
  hiddenAuthors: string[]
}

export const DEFAULT_REVISION_FILTER: RevisionFilter = {
  insDel: true,
  format: true,
  hiddenAuthors: [],
}

const FORMAT_KINDS = new Set<RevisionRange['kind']>([
  'rPrChange',
  'pPrChange',
  'paraMarkRPrChange',
  'tblPrChange',
  'tblGridChange',
  'trPrChange',
  'tblPrExChange',
  'tcPrChange',
  'cellMerge',
  'numberingChange',
  'sectPrChange',
])

export function revisionGroup(kind: RevisionRange['kind']): RevisionGroup {
  return FORMAT_KINDS.has(kind) ? 'format' : 'insDel'
}

export function revisionShown(
  filter: RevisionFilter,
  rev: Pick<RevisionRange, 'kind' | 'author'>,
): boolean {
  if (filter.hiddenAuthors.includes(rev.author)) return false
  return revisionGroup(rev.kind) === 'format' ? filter.format : filter.insDel
}

export function filterRevisions<T extends Pick<RevisionRange, 'kind' | 'author'>>(
  ranges: T[],
  filter: RevisionFilter,
): T[] {
  return ranges.filter((r) => revisionShown(filter, r))
}

/** revisions touching a (non-empty) selection; a caret inside a range counts */
export function revisionsInSelection<T extends Pick<RevisionRange, 'from' | 'to'>>(
  ranges: T[],
  from: number,
  to: number,
): T[] {
  return ranges.filter((r) =>
    from === to ? from >= r.from && from <= r.to : r.from < to && r.to > from,
  )
}

/** authors in order of first appearance; ink index = position mod REVISION_INK_COUNT */
export function revisionAuthors(ranges: Array<Pick<RevisionRange, 'author'>>): string[] {
  const out: string[] = []
  for (const r of ranges) if (!out.includes(r.author)) out.push(r.author)
  return out
}

export function authorInkIndex(authors: string[], author: string): number {
  const i = authors.indexOf(author)
  return i < 0 ? 0 : i % REVISION_INK_COUNT
}

const cssString = (s: string) => `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`

/**
 * Stylesheet mapping each author's marks to their ink and hiding unticked
 * authors as if accepted. Author index 0 keeps the default ink, so a
 * single-author document renders exactly as before. Moves keep Word's move
 * color whoever made them.
 */
export function revisionAuthorCss(
  authors: string[],
  hiddenAuthors: string[],
  movesByAuthor: { from: boolean; to: boolean } = { from: false, to: false },
): string {
  const rules: string[] = []
  const insSel = movesByAuthor.to ? '' : ':not(.doc-move-to)'
  const delSel = movesByAuthor.from ? '' : ':not(.doc-move-from)'
  authors.forEach((author, i) => {
    const idx = i % REVISION_INK_COUNT
    if (idx === 0) return
    const a = cssString(author)
    rules.push(
      `.doc-ins[data-ins-author=${a}]${insSel},.doc-del[data-del-author=${a}]${delSel},.has-rpr-change[data-rpr-author=${a}]{--docs-paper-rev-ink:var(--docs-paper-rev-ink-${idx})}`,
    )
  })
  for (const author of hiddenAuthors) {
    const a = cssString(author)
    rules.push(
      `.doc-ins[data-ins-author=${a}]{color:inherit;text-decoration:none;font-weight:inherit;font-style:inherit}`,
      `.doc-del[data-del-author=${a}]{display:none}`,
      `.doc-del-marker[data-del-author=${a}]{display:none}`,
      `.has-rpr-change[data-rpr-author=${a}]{color:inherit;text-decoration:none;font-weight:inherit;font-style:inherit}`,
      `.has-move-from[data-rev-author=${a}]{display:none}`,
      `.has-move-to[data-rev-author=${a}]{background:none}`,
      `.has-ppr-change[data-ppr-change-author=${a}]{border-left:none;padding-left:0}`,
      `.has-ppr-change[data-ppr-change-author=${a}]::after{content:none}`,
    )
  }
  return rules.join('\n')
}

export interface ReviewEntry {
  rev: RevisionRange
  index: number
  /** first characters of the changed text; empty for row/cell/paragraph-only changes */
  excerpt: string
  inkIndex: number
}

export const EXCERPT_CHARS = 60

export function revisionExcerpt(doc: PmNode, rev: Pick<RevisionRange, 'from' | 'to'>): string {
  const from = Math.max(0, Math.min(rev.from, doc.content.size))
  const to = Math.max(from, Math.min(rev.to, doc.content.size))
  const text = doc.textBetween(from, to, ' ', ' ').replace(/\s+/g, ' ').trim()
  return text.length > EXCERPT_CHARS ? `${text.slice(0, EXCERPT_CHARS - 1)}…` : text
}

export function buildReviewEntries(
  doc: PmNode,
  ranges: RevisionRange[],
  authors = revisionAuthors(ranges),
): ReviewEntry[] {
  return ranges.map((rev, index) => ({
    rev,
    index,
    excerpt: rev.story ? rev.story.text.slice(0, EXCERPT_CHARS) : revisionExcerpt(doc, rev),
    inkIndex: authorInkIndex(authors, rev.author),
  }))
}

/** per-kind tally; a move whose source is listed counts once, under moveFrom */
export function countByKind(
  ranges: Array<Pick<RevisionRange, 'kind' | 'moveId'>>,
): Map<RevisionRange['kind'], number> {
  const out = new Map<RevisionRange['kind'], number>()
  const sources = new Set(
    ranges.filter((r) => r.kind === 'moveFrom' && r.moveId).map((r) => r.moveId),
  )
  for (const r of ranges) {
    if (r.kind === 'moveTo' && r.moveId && sources.has(r.moveId)) continue
    out.set(r.kind, (out.get(r.kind) ?? 0) + 1)
  }
  return out
}

export interface RevisionSummary {
  insertions: number
  deletions: number
  moves: number
  formatting: number
}

const INSERT_KINDS = new Set<RevisionRange['kind']>([
  'ins',
  'paraMarkIns',
  'rowIns',
  'cellIns',
  'blockIns',
])

/** Word's Reviewing Pane summary lines; a move's two halves count once */
export function summarizeRevisions(
  ranges: Array<Pick<RevisionRange, 'kind' | 'moveId'>>,
): RevisionSummary {
  const out: RevisionSummary = { insertions: 0, deletions: 0, moves: 0, formatting: 0 }
  const moveIds = new Set<string>()
  for (const r of ranges) {
    if (r.kind === 'moveFrom' || r.kind === 'moveTo') {
      if (r.moveId) moveIds.add(r.moveId)
      else out.moves++
    } else if (INSERT_KINDS.has(r.kind)) out.insertions++
    else if (FORMAT_KINDS.has(r.kind)) out.formatting++
    else out.deletions++
  }
  out.moves += moveIds.size
  return out
}
