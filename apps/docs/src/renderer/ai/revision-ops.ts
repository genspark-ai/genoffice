import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  applyRevisions,
  collectRevisions,
  PARAGRAPH_FORMAT_FIELDS,
  TEXT_STYLE_FIELDS,
  type RevisionRange,
} from '../editor/revisions'

export type RevisionType = 'insertion' | 'deletion' | 'formatting' | 'move'

export const REVISION_TYPES: readonly RevisionType[] = [
  'insertion',
  'deletion',
  'formatting',
  'move',
]

export interface RevisionEntry {
  /** positional id (r1 = first pending change in document order); renumbered after every edit */
  id: string
  type: RevisionType
  kind: RevisionRange['kind']
  author: string
  date?: string
  blockIndex: number
  /** affected text (what disappears on accept for deletions, what stays for insertions) */
  text: string
  /** formatting changes: `field: old → new` pairs */
  change?: string
  from: number
  to: number
}

export interface RevisionSelector {
  all?: boolean
  ids?: string[]
  author?: string
  type?: RevisionType
  blockIndex?: number
  blockRange?: [number, number]
  /** ISO date: only changes recorded strictly before it */
  before?: string
}

const TYPE_OF: Record<RevisionRange['kind'], RevisionType> = {
  ins: 'insertion',
  blockIns: 'insertion',
  rowIns: 'insertion',
  cellIns: 'insertion',
  // ins+del marks on the same text count as a deletion: both accept and reject remove it
  both: 'deletion',
  del: 'deletion',
  blockDel: 'deletion',
  rowDel: 'deletion',
  cellDel: 'deletion',
  pPrChange: 'formatting',
  rPrChange: 'formatting',
  moveFrom: 'move',
  moveTo: 'move',
}

function blockIndexOfPos(doc: PmNode, pos: number): number {
  let index = 0
  let found = 0
  doc.forEach((node, offset) => {
    if (pos >= offset && pos < offset + node.nodeSize) found = index
    index++
  })
  return found
}

const show = (v: unknown): string =>
  v === null || v === undefined || v === false ? 'none' : v === true ? 'yes' : String(v)

function diffPairs(
  old: Record<string, unknown>,
  cur: Record<string, unknown>,
  fields: readonly string[],
): string[] {
  const out: string[] = []
  const norm = (v: unknown) => (v === undefined || v === false ? null : v)
  for (const f of fields) {
    const a = norm(old[f])
    const b = norm(cur[f])
    if (JSON.stringify(a) !== JSON.stringify(b)) out.push(`${f}: ${show(a)} → ${show(b)}`)
  }
  return out
}

function runFormatChange(doc: PmNode, r: RevisionRange): string | undefined {
  let text: PmNode | null = null
  doc.nodesBetween(r.from, r.to, (node) => {
    if (text || !node.isText) return
    text = node
  })
  if (!text) return undefined
  const t = text as PmNode
  const rpr = t.marks.find((m) => m.type.name === 'rprChange')
  const old = (rpr?.attrs.old ?? {}) as Record<string, unknown>
  const ts = t.marks.find((m) => m.type.name === 'docTextStyle')
  const cur: Record<string, unknown> = { ...(ts?.attrs ?? {}) }
  for (const flag of ['bold', 'italic', 'underline', 'strike']) {
    cur[flag] = t.marks.some((m) => m.type.name === flag)
  }
  const pairs = diffPairs(old, cur, ['bold', 'italic', 'underline', 'strike', ...TEXT_STYLE_FIELDS])
  return pairs.length ? pairs.join('; ') : undefined
}

function paragraphFormatChange(doc: PmNode, r: RevisionRange): string | undefined {
  const node = doc.nodeAt(r.from)
  const raw = node?.attrs?.pPrChange as string | null | undefined
  if (!node || !raw) return undefined
  let old: Record<string, unknown>
  try {
    const info = JSON.parse(raw)
    old = (info?.old ?? {}) as Record<string, unknown>
    if (old.format && typeof old.format === 'object') old = old.format as Record<string, unknown>
    if (info?.oldAlign !== undefined) old = { ...old, align: info.oldAlign }
  } catch {
    return undefined
  }
  const pairs = diffPairs(old, node.attrs as Record<string, unknown>, PARAGRAPH_FORMAT_FIELDS)
  return pairs.length ? pairs.join('; ') : undefined
}

/** Every pending tracked change with a positional id and enough context to pick it. */
export function listRevisionEntries(doc: PmNode): RevisionEntry[] {
  return collectRevisions(doc).map((r, i) => {
    const change =
      r.kind === 'rPrChange'
        ? runFormatChange(doc, r)
        : r.kind === 'pPrChange'
          ? paragraphFormatChange(doc, r)
          : undefined
    return {
      id: `r${i + 1}`,
      type: TYPE_OF[r.kind],
      kind: r.kind,
      // moves and some importers leave no author; one spelling so `author: "unknown"` selects them
      author: r.author || 'unknown',
      ...(r.date ? { date: r.date } : {}),
      blockIndex: blockIndexOfPos(doc, r.from),
      text: doc.textBetween(r.from, r.to, '\n', ' ').replace(/\s+/g, ' ').trim(),
      ...(change ? { change } : {}),
      from: r.from,
      to: r.to,
    }
  })
}

function tally(entries: RevisionEntry[], key: 'author' | 'type'): string {
  const counts = new Map<string, number>()
  for (const e of entries) {
    counts.set(e[key], (counts.get(e[key]) ?? 0) + 1)
  }
  return [...counts].map(([k, n]) => `${k} (${n})`).join(', ')
}

/** one line an agent can act on: what is pending, by whom, of which type */
export function describePending(entries: RevisionEntry[]): string {
  if (entries.length === 0) return 'the document has no pending tracked changes'
  return `${entries.length} pending: authors ${tally(entries, 'author')}; types ${tally(entries, 'type')}; ids ${entries[0]!.id}–${entries[entries.length - 1]!.id}`
}

export function validateSelector(
  input: Record<string, unknown>,
): RevisionSelector | { error: string } {
  const sel: RevisionSelector = {}
  if (input.all !== undefined) {
    if (input.all !== true) return { error: 'all must be true when given' }
    sel.all = true
  }
  if (input.ids !== undefined) {
    if (!Array.isArray(input.ids) || !input.ids.every((x) => typeof x === 'string' && x.trim()))
      return { error: 'ids must be an array of revision ids such as ["r1", "r4"]' }
    sel.ids = input.ids.map((x) => String(x).trim())
  }
  if (input.author !== undefined) {
    if (typeof input.author !== 'string' || !input.author.trim())
      return { error: 'author must be a non-empty string' }
    sel.author = input.author.trim()
  }
  if (input.type !== undefined) {
    if (!REVISION_TYPES.includes(input.type as RevisionType))
      return { error: `type must be one of ${REVISION_TYPES.join(', ')}` }
    sel.type = input.type as RevisionType
  }
  if (input.blockIndex !== undefined) {
    if (!Number.isInteger(input.blockIndex) || (input.blockIndex as number) < 0)
      return { error: 'blockIndex must be a non-negative integer' }
    sel.blockIndex = input.blockIndex as number
  }
  if (input.blockRange !== undefined) {
    const r = input.blockRange
    if (
      !Array.isArray(r) ||
      r.length !== 2 ||
      !r.every((x) => Number.isInteger(x) && x >= 0) ||
      r[0] > r[1]
    )
      return { error: 'blockRange must be [from, to] block indexes with from ≤ to' }
    sel.blockRange = [r[0], r[1]]
  }
  if (input.before !== undefined) {
    if (typeof input.before !== 'string' || Number.isNaN(Date.parse(input.before)))
      return { error: 'before must be an ISO date such as 2026-09-01 or 2026-09-01T12:00:00Z' }
    sel.before = input.before
  }
  if (Object.keys(sel).length === 0)
    return {
      error:
        'give a selector: all: true, ids: [...], or any of author / type / blockIndex / blockRange / before',
    }
  return sel
}

/** The pending changes a selector picks; an error names what exists when nothing matches. */
export function selectRevisions(
  doc: PmNode,
  sel: RevisionSelector,
): { entries: RevisionEntry[] } | { error: string } {
  const all = listRevisionEntries(doc)
  if (all.length === 0) return { error: 'the document has no pending tracked changes' }
  if (sel.ids) {
    const known = new Map(all.map((e) => [e.id, e]))
    const unknown = sel.ids.filter((id) => !known.has(id))
    if (unknown.length)
      return {
        error: `unknown revision id(s) ${unknown.join(', ')}; valid ids are r1–r${all.length} (call read_revisions after every edit, ids are positional)`,
      }
  }
  const beforeMs = sel.before ? Date.parse(sel.before) : null
  const entries = all.filter((e) => {
    if (sel.ids && !sel.ids.includes(e.id)) return false
    if (sel.author !== undefined && e.author.toLowerCase() !== sel.author.toLowerCase())
      return false
    if (sel.type !== undefined && e.type !== sel.type) return false
    if (sel.blockIndex !== undefined && e.blockIndex !== sel.blockIndex) return false
    if (sel.blockRange && (e.blockIndex < sel.blockRange[0] || e.blockIndex > sel.blockRange[1]))
      return false
    if (beforeMs !== null && (!e.date || Date.parse(e.date) >= beforeMs)) return false
    return true
  })
  if (entries.length === 0) {
    const asked = Object.entries(sel)
      .filter(([k]) => k !== 'all')
      .map(([k, v]) => `${k}=${JSON.stringify(v)}`)
      .join(', ')
    return {
      error: `no tracked change matches ${asked || 'the selector'}; ${describePending(all)}`,
    }
  }
  return { entries }
}

/** accept or reject the selected changes in one transaction, last position first */
export function applyRevisionSelection(
  editor: Editor,
  sel: RevisionSelector,
  mode: 'accept' | 'reject',
): { entries: RevisionEntry[] } | { error: string } {
  const picked = selectRevisions(editor.state.doc, sel)
  if ('error' in picked) return picked
  applyRevisions(
    editor,
    picked.entries.map((e) => ({
      from: e.from,
      to: e.to,
      kind: e.kind,
      author: e.author,
      ...(e.date ? { date: e.date } : {}),
    })),
    mode,
  )
  return picked
}
