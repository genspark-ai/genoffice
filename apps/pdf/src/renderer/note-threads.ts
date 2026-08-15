import type { DrawingInput } from '../shared/ipc'

/** pdf.js AnnotationType.TEXT (sticky-note comments) */
export const PDFJS_ANNOT_TEXT = 1

export type NoteInput = Extract<DrawingInput, { kind: 'note' }>

/** A note (Text) annotation already saved in the file (read via pdf.js getAnnotations) */
export interface SavedNoteAnnot {
  pageIndex: number
  /** PDF object number (pdf.js id "123R" → 123) */
  objNum: number
  type: 'note'
  rect: [number, number, number, number]
  /** rgb 0-1; null when the annotation has no /C */
  color: [number, number, number] | null
  /** /T; '' when absent */
  author: string
  contents: string
  /** /CreationDate (falling back to /M) as epoch ms; null when absent/unparseable */
  timeMs: number | null
  /** Parent note object number (/IRT); null = thread root */
  inReplyTo: number | null
}

/** One comment in a thread: either saved in the file or pending in this session */
export interface NoteThreadItem {
  /** 'S<objNum>' for saved, 'P<drawing id>' for pending */
  key: string
  saved: SavedNoteAnnot | null
  /** LocalDrawing id (== note input localId) for pending items */
  pendingId: string | null
  author: string
  contents: string
  timeMs: number | null
  color: [number, number, number] | null
  at: [number, number]
  replies: NoteThreadItem[]
}

export const savedNoteKey = (objNum: number): string => `S${objNum}`
export const pendingNoteKey = (id: string): string => `P${id}`

/** PDF date string (D:YYYYMMDDHHmmSS±hh'mm') → epoch ms; null when unparseable */
export function parsePdfDate(s: string | null | undefined): number | null {
  if (!s) return null
  const m =
    /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?(?:([+\-Z])(?:(\d{2})'?(\d{2})?'?)?)?/.exec(
      s.trim(),
    )
  if (!m) return null
  const [, y, mo, d, h, mi, se, tzSign, tzH, tzM] = m
  const year = Number(y)
  const month = Number(mo ?? '1') - 1
  const day = Number(d ?? '1')
  const hour = Number(h ?? '0')
  const min = Number(mi ?? '0')
  const sec = Number(se ?? '0')
  if (month < 0 || month > 11 || day < 1 || day > 31 || hour > 23 || min > 59 || sec > 59)
    return null
  let ms = Date.UTC(year, month, day, hour, min, sec)
  if (tzSign === '+' || tzSign === '-') {
    const offMin = Number(tzH ?? '0') * 60 + Number(tzM ?? '0')
    ms += (tzSign === '+' ? -1 : 1) * offMin * 60_000
  } else if (tzSign !== 'Z') {
    // No timezone: treat as local time
    ms = new Date(year, month, day, hour, min, sec).getTime()
  }
  return Number.isFinite(ms) ? ms : null
}

/** Fields of a pdf.js getAnnotations() item this module reads */
export interface PdfJsAnnotData {
  id: string
  annotationType: number
  rect: number[]
  color?: { [i: number]: number; length: number } | null
  titleObj?: { str: string } | null
  contentsObj?: { str: string } | null
  creationDate?: string | null
  modificationDate?: string | null
  inReplyTo?: string | null
  replyType?: string
  hidden?: boolean
}

/**
 * pdf.js annotation → SavedNoteAnnot. Null for non-Text annots, hidden ones, annots
 * not backed by an object ref (unaddressable for delete/reply), and /RT Group members
 * (grouping, not a reply — pdf.js already mirrors the parent's contents onto them).
 */
export function toSavedNote(a: PdfJsAnnotData, pageIndex: number): SavedNoteAnnot | null {
  if (a.annotationType !== PDFJS_ANNOT_TEXT || a.hidden) return null
  if (a.replyType === 'Group') return null
  const objNum = /^(\d+)R$/.exec(a.id)
  if (!objNum || !Array.isArray(a.rect) || a.rect.length !== 4) return null
  const parent = a.inReplyTo ? /^(\d+)R$/.exec(a.inReplyTo) : null
  const color =
    a.color && a.color.length >= 3
      ? ([a.color[0]! / 255, a.color[1]! / 255, a.color[2]! / 255] as [number, number, number])
      : null
  return {
    pageIndex,
    objNum: Number(objNum[1]),
    type: 'note',
    rect: [a.rect[0]!, a.rect[1]!, a.rect[2]!, a.rect[3]!],
    color,
    author: a.titleObj?.str ?? '',
    contents: a.contentsObj?.str ?? '',
    timeMs: parsePdfDate(a.creationDate) ?? parsePdfDate(a.modificationDate),
    inReplyTo: parent ? Number(parent[1]) : null,
  }
}

const byTime = (a: NoteThreadItem, b: NoteThreadItem): number => {
  if (a.timeMs === null && b.timeMs === null) return 0
  if (a.timeMs === null) return 1
  if (b.timeMs === null) return -1
  return a.timeMs - b.timeMs
}

/**
 * Assemble saved + pending notes of one page into threads. Roots are notes without a
 * resolvable parent — orphaned replies (parent deleted by another tool) are promoted
 * to roots rather than dropped, and /IRT cycles in malformed files can't hide notes.
 * Replies are sorted by creation time (unknown times last); root order is input order.
 */
export function buildNoteThreads(
  saved: SavedNoteAnnot[],
  pending: { id: string; input: NoteInput }[],
): NoteThreadItem[] {
  const savedItems = new Map<number, NoteThreadItem>()
  for (const s of saved) {
    savedItems.set(s.objNum, {
      key: savedNoteKey(s.objNum),
      saved: s,
      pendingId: null,
      author: s.author,
      contents: s.contents,
      timeMs: s.timeMs,
      color: s.color,
      at: [s.rect[0], Math.max(s.rect[1], s.rect[3])],
      replies: [],
    })
  }
  const roots: NoteThreadItem[] = []
  const attached = new Set<number>()
  for (const s of saved) {
    const item = savedItems.get(s.objNum)!
    const parent = s.inReplyTo !== null ? savedItems.get(s.inReplyTo) : undefined
    if (parent && parent !== item) {
      parent.replies.push(item)
      attached.add(s.objNum)
    } else {
      roots.push(item)
    }
  }
  // Cycle guard: attached items unreachable from any root are promoted to roots,
  // detaching them from their (cyclic) parent so no item appears twice
  const reachable = new Set<string>()
  const visit = (item: NoteThreadItem) => {
    if (reachable.has(item.key)) return
    reachable.add(item.key)
    for (const r of item.replies) visit(r)
  }
  for (const r of roots) visit(r)
  for (const objNum of attached) {
    const item = savedItems.get(objNum)!
    if (reachable.has(item.key)) continue
    for (const other of savedItems.values()) {
      if (other !== item) other.replies = other.replies.filter((r) => r !== item)
    }
    roots.push(item)
    visit(item)
  }

  const pendingItems = new Map<string, NoteThreadItem>()
  for (const p of pending) {
    pendingItems.set(p.id, {
      key: pendingNoteKey(p.id),
      saved: null,
      pendingId: p.id,
      author: p.input.author ?? '',
      contents: p.input.contents,
      timeMs: p.input.createdMs ?? null,
      color: p.input.color,
      at: p.input.at,
      replies: [],
    })
  }
  for (const p of pending) {
    const item = pendingItems.get(p.id)!
    const parent = p.input.replyToLocalId
      ? pendingItems.get(p.input.replyToLocalId)
      : p.input.replyToSaved
        ? savedItems.get(p.input.replyToSaved.objNum)
        : undefined
    if (parent && parent !== item) parent.replies.push(item)
    else roots.push(item)
  }

  const sortReplies = (item: NoteThreadItem) => {
    item.replies.sort(byTime)
    for (const r of item.replies) sortReplies(r)
  }
  for (const r of roots) sortReplies(r)
  return roots
}

/** DFS flatten of a thread (root first), with the depth of each item */
export function flattenThread(root: NoteThreadItem): { item: NoteThreadItem; depth: number }[] {
  const out: { item: NoteThreadItem; depth: number }[] = []
  const walk = (item: NoteThreadItem, depth: number) => {
    out.push({ item, depth })
    for (const r of item.replies) walk(r, depth + 1)
  }
  walk(root, 0)
  return out
}

/** An item plus everything under it — what a cascade delete must remove */
export function threadSubtree(item: NoteThreadItem): {
  saved: SavedNoteAnnot[]
  pendingIds: string[]
} {
  const saved: SavedNoteAnnot[] = []
  const pendingIds: string[] = []
  for (const { item: node } of flattenThread(item)) {
    if (node.saved) saved.push(node.saved)
    if (node.pendingId !== null) pendingIds.push(node.pendingId)
  }
  return { saved, pendingIds }
}

/** Find the root whose subtree contains `key`; null when the key is gone */
export function findThreadRoot(roots: NoteThreadItem[], key: string): NoteThreadItem | null {
  for (const root of roots) {
    if (flattenThread(root).some(({ item }) => item.key === key)) return root
  }
  return null
}
