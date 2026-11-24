/**
 * Tools > Combine Documents: the differences between the open document and a
 * revised one, written into the open document as ordinary tracked changes.
 *
 * Compared: body paragraphs (aligned paragraph by paragraph, then word by
 * word inside a changed paragraph) and tables, row by row and cell by cell
 * (column by column when the table gained or lost one); the formatting of
 * paragraphs and text; inline pictures and equations. Not compared, and
 * counted in `skipped` when they differ: a table with a column added or
 * removed across a merged cell, block-level objects, notes, fields.
 */
import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import type { Mark, Node as PmNode, Schema } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { TableMap } from '@tiptap/pm/tables'
import { columnWidthsPx, pctOf, reflowColumns } from '../ai/table-ops'
import {
  alignTexts,
  diffText,
  diffTrackedText,
  type CompareGranularity,
  type TextAlignment,
  type TextCompareIgnore,
  type TrackedState,
} from '@genoffice/docx-engine'
import type { PmNode as JsonNode } from './convert'
import {
  paragraphSnapshotOf,
  runFormatSnapshotOf,
  TEXT_STYLE_FIELDS,
  TRACK_IGNORE,
  withRowRecord,
  withTrackMarker,
  type ParagraphSnapshot,
} from './revisions'

export interface CombineOptions {
  author: string
  /** ISO timestamp recorded on every change */
  date: string
  /** a paragraph style the open document defines; an inserted paragraph keeps only those */
  hasStyle?: (styleId: string) => boolean
  /** default "word" */
  granularity?: CompareGranularity
  /** differences that are not changes */
  ignore?: TextCompareIgnore
  /** false leaves every table as it is */
  tables?: boolean
  /** paragraphs and text that differ in their formatting alone are tracked formatting changes */
  formatting?: boolean
  /**
   * The open document is the revised one and the other the original: what only
   * the open document has is marked inserted, what only the other has comes in
   * marked deleted.
   */
  openIsRevised?: boolean
  /**
   * Combine: tracked changes the other document already holds come in under
   * their own author and date, and the open document's own are left as they
   * are. Off, the other document is read with its changes accepted.
   */
  keepRevisions?: boolean
}

export interface CombineSummary {
  /** inserted text regions and inserted paragraphs */
  insertions: number
  /** deleted text regions and deleted paragraphs */
  deletions: number
  /** text and paragraphs whose formatting alone changed */
  formatChanges: number
  /** tables, objects and inline objects that differ and were left as they are */
  skipped: number
  /** false when the two documents share no paragraph: nothing is written then */
  related: boolean
  /** keepRevisions: tracked changes of the other document that came in under their own author */
  carried?: number
}

const PARAGRAPH_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])
const TABLE_TYPES = new Set(['docTable', 'docTableRow', 'docTableCell', 'docTableHeader'])
const ATOM = '\uFFFC'
const ATOM_FIRST = 0xe000
const ATOM_LAST = 0xefff
const isAtom = (ch: string | undefined): boolean =>
  ch === ATOM ||
  (ch !== undefined && ch.charCodeAt(0) >= ATOM_FIRST && ch.charCodeAt(0) <= ATOM_LAST)

/**
 * One character for each distinct picture (by its bytes) and equation (by its
 * OMML) of a combine, so the same object on both sides is kept and a different
 * one is a deletion and an insertion. Any other inline object reads as ATOM.
 */
let atomChars = new Map<string, string>()
function atomChar(type: string, attrs: Record<string, unknown> | null | undefined): string {
  const identity =
    type === 'docInlineMath'
      ? `math ${String(attrs?.omml ?? '')}`
      : type === 'docInlineImage'
        ? `picture ${String(attrs?.dataUrl ?? '')}`
        : null
  if (identity === null) return ATOM
  let ch = atomChars.get(identity)
  if (!ch) {
    const code = ATOM_FIRST + atomChars.size
    ch = code > ATOM_LAST ? ATOM : String.fromCharCode(code)
    atomChars.set(identity, ch)
  }
  return ch
}
const SIMPLE_PICTURE = /^data:image\/(?:png|jpeg|gif);base64,/
const CARRIED_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])
const TOGGLE_MARKS = ['bold', 'italic', 'underline', 'strike'] as const
/** attrs of a revised paragraph that point into its own package or record its own history */
const FOREIGN_PARAGRAPH_ATTRS = [
  'docxIndex',
  'bookmarks',
  'hiddenBookmarks',
  'commentStarts',
  'commentEnds',
  'sdtShell',
  'moveRevision',
  'pPrChange',
  'paraMarkDel',
  'paraMarkIns',
  'paraMarkRPrChange',
  'numberingChange',
  'sectPrChange',
  'blockRevision',
  'ruleOverlays',
]

const breakChar = (attrs: Record<string, unknown> | null | undefined): string =>
  attrs?.pageBreak ? '\f' : attrs?.colBreak ? '\v' : attrs?.wrapBreak ? '\u001e' : '\n'

const hasMark = (marks: ReadonlyArray<{ type: string | { name: string } }>, name: string) =>
  marks.some((m) => (typeof m.type === 'string' ? m.type : m.type.name) === name)

/** one compared character of the open document's paragraph and where it sits */
interface LeftText {
  text: string
  /** document position of each character */
  at: number[]
  /** document position right after each character */
  end: number[]
  /** the tracked change the character sits in */
  tracked: TrackedState[]
  /** the inline node each character belongs to */
  node: PmNode[]
}

const trackedState = (marks: ReadonlyArray<{ type: string | { name: string } }>): TrackedState =>
  hasMark(marks, 'ins') ? 'ins' : hasMark(marks, 'del') ? 'del' : ''

/** the paragraph's text with its pending deletions left out, or with `markup` as it stands */
function leftText(para: PmNode, pos: number, markup = false): LeftText {
  let text = ''
  const at: number[] = []
  const end: number[] = []
  const tracked: TrackedState[] = []
  const node: PmNode[] = []
  para.forEach((child, offset) => {
    const start = pos + 1 + offset
    if (!markup && hasMark(child.marks, 'del')) return
    const flag = trackedState(child.marks)
    if (child.isText) {
      for (let i = 0; i < child.text!.length; i++) {
        at.push(start + i)
        end.push(start + i + 1)
        tracked.push(flag)
        node.push(child)
      }
      text += child.text
    } else {
      at.push(start)
      end.push(start + child.nodeSize)
      tracked.push(flag)
      node.push(child)
      text +=
        child.type.name === 'hardBreak' && !child.attrs.mathBreak
          ? breakChar(child.attrs)
          : atomChar(child.type.name, child.attrs)
    }
  })
  return { text, at, end, tracked, node }
}

interface RightText {
  text: string
  /** the inline node each character comes from */
  source: JsonNode[]
}

function rightText(para: JsonNode, markup = false): RightText {
  let text = ''
  const source: JsonNode[] = []
  for (const child of para.content ?? []) {
    if (!markup && hasMark(child.marks ?? [], 'del')) continue
    const piece =
      child.type === 'text'
        ? (child.text ?? '')
        : child.type === 'hardBreak' && !child.attrs?.mathBreak
          ? breakChar(child.attrs)
          : atomChar(child.type, child.attrs)
    for (let i = 0; i < piece.length; i++) source.push(child)
    text += piece
  }
  return { text, source }
}

/** the text every paragraph under a table contributes, cell after cell */
const TABLE_TAG = '\u0001table '
function tableText(paragraphs: string[]): string {
  return TABLE_TAG + paragraphs.join('\n')
}

function leftParagraphs(node: PmNode, pos: number): Array<{ node: PmNode; pos: number }> {
  const out: Array<{ node: PmNode; pos: number }> = []
  node.descendants((child, offset) => {
    if (!PARAGRAPH_TYPES.has(child.type.name)) return true
    out.push({ node: child, pos: pos + 1 + offset })
    return false
  })
  return out
}

function rightParagraphs(node: JsonNode): JsonNode[] {
  const out: JsonNode[] = []
  const walk = (n: JsonNode) => {
    if (PARAGRAPH_TYPES.has(n.type)) out.push(n)
    else for (const child of n.content ?? []) walk(child)
  }
  walk(node)
  return out
}

/** rows, cells and the kind of every block in each cell, so equal shapes pair paragraph by paragraph */
function leftShape(node: PmNode): string {
  if (PARAGRAPH_TYPES.has(node.type.name)) return 'p'
  const parts: string[] = []
  node.forEach((child) => parts.push(leftShape(child)))
  return `${node.type.name}(${parts.join(',')})`
}

function rightShape(node: JsonNode): string {
  if (PARAGRAPH_TYPES.has(node.type)) return 'p'
  return `${node.type}(${(node.content ?? []).map(rightShape).join(',')})`
}

type Kind = 'text' | 'table' | 'object'

interface Unit {
  kind: Kind
  text: string
}

const objectText = (type: string, attrs: Record<string, unknown> | null | undefined) =>
  `\u0001object ${type} ${String(attrs?.blockType ?? '')} ${String(attrs?.previewText ?? '')} ${String(attrs?.imageDataUrl ?? '')}`

interface Op {
  pos: number
  run: (tr: Transaction) => void
}

/**
 * Write the differences into `tr` (positions of the document it starts from).
 * Nothing is written when the documents are unrelated.
 */
export function combineIntoTransaction(
  tr: Transaction,
  revised: JsonNode[],
  options: CombineOptions,
): CombineSummary {
  atomChars = new Map()
  const doc = tr.doc
  const schema: Schema = doc.type.schema
  const record = { author: options.author, date: options.date }
  const flipped = options.openIsRevised === true
  // what comes in from the other document, and what only this document has
  const comingKind = flipped ? 'del' : 'ins'
  const goingKind = flipped ? 'ins' : 'del'
  const comingMark = schema.marks[comingKind]!.create(record)
  const goingMark = schema.marks[goingKind]!.create(record)
  const comingMarkAttr = flipped ? 'paraMarkDel' : 'paraMarkIns'
  const goingMarkAttr = flipped ? 'paraMarkIns' : 'paraMarkDel'
  const ignore = options.ignore ?? {}
  const tables = options.tables !== false
  const keep = options.keepRevisions === true
  const markRecord = JSON.stringify(record)
  const formatting = options.formatting === true
  const summary: CombineSummary = {
    insertions: 0,
    deletions: 0,
    formatChanges: 0,
    skipped: 0,
    related: true,
  }
  const ops: Op[] = []

  const left: Array<Unit & { node: PmNode; pos: number }> = []
  doc.forEach((node, pos) => {
    const name = node.type.name
    // a block already pending deletion is not part of what the document says
    if ((node.attrs.blockRevision as { kind?: string } | null)?.kind === 'del') return
    if (PARAGRAPH_TYPES.has(name))
      left.push({ kind: 'text', text: leftText(node, pos, keep).text, node, pos })
    else if (name === 'docTable')
      left.push({
        kind: 'table',
        text: tableText(leftParagraphs(node, pos).map((p) => leftText(p.node, p.pos, keep).text)),
        node,
        pos,
      })
    else left.push({ kind: 'object', text: objectText(name, node.attrs), node, pos })
  })
  const right: Array<Unit & { json: JsonNode }> = []
  for (const json of revised) {
    if ((json.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') continue
    if (PARAGRAPH_TYPES.has(json.type))
      right.push({ kind: 'text', text: rightText(json, keep).text, json })
    else if (json.type === 'docTable')
      right.push({
        kind: 'table',
        text: tableText(rightParagraphs(json).map((p) => rightText(p, keep).text)),
        json,
      })
    else right.push({ kind: 'object', text: objectText(json.type, json.attrs), json })
  }

  // between two blocks both documents have, what goes away reads before what comes in
  const aligned: TextAlignment[] = []
  let incoming: TextAlignment[] = []
  for (const entry of alignTexts(
    left.map((u) => u.text),
    right.map((u) => u.text),
    ignore,
  )) {
    if (entry.kind === 'added') incoming.push(entry)
    else if (entry.kind === 'removed') aligned.push(entry)
    else {
      aligned.push(...incoming, entry)
      incoming = []
    }
  }
  aligned.push(...incoming)
  const hasWords = (u: Unit) =>
    u.kind !== 'object' && /\S/.test(u.kind === 'table' ? u.text.slice(TABLE_TAG.length) : u.text)
  const hasText = (units: Unit[]) => units.some(hasWords)
  const shared = aligned.some(
    (entry) =>
      entry.kind === 'pair' &&
      left[entry.left]!.kind === right[entry.right]!.kind &&
      left[entry.left]!.kind !== 'object' &&
      hasWords(left[entry.left]!),
  )
  if (hasText(left) && hasText(right) && !shared) return { ...summary, related: false }

  /** marks an inserted piece of text keeps from the revised run; links, fields and comments stay behind */
  const carried = (child: JsonNode): Mark[] => {
    const marks: Mark[] = []
    for (const mark of child.marks ?? []) {
      if (!CARRIED_MARKS.has(mark.type) || !schema.marks[mark.type]) continue
      const attrs =
        mark.type === 'docTextStyle' ? { ...mark.attrs, rawRPr: null, styleId: null } : mark.attrs
      marks.push(schema.marks[mark.type]!.create(attrs))
    }
    return marks
  }

  // a move of the other document comes in under a name this document does not use yet
  const moveNames = new Set<string>()
  doc.descendants((node) => {
    for (const mark of node.marks) if (mark.attrs.move) moveNames.add(String(mark.attrs.move))
    return true
  })
  const carriedMoves = new Map<string, string>()
  const carriedMove = (name: unknown): string | null => {
    if (typeof name !== 'string' || !name) return null
    let mine = carriedMoves.get(name)
    for (let n = 1; !mine; n++) {
      if (moveNames.has(`move${n}`)) continue
      mine = `move${n}`
      moveNames.add(mine)
      carriedMoves.set(name, mine)
    }
    return mine
  }

  /** the tracked change a piece of the other document already sits in, under an id minted on save */
  const carriedChanges = new Set<string>()
  const ownChange = (child: JsonNode): Mark[] =>
    (child.marks ?? [])
      .filter((mark) => mark.type === 'ins' || mark.type === 'del')
      .map((mark) => {
        carriedChanges.add(
          JSON.stringify([mark.type, mark.attrs?.author, mark.attrs?.date, mark.attrs?.id]),
        )
        return schema.marks[mark.type]!.create({
          ...mark.attrs,
          id: null,
          move: carriedMove(mark.attrs?.move),
        })
      })
  const comingMarks = (child: JsonNode): Mark[] => {
    const own = keep ? ownChange(child) : []
    return own.length ? own : [comingMark]
  }

  /**
   * The revised characters [from, to) as nodes of this schema, each marked
   * inserted (or, kept, in the change it already sits in); objects stay behind.
   */
  const insertedInline = (text: RightText, from: number, to: number): PmNode[] => {
    const nodes: PmNode[] = []
    let i = from
    while (i < to) {
      const child = text.source[i]!
      let end = i + 1
      while (end < to && text.source[end] === child) end++
      if (child.type === 'text') {
        nodes.push(schema.text(text.text.slice(i, end), [...carried(child), ...comingMarks(child)]))
      } else if (child.type === 'hardBreak') {
        nodes.push(schema.nodes.hardBreak!.create(child.attrs ?? null, null, comingMarks(child)))
      } else if (child.type === 'docInlineMath' && schema.nodes.docInlineMath) {
        nodes.push(schema.nodes.docInlineMath.create(child.attrs ?? null, null, comingMarks(child)))
      } else if (
        child.type === 'docInlineImage' &&
        schema.nodes.docInlineImage &&
        SIMPLE_PICTURE.test(String(child.attrs?.dataUrl ?? '')) &&
        Number(child.attrs?.widthPx) > 0 &&
        Number(child.attrs?.heightPx) > 0
      ) {
        // the bytes and the size come over; the drawing is written anew for this package on save
        const { dataUrl, widthPx, heightPx } = child.attrs!
        nodes.push(
          schema.nodes.docInlineImage.create(
            { dataUrl, widthPx, heightPx },
            null,
            comingMarks(child),
          ),
        )
      } else summary.skipped += end - i
      i = end
    }
    return nodes
  }

  /** strike [from, to) of the compared text; a picture or an equation goes with it, any other object stays */
  const strike = (text: LeftText, from: number, to: number): boolean => {
    let struck = false
    let i = from
    while (i < to) {
      if (text.text[i] === ATOM) {
        summary.skipped++
        i++
        continue
      }
      let end = i + 1
      while (end < to && text.text[end] !== ATOM && text.at[end] === text.end[end - 1]) end++
      const a = text.at[i]!
      const b = text.end[end - 1]!
      ops.push({ pos: a, run: (t) => t.addMark(a, b, goingMark) })
      struck = true
      i = end
    }
    return struck
  }

  const otherTracked = (text: RightText): TrackedState[] =>
    text.source.map((child) => trackedState(child.marks ?? []))

  const listNumIds = new Set<string>()
  doc.descendants((node) => {
    if (node.type.name === 'docListItem' && node.attrs.numId != null)
      listNumIds.add(String(node.attrs.numId))
    return !node.isTextblock
  })
  const knownStyle = (styleId: unknown): boolean =>
    !styleId || !options.hasStyle || options.hasStyle(String(styleId))

  /** another paragraph's formatting as this document can hold it: its style and list only when they exist here */
  const heldHere = (theirs: ParagraphSnapshot, own: ParagraphSnapshot): ParagraphSnapshot => {
    const styleId = knownStyle(theirs.styleId) ? theirs.styleId : own.styleId
    if (theirs.type !== 'docListItem' || listNumIds.has(String(theirs.numId)))
      return { ...theirs, styleId }
    const { type, level, kind, numId, ilvl } = own
    return { ...theirs, styleId, type, level, kind, numId, ilvl }
  }
  const listItem = (snapshot: ParagraphSnapshot) => snapshot.type === 'docListItem'
  // a list is compared by its kind and level: its id belongs to one package
  const formatKey = ({ format, ...rest }: ParagraphSnapshot): string =>
    JSON.stringify({
      ...rest,
      numId: null,
      // an attr the node defaults to off reads the same as one a record leaves out
      format: {
        ...format,
        pageBreakBefore: format.pageBreakBefore || null,
        bidi: format.bidi || null,
      },
    })
  const parsedRecord = (raw: unknown): Record<string, unknown> | null => {
    if (typeof raw !== 'string') return null
    try {
      return JSON.parse(raw) as Record<string, unknown>
    } catch {
      return null
    }
  }
  /** what a w:pPrChange record says the paragraph was, in the snapshot's shape */
  const recordedSnapshot = (record: Record<string, unknown>): ParagraphSnapshot => {
    const old = (record.old ?? {}) as Record<string, unknown>
    const flat = old.format && typeof old.format === 'object' ? { ...old, ...old.format } : old
    return paragraphSnapshotOf(typeof old.type === 'string' ? old.type : 'docParagraph', flat)
  }

  /**
   * A paragraph both documents have, formatted differently: this document
   * takes the revised formatting and records what the original had. A change
   * the other document already tracks on it comes in under its own author.
   */
  const formatParagraph = (node: PmNode, pos: number, json: JsonNode): void => {
    if (!formatting || (keep && node.attrs.pPrChange)) return
    const own = paragraphSnapshotOf(node.type.name, node.attrs)
    const theirs = heldHere(paragraphSnapshotOf(json.type, json.attrs ?? {}), own)
    const theirRecord = keep ? parsedRecord(json.attrs?.pPrChange) : null
    const same = formatKey(own) === formatKey(theirs)
    if (same && !theirRecord) return
    const tracked =
      theirRecord &&
      (same || formatKey(heldHere(recordedSnapshot(theirRecord), own)) === formatKey(own))
    const credit = tracked
      ? { author: theirRecord.author ?? '', date: theirRecord.date ?? null }
      : record
    const old = same
      ? heldHere(recordedSnapshot(theirRecord!), own)
      : flipped && !tracked
        ? theirs
        : own
    const next = flipped && !tracked ? own : theirs
    if (tracked) carriedChanges.add(JSON.stringify(['pPrChange', credit.author, credit.date, pos]))
    else summary.formatChanges++
    ops.push({
      pos,
      run: (t) => {
        const live = t.doc.nodeAt(pos)!
        t.setNodeMarkup(pos, schema.nodes[next.type] ?? live.type, {
          ...live.attrs,
          ...next.format,
          styleId: next.styleId,
          level: next.level,
          kind: next.kind,
          numId: listItem(next) && listItem(own) && next.kind === own.kind ? own.numId : next.numId,
          ilvl: next.ilvl,
          pPrChange: JSON.stringify({ ...credit, old }),
        })
      },
    })
  }

  const runFormat = (marks: JsonNode['marks'] | PmNode['marks']): Record<string, unknown> => {
    const snapshot = runFormatSnapshotOf(marks ?? [])
    if (!knownStyle(snapshot.styleId)) delete snapshot.styleId
    return snapshot
  }

  /**
   * Text both documents have, formatted differently, over the characters
   * [from, from + length) of this paragraph and the ones at `other` of the
   * revised one: a tracked formatting change that records the original's.
   */
  const formatRuns = (l: LeftText, r: RightText, from: number, other: number, length: number) => {
    if (!formatting || !schema.marks.rprChange) return
    let i = 0
    while (i < length) {
      const mine = l.node[from + i]!
      const child = r.source[other + i]!
      let end = i + 1
      while (end < length && l.node[from + end] === mine && r.source[other + end] === child) end++
      const a = l.at[from + i]!
      const b = l.end[from + end - 1]!
      i = end
      if (!mine.isText || child.type !== 'text') continue
      const tracksFormat = mine.marks.some((m) => m.type.name === 'rprChange')
      if (keep && tracksFormat) continue
      const own = runFormat(mine.marks)
      const theirs = runFormat(child.marks)
      const theirRecord = keep
        ? (child.marks ?? []).find((m) => m.type === 'rprChange')?.attrs
        : undefined
      const same = JSON.stringify(own) === JSON.stringify(theirs)
      if (same && !theirRecord) continue
      const recorded = (theirRecord?.old ?? {}) as Record<string, unknown>
      const theirOld = runFormat([
        ...TOGGLE_MARKS.filter((name) => recorded[name]).map((name) => ({ type: name })),
        { type: 'docTextStyle', attrs: recorded },
      ])
      const tracked = !!theirRecord && (same || JSON.stringify(theirOld) === JSON.stringify(own))
      const credit = tracked
        ? { author: theirRecord.author ?? '', date: theirRecord.date ?? null }
        : record
      const takeTheirs = !same && !(flipped && !tracked)
      const old = same ? theirOld : takeTheirs ? own : theirs
      if (tracked)
        carriedChanges.add(
          JSON.stringify(['rPrChange', credit.author, credit.date, theirRecord.id]),
        )
      else summary.formatChanges++
      const change = schema.marks.rprChange.create({ ...credit, old })
      ops.push({
        pos: a,
        run: (t) => {
          if (takeTheirs) {
            for (const name of TOGGLE_MARKS) {
              const type = schema.marks[name]
              if (!type) continue
              if (theirs[name]) t.addMark(a, b, type.create())
              else t.removeMark(a, b, type)
            }
            const style = schema.marks.docTextStyle
            const mineStyle = mine.marks.find((m) => m.type.name === 'docTextStyle')
            const fields = Object.fromEntries(
              TEXT_STYLE_FIELDS.map((field) => [field, theirs[field] ?? null]),
            )
            if (style && (mineStyle || Object.values(fields).some((v) => v !== null)))
              t.addMark(a, b, style.create({ ...mineStyle?.attrs, ...fields }))
          }
          t.addMark(a, b, change)
        },
      })
    }
  }

  /** both read with their markup: each side's own changes stay, only plain differences are new */
  const mergeParagraph = (node: PmNode, pos: number, json: JsonNode): void => {
    const l = leftText(node, pos, true)
    const r = rightText(json, true)
    const theirs = otherTracked(r)
    formatParagraph(node, pos, json)
    if (l.text === r.text && !theirs.some((flag, i) => flag && !l.tracked[i])) {
      formatRuns(l, r, 0, 0, l.text.length)
      return
    }
    let consumed = 0
    const chunks = diffTrackedText(
      l.text,
      r.text,
      { left: l.tracked, right: theirs, carryFrom: 'right' },
      options.granularity,
      ignore,
    )
    for (const chunk of chunks) {
      if (chunk.kind === 'keep') {
        // text the other document tracks and this one has plain takes that change on
        for (let i = 0; chunk.carry && i < chunk.left![1] - chunk.left![0]; i++) {
          const mine = chunk.left![0] + i
          const child = r.source[chunk.right![0] + i]!
          if (l.tracked[mine] || isAtom(l.text[mine])) continue
          const a = l.at[mine]!
          const b = l.end[mine]!
          for (const mark of ownChange(child))
            ops.push({ pos: a, run: (t) => t.addMark(a, b, mark) })
        }
        if (chunk.left![1] - chunk.left![0] === chunk.right![1] - chunk.right![0])
          formatRuns(l, r, chunk.left![0], chunk.right![0], chunk.left![1] - chunk.left![0])
        consumed = chunk.left![1]
      } else if (chunk.kind === 'del') {
        let struck = false
        let i = chunk.left![0]
        while (i < chunk.left![1]) {
          let end = i
          while (end < chunk.left![1] && l.tracked[end] === l.tracked[i]) end++
          if (!l.tracked[i] && strike(l, i, end)) struck = true
          i = end
        }
        if (struck) summary.deletions++
        consumed = chunk.left![1]
      } else {
        const nodes = insertedInline(r, chunk.right![0], chunk.right![1])
        if (nodes.length === 0) continue
        const at = consumed === 0 ? pos + 1 : l.end[consumed - 1]!
        ops.push({ pos: at, run: (t) => t.insert(at, nodes) })
        if (theirs.slice(chunk.right![0], chunk.right![1]).some((flag) => !flag))
          summary.insertions++
      }
    }
  }

  const diffParagraph = (node: PmNode, pos: number, json: JsonNode): void => {
    if (keep) return mergeParagraph(node, pos, json)
    const l = leftText(node, pos)
    const r = rightText(json)
    formatParagraph(node, pos, json)
    if (l.text === r.text) {
      formatRuns(l, r, 0, 0, l.text.length)
      return
    }
    let consumed = 0
    for (const chunk of diffText(l.text, r.text, options.granularity, ignore)) {
      if (chunk.kind === 'keep') {
        if (chunk.left![1] - chunk.left![0] === chunk.right![1] - chunk.right![0])
          formatRuns(l, r, chunk.left![0], chunk.right![0], chunk.left![1] - chunk.left![0])
        consumed = chunk.left![1]
      } else if (chunk.kind === 'del') {
        if (strike(l, chunk.left![0], chunk.left![1])) summary.deletions++
        consumed = chunk.left![1]
      } else {
        const nodes = insertedInline(r, chunk.right![0], chunk.right![1])
        if (nodes.length === 0) continue
        const at = consumed === 0 ? pos + 1 : l.end[consumed - 1]!
        ops.push({ pos: at, run: (t) => t.insert(at, nodes) })
        summary.insertions++
      }
    }
  }

  const removeBlock = (unit: (typeof left)[number]): void => {
    const { node, pos } = unit
    // a paragraph that is itself a pending insertion keeps that record: its text
    // is struck on top of it, as inserted text inside a paragraph is
    const pendingInsertion = !!node.attrs.blockRevision
    if (keep && unit.kind === 'text') {
      // a paragraph this document inserted or deleted as a tracked change stays that change
      const own = leftText(node, pos, true)
      const marked = !!(node.attrs.paraMarkIns || node.attrs.paraMarkDel)
      if (own.text === '' ? marked : own.tracked.every(Boolean)) return
    }
    if (unit.kind === 'table') return removeTable(unit)
    const empty = leftText(node, pos).text === ''
    if (unit.kind !== 'text' || (pendingInsertion && empty)) {
      summary.skipped++
      return
    }
    if (empty) {
      if (node.attrs[goingMarkAttr]) return
      ops.push({
        pos,
        run: (t) => t.setNodeMarkup(pos, undefined, { ...node.attrs, [goingMarkAttr]: markRecord }),
      })
    } else {
      ops.push({
        pos,
        run: (t) => {
          if (!pendingInsertion)
            t.setNodeMarkup(pos, undefined, {
              ...node.attrs,
              blockRevision: { kind: goingKind, ...record },
            })
          // text someone already struck keeps that record
          node.forEach((child, offset) => {
            if (hasMark(child.marks, 'del')) return
            t.addMark(pos + 1 + offset, pos + 1 + offset + child.nodeSize, goingMark)
          })
        },
      })
    }
    summary.deletions++
  }

  /**
   * A paragraph of the other document as a node of this one, its text coming
   * in as a change; null when it holds nothing but objects that stay behind.
   * `inside` a row or table that is itself the change, the paragraph carries no record of its own.
   */
  const incomingParagraph = (json: JsonNode, at: number, inside = false): PmNode | null => {
    const text = rightText(json, keep)
    const content = insertedInline(text, 0, text.text.length)
    const theirs = keep ? otherTracked(text) : []
    if (content.length === 0 && text.text.length > 0) return null
    const attrs: Record<string, unknown> = { ...json.attrs }
    for (const key of FOREIGN_PARAGRAPH_ATTRS) attrs[key] = null
    const styleId = attrs.styleId as string | null | undefined
    if (styleId && options.hasStyle && !options.hasStyle(styleId)) attrs.styleId = null
    // a list this document does not define would number from another package's definition
    const type =
      json.type === 'docListItem' && !listNumIds.has(String(json.attrs?.numId))
        ? 'docParagraph'
        : json.type
    // a paragraph the other document inserted or deleted as a tracked change comes in as that change
    const whole =
      keep &&
      (theirs.length > 0
        ? theirs.every(Boolean)
        : !!(json.attrs?.paraMarkIns || json.attrs?.paraMarkDel))
    if (whole) {
      attrs.paraMarkIns = json.attrs?.paraMarkIns ?? null
      attrs.paraMarkDel = json.attrs?.paraMarkDel ?? null
      attrs.moveRevision = json.attrs?.moveRevision ?? null
      if (theirs.length === 0)
        carriedChanges.add(JSON.stringify(['mark', attrs.paraMarkIns, attrs.paraMarkDel, at]))
    } else if (inside) {
      if (content.length === 0) return schema.nodes[type]!.create(attrs)
    } else if (content.length === 0) attrs[comingMarkAttr] = markRecord
    else attrs.blockRevision = { kind: comingKind, ...record }
    if (!whole && !inside) summary.insertions++
    return schema.nodes[type]!.create(attrs, content)
  }

  /** a table, row or cell of the other document as a node of this one, every paragraph in it coming in */
  const incomingNode = (json: JsonNode, at: number): PmNode | null => {
    if (PARAGRAPH_TYPES.has(json.type)) return incomingParagraph(json, at, true)
    const type = schema.nodes[json.type]
    if (!type || !TABLE_TYPES.has(json.type)) return null
    const children = (json.content ?? [])
      .map((child) => incomingNode(child, at))
      .filter((node): node is PmNode => node !== null)
    const attrs = {
      ...json.attrs,
      ...(json.attrs && 'docxIndex' in json.attrs ? { docxIndex: null } : {}),
    }
    return type.createAndFill(attrs, children)
  }

  type Revision = { kind?: string; author?: string; date?: string | null } | null | undefined
  const rowRecord = (
    kind: 'ins' | 'del',
    by: { author: string; date: string },
    rawTrPr: unknown,
  ) => ({
    rowRevision: { kind, author: by.author, date: by.date },
    rawTrPr: withRowRecord(rawTrPr, kind, by.author, by.date),
  })

  /** a row only the other document has */
  const incomingRow = (json: JsonNode, at: number): PmNode | null => {
    const row = incomingNode(json, at)
    if (!row) return null
    // a row the other document already tracks comes in as that change
    if (keep && json.attrs?.rowRevision) {
      carriedChanges.add(JSON.stringify(['row', json.attrs.rowRevision, at, carriedChanges.size]))
      return row
    }
    summary.insertions++
    return row.type.create(
      { ...row.attrs, ...rowRecord(comingKind, record, row.attrs.rawTrPr) },
      row.content,
    )
  }

  const addBlock = (unit: (typeof right)[number], at: number): void => {
    const { json } = unit
    if (unit.kind === 'table') {
      const theirs = json.attrs?.blockRevision as Revision
      const table = incomingNode(json, at)
      if (!table) {
        summary.skipped++
        return
      }
      const whole = keep && !!theirs
      if (whole) carriedChanges.add(JSON.stringify(['table', theirs, at]))
      else summary.insertions++
      const blockRevision = whole ? theirs : { kind: comingKind, ...record }
      const node = table.type.create({ ...table.attrs, blockRevision }, table.content)
      ops.push({ pos: at, run: (t) => t.insert(at, node) })
      return
    }
    if (unit.kind !== 'text') {
      summary.skipped++
      return
    }
    const node = incomingParagraph(json, at)
    if (node) ops.push({ pos: at, run: (t) => t.insert(at, node) })
  }

  /** every piece of text in [from, to) that is not already struck goes */
  const strikeRange = (from: number, to: number): void => {
    ops.push({
      pos: from,
      run: (t) =>
        t.doc.nodesBetween(from, to, (child, at) => {
          if (child.isInline && !hasMark(child.marks, 'del'))
            t.addMark(Math.max(at, from), Math.min(at + child.nodeSize, to), goingMark)
          return true
        }),
    })
  }

  const removeTable = (unit: (typeof left)[number]): void => {
    const { node, pos } = unit
    if (node.attrs.blockRevision) {
      if (!keep) summary.skipped++
      return
    }
    ops.push({
      pos,
      run: (t) =>
        t.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          blockRevision: { kind: goingKind, ...record },
        }),
    })
    strikeRange(pos + 1, pos + node.nodeSize - 1)
    summary.deletions++
  }

  interface LeftRow {
    node: PmNode
    pos: number
    text: string
  }
  const leftRows = (table: PmNode, pos: number): LeftRow[] => {
    const rows: LeftRow[] = []
    table.forEach((row, offset) => {
      const at = pos + 1 + offset
      const text = leftParagraphs(row, at)
        .map((p) => leftText(p.node, p.pos, keep).text)
        .join('\n')
      rows.push({ node: row, pos: at, text })
    })
    return rows
  }
  const rowText = (row: JsonNode): string =>
    rightParagraphs(row)
      .map((p) => rightText(p, keep).text)
      .join('\n')
  const revisionKind = (attrs: Record<string, unknown> | null | undefined): string | undefined =>
    (attrs?.rowRevision as Revision)?.kind

  const cellText = (cell: PmNode, at: number): string =>
    leftParagraphs(cell, at)
      .map((p) => leftText(p.node, p.pos, keep).text)
      .join('\n')

  /** a cell the other document tracks as inserted or deleted, and this one does not, takes that change on */
  const carryCellRecord = (cell: PmNode, pos: number, other: JsonNode): void => {
    const theirs = other.attrs?.cellRevision as Revision
    if (!keep || !theirs?.kind || cell.attrs.cellRevision) return
    const kind = theirs.kind === 'del' ? 'del' : 'ins'
    const by = { author: theirs.author ?? '', date: theirs.date ?? options.date }
    carriedChanges.add(JSON.stringify(['cell', theirs, pos]))
    ops.push({
      pos,
      run: (t) => {
        const live = t.doc.nodeAt(pos)!
        t.setNodeMarkup(pos, undefined, {
          ...live.attrs,
          cellRevision: { kind, ...by },
          rawTcPr: withTrackMarker(
            (live.attrs.rawTcPr as string | null) ?? null,
            kind === 'ins' ? 'w:cellIns' : 'w:cellDel',
            'w:tcPr',
            by.author,
            by.date,
          ),
        })
      },
    })
  }

  /** a cell both rows have */
  const diffCell = (cell: PmNode, at: number, other: JsonNode): void => {
    carryCellRecord(cell, at, other)
    if (leftShape(cell) === rightShape(other)) {
      const theirs = rightParagraphs(other)
      leftParagraphs(cell, at).forEach((p, k) => diffParagraph(p.node, p.pos, theirs[k]!))
    } else if (cellText(cell, at) !== rowText(other)) summary.skipped++
  }

  const cellRecord = (kind: 'ins' | 'del', rawTcPr: unknown) => ({
    cellRevision: { kind, ...record },
    rawTcPr: withTrackMarker(
      (rawTcPr as string | null) ?? null,
      kind === 'ins' ? 'w:cellIns' : 'w:cellDel',
      'w:tcPr',
      record.author,
      record.date,
    ),
  })

  /** an empty cell shaped like `like`, standing in a row for a column the row's own table does not have */
  const standInCell = (like: PmNode, kind: 'ins' | 'del'): PmNode => {
    const width = like.attrs.colwidth as number[] | null
    return like.type.create(
      {
        ...like.attrs,
        colspan: 1,
        rowspan: 1,
        colwidth: width ? width.slice(0, 1) : null,
        cellMerge: null,
        tcPrChange: null,
        ...cellRecord(kind, null),
      },
      schema.nodes.docParagraph!.create(),
    )
  }

  /** a cell by the entries of the column alignment it covers */
  interface PlacedCell {
    /** -1: the row is covered here by a cell merged down from above */
    index: number
    from: number
    span: number
    kind: TextAlignment['kind']
    shape: string
    merge: 'restart' | 'continue' | null
  }
  /**
   * How the columns of two tables stand for each other, when one was added or
   * removed. Each table fills a grid (merged cells count by what they cover),
   * and no merged cell reaches across a column only the other table has.
   */
  interface Columns {
    aligned: TextAlignment[]
    /** a cell of the other table per column of its own */
    theirs: JsonNode[]
    /** a cell of this table per column of its own */
    mine: PmNode[]
    placedMine: Map<LeftRow, PlacedCell[]>
    placedTheirs: Map<JsonNode, PlacedCell[]>
    /** cells merged down among rows only one table has: what each merge covers while the changes are pending */
    tall?: {
      rowspans: Map<PmNode | JsonNode, number>
      /** entries of a row only one table has where a merge of the other runs through it */
      covered: Map<LeftRow | JsonNode, Set<number>>
    }
  }
  interface Spot {
    index: number
    start: number
    colspan: number
    rowspan: number
  }
  /** where each cell of each row stands in the grid; null for a grid the rows do not all fill */
  const layOut = (
    rows: Array<Array<Record<string, unknown> | null | undefined>>,
  ): { width: number; rows: Spot[][]; tall: boolean } | null => {
    const taken: Array<Set<number>> = rows.map(() => new Set())
    const out: Spot[][] = []
    let tall = false
    for (const [r, cells] of rows.entries()) {
      const spots: Spot[] = []
      let cursor = 0
      for (const [index, attrs] of cells.entries()) {
        if (attrs?.gridGap) return null
        while (taken[r]!.has(cursor)) cursor++
        const colspan = Math.max(1, Number(attrs?.colspan) || 1)
        const rowspan = Math.max(1, Number(attrs?.rowspan) || 1)
        if (r + rowspan > rows.length) return null
        tall ||= rowspan > 1
        for (let down = 0; down < rowspan; down++)
          for (let k = 0; k < colspan; k++) taken[r + down]!.add(cursor + k)
        spots.push({ index, start: cursor, colspan, rowspan })
        cursor += colspan
      }
      out.push(spots)
    }
    const width = taken[0]?.size ?? 0
    const filled = taken.every((row) => row.size === width && Math.max(...row) === width - 1)
    return width > 0 && filled ? { width, rows: out, tall } : null
  }
  const alignColumns = (
    mine: LeftRow[],
    theirs: JsonNode[],
    pairs: Array<{ left: number; right: number }>,
    /** how the rows stand for each other, when some are in one table only */
    unpaired: TextAlignment[] | null,
  ): Columns | null => {
    const attrsOf = (row: PmNode) => {
      const cells: Array<Record<string, unknown>> = []
      row.forEach((cell) => cells.push(cell.attrs))
      return cells
    }
    const left = layOut(mine.map((row) => attrsOf(row.node)))
    const right = layOut(theirs.map((row) => (row.content ?? []).map((cell) => cell.attrs)))
    if (!left || !right) return null
    // a cell merged down takes part in every row it covers
    const tall = !!unpaired && (left.tall || right.tall)
    const cellAt = (row: LeftRow, index: number) => {
      let at = row.pos + 1
      for (let i = 0; i < index; i++) at += row.node.child(i).nodeSize
      return cellText(row.node.child(index), at)
    }
    // a column reads as the cells that start in it
    const leftTexts = Array.from({ length: left.width }, (_, col) =>
      pairs
        .map((pair) => {
          const spot = left.rows[pair.left]!.find((cell) => cell.start === col)
          return spot ? cellAt(mine[pair.left]!, spot.index) : ''
        })
        .join('\n'),
    )
    const rightTexts = Array.from({ length: right.width }, (_, col) =>
      pairs
        .map((pair) => {
          const spot = right.rows[pair.right]!.find((cell) => cell.start === col)
          return spot ? rowText(theirs[pair.right]!.content![spot.index]!) : ''
        })
        .join('\n'),
    )
    const aligned: TextAlignment[] =
      tall && left.width === right.width
        ? leftTexts.map((_, i) => ({ kind: 'pair', left: i, right: i }))
        : alignTexts(leftTexts, rightTexts, ignore)
    if (!aligned.some((entry) => entry.kind === 'pair')) return null

    const entryOf = { left: new Array<number>(left.width), right: new Array<number>(right.width) }
    aligned.forEach((entry, i) => {
      if (entry.kind !== 'added') entryOf.left[entry.left] = i
      if (entry.kind !== 'removed') entryOf.right[entry.right] = i
    })
    /** a row's cells by the entries they cover; null when one reaches across a column of the other table */
    const place = (spots: Spot[], at: number[]): PlacedCell[] | null => {
      const out: PlacedCell[] = []
      for (const spot of spots) {
        const from = at[spot.start]!
        const kind = aligned[from]!.kind
        for (let k = 1; k < spot.colspan; k++)
          if (at[spot.start + k] !== from + k || aligned[from + k]!.kind !== kind) return null
        const merge = spot.index < 0 ? 'continue' : spot.rowspan > 1 ? 'restart' : null
        out.push({
          index: spot.index,
          from,
          span: spot.colspan,
          kind,
          merge,
          shape: tall ? `${spot.colspan}:${merge}` : `${spot.colspan}x${spot.rowspan}`,
        })
      }
      return out
    }
    /** each row with the cells merged down into it from above */
    const covering = (rows: Spot[][]): Spot[][] => {
      if (!tall) return rows
      const out = rows.map((row) => [...row])
      rows.forEach((row, r) => {
        for (const spot of row)
          for (let down = 1; down < spot.rowspan; down++)
            out[r + down]!.push({ ...spot, index: -1, rowspan: 1 })
      })
      return out.map((row) => row.sort((a, b) => a.start - b.start))
    }
    const placedMine = new Map<LeftRow, PlacedCell[]>()
    for (const [i, spots] of covering(left.rows).entries()) {
      const cells = place(spots, entryOf.left)
      if (!cells) return null
      placedMine.set(mine[i]!, cells)
    }
    const placedTheirs = new Map<JsonNode, PlacedCell[]>()
    for (const [i, spots] of covering(right.rows).entries()) {
      const cells = place(spots, entryOf.right)
      if (!cells) return null
      placedTheirs.set(theirs[i]!, cells)
    }
    // the cells two paired rows share are merged alike
    const shared = (cells: PlacedCell[]) =>
      cells
        .filter((cell) => cell.kind === 'pair')
        .map((cell) => `${cell.from}:${cell.shape}`)
        .join(' ')
    for (const pair of pairs) {
      if (
        shared(placedMine.get(mine[pair.left]!)!) !== shared(placedTheirs.get(theirs[pair.right]!)!)
      )
        return null
    }
    /** a cell of each column: an unmerged one that starts in it, else whatever does */
    const sample = <T>(
      grid: { width: number; rows: Spot[][] },
      cell: (row: number, index: number) => T,
    ): T[] =>
      Array.from({ length: grid.width }, (_, col) => {
        const found = (plain: boolean) => {
          for (const [r, spots] of grid.rows.entries()) {
            const spot = spots.find((s) =>
              plain
                ? s.start === col && s.colspan === 1
                : s.start <= col && col < s.start + s.colspan,
            )
            if (spot) return cell(r, spot.index)
          }
          return undefined
        }
        return (found(true) ?? found(false))!
      })
    const columns: Columns = {
      aligned,
      theirs: sample(right, (r, index) => theirs[r]!.content![index]!),
      mine: sample(left, (r, index) => mine[r]!.node.child(index)),
      placedMine,
      placedTheirs,
    }
    if (!tall) return columns
    const merges = pendingMerges(columns, unpaired, mine, theirs)
    return merges ? { ...columns, tall: merges } : null
  }

  /**
   * Cells merged down in a table some of whose rows are in one table only:
   * how many rows each merge covers while the changes are pending, and where
   * a merge of one table runs through a row of the other. null when a row of
   * one table would cut a merge, or starts one that the rows below continue.
   */
  const pendingMerges = (
    columns: Columns,
    order: TextAlignment[],
    mine: LeftRow[],
    theirs: JsonNode[],
  ): NonNullable<Columns['tall']> | null => {
    const written = order.map((entry) => ({
      left: entry.kind === 'added' ? undefined : mine[entry.left]!,
      right: entry.kind === 'removed' ? undefined : theirs[entry.right]!,
    }))
    type Written = (typeof written)[number]
    const cellsOf = (row: Written, side: 'left' | 'right') =>
      side === 'left'
        ? row.left && columns.placedMine.get(row.left)
        : row.right && columns.placedTheirs.get(row.right)
    /** the table a column of the comparison is read from in this row; a column both have, from whichever is there */
    const sideOf = (row: Written, from: number): 'left' | 'right' => {
      const kind = columns.aligned[from]!.kind
      return kind === 'removed' || (kind === 'pair' && row.left) ? 'left' : 'right'
    }
    const covered = new Map<LeftRow | JsonNode, Set<number>>()
    /** whether the row continues, at this entry, a merge that started above */
    const continues = (k: number, from: number, span: number): boolean => {
      const row = written[k]
      if (!row) return false
      const side = sideOf(row, from)
      const cells = cellsOf(row, side)
      if (cells)
        return cells.some((c) => c.from === from && c.span === span && c.merge === 'continue')
      // a row the column's table does not have: the merge runs through it when it goes on below
      const below = written.slice(k + 1).find((next) => cellsOf(next, side))
      return !!below && continues(written.indexOf(below), from, span)
    }
    const rowspans = new Map<PmNode | JsonNode, number>()
    for (const [k, row] of written.entries()) {
      const single = !(row.left && row.right)
      for (const side of ['left', 'right'] as const) {
        const cells = cellsOf(row, side)
        if (!cells) continue
        for (const cell of cells) {
          if (side !== sideOf(row, cell.from)) continue
          if (cell.merge === 'continue') {
            if (cell.kind !== 'pair') continue
            const above = written[k - 1]
            const up = above && cellsOf(above, sideOf(above, cell.from))
            if (!up?.some((c) => c.from === cell.from && c.span === cell.span && c.merge))
              return null
            continue
          }
          if (single && cell.kind === 'pair' && cell.merge === 'restart') return null
          let rows = 1
          while (continues(k + rows, cell.from, cell.span)) rows++
          const node =
            side === 'left' ? row.left!.node.child(cell.index) : row.right!.content![cell.index]!
          const own = Math.max(1, Number(node.attrs?.rowspan) || 1)
          if (rows !== own) rowspans.set(node, rows)
        }
      }
      if (!single) continue
      const other = row.left ? 'right' : 'left'
      const below = written.slice(k + 1).find((next) => cellsOf(next, other))
      const through = new Set<number>()
      for (const cell of (below && cellsOf(below, other)) ?? []) {
        if (cell.merge !== 'continue' || cell.kind === 'pair') continue
        for (let i = 0; i < cell.span; i++) through.add(cell.from + i)
      }
      if (through.size > 0) covered.set((row.left ?? row.right)!, through)
    }
    return { rowspans, covered }
  }

  /** a cell of the other table as it comes in: over the rows its merge covers while the changes are pending */
  const incomingCell = (json: JsonNode, at: number, columns: Columns, marked: boolean) => {
    const cell = incomingNode(json, at)
    if (!cell) return null
    const rowspan = columns.tall?.rowspans.get(json)
    return cell.type.create(
      {
        ...cell.attrs,
        ...(rowspan ? { rowspan } : {}),
        ...(marked ? cellRecord(comingKind, cell.attrs.rawTcPr) : {}),
      },
      cell.content,
    )
  }

  const startsAt = (cells: PlacedCell[] | undefined) =>
    new Map((cells ?? []).filter((cell) => cell.index >= 0).map((cell) => [cell.from, cell]))

  /** a row both tables have, cell by cell along the columns: a column only one has is its cell tracked */
  const diffRowColumns = (mine: LeftRow, json: JsonNode, columns: Columns): void => {
    const cells = json.content ?? []
    const own = startsAt(columns.placedMine.get(mine))
    const other = startsAt(columns.placedTheirs.get(json))
    let at = mine.pos + 1
    columns.aligned.forEach((entry, i) => {
      if (entry.kind === 'added') {
        const spot = other.get(i)
        if (!spot) return
        const pos = at
        const node = incomingCell(cells[spot.index]!, pos, columns, true)
        if (node) ops.push({ pos, run: (t) => t.insert(pos, node) })
        return
      }
      const spot = own.get(i)
      if (!spot) return
      const cell = mine.node.child(spot.index)
      const pos = at
      at += cell.nodeSize
      if (entry.kind === 'pair') {
        diffCell(cell, pos, cells[other.get(i)!.index]!)
        return
      }
      if (cell.attrs.cellRevision) return
      ops.push({
        pos,
        run: (t) => {
          const live = t.doc.nodeAt(pos)!
          t.setNodeMarkup(pos, undefined, {
            ...live.attrs,
            ...cellRecord(goingKind, live.attrs.rawTcPr),
          })
        },
      })
      strikeRange(pos + 1, pos + cell.nodeSize - 1)
    })
  }

  /** a row only this table has gets a cell for each column only the other table has */
  const standInsForRow = (row: LeftRow, columns: Columns): void => {
    const own = startsAt(columns.placedMine.get(row))
    const through = columns.tall?.covered.get(row)
    let at = row.pos + 1
    columns.aligned.forEach((entry, i) => {
      if (entry.kind !== 'added') {
        const spot = own.get(i)
        if (!spot) return
        const cell = row.node.child(spot.index)
        const pos = at
        at += cell.nodeSize
        // among cells merged down a column goes or stays by its cells alone: the row's own cell is marked too
        if (columns.tall && entry.kind === 'removed' && !cell.attrs.cellRevision)
          ops.push({
            pos,
            run: (t) => {
              const live = t.doc.nodeAt(pos)!
              t.setNodeMarkup(pos, undefined, {
                ...live.attrs,
                ...cellRecord(goingKind, live.attrs.rawTcPr),
              })
            },
          })
        return
      }
      if (through?.has(i)) return
      const pos = at
      const like = incomingNode(columns.theirs[entry.right]!, pos)
      if (like) ops.push({ pos, run: (t) => t.insert(pos, standInCell(like, comingKind)) })
    })
  }

  /** a row only the other table has gets a cell for each column only this table has */
  const withStandIns = (row: PmNode, json: JsonNode, columns: Columns, at: number): PmNode => {
    const own = startsAt(columns.placedTheirs.get(json))
    const through = columns.tall?.covered.get(json)
    const cells: PmNode[] = []
    columns.aligned.forEach((entry, i) => {
      if (entry.kind === 'removed') {
        if (!through?.has(i)) cells.push(standInCell(columns.mine[entry.left]!, goingKind))
      } else if (own.has(i)) {
        const index = own.get(i)!.index
        const cell = columns.tall
          ? incomingCell(json.content![index]!, at, columns, entry.kind === 'added')
          : null
        cells.push(cell ?? row.child(index))
      }
    })
    return row.type.create(row.attrs, cells)
  }

  /** the table's columns as wide as they were, a column that came in as wide as the one before it */
  const reflowMerged = (unit: { node: PmNode; pos: number }, columns: Columns): void => {
    const map = TableMap.get(unit.node)
    const own = columnWidthsPx(unit.node, map)
    const ownPct = pctOf(unit.node, map, own)
    const spread = <T>(values: T[]): T[] => {
      const out: T[] = []
      for (const entry of columns.aligned)
        out.push(entry.kind === 'added' ? (out[out.length - 1] ?? values[0]!) : values[entry.left]!)
      return out
    }
    const widths = own ? spread(own) : null
    const pct = spread(ownPct)
    ops.push({ pos: unit.pos, run: (t) => reflowColumns(t, unit.pos, widths, pct) })
  }

  /** a row both tables have: its cells stand for each other one by one */
  const diffRow = (mine: LeftRow, json: JsonNode, columns: Columns | null): void => {
    if (columns) diffRowColumns(mine, json, columns)
    else {
      const cells = json.content ?? []
      let i = 0
      mine.node.forEach((cell, offset) => diffCell(cell, mine.pos + 1 + offset, cells[i++]!))
    }
    // a row change the other document tracks and this one does not show
    const theirs = json.attrs?.rowRevision as Revision
    if (!keep || !theirs?.kind || mine.node.attrs.rowRevision) return
    const by = { author: theirs.author ?? '', date: theirs.date ?? options.date }
    carriedChanges.add(JSON.stringify(['row', theirs, mine.pos]))
    ops.push({
      pos: mine.pos,
      run: (t) => {
        const live = t.doc.nodeAt(mine.pos)!
        t.setNodeMarkup(mine.pos, undefined, {
          ...live.attrs,
          ...rowRecord(theirs.kind === 'del' ? 'del' : 'ins', by, live.attrs.rawTrPr),
        })
      },
    })
  }

  const removeRow = (row: LeftRow): void => {
    // a row this document already tracks as a change stays that change
    if (row.node.attrs.rowRevision) {
      if (!keep) summary.skipped++
      return
    }
    ops.push({
      pos: row.pos,
      run: (t) => {
        const live = t.doc.nodeAt(row.pos)!
        t.setNodeMarkup(row.pos, undefined, {
          ...live.attrs,
          ...rowRecord(goingKind, record, live.attrs.rawTrPr),
        })
      },
    })
    // cell by cell: a cell that comes into the row for a column of the other table is not struck
    row.node.forEach((cell, offset) =>
      strikeRange(row.pos + 1 + offset + 1, row.pos + 1 + offset + cell.nodeSize - 1),
    )
    summary.deletions++
  }

  const diffTable = (unit: (typeof left)[number], json: JsonNode, sameText: boolean): void => {
    const allMine = leftRows(unit.node, unit.pos)
    const allTheirs = json.content ?? []
    // with every change accepted, a row pending deletion is not part of the table
    const struck = (rows: Array<{ attrs?: Record<string, unknown> | null }>) =>
      !keep && rows.some((row) => revisionKind(row.attrs) === 'del')
    const inPlace =
      leftShape(unit.node) === rightShape(json) &&
      !struck(allMine.map((row) => row.node)) &&
      !struck(allTheirs)
    if (sameText && inPlace && !formatting && !keep) return
    const mine = allMine.filter((row) => inPlace || keep || revisionKind(row.node.attrs) !== 'del')
    const theirs = allTheirs.filter((row) => inPlace || keep || revisionKind(row.attrs) !== 'del')
    const rows: TextAlignment[] = inPlace
      ? mine.map((_, i) => ({ kind: 'pair', left: i, right: i }))
      : alignTexts(
          mine.map((row) => row.text),
          theirs.map(rowText),
          ignore,
        )
    const pairs = rows.filter((entry) => entry.kind === 'pair')
    const stands = pairs.every(
      (entry) => mine[entry.left]!.node.childCount === (theirs[entry.right]!.content ?? []).length,
    )
    // a column added or removed: the cells of a row stand for each other by their column
    const whole =
      pairs.length === rows.length &&
      mine.length === allMine.length &&
      theirs.length === allTheirs.length
    const merged = (attrs: Record<string, unknown> | null | undefined) => Number(attrs?.rowspan) > 1
    const tall =
      !whole &&
      (mine.some((row) => {
        let any = false
        row.node.forEach((cell) => (any ||= merged(cell.attrs)))
        return any
      }) ||
        theirs.some((row) => (row.content ?? []).some((cell) => merged(cell.attrs))))
    const columns =
      pairs.length === 0
        ? null
        : tall
          ? alignColumns(mine, theirs, pairs, rows)
          : stands
            ? null
            : alignColumns(mine, theirs, pairs, null)
    // no row in common, or columns that do not line up: the table is left as it is
    if (pairs.length === 0 || ((tall || !stands) && !columns)) {
      if (!sameText) summary.skipped++
      return
    }
    if (columns) {
      for (const entry of columns.aligned) {
        if (entry.kind === 'added') summary.insertions++
        else if (entry.kind === 'removed') summary.deletions++
      }
      if (columns.aligned.some((entry) => entry.kind !== 'pair')) reflowMerged(unit, columns)
    }
    let after = unit.pos + 1
    for (const entry of rows) {
      if (entry.kind === 'added') {
        const at = after
        const row = incomingRow(theirs[entry.right]!, at)
        const filled = row && columns ? withStandIns(row, theirs[entry.right]!, columns, at) : row
        if (filled) ops.push({ pos: at, run: (t) => t.insert(at, filled) })
        continue
      }
      const row = mine[entry.left]!
      after = row.pos + row.node.nodeSize
      if (entry.kind === 'removed') {
        removeRow(row)
        if (columns) standInsForRow(row, columns)
      } else diffRow(row, theirs[entry.right]!, columns)
    }
    // last, so they run before a cell comes in at the same place
    for (const row of mine) {
      row.node.forEach((cell, offset) => {
        const rowspan = columns?.tall?.rowspans.get(cell)
        const pos = row.pos + 1 + offset
        if (rowspan) ops.push({ pos, run: (t) => t.setNodeAttribute(pos, 'rowspan', rowspan) })
      })
    }
  }

  // where a block only the revised document has goes: after the last block of this document before it
  let boundary = 0
  for (const entry of aligned) {
    if (entry.kind === 'added') {
      if (tables || right[entry.right]!.kind !== 'table') addBlock(right[entry.right]!, boundary)
      continue
    }
    const mine = left[entry.left]!
    boundary = mine.pos + mine.node.nodeSize
    if (!tables && mine.kind === 'table') continue
    if (entry.kind === 'removed') {
      removeBlock(mine)
      continue
    }
    const theirs = right[entry.right]!
    if (!tables && theirs.kind === 'table') continue
    if (mine.kind !== theirs.kind) {
      removeBlock(mine)
      addBlock(theirs, boundary)
    } else if (mine.kind === 'text') diffParagraph(mine.node, mine.pos, theirs.json)
    else if (mine.kind === 'table') diffTable(mine, theirs.json, mine.text === theirs.text)
    else if (mine.text !== theirs.text) summary.skipped++
  }

  // last position first, and at one position the later op first, so every
  // position still names the document the ops were planned against
  ops
    .map((op, order) => ({ op, order }))
    .sort((a, b) => b.op.pos - a.op.pos || b.order - a.order)
    .forEach(({ op }) => op.run(tr))
  settleMoves(tr, new Set(carriedMoves.values()))
  if (keep) summary.carried = carriedChanges.size
  return flipped
    ? { ...summary, insertions: summary.deletions, deletions: summary.insertions }
    : summary
}

/** A carried move whose other half did not come along reads as a plain deletion or insertion. */
function settleMoves(tr: Transaction, names: ReadonlySet<string>): void {
  if (names.size === 0) return
  const sides = new Map<string, Set<string>>()
  const spans: Array<{ from: number; to: number; mark: Mark; para: number }> = []
  tr.doc.descendants((node, pos, parent) => {
    for (const mark of node.marks) {
      const name = mark.attrs.move as string | null
      if (!name || !names.has(name)) continue
      sides.set(name, (sides.get(name) ?? new Set()).add(mark.type.name))
      const para = parent?.isTextblock ? tr.doc.resolve(pos).before() : -1
      spans.push({ from: pos, to: pos + node.nodeSize, mark, para })
    }
    return true
  })
  for (const { from, to, mark, para } of spans) {
    if (sides.get(mark.attrs.move as string)!.size > 1) continue
    tr.removeMark(from, to, mark)
    tr.addMark(from, to, mark.type.create({ ...mark.attrs, move: null }))
    const host = para >= 0 ? tr.doc.nodeAt(para) : null
    if (host?.attrs.moveRevision)
      tr.setNodeMarkup(para, undefined, { ...host.attrs, moveRevision: null })
  }
}

/** Combine as one undoable edit; the document is untouched when nothing differs or nothing is shared. */
export function combineDocuments(
  editor: Editor,
  revised: JsonNode[],
  options: CombineOptions,
): CombineSummary {
  const tr = editor.state.tr
  const summary = combineIntoTransaction(tr, revised, options)
  if (!tr.docChanged) return summary
  tr.setMeta(TRACK_IGNORE, true)
  editor.view.dispatch(closeHistory(tr))
  return summary
}
