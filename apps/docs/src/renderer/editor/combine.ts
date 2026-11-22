/**
 * Tools > Combine Documents: the differences between the open document and a
 * revised one, written into the open document as ordinary tracked changes.
 *
 * Compared: the text of body paragraphs (aligned paragraph by paragraph, then
 * word by word inside a changed paragraph) and of the cells of a table that
 * has the same rows, cells and paragraphs on both sides. Not compared, and
 * counted in `skipped` when they differ: formatting, tables that were added,
 * removed or reshaped, pictures and other objects, notes, fields.
 */
import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import type { Mark, Node as PmNode, Schema } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
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
import { TRACK_IGNORE } from './revisions'

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
  /** tables, objects and inline objects that differ and were left as they are */
  skipped: number
  /** false when the two documents share no paragraph: nothing is written then */
  related: boolean
  /** keepRevisions: tracked changes of the other document that came in under their own author */
  carried?: number
}

const PARAGRAPH_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])
const ATOM = '\uFFFC'
const CARRIED_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])
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
}

const trackedState = (marks: ReadonlyArray<{ type: string | { name: string } }>): TrackedState =>
  hasMark(marks, 'ins') ? 'ins' : hasMark(marks, 'del') ? 'del' : ''

/** the paragraph's text with its pending deletions left out, or with `markup` as it stands */
function leftText(para: PmNode, pos: number, markup = false): LeftText {
  let text = ''
  const at: number[] = []
  const end: number[] = []
  const tracked: TrackedState[] = []
  para.forEach((child, offset) => {
    const start = pos + 1 + offset
    if (!markup && hasMark(child.marks, 'del')) return
    const flag = trackedState(child.marks)
    if (child.isText) {
      for (let i = 0; i < child.text!.length; i++) {
        at.push(start + i)
        end.push(start + i + 1)
        tracked.push(flag)
      }
      text += child.text
    } else {
      at.push(start)
      end.push(start + child.nodeSize)
      tracked.push(flag)
      text +=
        child.type.name === 'hardBreak' && !child.attrs.mathBreak ? breakChar(child.attrs) : ATOM
    }
  })
  return { text, at, end, tracked }
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
          : ATOM
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
  const summary: CombineSummary = { insertions: 0, deletions: 0, skipped: 0, related: true }
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

  /** the tracked change a piece of the other document already sits in, under an id minted on save */
  const carriedChanges = new Set<string>()
  const ownChange = (child: JsonNode): Mark[] =>
    (child.marks ?? [])
      .filter((mark) => mark.type === 'ins' || mark.type === 'del')
      .map((mark) => {
        carriedChanges.add(
          JSON.stringify([mark.type, mark.attrs?.author, mark.attrs?.date, mark.attrs?.id]),
        )
        return schema.marks[mark.type]!.create({ ...mark.attrs, id: null, move: null })
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
      } else summary.skipped += end - i
      i = end
    }
    return nodes
  }

  /** strike [from, to) of the compared text; an object in it stays, text around it goes */
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

  /** both read with their markup: each side's own changes stay, only plain differences are new */
  const mergeParagraph = (node: PmNode, pos: number, json: JsonNode): void => {
    const l = leftText(node, pos, true)
    const r = rightText(json, true)
    const theirs = otherTracked(r)
    if (l.text === r.text && !theirs.some((flag, i) => flag && !l.tracked[i])) return
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
          if (l.tracked[mine] || l.text[mine] === ATOM) continue
          const a = l.at[mine]!
          const b = l.end[mine]!
          for (const mark of ownChange(child))
            ops.push({ pos: a, run: (t) => t.addMark(a, b, mark) })
        }
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
    if (l.text === r.text) return
    let consumed = 0
    for (const chunk of diffText(l.text, r.text, options.granularity, ignore)) {
      if (chunk.kind === 'keep') {
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

  const listNumIds = new Set<string>()
  doc.descendants((node) => {
    if (node.type.name === 'docListItem' && node.attrs.numId != null)
      listNumIds.add(String(node.attrs.numId))
    return !node.isTextblock
  })

  const addBlock = (unit: (typeof right)[number], at: number): void => {
    const { json } = unit
    if (unit.kind !== 'text') {
      summary.skipped++
      return
    }
    const text = rightText(json, keep)
    const content = insertedInline(text, 0, text.text.length)
    const theirs = keep ? otherTracked(text) : []
    // nothing but objects that stay behind: no paragraph to bring over
    if (content.length === 0 && text.text.length > 0) return
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
      if (theirs.length === 0)
        carriedChanges.add(JSON.stringify(['mark', attrs.paraMarkIns, attrs.paraMarkDel, at]))
    } else if (content.length === 0) attrs[comingMarkAttr] = markRecord
    else attrs.blockRevision = { kind: comingKind, ...record }
    const node = schema.nodes[type]!.create(attrs, content)
    ops.push({ pos: at, run: (t) => t.insert(at, node) })
    if (!whole) summary.insertions++
  }

  const diffTable = (unit: (typeof left)[number], json: JsonNode, sameText: boolean): void => {
    if (sameText) return
    if (leftShape(unit.node) !== rightShape(json)) {
      summary.skipped++
      return
    }
    const mine = leftParagraphs(unit.node, unit.pos)
    const theirs = rightParagraphs(json)
    mine.forEach((p, i) => diffParagraph(p.node, p.pos, theirs[i]!))
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
  if (keep) summary.carried = carriedChanges.size
  return flipped
    ? { ...summary, insertions: summary.deletions, deletions: summary.insertions }
    : summary
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
