/**
 * Word revisions (track changes) support for the TipTap editor.
 *
 * - TrackChangesExtension records edits as ins / del marks while enabled
 *   (Word: Review → Track Changes).
 * - collectRevisions / accept* / reject* implement Accept/Reject Changes.
 *
 * Semantics of a text range carrying BOTH marks (inserted while tracking,
 * then deleted while tracking): the text never existed in the base document,
 * so both accepting and rejecting remove it.
 */
import { Extension, type Editor } from '@tiptap/core'
import { Fragment, type Mark as PmMark, type Node as PmNode } from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import { TableMap, removeColumn } from '@tiptap/pm/tables'
import {
  AddMarkStep,
  canJoin,
  Mapping,
  RemoveMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
} from '@tiptap/pm/transform'
import { authorName } from '../author-name'
import { charScaleEm, wordKerns } from '../line-metrics'
import type { StoryLocation } from './story-revisions'
import {
  gridColTwips,
  tblPrRestoreAttrs,
  tcPrRestoreAttrs,
  trPrRestoreAttrs,
} from './prop-change-restore'
import { columnWidthsPx, pctOf, reflowColumns } from '../ai/table-ops'

export interface RevisionRange {
  from: number
  to: number
  kind:
    | 'ins'
    | 'del'
    | 'pPrChange'
    | 'paraMarkDel'
    | 'paraMarkIns'
    | 'paraMarkRPrChange'
    | 'moveFrom'
    | 'moveTo'
    | 'rPrChange'
    | 'rowIns'
    | 'rowDel'
    | 'cellIns'
    | 'cellDel'
    | 'blockIns'
    | 'blockDel'
    | 'tblPrChange'
    | 'tblGridChange'
    | 'trPrChange'
    | 'tblPrExChange'
    | 'tcPrChange'
    | 'cellMerge'
    | 'numberingChange'
    | 'sectPrChange'
  author: string
  date?: string
  /** text carrying both ins and del marks: hidden in every view but All Markup */
  overlapped?: true
  /** w:name shared by the source (moveFrom) and destination (moveTo) of one move */
  moveId?: string
  /** set when the change lives outside the body (note, header / footer, textbox); from / to only order it */
  story?: StoryLocation
}

/** transactions carrying this meta are never recorded as revisions */
export const TRACK_IGNORE = 'trackIgnore'
/** meta: record this transaction's edits under the given author instead of the user */
export const TRACK_AUTHOR = 'trackAuthor'

const recorderKey = new PluginKey('trackChangesRecorder')

/** the recorder's author while tracking is on, else null (for commands that record themselves) */
export function activeTracking(state: EditorState): { author: string } | null {
  const storage = (recorderKey.get(state)?.spec as { storage?: TrackChangesStorage } | undefined)
    ?.storage
  return storage?.enabled ? { author: storage.author } : null
}

export const revisionDateNow = () => new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
const markRecord = (author: string, date: string) => JSON.stringify({ author, date })
const recordAuthor = (raw: unknown): string | null => {
  if (typeof raw !== 'string' || !raw) return null
  try {
    return String(JSON.parse(raw)?.author ?? '')
  } catch {
    return null
  }
}

/** trPr / tcPr bytes carrying a tracked-change marker (appended to existing properties) */
export function withTrackMarker(
  raw: string | null,
  tag: string,
  container: string,
  author: string,
  date: string,
): string {
  const marker = `<${tag} w:id="" w:author="${xmlAttr(author)}" w:date="${date}"/>`
  if (raw && raw.endsWith(`</${container}>`))
    return raw.replace(`</${container}>`, `${marker}</${container}>`)
  return `<${container}>${marker}</${container}>`
}

const TRACKED_FORMAT_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])

export const TEXT_STYLE_FIELDS = [
  'color',
  'sizeHalfPoints',
  'font',
  'fontAscii',
  'charSpacingTwips',
  'charScaleEm',
  'highlight',
  'vertAlign',
  'styleId',
] as const

const PARAGRAPH_NODE_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])
export const PARAGRAPH_FORMAT_FIELDS = [
  'align',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'indentFirstLineChars',
  'spaceBefore',
  'spaceAfter',
  'pageBreakBefore',
  'shadingFill',
  'borders',
  'tabStops',
  'dropCap',
  'bidi',
] as const
const PARAGRAPH_STRUCTURE_FIELDS = ['styleId', 'level', 'kind', 'numId', 'ilvl'] as const
/** recorded and restored with a paragraph-format change, but not what starts one while typing */
const PARAGRAPH_FLOW_FIELDS = [
  'keepNext',
  'keepLines',
  'widowControl',
  'contextualSpacing',
] as const
const xmlAttr = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')

export interface ParagraphSnapshot {
  type: string
  format: Record<string, unknown>
  styleId: unknown
  level: unknown
  kind: unknown
  numId: unknown
  ilvl: unknown
}

/** what a paragraph-format change records of a paragraph, from its node type and attrs */
export function paragraphSnapshotOf(
  type: string,
  attrs: Record<string, unknown>,
): ParagraphSnapshot {
  const format: Record<string, unknown> = {}
  for (const field of [...PARAGRAPH_FORMAT_FIELDS, ...PARAGRAPH_FLOW_FIELDS])
    format[field] = attrs[field] ?? null
  return {
    type,
    format,
    styleId: attrs.styleId ?? null,
    level: attrs.level ?? null,
    kind: attrs.kind ?? null,
    numId: attrs.numId ?? null,
    ilvl: attrs.ilvl ?? null,
  }
}

const paragraphSnapshot = (node: PmNode): ParagraphSnapshot =>
  paragraphSnapshotOf(node.type.name, node.attrs)

function paragraphPropertiesChanged(before: PmNode, after: PmNode): boolean {
  if (before.type !== after.type) return true
  return [...PARAGRAPH_FORMAT_FIELDS, ...PARAGRAPH_STRUCTURE_FIELDS].some(
    (field) => (before.attrs[field] ?? null) !== (after.attrs[field] ?? null),
  )
}

type AnyMark = { type: string | { name: string }; attrs?: Record<string, unknown> | null }

/** what a run-format change records of a piece of text, from its marks (a node's or its JSON's) */
export function runFormatSnapshotOf(marks: readonly AnyMark[]): Record<string, unknown> {
  const named = (name: string) =>
    marks.find((m) => (typeof m.type === 'string' ? m.type : m.type.name) === name)
  const textStyle = named('docTextStyle')
  const snapshot: Record<string, unknown> = {
    bold: !!named('bold'),
    italic: !!named('italic'),
    underline: !!named('underline'),
    strike: !!named('strike'),
  }
  for (const field of TEXT_STYLE_FIELDS) {
    const value = textStyle?.attrs?.[field]
    if (value !== null && value !== undefined) snapshot[field] = value
  }
  return snapshot
}

const formatSnapshot = (node: PmNode): Record<string, unknown> => runFormatSnapshotOf(node.marks)

/**
 * docTextStyle attrs for text whose format revision is rejected: every attr
 * reset, then the recorded previous formatting (the parser's modeled subset).
 */
function rejectedTextStyleAttrs(
  specAttrs: Record<string, unknown>,
  old: Record<string, unknown>,
  text: string,
): Record<string, unknown> {
  const next: Record<string, unknown> = {}
  for (const key of Object.keys(specAttrs)) next[key] = null
  for (const field of TEXT_STYLE_FIELDS) next[field] = old[field] ?? null
  for (const field of [
    'charScalePct',
    'caps',
    'dstrike',
    'vanishOwn',
    'snapToGridOwn',
    'kernHalfPoints',
    'positionHalfPoints',
    'shading',
    'rtl',
    'themeColor',
  ])
    next[field] = old[field] ?? null
  if (old.vanishOwn != null) next.vanish = old.vanishOwn
  next.offGrid = old.snapToGridOwn === false ? true : null
  next.cs = old.rtl ?? null
  const kern = old.kernHalfPoints as number | undefined
  const size = old.sizeHalfPoints as number | undefined
  next.kern = wordKerns(kern ?? undefined, size ?? undefined) ?? null
  const scale = old.charScalePct as number | undefined
  next.charScaleEm = scale
    ? charScaleEm(text, scale, {
        fontAscii: old.fontAscii as string | null,
        font: old.font as string | null,
        sizeHalfPoints: size ?? null,
        bold: old.bold as boolean | null,
        italic: old.italic as boolean | null,
      })
    : null
  return next
}

/** contiguous revision ranges in document order (adjacent same-kind ranges merged) */
export function collectRevisions(doc: PmNode): RevisionRange[] {
  const out: RevisionRange[] = []

  // Table row / cell revisions (trPr w:ins/w:del, tcPr w:cellIns/w:cellDel).
  // Text ins/del marks inside such a row/cell belong to the same revision and
  // are suppressed in the text pass (Word: one revision per inserted/deleted row).
  const tableRevs: RevisionRange[] = []
  const suppressed: Array<{ from: number; to: number; kind: 'ins' | 'del' }> = []
  doc.forEach((node, pos) => {
    const revision = node.attrs?.blockRevision as {
      kind?: 'ins' | 'del'
      author?: string
      date?: string
    } | null
    if (!revision?.kind) return
    out.push({
      from: pos,
      to: pos + node.nodeSize,
      kind: revision.kind === 'ins' ? 'blockIns' : 'blockDel',
      author: revision.author ?? '',
      ...(revision.date ? { date: revision.date } : {}),
    })
    suppressed.push({ from: pos, to: pos + node.nodeSize, kind: revision.kind })
  })
  const propRev = (record: unknown, kind: RevisionRange['kind'], from: number, to: number) => {
    const info = pendingPropChange(record)
    if (info)
      tableRevs.push({
        from,
        to,
        kind,
        author: info.author,
        ...(info.date ? { date: info.date } : {}),
      })
  }
  doc.descendants((node, pos) => {
    const name = node.type.name
    const end = pos + node.nodeSize
    if (name === 'docTable') {
      propRev(node.attrs?.tblPrChange, 'tblPrChange', pos, end)
      // w:tblGridChange carries only an id: Word writes it alongside the tblPrChange, whose author it shares
      const grid = pendingPropChange(node.attrs?.tblGridChange)
      const credit = pendingPropChange(node.attrs?.tblPrChange)
      if (grid)
        tableRevs.push({
          from: pos,
          to: end,
          kind: 'tblGridChange',
          author: grid.author || credit?.author || '',
          ...((grid.date ?? credit?.date) ? { date: grid.date ?? credit!.date } : {}),
        })
      return
    }
    if (name !== 'docTableRow' && name !== 'docTableCell' && name !== 'docTableHeader') return
    const isRow = name === 'docTableRow'
    if (isRow) {
      propRev(node.attrs?.trPrChange, 'trPrChange', pos, end)
      propRev(node.attrs?.tblPrExChange, 'tblPrExChange', pos, end)
    } else {
      propRev(node.attrs?.tcPrChange, 'tcPrChange', pos, end)
      propRev(node.attrs?.cellMerge, 'cellMerge', pos, end)
    }
    const rev = (isRow ? node.attrs?.rowRevision : node.attrs?.cellRevision) as {
      kind?: 'ins' | 'del'
      author?: string
      date?: string
    } | null
    if (!rev?.kind) return
    tableRevs.push({
      from: pos,
      to: pos + node.nodeSize,
      kind: isRow
        ? rev.kind === 'ins'
          ? 'rowIns'
          : 'rowDel'
        : rev.kind === 'ins'
          ? 'cellIns'
          : 'cellDel',
      author: rev.author ?? '',
      ...(rev.date ? { date: rev.date } : {}),
    })
    suppressed.push({ from: pos, to: pos + node.nodeSize, kind: rev.kind })
  })
  const isSuppressed = (pos: number, kind: 'ins' | 'del') =>
    suppressed.some((s) => pos >= s.from && pos < s.to && s.kind === kind)

  // First pass: paragraph-level revisions at any depth, including table cells.
  // A moved paragraph's text carries plain ins / del marks (the parser maps
  // moveTo / moveFrom onto them); the paragraph attr only says they are a move.
  const moveParas: Array<{
    from: number
    to: number
    kind: 'moveFrom' | 'moveTo'
    hasRuns?: boolean
    /** its moved runs name their move: the other ins / del in it are ordinary edits */
    named: boolean
  }> = []
  const credit = (raw: string | null): { author: string; date?: string } => {
    try {
      const parsed = raw ? JSON.parse(raw) : null
      return { author: parsed?.author ?? '', ...(parsed?.date ? { date: parsed.date } : {}) }
    } catch {
      return { author: '' }
    }
  }
  doc.descendants((node, pos) => {
    const range = { from: pos, to: pos + node.nodeSize }
    const sectPrChange = node.attrs?.sectPrChange as string | null | undefined
    if (sectPrChange && !sectPrChange.includes('"resolved"'))
      out.push({ ...range, kind: 'sectPrChange', ...credit(sectPrChange) })
    if (!PARAGRAPH_NODE_TYPES.has(node.type.name)) return
    const numberingChange = node.attrs?.numberingChange as string | null
    if (numberingChange) out.push({ ...range, kind: 'numberingChange', ...credit(numberingChange) })
    const pPrChange = node.attrs?.pPrChange as string | null
    const paraMarkDel = node.attrs?.paraMarkDel as string | null
    const paraMarkIns = node.attrs?.paraMarkIns as string | null
    const paraMarkRPrChange = node.attrs?.paraMarkRPrChange as string | null
    const moveRevision = node.attrs?.moveRevision as string | null
    if (pPrChange) out.push({ ...range, kind: 'pPrChange', ...credit(pPrChange) })
    if (paraMarkDel) out.push({ ...range, kind: 'paraMarkDel', ...credit(paraMarkDel) })
    if (paraMarkIns) out.push({ ...range, kind: 'paraMarkIns', ...credit(paraMarkIns) })
    if (paraMarkRPrChange && !paraMarkRPrChange.includes('"resolved"'))
      out.push({ ...range, kind: 'paraMarkRPrChange', ...credit(paraMarkRPrChange) })
    if (moveRevision === 'from' || moveRevision === 'to') {
      let named = false
      node.descendants((child) => {
        named ||= child.marks.some((m) => m.attrs.move && /^(?:ins|del)$/.test(m.type.name))
      })
      moveParas.push({
        ...range,
        kind: moveRevision === 'from' ? 'moveFrom' : 'moveTo',
        // a moved paragraph mark is handled as that mark's revision
        hasRuns: named || !!(paraMarkDel || paraMarkIns),
        named,
      })
    }
  })
  const moveKindAt = (pos: number, kind: 'ins' | 'del'): 'moveFrom' | 'moveTo' | null => {
    const para = moveParas.find((m) => pos >= m.from && pos < m.to)
    if (!para || para.named) return null
    const match = para.kind === 'moveFrom' ? kind === 'del' : kind === 'ins'
    if (match) para.hasRuns = true
    return match ? para.kind : null
  }

  // Second pass: collect run-level inline revisions (ins/del marks). A format
  // change on inserted text is its own revision (Word lists both), so the
  // rPrChange ranges merge among themselves and join the list at the end.
  const formatRevs: RevisionRange[] = []
  doc.descendants((node, pos) => {
    if (!node.isInline) return
    const rpr = node.isText ? node.marks.find((m) => m.type.name === 'rprChange') : undefined
    if (rpr) {
      const author = String(rpr.attrs.author ?? '')
      const prev = formatRevs[formatRevs.length - 1]
      if (prev && prev.to === pos && prev.author === author) {
        prev.to = pos + node.nodeSize
      } else {
        formatRevs.push({
          from: pos,
          to: pos + node.nodeSize,
          kind: 'rPrChange',
          author,
          ...(rpr.attrs.date ? { date: String(rpr.attrs.date) } : {}),
        })
      }
    }
    // text inserted by one author and then deleted by another carries both
    // marks: two revisions, as in Word (rejecting the deletion restores the
    // text as a pending insertion). The deletion comes first so the current-
    // revision actions at the caret target the later change.
    const marks: Array<['ins' | 'del', PmMark]> = []
    const del = node.marks.find((m) => m.type.name === 'del')
    const ins = node.marks.find((m) => m.type.name === 'ins')
    if (del) marks.push(['del', del])
    if (ins) marks.push(['ins', ins])
    const overlapped = marks.length === 2
    for (const [textKind, src] of marks) {
      if (isSuppressed(pos, textKind)) continue
      const moveId = src.attrs.move ? String(src.attrs.move) : undefined
      const paraKind = moveKindAt(pos, textKind)
      const kind = moveId ? (textKind === 'del' ? 'moveFrom' : 'moveTo') : (paraKind ?? textKind)
      const author = String(src.attrs.author ?? '')
      const prev = out
        .slice(-2)
        .find(
          (r) =>
            r.to === pos &&
            r.kind === kind &&
            r.author === author &&
            !!r.overlapped === overlapped &&
            r.moveId === moveId,
        )
      if (prev) {
        prev.to = pos + node.nodeSize
      } else {
        out.push({
          from: pos,
          to: pos + node.nodeSize,
          kind,
          author,
          ...(src.attrs.date ? { date: String(src.attrs.date) } : {}),
          ...(overlapped ? { overlapped: true } : {}),
          ...(moveId ? { moveId } : {}),
        })
      }
    }
  })
  // a moved paragraph with no run-level marks (empty, or the mark alone moved) is the revision itself
  for (const m of moveParas)
    if (!m.hasRuns) out.push({ from: m.from, to: m.to, kind: m.kind, author: '' })
  out.push(...formatRevs, ...tableRevs)
  return out.sort((a, b) => a.from - b.from)
}

/** a table / row / cell property-change record still awaiting review */
function pendingPropChange(
  record: unknown,
): { author: string; date?: string; oldXml?: string } | null {
  if (!record || typeof record !== 'object') return null
  const r = record as { author?: string; date?: string; oldXml?: string; resolved?: string }
  if (r.resolved) return null
  return {
    author: r.author ?? '',
    ...(r.date ? { date: r.date } : {}),
    ...(r.oldXml ? { oldXml: r.oldXml } : {}),
  }
}

const PROP_CHANGE_ATTR: Partial<Record<RevisionRange['kind'], string>> = {
  tblPrChange: 'tblPrChange',
  tblGridChange: 'tblGridChange',
  trPrChange: 'trPrChange',
  tblPrExChange: 'tblPrExChange',
  tcPrChange: 'tcPrChange',
  cellMerge: 'cellMerge',
}

/**
 * trPr with a row ins/del record in CT_TrPr position (after the base
 * properties, before trPrChange). w:id stays empty: the save mints it from the
 * package-wide revision counter, as it does for run-level ins/del.
 */
export function withRowRecord(
  raw: unknown,
  kind: 'ins' | 'del',
  author: string,
  date: string,
): string {
  const record = `<w:${kind} w:id="" w:author="${xmlAttr(author)}" w:date="${date}"/>`
  const base =
    typeof raw === 'string' && /^<w:trPr(?:\s[^>]*)?>[\s\S]*<\/w:trPr>$/.test(raw)
      ? raw.replace(/<w:(?:ins|del)\b[^>]*\/>|<w:(?:ins|del)\b[^>]*>[\s\S]*?<\/w:(?:ins|del)>/g, '')
      : '<w:trPr></w:trPr>'
  const at = base.search(/<w:trPrChange\b|<\/w:trPr>/)
  return base.slice(0, at) + record + base.slice(at)
}

/** strip a tracked-change marker element from raw trPr/tcPr bytes; empty container → null */
function stripTrackMarker(raw: string | null, tag: string, container: string): string | null {
  if (!raw) return raw
  const out = raw
    .replace(new RegExp(`<${tag}[^>]*/>`, 'g'), '')
    .replace(new RegExp(`<${tag}[^>]*>[\\s\\S]*?</${tag}>`, 'g'), '')
  if (new RegExp(`^<${container}(?:\\s[^>]*)?>\\s*</${container}>$`).test(out)) return null
  return out
}

/** revisions as Word counts them: the two sides of a move are one change */
export function revisionCount(ranges: Array<Pick<RevisionRange, 'moveId'>>): number {
  const moves = new Set<string>()
  let n = 0
  for (const r of ranges) {
    if (r.moveId) moves.add(r.moveId)
    else n++
  }
  return n + moves.size
}

/** the given ranges plus the other side of every move among them (accept / reject act on the pair) */
function withMovePairs(doc: PmNode, ranges: RevisionRange[]): RevisionRange[] {
  if (!ranges.some((r) => r.kind === 'moveFrom' || r.kind === 'moveTo')) return ranges
  const live = collectRevisions(doc)
  const same = (a: RevisionRange, b: RevisionRange) =>
    a.kind === b.kind && a.from === b.from && a.to === b.to
  const ids = new Set<string>()
  for (const r of ranges) {
    const id = r.moveId ?? live.find((l) => same(l, r))?.moveId
    if (id) ids.add(id)
  }
  if (ids.size === 0) return ranges
  const out = [...ranges]
  for (const l of live)
    if (l.moveId && ids.has(l.moveId) && !out.some((o) => same(o, l))) out.push(l)
  return out
}

export function applyRevisions(
  editor: Editor,
  ranges: RevisionRange[],
  mode: 'accept' | 'reject',
): void {
  if (ranges.length === 0) return
  if (mode === 'reject') readShiftsFromGrid(editor, ranges)
  const { state } = editor
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  // cells Shift cells down added in new bottom rows: they go with their row, once its moved cells are back
  const bottomRowCells: RevisionRange[] = []
  const shiftedTables: Array<{ pos: number; step: number }> = []
  // reverse order keeps earlier positions valid while later ranges mutate
  for (const r of withMovePairs(state.doc, ranges).sort((a, b) => b.from - a.from)) {
    if (r.kind === 'blockIns' || r.kind === 'blockDel') {
      const revisionKind = r.kind === 'blockIns' ? 'ins' : 'del'
      const removeNode = mode === 'accept' ? revisionKind === 'del' : revisionKind === 'ins'
      const nodePos = tr.mapping.map(r.from, -1)
      const node = tr.doc.nodeAt(nodePos)
      if (!node) continue
      if (removeNode) {
        if (tr.doc.childCount > 1) tr.delete(nodePos, nodePos + node.nodeSize)
      } else {
        tr.setNodeMarkup(nodePos, undefined, {
          ...node.attrs,
          blockRevision: null,
          paraMarkIns: null,
          paraMarkDel: null,
        })
        // The paragraph carries the same write's run marks (a redline marks
        // both); keeping the block keeps its text, so its marks must go too —
        // otherwise a change the editor no longer lists stays pending forever
        const markType = state.schema.marks[revisionKind]
        if (markType) {
          const from = nodePos
          const to = nodePos + node.nodeSize
          if (tr.doc.rangeHasMark(from, to, markType)) tr.removeMark(from, to, markType)
        }
      }
      continue
    }

    if (r.kind === 'numberingChange' || r.kind === 'sectPrChange') {
      // numberingChange records only the old marker text: nothing to restore, the record goes either way.
      // sectPrChange: the save path restores the previous sectPr bytes from the record
      // -1: a record resolved on this same node just replaced its open token
      const pos = tr.mapping.map(r.from, -1)
      const node = tr.doc.nodeAt(pos)
      const raw = node?.attrs?.[r.kind] as string | null | undefined
      if (!node || !raw) continue
      // the trailing section record has no block of its own: its verdict stays on the host
      const keep = mode === 'reject' || raw.includes('"trailing":true')
      const next =
        !keep || r.kind === 'numberingChange'
          ? null
          : JSON.stringify({ ...JSON.parse(raw), resolved: mode })
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, [r.kind]: next })
      continue
    }

    const propAttr = PROP_CHANGE_ATTR[r.kind]
    if (propAttr) {
      const mapped = tr.mapping.mapResult(r.from, -1)
      const node = tr.doc.nodeAt(mapped.pos)
      const record = node?.attrs?.[propAttr] as { oldXml?: string } | null | undefined
      if (mapped.deleted || !node || !record) continue
      const pos = mapped.pos
      if (mode === 'accept') {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, [propAttr]: null })
        continue
      }
      // reject: the record stays (resolved) so the save restores the old bytes; the
      // modeled subset of the old properties comes back on the node right away
      const resolved = { ...record, resolved: 'reject' }
      const oldXml = record.oldXml ?? ''
      if (r.kind === 'tblPrChange') {
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          ...tblPrRestoreAttrs(oldXml),
          [propAttr]: resolved,
        })
      } else if (r.kind === 'tblGridChange') {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, [propAttr]: resolved })
        const twips = gridColTwips(oldXml)
        const map = TableMap.get(node)
        if (twips.length === map.width) {
          const widths = twips.map((w) => Math.max(1, Math.round(w / 15)))
          reflowColumns(tr, pos, widths, pctOf(node, map, widths))
        }
      } else if (r.kind === 'trPrChange') {
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          ...trPrRestoreAttrs(oldXml),
          [propAttr]: resolved,
        })
      } else if (r.kind === 'tcPrChange') {
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          ...tcPrRestoreAttrs(oldXml),
          [propAttr]: resolved,
        })
      } else if (r.kind === 'cellMerge') {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, [propAttr]: resolved })
        unmergeCellRows(tr, pos)
      } else {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, [propAttr]: resolved })
      }
      continue
    }

    if (r.kind === 'pPrChange') {
      const $pos = state.doc.resolve(r.from + 1)
      if ($pos.parent.attrs && $pos.parent.attrs.pPrChange != null) {
        const nodePos = $pos.before($pos.depth)
        if (mode === 'reject') {
          const oldFormatStr = $pos.parent.attrs.pPrChange as string
          let old: Record<string, unknown> = {}
          try {
            const info = JSON.parse(oldFormatStr)
            old = (info?.old ?? {}) as Record<string, unknown>
            if (info?.oldAlign !== undefined) old = { ...old, align: info.oldAlign }
          } catch {
            /* ignore */
          }
          const oldFormat = (
            old.format && typeof old.format === 'object' ? old.format : old
          ) as Record<string, unknown>
          const restore: Record<string, unknown> = {}
          for (const field of PARAGRAPH_FORMAT_FIELDS) restore[field] = oldFormat[field] ?? null
          // a record read from a file lists every property the paragraph had; an older snapshot may not
          const fromFile = oldFormat === old
          for (const field of PARAGRAPH_FLOW_FIELDS)
            if (fromFile || field in oldFormat) restore[field] = oldFormat[field] ?? null
          const hasStructureSnapshot = typeof old.type === 'string'
          if (hasStructureSnapshot) {
            for (const field of PARAGRAPH_STRUCTURE_FIELDS) restore[field] = old[field] ?? null
          }
          const oldTypeName = typeof old.type === 'string' ? old.type : $pos.parent.type.name
          const oldType = state.schema.nodes[oldTypeName] ?? $pos.parent.type
          tr.setNodeMarkup(nodePos, oldType, {
            ...$pos.parent.attrs,
            pPrChange: null,
            ...restore,
          })
        } else {
          // accept: just clear the pPrChange attr (keep current format)
          tr.setNodeMarkup(nodePos, undefined, { ...$pos.parent.attrs, pPrChange: null })
        }
      }
      continue
    }

    if (r.kind === 'rPrChange') {
      // accept = keep the current format; reject = restore the pre-revision modeled format subset.
      // Both clear the rprChange mark and strip the w:rPrChange record from rawRPr (takes effect on save)
      // an insertion or deletion on the same text may already have moved or removed it
      state.doc.nodesBetween(r.from, r.to, (node, pos) => {
        if (!node.isText) return
        const from = tr.mapping.map(Math.max(pos, r.from), 1)
        const to = tr.mapping.map(Math.min(pos + node.nodeSize, r.to), -1)
        if (from >= to) return
        const rpr = node.marks.find((m) => m.type.name === 'rprChange')
        const ts = node.marks.find((m) => m.type.name === 'docTextStyle')
        if (mode === 'reject' && rpr) {
          const old = (rpr.attrs.old ?? {}) as Record<string, unknown>
          for (const flag of ['bold', 'italic', 'underline', 'strike'] as const) {
            const markType = state.schema.marks[flag]
            if (!markType) continue
            if (old[flag]) tr.addMark(from, to, markType.create())
            else tr.removeMark(from, to, markType)
          }
          const tsType = state.schema.marks.docTextStyle
          if (tsType) {
            const nextAttrs = rejectedTextStyleAttrs(tsType.spec.attrs ?? {}, old, node.text ?? '')
            const oldRPr = rpr.attrs.oldRPr as string | null
            // the recorded previous rPr is what the save writes back; the modeled
            // fields above agree with it, so the merge keeps its bytes
            if (oldRPr != null) nextAttrs.rawRPr = oldRPr ? `<w:rPr>${oldRPr}</w:rPr>` : null
            if (ts) nextAttrs.tblRun = ts.attrs.tblRun ?? null
            const hasAny = Object.values(nextAttrs).some((v) => v !== null && v !== undefined)
            if (ts || hasAny) tr.addMark(from, to, tsType.create(nextAttrs))
          }
        }
        if (ts?.attrs.rawRPr && /<w:rPrChange/.test(String(ts.attrs.rawRPr))) {
          const raw = String(ts.attrs.rawRPr).replace(
            /<w:rPrChange[\s\S]*?<\/w:rPrChange>|<w:rPrChange[^>]*\/>/,
            '',
          )
          // the reject branch already rebuilt docTextStyle above; strip the record here only on accept or when nothing was rebuilt
          const cur = mode === 'reject' ? null : ts
          if (cur) tr.addMark(from, to, cur.type.create({ ...cur.attrs, rawRPr: raw }))
          else if (mode === 'reject') {
            // reject: drop the old format's rawRPr too (modeled fields restored, unmodeled attrs abandoned)
          }
        }
      })
      const mf = tr.mapping.map(r.from, 1)
      const mt = tr.mapping.map(r.to, -1)
      if (mf < mt) tr.removeMark(mf, mt, state.schema.marks.rprChange)
      continue
    }

    if (r.kind === 'paraMarkRPrChange') {
      const pos = tr.mapping.map(r.from, 1)
      const node = tr.doc.nodeAt(pos)
      const raw = node?.attrs?.paraMarkRPrChange as string | null | undefined
      if (!node || !raw) continue
      // the mark's formatting is not modeled: the save path resolves the record in the raw pPr
      const next =
        mode === 'accept' ? null : JSON.stringify({ ...JSON.parse(raw), resolved: 'reject' })
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, paraMarkRPrChange: next })
      continue
    }

    if (r.kind === 'paraMarkDel' || r.kind === 'paraMarkIns') {
      // the mark itself is the revision: accepting a deletion / rejecting an
      // insertion joins the paragraph into the next one (Word keeps the first
      // paragraph's properties); the record leaves either way
      const pos = tr.mapping.map(r.from, 1)
      const node = tr.doc.nodeAt(pos)
      if (!node || !node.attrs?.[r.kind]) continue
      // a moved mark (w:pPr/w:rPr/w:moveFrom|moveTo) also flagged the paragraph as a move
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, [r.kind]: null, moveRevision: null })
      const join = (mode === 'accept') === (r.kind === 'paraMarkDel')
      const after = pos + node.nodeSize
      if (join && tr.doc.resolve(after).nodeAfter?.isTextblock && canJoin(tr.doc, after))
        tr.join(after)
      continue
    }

    if (
      r.kind === 'rowIns' ||
      r.kind === 'rowDel' ||
      r.kind === 'cellIns' ||
      r.kind === 'cellDel'
    ) {
      const isRow = r.kind === 'rowIns' || r.kind === 'rowDel'
      const revKind = r.kind === 'rowIns' || r.kind === 'cellIns' ? 'ins' : 'del'
      const removeNode = mode === 'accept' ? revKind === 'del' : revKind === 'ins'
      // a whole-column removal below shifts the rest of the table: positions must be mapped
      const mapped = tr.mapping.mapResult(r.from, 1)
      const from = mapped.pos
      const to = tr.mapping.map(r.to, -1)
      const node = tr.doc.nodeAt(from)
      if (mapped.deleted || !node || to <= from) continue
      if (removeNode && isRow) {
        // accepting a deleted row / rejecting an inserted row = remove the whole row; if it is the only row, remove the whole table
        const $pos = tr.doc.resolve(from)
        if ($pos.parent.childCount <= 1) tr.delete($pos.before(), $pos.after())
        else if (!removeRowAmongMerges(tr, $pos.before(), $pos.index())) tr.delete(from, to)
      } else if (removeNode) {
        const cellPos = from
        const cell = node
        if (!cell.attrs.cellRevision) continue
        const $cell = tr.doc.resolve(cellPos)
        const tablePos = $cell.before($cell.depth - 1)
        const table = $cell.node($cell.depth - 1)
        const map = TableMap.get(table)
        const rect = map.findCell(cellPos - tablePos - 1)
        const shift = (cell.attrs.cellRevision as { shift?: string }).shift
        const wholeColumn =
          !shift &&
          map.width > 1 &&
          rect.right - rect.left === 1 &&
          Array.from({ length: map.height }, (_, row) => row).every(
            (row) =>
              // a cell of a row that is itself inserted or deleted goes or stays with that row
              (table.child(row).attrs.rowRevision as { kind?: string } | null)?.kind === revKind ||
              (
                table.nodeAt(map.map[row * map.width + rect.left]!)?.attrs.cellRevision as {
                  kind?: string
                } | null
              )?.kind === revKind,
          )
        if (wholeColumn) {
          const widths = columnWidthsPx(table, map)
          const pct = pctOf(table, map, widths)
          removeColumn(tr, { ...rect, map, table, tableStart: tablePos + 1 }, rect.left)
          reflowColumns(
            tr,
            tablePos,
            widths?.filter((_, col) => col !== rect.left) ?? null,
            pct.filter((_, col) => col !== rect.left),
          )
          continue
        }
        if (shift === 'row') {
          bottomRowCells.push(r)
          continue
        }
        if (shift === 'down' && closeColumnOver(tr, tablePos, cellPos)) {
          shiftedTables.push({ pos: tablePos, step: tr.steps.length })
          continue
        }
        if (shift === 'up' && closeColumnOver(tr, tablePos, cellPos)) continue
        if (!shift && closeRowOver(tr, tablePos, cellPos)) continue
        // a merged cell is in the way of the others: the cell stays, emptied, without its mark
        clearCellInPlace(tr, cellPos, revKind)
      } else {
        // keep the content: clear the revision marker and strip the record from the raw trPr/tcPr bytes (takes effect on save)
        const attrs = isRow
          ? {
              ...node.attrs,
              rowRevision: null,
              rawTrPr: stripTrackMarker(
                node.attrs.rawTrPr as string | null,
                revKind === 'ins' ? 'w:ins' : 'w:del',
                'w:trPr',
              ),
            }
          : {
              ...node.attrs,
              cellRevision: null,
              rawTcPr: stripTrackMarker(
                node.attrs.rawTcPr as string | null,
                revKind === 'ins' ? 'w:cellIns' : 'w:cellDel',
                'w:tcPr',
              ),
            }
        tr.setNodeMarkup(from, undefined, attrs)
        const markType = state.schema.marks[revKind]
        if (markType) tr.removeMark(from + 1, to - 1, markType)
      }
      continue
    }

    // moved text is a deletion at its source and an insertion at its destination
    const textKind = r.kind === 'moveFrom' ? 'del' : r.kind === 'moveTo' ? 'ins' : r.kind
    if (r.kind === 'moveFrom' || r.kind === 'moveTo') {
      const block = state.doc.nodeAt(r.from)
      if (block && PARAGRAPH_NODE_TYPES.has(block.type.name) && r.to === r.from + block.nodeSize) {
        // whole-paragraph move with no run-level marks: the paragraph goes or stays as a block
        const pos = tr.mapping.map(r.from, 1)
        const node = tr.doc.nodeAt(pos)
        if (!node?.attrs?.moveRevision) continue
        const drop = (mode === 'accept') === (r.kind === 'moveFrom')
        if (drop && tr.doc.resolve(pos).parent.childCount > 1) tr.delete(pos, pos + node.nodeSize)
        else tr.setNodeMarkup(pos, undefined, { ...node.attrs, moveRevision: null })
        continue
      }
      const $para = state.doc.resolve(r.from)
      if ($para.parent.attrs?.moveRevision)
        tr.setNodeMarkup($para.before(), undefined, { ...$para.parent.attrs, moveRevision: null })
    }
    // an ins and a del revision can share the same text: map through earlier steps
    const from = tr.mapping.map(r.from, 1)
    const to = tr.mapping.map(r.to, -1)
    if (from >= to) continue
    const removeText = mode === 'accept' ? textKind !== 'ins' : textKind !== 'del'
    if (removeText) {
      // a revision covering a block's entire content removes the block itself
      // (Word: rejecting an inserted paragraph deletes the paragraph, not just
      // its text), unless it is the document's last remaining block
      const $from = tr.doc.resolve(from)
      const $to = tr.doc.resolve(to)
      if (
        $from.parent.isTextblock &&
        $from.sameParent($to) &&
        from === $from.start() &&
        to === $to.end() &&
        $from.depth === 1 &&
        tr.doc.childCount > 1
      ) {
        tr.delete($from.before(), $to.after())
      } else {
        tr.delete(from, to)
      }
    } else if (mode === 'accept') {
      tr.removeMark(from, to, state.schema.marks.ins)
    } else {
      tr.removeMark(from, to, state.schema.marks.del)
    }
  }
  for (const r of bottomRowCells) {
    const mapped = tr.mapping.mapResult(r.from, 1)
    const cell = tr.doc.nodeAt(mapped.pos)
    if (mapped.deleted || !cell?.attrs.cellRevision) continue
    const $cell = tr.doc.resolve(mapped.pos)
    const tablePos = $cell.before($cell.depth - 1)
    const row = $cell.index($cell.depth - 1)
    if (dropInsertedRow(tr, tablePos, row)) continue
    const start = $cell.start()
    const rest: number[] = []
    $cell.parent.forEach((other, offset) => {
      const rev = other.attrs.cellRevision as { kind?: string; shift?: string } | null
      if (rev?.kind === 'ins' && rev.shift === 'row') rest.unshift(start + offset)
    })
    for (const pos of rest) clearCellInPlace(tr, pos, 'ins')
  }
  for (const { pos, step } of shiftedTables)
    settleShiftedRows(tr, tr.mapping.slice(step).map(pos, 1))
  releaseHeldNames(tr)
  editor.view.dispatch(tr)
}

/**
 * A rejected conversion: a paragraph that is no longer struck takes back the
 * bookmarks it handed to the converted content, each once that content no
 * longer has it; until then it keeps holding the name.
 */
function releaseHeldNames(tr: Transaction): void {
  const holders: Array<{ node: PmNode; pos: number }> = []
  const used = new Set<string>()
  const walk = (node: PmNode, pos: number, struck: boolean): void => {
    node.forEach((child, offset) => {
      const at = pos + offset
      const gone =
        struck ||
        !!child.attrs.paraMarkDel ||
        (child.attrs.rowRevision as { kind?: string } | null)?.kind === 'del'
      for (const attr of ['bookmarks', 'hiddenBookmarks'] as const)
        for (const name of (child.attrs[attr] as string[] | null) ?? []) used.add(name)
      if (child.attrs.bookmarksHeld && !gone) holders.push({ node: child, pos: at })
      if (!child.isTextblock) walk(child, at + 1, gone)
    })
  }
  walk(tr.doc, 0, false)
  for (const { node, pos } of holders) {
    const held = node.attrs.bookmarksHeld as Record<string, string[] | undefined>
    const kept: Record<string, string[]> = {}
    for (const attr of ['bookmarks', 'hiddenBookmarks'] as const) {
      const names = held[attr] ?? []
      const back = names.filter((name) => !used.has(name))
      const waiting = names.filter((name) => used.has(name))
      if (back.length > 0) {
        tr.setNodeAttribute(pos, attr, [...((node.attrs[attr] as string[] | null) ?? []), ...back])
        for (const name of back) used.add(name)
      }
      if (waiting.length > 0) kept[attr] = waiting
    }
    tr.setNodeAttribute(pos, 'bookmarksHeld', Object.keys(kept).length > 0 ? kept : null)
  }
}

/**
 * A row that goes from among cells merged down: a cell that runs through it is
 * one row shorter, and one that starts in it is left to the rows below. False
 * when no merged cell crosses the row, which then goes as it is.
 */
function removeRowAmongMerges(tr: Transaction, tablePos: number, row: number): boolean {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  if (row >= map.height) return false
  const crossed = Array.from({ length: map.width }, (_, col) => {
    const rect = map.findCell(map.map[row * map.width + col]!)
    return rect.bottom - rect.top > 1
  }).some(Boolean)
  if (!crossed) return false
  const tableStart = tablePos + 1
  let rowPos = tableStart
  for (let i = 0; i < row; i++) rowPos += table.child(i).nodeSize
  const mapFrom = tr.mapping.maps.length
  tr.delete(rowPos, rowPos + table.child(row).nodeSize)
  const seen = new Set<number>()
  for (let col = 0; col < map.width; col++) {
    const index = row * map.width + col
    const offset = map.map[index]!
    if (seen.has(offset)) continue
    seen.add(offset)
    const cell = table.nodeAt(offset)!
    const rowspan = (cell.attrs.rowspan as number) - 1
    if (row > 0 && offset === map.map[index - map.width]) {
      // an attribute step: positions other pending changes hold on the cell stay as they are
      tr.setNodeAttribute(tr.mapping.slice(mapFrom).map(tableStart + offset), 'rowspan', rowspan)
    } else if (row + 1 < map.height && offset === map.map[index + map.width]) {
      // what the row held goes with it: the rows below keep the merged cell, empty
      const rest = cell.type.createAndFill({ ...cell.attrs, rowspan })
      const at = tableStart + map.positionAt(row + 1, col, table)
      if (rest) tr.insert(tr.mapping.slice(mapFrom).map(at), rest)
    }
  }
  return true
}

const shiftOf = (cell: PmNode): string | undefined =>
  (cell.attrs.cellRevision as { shift?: string } | null)?.shift

type CellRevision = { kind?: string; author?: string; date?: string; shift?: string }

/**
 * Which way inserted cells pushed the others is not in the file, but Shift
 * cells down leaves a shape of its own: rows added at the bottom whose cells
 * are all inserted except in the columns that moved, and as many inserted
 * cells higher up in exactly those columns. Null when the grid is not that.
 */
export function downShiftsOf(table: PmNode): Map<number, 'down' | 'row'> | null {
  const map = TableMap.get(table)
  const revOf = (offset: number) => table.nodeAt(offset)!.attrs.cellRevision as CellRevision | null
  const inserted = (offset: number) => revOf(offset)?.kind === 'ins'
  const at = (row: number, col: number) => map.map[row * map.width + col]!
  const pending = [...new Set(map.map)].filter(inserted)
  const first = pending.length > 0 ? revOf(pending[0]!)! : null
  if (
    !first ||
    pending.some((offset) => {
      const rev = revOf(offset)!
      return rev.shift || rev.author !== first.author || rev.date !== first.date
    })
  )
    return null
  const bottom = map.height - 1
  const moved = Array.from({ length: map.width }, (_, col) => col).filter(
    (col) => !inserted(at(bottom, col)) && !isGap(table.nodeAt(at(bottom, col))),
  )
  const contiguous = moved.every((col, i) => i === 0 || col === moved[i - 1]! + 1)
  if (moved.length === 0 || moved.length === map.width || !contiguous) return null
  const left = moved[0]!
  const right = left + moved.length
  // an added row: its own cells only, inserted (or a placeholder) outside the moved columns
  const added = (row: number): boolean => {
    let any = false
    for (let col = 0; col < map.width; col++) {
      const rect = map.findCell(at(row, col))
      if (rect.bottom - rect.top !== 1) return false
      const fresh = inserted(at(row, col))
      if (col >= left && col < right) {
        if (fresh || rect.left < left || rect.right > right) return false
      } else if (fresh) any = true
      else if (!isGap(table.nodeAt(at(row, col)))) return false
    }
    return any
  }
  let height = 0
  while (height < map.height - 1 && added(bottom - height)) height++
  if (height === 0) return null
  const out = new Map<number, 'down' | 'row'>()
  let rows: number[] | null = null
  for (let col = 0; col < map.width; col++) {
    const own: number[] = []
    for (let row = 0; row <= bottom - height; row++) {
      if (!inserted(at(row, col))) continue
      const rect = map.findCell(at(row, col))
      if (rect.bottom - rect.top !== 1 || rect.left < left || rect.right > right) return null
      own.push(row)
      out.set(at(row, col), 'down')
    }
    if (col < left || col >= right) continue
    const block = own.length === height && own.every((row, i) => row === own[0]! + i)
    if (!block || (rows && rows[0] !== own[0])) return null
    rows = own
  }
  for (let row = bottom - height + 1; row <= bottom; row++)
    for (let col = 0; col < map.width; col++)
      if (inserted(at(row, col))) out.set(at(row, col), 'row')
  return out
}

/** before a rejection: the cells of a reopened file are told the way they came, as far as their table shows it */
function readShiftsFromGrid(editor: Editor, ranges: RevisionRange[]): void {
  const { doc } = editor.state
  const tr = editor.state.tr
  const seen = new Set<number>()
  for (const r of ranges) {
    if (r.kind !== 'cellIns') continue
    const cell = doc.nodeAt(r.from)
    if (!cell?.attrs.cellRevision || shiftOf(cell)) continue
    const $cell = doc.resolve(r.from)
    const tablePos = $cell.before($cell.depth - 1)
    if (seen.has(tablePos)) continue
    seen.add(tablePos)
    for (const [offset, shift] of downShiftsOf($cell.node($cell.depth - 1)) ?? []) {
      const pos = tablePos + 1 + offset
      const node = doc.nodeAt(pos)!
      tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        cellRevision: { ...node.attrs.cellRevision, shift },
      })
    }
  }
  // its own step, out of the history: the rejection below maps positions through its own steps only
  if (tr.docChanged)
    editor.view.dispatch(tr.setMeta(TRACK_IGNORE, true).setMeta('addToHistory', false))
}

/**
 * Once no cell shifted in downward is left pending, the rows the insertion
 * added go where nothing else is in them, and what stays of them is plain.
 */
function settleShiftedRows(tr: Transaction, tablePos: number): void {
  const table = tr.doc.nodeAt(tablePos)
  if (table?.type.name !== 'docTable') return
  let down = false
  table.forEach((row) => row.forEach((cell) => (down ||= shiftOf(cell) === 'down')))
  if (down) return
  for (let row = table.childCount - 1; row >= 0; row--) {
    const live = tr.doc.nodeAt(tablePos)!
    if (row >= live.childCount) continue
    let at = tablePos + 1
    for (let i = 0; i < row; i++) at += live.child(i).nodeSize
    const cells: number[] = []
    live.child(row).forEach((cell, offset) => {
      if (shiftOf(cell) === 'row') cells.unshift(at + 1 + offset)
    })
    if (cells.length === 0 || dropInsertedRow(tr, tablePos, row)) continue
    for (const pos of cells) clearCellInPlace(tr, pos, 'ins')
  }
}

function clearCellInPlace(tr: Transaction, cellPos: number, revKind: 'ins' | 'del'): void {
  const cell = tr.doc.nodeAt(cellPos)!
  tr.setNodeMarkup(cellPos, undefined, {
    ...cell.attrs,
    cellRevision: null,
    rawTcPr: stripTrackMarker(
      cell.attrs.rawTcPr as string | null,
      revKind === 'ins' ? 'w:cellIns' : 'w:cellDel',
      'w:tcPr',
    ),
  })
  const empty = cell.type.schema.nodes.docParagraph.createAndFill()
  if (empty) tr.replaceWith(cellPos + 1, cellPos + cell.nodeSize - 1, empty)
}

const isGap = (cell: PmNode | null | undefined): boolean => cell?.attrs.gridGap === true

/** a grid row no cell reaches into from above and none leaves downward */
function ownRow(table: PmNode, map: TableMap, row: number): PmNode[] | null {
  const cells: PmNode[] = []
  for (let col = 0; col < map.width;) {
    const rect = map.findCell(map.map[row * map.width + col]!)
    if (rect.top !== row || rect.bottom !== row + 1) return null
    cells.push(table.nodeAt(map.map[row * map.width + col]!)!)
    col = rect.right
  }
  return cells
}

/** the row goes when nothing but inserted cells and placeholders is left in it */
function dropInsertedRow(tr: Transaction, tablePos: number, row: number): boolean {
  const table = tr.doc.nodeAt(tablePos)!
  const cells = ownRow(table, TableMap.get(table), row)
  const inserted = (cell: PmNode) =>
    isGap(cell) || (cell.attrs.cellRevision as { kind?: string } | null)?.kind === 'ins'
  if (!cells || table.childCount < 2 || !cells.every(inserted) || cells.every(isGap)) return false
  let pos = tablePos + 1
  for (let i = 0; i < row; i++) pos += table.child(i).nodeSize
  tr.delete(pos, pos + table.child(row).nodeSize)
  return true
}

/**
 * A rejected inserted cell (or an accepted deleted one) leaves its row, as
 * the w:tc leaves the w:tr: the cells after it close up and a placeholder
 * takes the row's end. False when a vertically merged cell is among them.
 */
function closeRowOver(tr: Transaction, tablePos: number, cellPos: number): boolean {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const rect = map.findCell(cellPos - tablePos - 1)
  for (let col = rect.left; col < map.width;) {
    const next = map.findCell(map.map[rect.top * map.width + col]!)
    if (next.top !== rect.top || next.bottom !== rect.top + 1) return false
    col = next.right
  }
  const cell = table.nodeAt(cellPos - tablePos - 1)!
  const span = rect.right - rect.left
  const widths = columnWidthsPx(table, map)
  const pct = pctOf(table, map, widths)
  const gap = cell.type.create(
    {
      ...cell.attrs,
      gridGap: true,
      fill: null,
      borders: null,
      cellRevision: null,
      rawTcPr: stripTrackMarker(
        stripTrackMarker(cell.attrs.rawTcPr as string | null, 'w:cellIns', 'w:tcPr'),
        'w:cellDel',
        'w:tcPr',
      ),
      colwidth: widths ? widths.slice(map.width - span) : null,
    },
    cell.type.schema.nodes.docParagraph.create(),
  )
  tr.insert(tr.doc.resolve(cellPos).end(), gap)
  tr.delete(cellPos, cellPos + cell.nodeSize)

  let shorter = tr.doc.nodeAt(tablePos)!
  const rowCells = ownRow(shorter, TableMap.get(shorter), rect.top)
  if (rowCells?.every(isGap)) {
    if (shorter.childCount < 2) {
      tr.delete(tablePos, tablePos + shorter.nodeSize)
      return true
    }
    let pos = tablePos + 1
    for (let i = 0; i < rect.top; i++) pos += shorter.child(i).nodeSize
    tr.delete(pos, pos + shorter.child(rect.top).nodeSize)
    shorter = tr.doc.nodeAt(tablePos)!
  }
  // columns every row now ends short of are the ones the insertion added
  const trailing = (t: PmNode, m: TableMap, row: number) => {
    let count = 0
    while (count < m.width && isGap(t.nodeAt(m.map[row * m.width + m.width - 1 - count]!))) count++
    return count
  }
  let spare = Infinity
  const shortMap = TableMap.get(shorter)
  for (let row = 0; row < shortMap.height; row++)
    spare = Math.min(spare, trailing(shorter, shortMap, row))
  if (spare === 0 || spare >= shortMap.width) return true
  for (let i = 0; i < spare; i++) {
    const t = tr.doc.nodeAt(tablePos)!
    const m = TableMap.get(t)
    removeColumn(
      tr,
      {
        top: 0,
        left: m.width - 1,
        bottom: m.height,
        right: m.width,
        map: m,
        table: t,
        tableStart: tablePos + 1,
      },
      m.width - 1,
    )
  }
  reflowColumns(
    tr,
    tablePos,
    widths?.slice(0, map.width - spare) ?? null,
    pct.slice(0, map.width - spare),
  )
  return true
}

/**
 * A cell Shift cells down put in goes back out: the cells below it move up
 * one row, and the row the insertion added at the bottom goes once nothing
 * but its own cells is left in it. False when a merged cell is in the column.
 */
function closeColumnOver(tr: Transaction, tablePos: number, cellPos: number): boolean {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const rect = map.findCell(cellPos - tablePos - 1)
  const column: number[] = []
  for (let row = rect.top; row < map.height; row++) {
    const offset = map.map[row * map.width + rect.left]!
    const each = map.findCell(offset)
    if (each.left !== rect.left || each.right !== rect.left + 1 || each.bottom - each.top !== 1)
      return false
    column.push(offset)
  }
  const cell = table.nodeAt(column[0]!)!
  const blank = cell.type.schema.nodes.docParagraph.create()
  const last = table.nodeAt(column[column.length - 1]!)!
  const deleted = (cell.attrs.cellRevision as { kind?: string } | null)?.kind === 'del'
  const filler = deleted
    ? // Delete Cells, shift up: the column keeps its length, a blank cell like the last takes the end
      last.type.create(
        {
          ...last.attrs,
          cellRevision: null,
          rawTcPr: stripTrackMarker(
            stripTrackMarker(last.attrs.rawTcPr as string | null, 'w:cellIns', 'w:tcPr'),
            'w:cellDel',
            'w:tcPr',
          ),
        },
        blank,
      )
    : // the cell left over at the bottom belongs to the row the insertion added
      cell.type.create(
        { ...cell.attrs, cellRevision: { ...cell.attrs.cellRevision, shift: 'row' } },
        blank,
      )
  for (let i = column.length - 1; i >= 0; i--) {
    const at = tablePos + 1 + column[i]!
    const next = i === column.length - 1 ? filler : table.nodeAt(column[i + 1]!)!
    tr.replaceWith(at, at + table.nodeAt(column[i]!)!.nodeSize, next)
  }
  if (!deleted) dropInsertedRow(tr, tablePos, map.height - 1)
  return true
}

/**
 * Rejecting a w:cellMerge: the vertical merge is undone — the restart cell keeps its
 * content at rowspan 1 and every row it spanned gets an empty cell back in its column.
 */
function unmergeCellRows(tr: Transaction, cellPos: number): void {
  const $cell = tr.doc.resolve(cellPos)
  const table = $cell.node(-1)
  const tableStart = $cell.start(-1)
  const map = TableMap.get(table)
  const rect = map.findCell(cellPos - tableStart)
  if (rect.bottom - rect.top <= 1) return
  const cell = tr.doc.nodeAt(cellPos)!
  tr.setNodeMarkup(cellPos, undefined, { ...cell.attrs, rowspan: 1 })
  const rowEnds: number[] = []
  let offset = 0
  table.forEach((row) => {
    offset += row.nodeSize
    rowEnds.push(tableStart + offset - 1)
  })
  // bottom-up so earlier insertions do not shift the later ones
  for (let row = rect.bottom - 1; row > rect.top; row--) {
    let insertAt = rowEnds[row]
    for (let col = rect.right; col < map.width; col++) {
      const idx = row * map.width + col
      const pos = map.map[idx]!
      if (row > 0 && map.map[idx - map.width] === pos) continue
      if (col > 0 && map.map[idx - 1] === pos) continue
      insertAt = tableStart + pos
      break
    }
    const empty = cell.type.createAndFill({
      colspan: rect.right - rect.left,
      colwidth: cell.attrs.colwidth,
    })
    if (empty) tr.insert(insertAt, empty)
  }
}

export function acceptAllRevisions(editor: Editor): void {
  applyRevisions(editor, collectRevisions(editor.state.doc), 'accept')
}

export function rejectAllRevisions(editor: Editor): void {
  applyRevisions(editor, collectRevisions(editor.state.doc), 'reject')
}

/** accept / reject only the revisions of one author (e.g. the AI assistant) */
export function applyRevisionsBy(editor: Editor, author: string, mode: 'accept' | 'reject'): void {
  applyRevisions(
    editor,
    collectRevisions(editor.state.doc).filter((r) => r.author === author),
    mode,
  )
}

/** the revision containing the selection head, else the next one, else the first */
function currentRevision(editor: Editor): RevisionRange | null {
  const ranges = collectRevisions(editor.state.doc)
  if (ranges.length === 0) return null
  const head = editor.state.selection.head
  return (
    ranges.find((r) => head >= r.from && head <= r.to) ??
    ranges.find((r) => r.from > head) ??
    ranges[0]
  )
}

export function acceptCurrentRevision(editor: Editor): boolean {
  const r = currentRevision(editor)
  if (!r) return false
  applyRevisions(editor, [r], 'accept')
  return true
}

export function rejectCurrentRevision(editor: Editor): boolean {
  const r = currentRevision(editor)
  if (!r) return false
  applyRevisions(editor, [r], 'reject')
  return true
}

/** the next / previous revision after the selection (wraps around) */
export function nextRevision(editor: Editor, dir: 1 | -1): RevisionRange | null {
  const ranges = collectRevisions(editor.state.doc)
  if (ranges.length === 0) return null
  // compare range starts so adjacent revisions (shared boundary) still advance
  const anchor = editor.state.selection.from
  return dir === 1
    ? (ranges.find((r) => r.from > anchor) ?? ranges[0])
    : ([...ranges].reverse().find((r) => r.from < anchor) ?? ranges[ranges.length - 1])
}

export function selectRevision(editor: Editor, target: RevisionRange): void {
  const tr = editor.state.tr
  // between: row/cell-level revision range endpoints are not text positions; snap to the nearest selectable text position
  tr.setSelection(TextSelection.between(tr.doc.resolve(target.from), tr.doc.resolve(target.to)))
  tr.scrollIntoView()
  tr.setMeta(TRACK_IGNORE, true)
  editor.view.dispatch(tr)
}

/** move the selection to the next / previous revision (wraps around) */
export function gotoRevision(editor: Editor, dir: 1 | -1): RevisionRange | null {
  const target = nextRevision(editor, dir)
  if (target) selectRevision(editor, target)
  return target
}

const SENTENCE_END = /[.!?\u2026\u3002\uff01\uff1f\uff0e]["'\u2019\u201d)\]\u300d\u300f\uff09]*\s*$/

/**
 * Word pairs a cut and a paste as a move only for whole sentences or paragraphs:
 * the range starts at a paragraph or sentence start and ends at a sentence or paragraph end.
 */
export function coversWholeSentences(doc: PmNode, from: number, to: number): boolean {
  const $from = doc.resolve(from)
  const $to = doc.resolve(to)
  // a range between blocks takes whole paragraphs
  if (!$from.parent.isTextblock || !$to.parent.isTextblock) return true
  const same = $from.sameParent($to)
  // a note mark or other inline object at a boundary is not text: the sentence still ends before it
  const textOf = (parent: PmNode, a: number, b: number) => parent.textBetween(a, b, undefined, '')
  const before = textOf($from.parent, 0, $from.parentOffset)
  if (before.trim() !== '' && !SENTENCE_END.test(before)) return false
  if (!same && $to.parentOffset === 0) return true
  const after = textOf($to.parent, $to.parentOffset, $to.parent.content.size)
  if (after.trim() === '') return true
  return SENTENCE_END.test(textOf($to.parent, same ? $from.parentOffset : 0, $to.parentOffset))
}

export interface TrackChangesStorage {
  enabled: boolean
  author: string
  /** Track moves: a cut + paste of whole sentences pairs as a move (off = plain deletion + insertion) */
  trackMoves: boolean
  /** Track Changes Options: live format edits record w:rPrChange / w:pPrChange */
  trackFormatting: boolean
}

declare module '@tiptap/core' {
  interface Storage {
    trackChanges: TrackChangesStorage
  }
}

/**
 * Records edits as tracked changes while `storage.enabled`:
 * - inserted content gets the ins mark
 * - deleted inline content is re-inserted struck through with the del mark
 *   (deleting your own tracked insertion really deletes it)
 *
 * Block-level deletions (paragraph merges) are not re-materialized — only the
 * inline text of the change is tracked. Accept/reject and save fidelity work
 * on run-level w:ins / w:del, matching the engine's revision model.
 */
export const TrackChangesExtension = Extension.create<object, TrackChangesStorage>({
  name: 'trackChanges',

  addStorage() {
    return { enabled: false, author: authorName(), trackMoves: true, trackFormatting: true }
  },

  addKeyboardShortcuts() {
    const storage = this.storage
    const moveCaret = (pos: number) =>
      this.editor.commands.command(({ tr, dispatch }) => {
        tr.setSelection(TextSelection.create(tr.doc, pos))
        if (dispatch) dispatch(tr)
        return true
      })
    /**
     * Backspace / Delete over text already marked deleted is a no-op (the
     * recorder puts it straight back), which would trap the caret. Word skips
     * such text, so step over the struck run and delete the live character
     * beyond it instead.
     */
    const skipStruckRun = (dir: -1 | 1): boolean => {
      if (!storage.enabled) return false
      const { state, view } = this.editor
      const { selection, schema } = state
      const delType = schema.marks.del
      if (!delType || !selection.empty) return false
      const $pos = selection.$from
      const size = $pos.parent.content.size
      if (dir === -1 ? $pos.parentOffset === 0 : $pos.parentOffset >= size) return false
      const isStruck = (offset: number): boolean => {
        const index = $pos.parent.childAfter(offset).index
        const child = $pos.parent.maybeChild(index)
        return !!child?.marks.some((mark) => mark.type === delType)
      }
      if (!isStruck(dir === -1 ? $pos.parentOffset - 1 : $pos.parentOffset)) return false
      const blockStart = $pos.start()
      let offset = $pos.parentOffset + dir
      if (dir === -1) while (offset > 0 && isStruck(offset - 1)) offset -= 1
      else while (offset < size && isStruck(offset)) offset += 1
      const runEdge = blockStart + offset
      if (offset === 0 || offset === size) {
        // nothing live left in that direction: park the caret at the run's edge
        view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, runEdge)))
        return true
      }
      moveCaret(runEdge)
      return dir === -1
        ? this.editor.commands.deleteRange({ from: runEdge - 1, to: runEdge })
        : this.editor.commands.deleteRange({ from: runEdge, to: runEdge + 1 })
    }
    /**
     * Backspace at the start of a paragraph / Delete at its predecessor's end
     * removes the first paragraph's mark. Under tracking that mark becomes a
     * pending deletion and both paragraphs stay; only this author's own pending
     * mark insertion really joins. A mark already struck is stepped over.
     */
    const joinParagraphMark = (dir: -1 | 1): boolean => {
      if (!storage.enabled) return false
      const { state, view } = this.editor
      const { selection } = state
      if (!selection.empty) return false
      const $pos = selection.$from
      if (!PARAGRAPH_NODE_TYPES.has($pos.parent.type.name)) return false
      const atEdge =
        dir === -1 ? $pos.parentOffset === 0 : $pos.parentOffset === $pos.parent.content.size
      if (!atEdge) return false
      const boundary = dir === -1 ? $pos.before() : $pos.after()
      const $boundary = state.doc.resolve(boundary)
      const first = dir === -1 ? $boundary.nodeBefore : $pos.parent
      const second = dir === -1 ? $pos.parent : $boundary.nodeAfter
      if (
        !first ||
        !second ||
        !PARAGRAPH_NODE_TYPES.has(first.type.name) ||
        !PARAGRAPH_NODE_TYPES.has(second.type.name)
      )
        return false
      const firstPos = boundary - first.nodeSize
      const firstEnd = boundary - 1
      if (recordAuthor(first.attrs.paraMarkIns) === storage.author) {
        if (!canJoin(state.doc, boundary)) return false
        const tr = state.tr.setMeta(TRACK_IGNORE, true)
        tr.join(boundary)
        // the surviving mark is the second paragraph's: its revision state carries over
        tr.setNodeMarkup(firstPos, undefined, {
          ...first.attrs,
          paraMarkIns: second.attrs.paraMarkIns,
          paraMarkDel: second.attrs.paraMarkDel,
          paraMarkRPrChange: second.attrs.paraMarkRPrChange,
        })
        tr.setSelection(TextSelection.create(tr.doc, firstEnd))
        view.dispatch(tr)
        return true
      }
      if (first.attrs.paraMarkDel) {
        if (dir === -1) {
          moveCaret(firstEnd)
          if (first.content.size === 0) return true
          return (
            skipStruckRun(-1) ||
            this.editor.commands.deleteRange({ from: firstEnd - 1, to: firstEnd })
          )
        }
        const secondStart = boundary + 1
        moveCaret(secondStart)
        if (second.content.size === 0) return true
        return (
          skipStruckRun(1) ||
          this.editor.commands.deleteRange({ from: secondStart, to: secondStart + 1 })
        )
      }
      const tr = state.tr.setMeta(TRACK_IGNORE, true)
      tr.setNodeMarkup(firstPos, undefined, {
        ...first.attrs,
        paraMarkDel: markRecord(storage.author, revisionDateNow()),
      })
      tr.setSelection(TextSelection.create(tr.doc, firstEnd))
      view.dispatch(tr)
      return true
    }
    return {
      Backspace: () => joinParagraphMark(-1) || skipStruckRun(-1),
      Delete: () => joinParagraphMark(1) || skipStruckRun(1),
    }
  },

  addProseMirrorPlugins() {
    const storage = this.storage
    const editor = this.editor
    // edits made while the IME is composing; recorded at compositionend so the
    // ins mark never rewrites the composition's DOM (Chromium aborts otherwise).
    // base/mapping snapshot the pre-composition doc: the flush diffs the final
    // span against it, because PM's IME DOM re-reads can replace far more than
    // the composed text (runs whose marks don't round-trip the DOM)
    let pendingIme: { from: number; to: number; base: PmNode; mapping: Mapping } | null = null
    // Word records cut + paste (and drag-drop) of the same content as a move:
    // the cut's struck text is remembered until the next edit; a paste of the
    // same content pairs it with the insertion under one move name
    let lastCut: {
      key: string
      delMark: PmMark
      ranges: Array<{ from: number; to: number }>
      /** the doc the ranges are positions in: later transactions map them forward */
      doc: PmNode
    } | null = null
    // the custom paste lanes dispatch without PM's uiEvent meta: the DOM event arms the next recording
    let armed: { event: 'cut' | 'paste' | 'drop'; at: number } | null = null
    const arm = (event: 'cut' | 'paste' | 'drop') => () => {
      armed = { event, at: Date.now() }
      return false
    }
    const sliceKey = (fragment: Fragment) =>
      fragment.textBetween(0, fragment.size, '\u2029', '\ufffc')
    const newMoveId = () => `move${Date.now() % 1e9}${Math.floor(Math.random() * 1e3)}`
    // inline content of [from,to) as 1 char per position, else null (not a
    // single textblock / exotic inline node): lets the flush diff by text
    const inlineText = (
      doc: PmNode,
      from: number,
      to: number,
    ): { text: string; children: PmNode[] } | null => {
      if (from >= to) return { text: '', children: [] }
      const $from = doc.resolve(from)
      if (!$from.sameParent(doc.resolve(to)) || !$from.parent.isTextblock) return null
      let ok = true
      let text = ''
      const children: PmNode[] = []
      doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isInline) return true
        if (node.isText) {
          const a = Math.max(from - pos, 0)
          const b = Math.min(to - pos, node.text!.length)
          const cut = a > 0 || b < node.text!.length ? node.cut(a, b) : node
          text += cut.text
          children.push(cut)
        } else if (node.isLeaf && node.nodeSize === 1) {
          text += '￼'
          children.push(node)
        } else ok = false
        return false
      })
      return ok ? { text, children } : null
    }
    return [
      new Plugin({
        key: recorderKey,
        storage,
        props: {
          handleDOMEvents: {
            cut: arm('cut'),
            paste: arm('paste'),
            drop: arm('drop'),
            copy: () => {
              lastCut = null
              return false
            },
            // PM keeps the doc in sync during composition and often dispatches
            // nothing at compositionend; poke a transaction to flush pendingIme
            compositionend: (view) => {
              setTimeout(() => {
                if (pendingIme && !view.isDestroyed && !view.composing) view.dispatch(view.state.tr)
              }, 0)
              return false
            },
          },
        },
        appendTransaction: (transactions, oldState, newState) => {
          if (!storage.enabled) {
            pendingIme = null
            return null
          }
          if (pendingIme) {
            for (const transaction of transactions) {
              pendingIme.from = transaction.mapping.map(pendingIme.from, -1)
              pendingIme.to = transaction.mapping.map(pendingIme.to, 1)
              pendingIme.mapping.appendMapping(transaction.mapping)
            }
          }
          if (lastCut) {
            for (const transaction of transactions) {
              if (transaction.before !== lastCut.doc) continue
              lastCut.ranges = lastCut.ranges.map((r) => ({
                from: transaction.mapping.map(r.from, -1),
                to: transaction.mapping.map(r.to, 1),
              }))
              lastCut.doc = transaction.doc
            }
          }
          const tracked = transactions.filter(
            (t) => t.docChanged && !t.getMeta(TRACK_IGNORE) && !t.getMeta('history$'),
          )
          let composing = false
          try {
            composing = Boolean(editor.view?.composing)
          } catch {
            /* headless editor */
          }
          // the IME commit lands after view.composing flips false (PM flushes the
          // final composition DOM records in a microtask from compositionend and
          // tags them); recording it as a plain edit struck the provisional text
          if (!composing && tracked.length > 0 && tracked.every((t) => t.getMeta('composition'))) {
            composing = true
          }
          const flush = composing ? null : pendingIme
          if (!composing) pendingIme = null
          if (tracked.length === 0 && !flush) return null

          const insType = newState.schema.marks.ins
          const delType = newState.schema.marks.del
          if (!insType || !delType) return null

          // flatten steps across transactions with a shared mapping to the final doc
          const steps = tracked.flatMap((t) =>
            t.steps.map((step, i) => ({ step, docBefore: t.docs[i] })),
          )
          const mapping = new Mapping(steps.map(({ step }) => step.getMap()))

          // Affected span in final-doc coordinates: the scans below only visit
          // blocks intersecting it (blocks outside the step ranges cannot change).
          let spanFrom = Infinity
          let spanTo = -Infinity
          let unbounded = false
          steps.forEach(({ step }, idx) => {
            const tail = mapping.slice(idx + 1)
            let ranged = false
            step.getMap().forEach((_oldStart, _oldEnd, from, to) => {
              ranged = true
              spanFrom = Math.min(spanFrom, tail.map(from, -1))
              spanTo = Math.max(spanTo, tail.map(to, 1))
            })
            if (ranged) return
            if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
              spanFrom = Math.min(spanFrom, tail.map(step.from, -1))
              spanTo = Math.max(spanTo, tail.map(step.to, 1))
            } else {
              // AttrStep and other rangeless steps expose pos; anything else falls
              // back to a whole-doc scan
              const pos = (step as unknown as { pos?: number }).pos
              if (typeof pos === 'number') {
                spanFrom = Math.min(spanFrom, tail.map(pos, -1))
                spanTo = Math.max(spanTo, tail.map(pos, 1) + 1)
              } else unbounded = true
            }
          })

          if (composing) {
            if (spanFrom <= spanTo && !unbounded) {
              const from = Math.max(0, spanFrom)
              const to = Math.min(newState.doc.content.size, spanTo)
              if (pendingIme) {
                pendingIme.from = Math.min(pendingIme.from, from)
                pendingIme.to = Math.max(pendingIme.to, to)
              } else {
                pendingIme = {
                  from,
                  to,
                  base: oldState.doc,
                  mapping: new Mapping(transactions.flatMap((t) => t.mapping.maps)),
                }
              }
            }
            return null
          }

          let newFrom = 0
          let newTo = 0
          let oldFrom = 0
          let oldTo = 0
          if (unbounded) {
            newTo = newState.doc.content.size
            oldTo = oldState.doc.content.size
          } else if (spanFrom <= spanTo) {
            newFrom = Math.max(0, spanFrom - 1)
            newTo = Math.min(newState.doc.content.size, spanTo + 1)
            // exact preimage: old/new scans must cover matching block sets, or the
            // anchor diff below would mistake one-sided coverage for a delete/insert
            const inverted = mapping.invert()
            oldFrom = Math.max(0, inverted.map(newFrom, -1))
            oldTo = Math.min(oldState.doc.content.size, inverted.map(newTo, 1))
          }

          const now = revisionDateNow()
          const author =
            tracked.map((t) => t.getMeta(TRACK_AUTHOR) as string | undefined).find(Boolean) ??
            storage.author

          const uiEvent =
            tracked.map((t) => t.getMeta('uiEvent') as string | undefined).find(Boolean) ??
            (armed && Date.now() - armed.at < 1000 ? armed.event : undefined)
          armed = null
          let removedKey = ''
          let insertedKey = ''
          let movable = true
          for (const { step, docBefore } of steps) {
            if (!(step instanceof ReplaceStep)) continue
            if (step.to > step.from) {
              removedKey += sliceKey(docBefore.slice(step.from, step.to).content)
              movable &&= coversWholeSentences(docBefore, step.from, step.to)
            }
            insertedKey += sliceKey(step.slice.content)
          }
          let moveId: string | null = null
          if (!storage.trackMoves) {
            /* moves record as deletion + insertion */
          } else if (uiEvent === 'drop' && movable && insertedKey && insertedKey === removedKey)
            moveId = newMoveId()
          else if (uiEvent === 'paste' && lastCut && insertedKey === lastCut.key)
            moveId = newMoveId()
          const pairCut = uiEvent === 'paste' && moveId ? lastCut : null
          if (uiEvent !== 'cut') lastCut = null
          let struck = false

          const insMark = insType.create({ author, date: now, move: moveId })
          const delMark = delType.create({
            author,
            date: now,
            move: uiEvent === 'drop' ? moveId : null,
          })

          const tr = newState.tr
          tr.setMeta(TRACK_IGNORE, true)
          let changed = false
          const recordedFormatRanges = new Set<string>()

          // Enter inside a paragraph: the inserted paragraph mark ends the first
          // half (Word credits that mark); the second half is the original's
          // continuation, not an inserted block
          const splitTails = new Set<number>()
          // A selection deleted from inside one paragraph into another: Word
          // strikes the removed text and marks the first paragraph's mark
          // deleted; the paragraphs stay apart (rebuilt below from the merged
          // result). Their old positions must not read as block deletions.
          const rebuiltOld = new Set<number>()
          const crossParagraphCut = (step: ReplaceStep, docBefore: PmNode) => {
            if (step.from === step.to) return null
            const $from = docBefore.resolve(step.from)
            const $to = docBefore.resolve(step.to)
            if (
              $from.depth !== $to.depth ||
              $from.sameParent($to) ||
              !PARAGRAPH_NODE_TYPES.has($from.parent.type.name) ||
              !PARAGRAPH_NODE_TYPES.has($to.parent.type.name) ||
              $from.start($from.depth - 1) !== $to.start($to.depth - 1)
            )
              return null
            let inlineInsert = true
            step.slice.content.forEach((child) => {
              if (!child.isInline) inlineInsert = false
            })
            if (
              step.slice.size > 0 &&
              (step.slice.openStart !== 0 || step.slice.openEnd !== 0 || !inlineInsert)
            )
              return null
            const depth = $from.depth
            const parentStart = $from.start(depth - 1)
            const paras: Array<{ node: PmNode; pos: number }> = []
            $from.node(depth - 1).forEach((child, offset) => {
              const pos = parentStart + offset
              if (pos >= $from.before(depth) && pos <= $to.before(depth))
                paras.push({ node: child, pos })
            })
            if (paras.some((p) => !PARAGRAPH_NODE_TYPES.has(p.node.type.name))) return null
            return { $from, $to, paras }
          }
          const isParagraphSplit = (step: ReplaceStep) => {
            const { slice } = step
            const head = slice.content.firstChild
            const second = slice.content.lastChild
            return (
              slice.openStart === 1 &&
              slice.openEnd === 1 &&
              slice.content.childCount === 2 &&
              !!head &&
              !!second &&
              PARAGRAPH_NODE_TYPES.has(head.type.name) &&
              PARAGRAPH_NODE_TYPES.has(second.type.name) &&
              head.content.size === 0 &&
              second.content.size === 0
            )
          }
          steps.forEach(({ step, docBefore }, idx) => {
            if (!(step instanceof ReplaceStep)) return
            const tail = mapping.slice(idx + 1)
            if (isParagraphSplit(step)) {
              if (PARAGRAPH_NODE_TYPES.has(docBefore.resolve(step.from).parent.type.name))
                splitTails.add(tail.map(step.from + 1, -1))
              return
            }
            const cut = crossParagraphCut(step, docBefore)
            if (!cut) return
            // Mapping.invert ignores slice bounds: build the prefix explicitly
            const toOld = new Mapping(steps.slice(0, idx).map((s) => s.step.getMap())).invert()
            for (const { pos } of cut.paras.slice(1)) rebuiltOld.add(toOld.map(pos, -1))
          })

          // Paragraph-property commands are SetNodeMarkup/ReplaceAround operations.
          // Comparing mapped textblocks catches alignment, spacing, indentation,
          // styles, heading conversions, and list level/type changes without
          // depending on the command's concrete ProseMirror step class.
          const fullMapping = new Mapping(
            tracked.flatMap((transaction) => transaction.steps.map((step) => step.getMap())),
          )
          // docxIndex anchors are NOT unique identifiers: ProseMirror's split
          // clones the attrs onto both halves, so every Enter mints a second
          // block with the same index. Collapsing them into single-slot Maps
          // misread a plain paragraph split as a block MOVE (the old position
          // no longer equals the surviving slot's position) and re-materialized
          // the whole original paragraph as a deleted block — a full visual
          // duplicate whose Reject corrupted the document. Group the anchors
          // instead and match the k-th old occurrence to the k-th new one.
          const groupAnchors = (doc: PmNode, from: number, to: number) => {
            const groups = new Map<number, Array<{ node: PmNode; pos: number }>>()
            doc.nodesBetween(from, to, (node, pos) => {
              const index = node.attrs?.docxIndex
              if (typeof index === 'number') {
                const group = groups.get(index)
                if (group) group.push({ node, pos })
                else groups.set(index, [{ node, pos }])
              }
              return false
            })
            return groups
          }
          const oldAnchors = groupAnchors(oldState.doc, oldFrom, oldTo)
          const newAnchors = groupAnchors(newState.doc, newFrom, newTo)
          /**
           * A paragraph JOIN keeps the block's content and only drops the
           * paragraph mark: its old content range maps onto an identical,
           * contiguous range inside one surviving textblock. A whole-block
           * deletion maps that range to a single point instead.
           */
          const survivesInsideTextblock = (node: PmNode, oldPos: number): boolean => {
            if (node.content.size === 0) return false
            const from = fullMapping.map(oldPos + 1, 1)
            const to = fullMapping.map(oldPos + 1 + node.content.size, -1)
            if (to - from !== node.content.size) return false
            const $from = newState.doc.resolve(from)
            return (
              $from.parent.isTextblock &&
              $from.sameParent(newState.doc.resolve(to)) &&
              newState.doc.slice(from, to).content.eq(node.content)
            )
          }
          const matchedNewPositions = new Set<number>()
          const movedNewPositions = new Set<number>()
          const deletedBlocks: Array<{ node: PmNode; pos: number }> = []
          const oldOrdinals = new Map<number, number>()
          oldState.doc.nodesBetween(oldFrom, oldTo, (before, oldPos) => {
            if (rebuiltOld.has(oldPos)) return false
            const index = before.attrs?.docxIndex
            if (typeof index === 'number') {
              const group = newAnchors.get(index) ?? []
              const ordinal = oldOrdinals.get(index) ?? 0
              oldOrdinals.set(index, ordinal + 1)
              const match = group[ordinal]
              if (match) {
                matchedNewPositions.add(match.pos)
                const mapped = fullMapping.map(oldPos, -1)
                if (
                  mapped !== match.pos &&
                  // only a 1:1 anchor at a new position is a move — with
                  // duplicated anchors (a split's halves) the position
                  // mismatch is the split itself, not a relocation
                  (oldAnchors.get(index)?.length ?? 0) === 1 &&
                  group.length === 1 &&
                  (before.attrs?.blockRevision as { kind?: string } | null)?.kind !== 'ins'
                ) {
                  deletedBlocks.push({ node: before, pos: mapped })
                  movedNewPositions.add(match.pos)
                }
              } else if (
                (before.attrs?.blockRevision as { kind?: string } | null)?.kind !== 'ins'
              ) {
                // No block left for this occurrence. After a paragraph JOIN
                // the content lives on inside a surviving textblock; re-inserting
                // it as a deleted block would show it twice and rejecting the
                // phantom would corrupt the document.
                if (!survivesInsideTextblock(before, oldPos))
                  deletedBlocks.push({ node: before, pos: fullMapping.map(oldPos, -1) })
              }
              return false
            }
            const mapped = fullMapping.map(oldPos, -1)
            const after = newState.doc.nodeAt(mapped)
            if (
              after &&
              (after.type === before.type ||
                (PARAGRAPH_NODE_TYPES.has(after.type.name) &&
                  PARAGRAPH_NODE_TYPES.has(before.type.name)))
            )
              matchedNewPositions.add(mapped)
            else if ((before.attrs?.blockRevision as { kind?: string } | null)?.kind !== 'ins') {
              deletedBlocks.push({ node: before, pos: mapped })
            }
            return false
          })
          // blocks that went from one place come back in their order: the later one first
          const inOrder = deletedBlocks.map((block, order) => ({ ...block, order }))
          for (const { node, pos } of inOrder.sort((a, b) => b.pos - a.pos || b.order - a.order)) {
            tr.insert(
              tr.mapping.map(pos, -1),
              node.type.create(
                {
                  ...node.attrs,
                  blockRevision: { kind: 'del', author: author, date: now },
                },
                node.content,
                node.marks,
              ),
            )
            changed = true
          }
          newState.doc.nodesBetween(newFrom, newTo, (node, pos) => {
            const index = node.attrs?.docxIndex
            const knownAnchor = typeof index === 'number' && oldAnchors.has(index)
            if (
              (!movedNewPositions.has(pos) && (knownAnchor || matchedNewPositions.has(pos))) ||
              node.attrs?.blockRevision ||
              splitTails.has(pos)
            )
              return false
            tr.setNodeMarkup(tr.mapping.map(pos, 1), undefined, {
              ...node.attrs,
              blockRevision: { kind: 'ins', author: author, date: now },
            })
            changed = true
            return false
          })

          // Native table row commands preserve the table anchor. Record the
          // single contiguous row span they insert/delete using trPr revisions.
          for (const [index, oldGroup] of oldAnchors) {
            // tables never split under Enter, so a table anchor group is 1:1
            const oldTable = oldGroup[0]!
            const nextTable = newAnchors.get(index)?.[0]
            if (oldTable.node.type.name !== 'docTable' || nextTable?.node.type.name !== 'docTable')
              continue
            const oldRows = Array.from({ length: oldTable.node.childCount }, (_, i) =>
              oldTable.node.child(i),
            )
            const newRows = Array.from({ length: nextTable.node.childCount }, (_, i) =>
              nextTable.node.child(i),
            )
            let prefix = 0
            while (
              prefix < oldRows.length &&
              prefix < newRows.length &&
              oldRows[prefix].eq(newRows[prefix])
            )
              prefix++
            let suffix = 0
            while (
              suffix < oldRows.length - prefix &&
              suffix < newRows.length - prefix &&
              oldRows[oldRows.length - 1 - suffix].eq(newRows[newRows.length - 1 - suffix])
            )
              suffix++
            const oldMiddle = oldRows.slice(prefix, oldRows.length - suffix)
            const newMiddle = newRows.slice(prefix, newRows.length - suffix)
            const rowPos = (rowIndex: number) => {
              let pos = nextTable.pos + 1
              for (let i = 0; i < rowIndex; i++) pos += newRows[i].nodeSize
              return pos
            }
            if (oldMiddle.length === 0 && newMiddle.length > 0) {
              for (let i = 0; i < newMiddle.length; i++) {
                const pos = tr.mapping.map(rowPos(prefix + i), 1)
                tr.setNodeMarkup(pos, undefined, {
                  ...newMiddle[i].attrs,
                  rowRevision: { kind: 'ins', author: author, date: now },
                  rawTrPr: withRowRecord(newMiddle[i].attrs.rawTrPr, 'ins', author, now),
                })
              }
              changed = true
            } else if (oldMiddle.length > 0 && newMiddle.length === 0) {
              const insertionPos = rowPos(prefix)
              for (const row of [...oldMiddle].reverse()) {
                tr.insert(
                  tr.mapping.map(insertionPos, -1),
                  row.type.create(
                    {
                      ...row.attrs,
                      rowRevision: { kind: 'del', author: author, date: now },
                      rawTrPr: withRowRecord(row.attrs.rawTrPr, 'del', author, now),
                    },
                    row.content,
                    row.marks,
                  ),
                )
              }
              changed = true
            }
          }

          oldState.doc.nodesBetween(oldFrom, oldTo, (before, oldPos) => {
            if (!storage.trackFormatting) return false
            if (!PARAGRAPH_NODE_TYPES.has(before.type.name) || rebuiltOld.has(oldPos)) return
            const newPos = fullMapping.map(oldPos, -1)
            const after = newState.doc.nodeAt(newPos)
            if (
              !after ||
              !PARAGRAPH_NODE_TYPES.has(after.type.name) ||
              !paragraphPropertiesChanged(before, after) ||
              after.attrs.pPrChange
            ) {
              return
            }
            tr.setNodeMarkup(newPos, undefined, {
              ...after.attrs,
              pPrChange: JSON.stringify({
                author: author,
                date: now,
                old: paragraphSnapshot(before),
              }),
            })
            changed = true
          })

          steps.forEach(({ step, docBefore }, idx) => {
            if (
              (step instanceof AddMarkStep || step instanceof RemoveMarkStep) &&
              TRACKED_FORMAT_MARKS.has(step.mark.type.name)
            ) {
              const rprType = newState.schema.marks.rprChange
              if (!rprType || !storage.trackFormatting) return
              const tail = mapping.slice(idx + 1)
              docBefore.nodesBetween(step.from, step.to, (node, pos) => {
                if (!node.isText) return
                const sourceFrom = Math.max(pos, step.from)
                const sourceTo = Math.min(pos + node.nodeSize, step.to)
                const from = tr.mapping.map(tail.map(sourceFrom, 1), 1)
                const to = tr.mapping.map(tail.map(sourceTo, -1), -1)
                if (from >= to) return
                const key = `${from}:${to}`
                if (recordedFormatRanges.has(key)) return
                recordedFormatRanges.add(key)
                const existing = node.marks.find((m) => m.type.name === 'rprChange')
                tr.addMark(
                  from,
                  to,
                  existing ??
                    rprType.create({
                      author: author,
                      date: now,
                      old: formatSnapshot(node),
                    }),
                )
                changed = true
              })
              return
            }

            if (!(step instanceof ReplaceStep) && !(step instanceof ReplaceAroundStep)) return
            if (step instanceof ReplaceAroundStep) {
              const before = docBefore.nodeAt(step.from)
              const tail = mapping.slice(idx + 1)
              const after = newState.doc.nodeAt(tail.map(step.from, -1))
              if (
                before &&
                after &&
                PARAGRAPH_NODE_TYPES.has(before.type.name) &&
                PARAGRAPH_NODE_TYPES.has(after.type.name) &&
                before.content.eq(after.content) &&
                paragraphPropertiesChanged(before, after)
              ) {
                return
              }
            }
            const from = step.from
            const to = step.to
            const slice = step.slice
            const tail = mapping.slice(idx + 1)
            let insertsWholeBlocks = slice.openStart === 0 && slice.content.childCount > 0
            slice.content.forEach((child) => {
              if (!child.isBlock) insertsWholeBlocks = false
            })

            if (step instanceof ReplaceStep && isParagraphSplit(step)) {
              const $split = docBefore.resolve(from)
              if (!PARAGRAPH_NODE_TYPES.has($split.parent.type.name)) return
              const headPos = tr.mapping.map(tail.map($split.before(), -1), -1)
              const head = tr.doc.nodeAt(headPos)
              if (!head || !PARAGRAPH_NODE_TYPES.has(head.type.name) || head.attrs.blockRevision)
                return
              // the split cloned the original mark's revision state onto both halves; it belongs to the second
              tr.setNodeMarkup(headPos, undefined, {
                ...head.attrs,
                paraMarkIns: markRecord(author, now),
                paraMarkDel: null,
                paraMarkRPrChange: null,
              })
              changed = true
              return
            }

            // mark inserted content
            if (slice.size > 0 && !insertsWholeBlocks) {
              const a = tr.mapping.map(tail.map(from, 1), 1)
              const b = tr.mapping.map(tail.map(from + slice.size, -1), -1)
              if (a < b) {
                tr.removeMark(a, b, delType)
                tr.addMark(a, b, insMark)
                changed = true
              }
            }

            const strike = (fragment: Fragment): PmNode[] => {
              const kept: PmNode[] = []
              fragment.forEach((child) => {
                if (child.marks.some((m) => m.type === insType && m.attrs.author === author)) return
                if (child.marks.some((m) => m.type === delType)) kept.push(child)
                else {
                  struck = true
                  kept.push(child.mark([...child.marks, delMark]))
                }
              })
              return kept
            }
            const cut = step instanceof ReplaceStep ? crossParagraphCut(step, docBefore) : null
            if (cut) {
              const headSize = from - cut.$from.start()
              const liveEnd = headSize + slice.size
              const lastHead = to - cut.$to.start()
              const mergedPos = tr.mapping.map(tail.map(cut.$from.before(), -1), -1)
              const merged = tr.doc.nodeAt(mergedPos)
              const lastNode = cut.paras[cut.paras.length - 1]!.node
              if (
                !merged ||
                !PARAGRAPH_NODE_TYPES.has(merged.type.name) ||
                merged.content.size !== liveEnd + lastNode.content.size - lastHead
              )
                return
              const parts = cut.paras.map(({ node }, i) => {
                const first = i === 0
                const last = i === cut.paras.length - 1
                const content: PmNode[] = []
                if (first) merged.content.cut(0, liveEnd).forEach((n) => content.push(n))
                content.push(
                  ...strike(
                    first
                      ? node.content.cut(headSize)
                      : last
                        ? node.content.cut(0, lastHead)
                        : node.content,
                  ),
                )
                if (last) merged.content.cut(liveEnd).forEach((n) => content.push(n))
                return { type: node.type, attrs: node.attrs, content, markDeleted: !last }
              })
              // a paragraph whose mark is this author's own pending insertion loses it for real
              const folded: typeof parts = []
              for (const part of parts) {
                const prev = folded[folded.length - 1]
                if (prev && recordAuthor(prev.attrs.paraMarkIns) === author) {
                  prev.content.push(...part.content)
                  prev.attrs = {
                    ...prev.attrs,
                    paraMarkIns: part.attrs.paraMarkIns,
                    paraMarkDel: part.attrs.paraMarkDel,
                    paraMarkRPrChange: part.attrs.paraMarkRPrChange,
                  }
                  prev.markDeleted = part.markDeleted
                } else folded.push({ ...part })
              }
              const nodes = folded.map((part) =>
                part.type.create(
                  part.markDeleted && !part.attrs.paraMarkDel
                    ? { ...part.attrs, paraMarkDel: markRecord(author, now) }
                    : part.attrs,
                  Fragment.from(part.content),
                ),
              )
              tr.replaceWith(mergedPos, mergedPos + merged.nodeSize, nodes)
              if (newState.selection.empty)
                tr.setSelection(TextSelection.create(tr.doc, mergedPos + 1 + liveEnd))
              changed = true
              return
            }

            // re-materialize deleted inline content, struck through
            if (to > from) {
              const removed = docBefore.slice(from, to)
              // only inline slices (deletions within one paragraph) are re-inserted
              let inlineOnly = removed.content.childCount > 0
              removed.content.forEach((child) => {
                if (!child.isInline) inlineOnly = false
              })
              if (inlineOnly && removed.openStart === 0 && removed.openEnd === 0) {
                const keep: PmNode[] = []
                removed.content.forEach((child) => {
                  const ownIns = child.marks.some(
                    (m) => m.type === insType && m.attrs.author === author,
                  )
                  if (ownIns) return // deleting your own insertion removes it outright
                  if (child.marks.some((m) => m.type === delType)) keep.push(child)
                  else {
                    struck = true
                    keep.push(child.mark([...child.marks, delMark]))
                  }
                })
                if (keep.length > 0) {
                  const pos = tr.mapping.map(tail.map(from + slice.size, -1), -1)
                  const end = keep.reduce((at, child) => {
                    tr.insert(at, child)
                    return at + child.nodeSize
                  }, pos)
                  // Leave the caret in front of the struck text, not behind it.
                  // Behind it, the next Backspace targets content that is already
                  // marked deleted — which this plugin restores verbatim — so the
                  // user could only ever strike one character per position.
                  if (newState.selection.empty && tr.mapping.map(newState.selection.from) === end) {
                    tr.setSelection(TextSelection.create(tr.doc, pos))
                  }
                  changed = true
                }
              }
            }
          })

          if (pairCut) {
            // the cut's struck text becomes the move's source
            const source: Array<{ from: number; to: number }> = []
            for (const r of pairCut.ranges) {
              const from = tr.mapping.map(r.from, 1)
              const to = tr.mapping.map(r.to, -1)
              if (from >= to) continue
              tr.doc.nodesBetween(from, to, (node, pos) => {
                if (node.isText && node.marks.some((m) => m.eq(pairCut.delMark)))
                  source.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) })
              })
            }
            for (const { from, to } of source)
              tr.addMark(from, to, delType.create({ ...pairCut.delMark.attrs, move: moveId }))
            if (source.length > 0) changed = true
          }

          if (flush) {
            const max = newState.doc.content.size
            const from = Math.max(0, Math.min(flush.from, max))
            const to = Math.max(from, Math.min(flush.to, max))
            const inv = flush.mapping.invert()
            const baseMax = flush.base.content.size
            const baseFrom = Math.max(0, Math.min(inv.map(from, -1), baseMax))
            const baseTo = Math.max(baseFrom, Math.min(inv.map(to, 1), baseMax))
            const cur = inlineText(newState.doc, from, to)
            const base = inlineText(flush.base, baseFrom, baseTo)
            // trim the common affix so only the middle the composition really
            // changed gets recorded (a one-character IME edit used to redline
            // the whole paragraph when PM rewrote the surrounding runs)
            let prefix = 0
            let suffix = 0
            let removed: PmNode[] = []
            if (cur && base) {
              const limit = Math.min(base.text.length, cur.text.length)
              while (prefix < limit && base.text[prefix] === cur.text[prefix]) prefix++
              while (
                suffix < limit - prefix &&
                base.text[base.text.length - 1 - suffix] === cur.text[cur.text.length - 1 - suffix]
              )
                suffix++
              const midTo = base.text.length - suffix
              let at = 0
              for (const child of base.children) {
                const len = child.isText ? child.text!.length : 1
                const a = Math.max(prefix - at, 0)
                const b = Math.min(midTo - at, len)
                if (a < b)
                  removed.push(child.isText && (a > 0 || b < len) ? child.cut(a, b) : child)
                at += len
              }
              removed = removed.filter(
                (child) =>
                  // deleting your own insertion removes it outright
                  !child.marks.some((m) => m.type === insType && m.attrs.author === author),
              )
            }
            const insFrom = tr.mapping.map(from + prefix, 1)
            const insTo = tr.mapping.map(to - suffix, -1)
            if (insFrom < insTo) {
              tr.removeMark(insFrom, insTo, delType)
              tr.addMark(insFrom, insTo, insMark)
              changed = true
            }
            if (removed.length > 0) {
              const at = Math.max(insFrom, insTo)
              const end = removed.reduce((acc, child) => {
                const struck = child.marks.some((m) => m.type === delType)
                  ? child
                  : child.mark([...child.marks, delMark])
                tr.insert(acc, struck)
                return acc + struck.nodeSize
              }, at)
              // same caret rule as above: keep it in front of the struck text
              if (newState.selection.empty && tr.mapping.map(newState.selection.from) === end) {
                tr.setSelection(TextSelection.create(tr.doc, at))
              }
              changed = true
            }
          }

          if (uiEvent === 'cut' && movable && struck && removedKey && changed) {
            // the struck text sits inside the recorded span; remember where, not just how it is marked
            const ranges: Array<{ from: number; to: number }> = []
            const from = Math.max(0, tr.mapping.map(newFrom, -1))
            const to = Math.min(tr.doc.content.size, tr.mapping.map(newTo, 1))
            tr.doc.nodesBetween(from, to, (node, pos) => {
              if (node.isText && node.marks.some((m) => m.eq(delMark)))
                ranges.push({ from: pos, to: pos + node.nodeSize })
            })
            if (ranges.length > 0) lastCut = { key: removedKey, delMark, ranges, doc: tr.doc }
          }

          return changed ? tr : null
        },
      }),
    ]
  },
})
