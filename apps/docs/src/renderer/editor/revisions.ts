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
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'
import {
  AddMarkStep,
  Mapping,
  RemoveMarkStep,
  ReplaceAroundStep,
  ReplaceStep,
} from '@tiptap/pm/transform'
import { t } from '../i18n/locale'

export interface RevisionRange {
  from: number
  to: number
  kind:
    | 'ins'
    | 'del'
    | 'both'
    | 'pPrChange'
    | 'moveFrom'
    | 'moveTo'
    | 'rPrChange'
    | 'rowIns'
    | 'rowDel'
    | 'cellIns'
    | 'cellDel'
    | 'blockIns'
    | 'blockDel'
  author: string
  date?: string
}

/** transactions carrying this meta are never recorded as revisions */
export const TRACK_IGNORE = 'trackIgnore'

const TRACKED_FORMAT_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])

const TEXT_STYLE_FIELDS = [
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
const PARAGRAPH_FORMAT_FIELDS = [
  'align',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
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
const xmlAttr = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')

interface ParagraphSnapshot {
  type: string
  format: Record<string, unknown>
  styleId: unknown
  level: unknown
  kind: unknown
  numId: unknown
  ilvl: unknown
}

function paragraphSnapshot(node: PmNode): ParagraphSnapshot {
  const format: Record<string, unknown> = {}
  for (const field of PARAGRAPH_FORMAT_FIELDS) format[field] = node.attrs[field] ?? null
  return {
    type: node.type.name,
    format,
    styleId: node.attrs.styleId ?? null,
    level: node.attrs.level ?? null,
    kind: node.attrs.kind ?? null,
    numId: node.attrs.numId ?? null,
    ilvl: node.attrs.ilvl ?? null,
  }
}

function paragraphPropertiesChanged(before: PmNode, after: PmNode): boolean {
  if (before.type !== after.type) return true
  return [...PARAGRAPH_FORMAT_FIELDS, ...PARAGRAPH_STRUCTURE_FIELDS].some(
    (field) => (before.attrs[field] ?? null) !== (after.attrs[field] ?? null),
  )
}

function formatSnapshot(node: PmNode): Record<string, unknown> {
  const textStyle = node.marks.find((m) => m.type.name === 'docTextStyle')
  const snapshot: Record<string, unknown> = {
    bold: node.marks.some((m) => m.type.name === 'bold'),
    italic: node.marks.some((m) => m.type.name === 'italic'),
    underline: node.marks.some((m) => m.type.name === 'underline'),
    strike: node.marks.some((m) => m.type.name === 'strike'),
  }
  for (const field of TEXT_STYLE_FIELDS) {
    const value = textStyle?.attrs[field]
    if (value !== null && value !== undefined) snapshot[field] = value
  }
  return snapshot
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
  doc.descendants((node, pos) => {
    const name = node.type.name
    if (name !== 'docTableRow' && name !== 'docTableCell' && name !== 'docTableHeader') return
    const isRow = name === 'docTableRow'
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
  const isSuppressed = (pos: number, kind: 'ins' | 'del' | 'both') =>
    suppressed.some((s) => pos >= s.from && pos < s.to && (kind === 'both' || s.kind === kind))

  // First pass: collect paragraph revisions at any depth, including table cells.
  doc.descendants((node, pos) => {
    if (!PARAGRAPH_NODE_TYPES.has(node.type.name)) return
    const pPrChange = node.attrs?.pPrChange as string | null
    const moveRevision = node.attrs?.moveRevision as string | null
    if (pPrChange) {
      let author = ''
      let date: string | undefined
      try {
        const parsed = JSON.parse(pPrChange)
        author = parsed?.author ?? ''
        date = parsed?.date
      } catch {
        /* ignore */
      }
      out.push({
        from: pos,
        to: pos + node.nodeSize,
        kind: 'pPrChange',
        author,
        ...(date ? { date } : {}),
      })
    } else if (moveRevision === 'from' || moveRevision === 'to') {
      out.push({
        from: pos,
        to: pos + node.nodeSize,
        kind: moveRevision === 'from' ? 'moveFrom' : 'moveTo',
        author: '',
      })
    }
  })

  // Second pass: collect run-level text revisions (ins/del marks)
  doc.descendants((node, pos) => {
    if (!node.isText) return
    const ins = node.marks.find((m) => m.type.name === 'ins')
    const del = node.marks.find((m) => m.type.name === 'del')
    if (!ins && !del) {
      const rpr = node.marks.find((m) => m.type.name === 'rprChange')
      if (!rpr) return
      const author = String(rpr.attrs.author ?? '')
      const prev = out[out.length - 1]
      if (prev && prev.to === pos && prev.kind === 'rPrChange' && prev.author === author) {
        prev.to = pos + node.nodeSize
      } else {
        out.push({
          from: pos,
          to: pos + node.nodeSize,
          kind: 'rPrChange',
          author,
          ...(rpr.attrs.date ? { date: String(rpr.attrs.date) } : {}),
        })
      }
      return
    }
    const kind: 'ins' | 'del' | 'both' = ins && del ? 'both' : ins ? 'ins' : 'del'
    if (isSuppressed(pos, kind)) return
    const src = del ?? ins!
    const author = String(src.attrs.author ?? '')
    const prev = out[out.length - 1]
    if (prev && prev.to === pos && prev.kind === kind && prev.author === author) {
      prev.to = pos + node.nodeSize
    } else {
      out.push({
        from: pos,
        to: pos + node.nodeSize,
        kind,
        author,
        ...(src.attrs.date ? { date: String(src.attrs.date) } : {}),
      })
    }
  })
  out.push(...tableRevs)
  return out.sort((a, b) => a.from - b.from)
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

function applyRevisions(editor: Editor, ranges: RevisionRange[], mode: 'accept' | 'reject'): void {
  if (ranges.length === 0) return
  const { state } = editor
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  // reverse order keeps earlier positions valid while later ranges mutate
  for (const r of [...ranges].sort((a, b) => b.from - a.from)) {
    if (r.kind === 'blockIns' || r.kind === 'blockDel') {
      const revisionKind = r.kind === 'blockIns' ? 'ins' : 'del'
      const removeNode = mode === 'accept' ? revisionKind === 'del' : revisionKind === 'ins'
      const nodePos = tr.mapping.map(r.from, -1)
      const node = tr.doc.nodeAt(nodePos)
      if (!node) continue
      if (removeNode) {
        if (tr.doc.childCount > 1) tr.delete(nodePos, nodePos + node.nodeSize)
      } else {
        tr.setNodeMarkup(nodePos, undefined, { ...node.attrs, blockRevision: null })
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
      state.doc.nodesBetween(r.from, r.to, (node, pos) => {
        if (!node.isText) return
        const from = Math.max(pos, r.from)
        const to = Math.min(pos + node.nodeSize, r.to)
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
            const baseAttrs = ts ? { ...ts.attrs } : {}
            const nextAttrs: Record<string, unknown> = { ...baseAttrs }
            for (const field of TEXT_STYLE_FIELDS) nextAttrs[field] = old[field] ?? null
            // Reject goes back to the modeled old format. rawRPr contains the new
            // format plus the revision record, so rebuild instead of preserving it.
            nextAttrs.rawRPr = null
            const hasAny = TEXT_STYLE_FIELDS.some(
              (field) => nextAttrs[field] !== null && nextAttrs[field] !== undefined,
            )
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
      tr.removeMark(r.from, r.to, state.schema.marks.rprChange)
      continue
    }

    if (r.kind === 'moveFrom' || r.kind === 'moveTo') {
      // moveFrom/moveTo: translate to del/ins logic using the same remove/keep semantics
      const effectiveKind = r.kind === 'moveFrom' ? 'del' : 'ins'
      const removeText = mode === 'accept' ? effectiveKind !== 'ins' : effectiveKind !== 'del'
      if (removeText) {
        const $from = state.doc.resolve(r.from)
        const $to = state.doc.resolve(r.to)
        if ($from.depth >= 1 && $from.sameParent($to) && state.doc.childCount > 1) {
          tr.delete($from.before(), $to.after())
        } else {
          tr.delete(r.from, r.to)
        }
      } else if (mode === 'accept') {
        tr.removeMark(r.from, r.to, state.schema.marks.ins)
      } else {
        tr.removeMark(r.from, r.to, state.schema.marks.del)
      }
      // Also clear the moveRevision attr from the node
      try {
        const $pos = state.doc.resolve(r.from + 1)
        const nodePos = $pos.before($pos.depth)
        if ($pos.parent.attrs?.moveRevision) {
          tr.setNodeMarkup(nodePos, undefined, { ...$pos.parent.attrs, moveRevision: null })
        }
      } catch {
        /* ignore if node is being deleted */
      }
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
      const node = state.doc.nodeAt(r.from)
      if (!node) continue
      if (removeNode && isRow) {
        // accepting a deleted row / rejecting an inserted row = remove the whole row; if it is the only row, remove the whole table
        const $pos = state.doc.resolve(r.from)
        if ($pos.parent.childCount <= 1) tr.delete($pos.before(), $pos.after())
        else tr.delete(r.from, r.to)
      } else if (removeNode) {
        // a cell cannot be removed outright (table grid self-repair would re-add an empty one): clear it in place and strip the marker
        tr.setNodeMarkup(r.from, undefined, {
          ...node.attrs,
          cellRevision: null,
          rawTcPr: stripTrackMarker(
            node.attrs.rawTcPr as string | null,
            revKind === 'ins' ? 'w:cellIns' : 'w:cellDel',
            'w:tcPr',
          ),
        })
        const empty = state.schema.nodes.docParagraph.createAndFill()
        if (empty) tr.replaceWith(r.from + 1, r.to - 1, empty)
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
        tr.setNodeMarkup(r.from, undefined, attrs)
        const markType = state.schema.marks[revKind]
        if (markType) tr.removeMark(r.from + 1, r.to - 1, markType)
      }
      continue
    }

    const removeText = mode === 'accept' ? r.kind !== 'ins' : r.kind !== 'del'
    if (removeText) {
      // a revision covering a block's entire content removes the block itself
      // (Word: rejecting an inserted paragraph deletes the paragraph, not just
      // its text), unless it is the document's last remaining block
      const $from = state.doc.resolve(r.from)
      const $to = state.doc.resolve(r.to)
      if (
        $from.parent.isTextblock &&
        $from.sameParent($to) &&
        r.from === $from.start() &&
        r.to === $to.end() &&
        $from.depth === 1 &&
        state.doc.childCount > 1
      ) {
        tr.delete($from.before(), $to.after())
      } else {
        tr.delete(r.from, r.to)
      }
    } else if (mode === 'accept') {
      tr.removeMark(r.from, r.to, state.schema.marks.ins)
    } else {
      tr.removeMark(r.from, r.to, state.schema.marks.del)
    }
  }
  editor.view.dispatch(tr)
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

/** move the selection to the next / previous revision (wraps around) */
export function gotoRevision(editor: Editor, dir: 1 | -1): boolean {
  const ranges = collectRevisions(editor.state.doc)
  if (ranges.length === 0) return false
  // compare range starts so adjacent revisions (shared boundary) still advance
  const anchor = editor.state.selection.from
  const target =
    dir === 1
      ? (ranges.find((r) => r.from > anchor) ?? ranges[0])
      : ([...ranges].reverse().find((r) => r.from < anchor) ?? ranges[ranges.length - 1])
  const tr = editor.state.tr
  // between: row/cell-level revision range endpoints are not text positions; snap to the nearest selectable text position
  tr.setSelection(TextSelection.between(tr.doc.resolve(target.from), tr.doc.resolve(target.to)))
  tr.scrollIntoView()
  tr.setMeta(TRACK_IGNORE, true)
  editor.view.dispatch(tr)
  return true
}

export interface TrackChangesStorage {
  enabled: boolean
  author: string
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
    return { enabled: false, author: t('editorDefaultAuthor') }
  },

  addKeyboardShortcuts() {
    const storage = this.storage
    /**
     * Backspace over text already marked deleted is a no-op (the recorder puts it
     * straight back), which would trap the caret. Word skips such text, so step
     * over the struck run and delete the live character in front of it instead.
     */
    const skipStruckRun = (): boolean => {
      if (!storage.enabled) return false
      const { state, view } = this.editor
      const { selection, schema } = state
      const delType = schema.marks.del
      if (!delType || !selection.empty) return false
      const $pos = selection.$from
      if ($pos.parentOffset === 0) return false
      const isStruck = (offset: number): boolean => {
        const index = $pos.parent.childAfter(offset).index
        const child = $pos.parent.maybeChild(index)
        return !!child?.marks.some((mark) => mark.type === delType)
      }
      if (!isStruck($pos.parentOffset - 1)) return false
      const blockStart = $pos.start()
      let offset = $pos.parentOffset - 1
      while (offset > 0 && isStruck(offset - 1)) offset -= 1
      if (offset === 0) {
        // nothing live left in this block: park the caret before the struck run
        view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, blockStart)))
        return true
      }
      const runStart = blockStart + offset
      this.editor.commands.command(({ tr, dispatch }) => {
        tr.setSelection(TextSelection.create(tr.doc, runStart))
        if (dispatch) dispatch(tr)
        return true
      })
      return this.editor.commands.deleteRange({ from: runStart - 1, to: runStart })
    }
    return { Backspace: skipStruckRun }
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
        key: new PluginKey('trackChangesRecorder'),
        props: {
          handleDOMEvents: {
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

          const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
          const insMark = insType.create({ author: storage.author, date: now })
          const delMark = delType.create({ author: storage.author, date: now })

          const tr = newState.tr
          tr.setMeta(TRACK_IGNORE, true)
          let changed = false
          const recordedFormatRanges = new Set<string>()

          // Paragraph-property commands are SetNodeMarkup/ReplaceAround operations.
          // Comparing mapped textblocks catches alignment, spacing, indentation,
          // styles, heading conversions, and list level/type changes without
          // depending on the command's concrete ProseMirror step class.
          const fullMapping = new Mapping(
            tracked.flatMap((transaction) => transaction.steps.map((step) => step.getMap())),
          )
          const oldAnchors = new Map<number, { node: PmNode; pos: number }>()
          oldState.doc.nodesBetween(oldFrom, oldTo, (node, pos) => {
            const index = node.attrs?.docxIndex
            if (typeof index === 'number') oldAnchors.set(index, { node, pos })
            return false
          })
          const newAnchors = new Map<number, { node: PmNode; pos: number }>()
          newState.doc.nodesBetween(newFrom, newTo, (node, pos) => {
            const index = node.attrs?.docxIndex
            if (typeof index === 'number') newAnchors.set(index, { node, pos })
            return false
          })
          const matchedNewPositions = new Set<number>()
          const movedNewPositions = new Set<number>()
          const deletedBlocks: Array<{ node: PmNode; pos: number }> = []
          oldState.doc.nodesBetween(oldFrom, oldTo, (before, oldPos) => {
            const index = before.attrs?.docxIndex
            if (typeof index === 'number') {
              const match = newAnchors.get(index)
              if (match) {
                matchedNewPositions.add(match.pos)
                const mapped = fullMapping.map(oldPos, -1)
                if (
                  mapped !== match.pos &&
                  (before.attrs?.blockRevision as { kind?: string } | null)?.kind !== 'ins'
                ) {
                  deletedBlocks.push({ node: before, pos: mapped })
                  movedNewPositions.add(match.pos)
                }
              } else if (
                (before.attrs?.blockRevision as { kind?: string } | null)?.kind !== 'ins'
              ) {
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
          for (const { node, pos } of deletedBlocks.sort((a, b) => b.pos - a.pos)) {
            tr.insert(
              tr.mapping.map(pos, -1),
              node.type.create(
                {
                  ...node.attrs,
                  blockRevision: { kind: 'del', author: storage.author, date: now },
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
              node.attrs?.blockRevision
            )
              return false
            tr.setNodeMarkup(tr.mapping.map(pos, 1), undefined, {
              ...node.attrs,
              blockRevision: { kind: 'ins', author: storage.author, date: now },
            })
            changed = true
            return false
          })

          // Native table row commands preserve the table anchor. Record the
          // single contiguous row span they insert/delete using trPr revisions.
          for (const [index, oldTable] of oldAnchors) {
            const nextTable = newAnchors.get(index)
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
                  rowRevision: { kind: 'ins', author: storage.author, date: now },
                  rawTrPr:
                    `<w:trPr><w:ins w:id="0" w:author="${xmlAttr(storage.author)}"` +
                    ` w:date="${now}"/></w:trPr>`,
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
                      rowRevision: { kind: 'del', author: storage.author, date: now },
                      rawTrPr:
                        `<w:trPr><w:del w:id="0" w:author="${xmlAttr(storage.author)}"` +
                        ` w:date="${now}"/></w:trPr>`,
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
            if (!PARAGRAPH_NODE_TYPES.has(before.type.name)) return
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
                author: storage.author,
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
              if (!rprType) return
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
                      author: storage.author,
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
                    (m) => m.type === insType && m.attrs.author === storage.author,
                  )
                  if (ownIns) return // deleting your own insertion removes it outright
                  if (child.marks.some((m) => m.type === delType)) keep.push(child)
                  else keep.push(child.mark([...child.marks, delMark]))
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
                  !child.marks.some((m) => m.type === insType && m.attrs.author === storage.author),
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

          return changed ? tr : null
        },
      }),
    ]
  },
})
