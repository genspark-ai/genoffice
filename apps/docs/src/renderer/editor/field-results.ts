/**
 * Caption numbers as results: a field saved with its result needs no `w:dirty`, so Word
 * does not ask to update on open.
 */
import type { Editor } from '@tiptap/core'
import type { Mark, Node as PmNode, Schema } from '@tiptap/pm/model'
import { Transform } from '@tiptap/pm/transform'
import {
  SeqCounter,
  parseSeqInstruction,
  rewriteSeqResultsXml,
  seqFieldsOfXml,
  unreliableSeqLabels,
  type Block,
  type FieldDisplay,
  type SeqInstruction,
} from '@genoffice/docx-engine'
import { blocksToPmDoc } from './convert'
import { collectFields, type FieldRange } from './field-codes'

/** the numbers a renumbering gave the SEQ fields of a block that keeps its place in the file */
export interface SeqRenumber {
  /** one per SEQ field of the block's XML; null where the file's own result stands */
  results: Array<string | null>
  /** the text the file's block shows with these numbers; null when that cannot be told */
  left: string | null
}

const unreliableOf = new WeakMap<Block[], Set<string>>()

/**
 * Labels whose numbers in the file are not what counting gives, read from its XML or as
 * the editor holds it: renumbering them would change captions nobody touched.
 */
export function unreliableSeqLabelsOf(blocks: Block[], schema: Schema): Set<string> {
  let labels = unreliableOf.get(blocks)
  if (!labels) {
    labels = unreliableSeqLabels(blocks.map((b) => b.originalXml))
    unreliableOf.set(blocks, labels)
    const captioned = blocks.filter((b) => b.originalXml && /SEQ\s/i.test(b.originalXml))
    if (captioned.length > 0) {
      try {
        const pristine = new Transform(schema.nodeFromJSON(blocksToPmDoc(blocks)))
        renumberSeqFields(pristine, blocks, { skip: labels, changedLabels: labels })
      } catch {
        // the file cannot be read back for the comparison: none of its labels is renumbered
        for (const b of captioned)
          for (const f of seqFieldsOfXml(b.originalXml!)) {
            const id = parseSeqInstruction(f.instr)?.id
            if (id) labels.add(id)
          }
      }
    }
  }
  return labels
}

const unescapeXml = (text: string): string =>
  text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')

const textOfXml = (xml: string): string =>
  unescapeXml([...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(''))

function baseXmlOf(
  node: PmNode,
  byIndex: Map<number, Block>,
): { xml: string; source: Block | null } {
  if (node.attrs.genXml) return { xml: String(node.attrs.genXml), source: null }
  const idx = node.attrs.docxIndex as number | null
  const source = idx === null || idx === undefined ? null : (byIndex.get(idx) ?? null)
  return { xml: source?.originalXml ?? '', source }
}

function blockIndex(blocks: Block[]): Map<number, Block> {
  const byIndex = new Map<number, Block>()
  for (const b of blocks) if (typeof b.docxIndex === 'number') byIndex.set(b.docxIndex, b)
  return byIndex
}

const offsetOfResult = (xml: string, at: number): number => textOfXml(xml.slice(0, at)).length

function leftOfRenumbered(source: Block, results: ReadonlyArray<string | null>): string | null {
  const xml = source.originalXml ?? ''
  const edits = seqFieldsOfXml(xml).flatMap((field, i) => {
    const to = results[i]
    return to === null || to === undefined || field.spans.length === 0 || to === field.result
      ? []
      : [{ at: offsetOfResult(xml, field.spans[0][0]), from: field.result, to }]
  })
  if (edits.length === 0) return source.fieldDisplay?.left ?? null
  return renumberedDisplay(source.fieldDisplay, edits)?.left ?? null
}

interface XmlSite {
  kind: 'xml'
  seq: SeqInstruction
  current: string
  dirty: boolean
  /** index into `nodes`, -1 for a block whose fields are counted but not written */
  owner: number
  field: number
  at: number
}

interface MarkSite {
  kind: 'mark'
  seq: SeqInstruction
  current: string
  dirty: boolean
  range: FieldRange
}

interface XmlOwner {
  pos: number
  node: PmNode
  /** the XML with the numbers of an earlier renumbering in it */
  xml: string
  source: Block | null
  count: number
  results: Array<string | null>
}

const BOOKMARK_START = /<w:bookmarkStart\b[^>]*?w:id="([^"]*)"[^>]*?w:name="([^"]*)"[^>]*>/g

function bookmarkTexts(xml: string): Map<string, string> {
  const out = new Map<string, string>()
  for (const m of xml.matchAll(BOOKMARK_START)) {
    const end = new RegExp(`<w:bookmarkEnd\\b[^>]*w:id="${m[1]}"[^>]*>`).exec(xml)
    const from = m.index + m[0].length
    out.set(m[2], textOfXml(xml.slice(from, end && end.index > from ? end.index : xml.length)))
  }
  return out
}

/** a REF whose result is the bookmark's text: `\h` and `\*` leave it so */
const plainRef = (instr: string | null): boolean =>
  !instr ||
  instr
    .replace(/^\s*(?:REF\s+)?(?:"[^"]+"|[^\s\\]+)/, '')
    .replace(/\\h\b|\\\*\s+\S+/g, '')
    .trim() === ''

function withDirty(marks: readonly Mark[], field: Mark, dirty: boolean): Mark[] {
  return marks.map((m) => (m.eq(field) ? m.type.create({ ...m.attrs, dirty }) : m))
}

function writeField(tr: Transform, range: FieldRange, text: string): void {
  const { schema } = tr.doc.type
  if (text === tr.doc.textBetween(range.from, range.to, '')) {
    // same result: only the flag goes, so formatting on part of it stays where it is
    const nodes: Array<{ pos: number; node: PmNode }> = []
    tr.doc.nodesBetween(range.from, range.to, (node, pos) => {
      if (node.isText) nodes.push({ pos, node })
    })
    for (const { pos, node } of nodes.reverse())
      tr.replaceWith(
        pos,
        pos + node.nodeSize,
        schema.text(node.text!, withDirty(node.marks, range.mark, false)),
      )
    return
  }
  const marks = tr.doc.nodeAt(range.from)?.marks ?? [range.mark]
  tr.replaceWith(
    range.from,
    range.to,
    schema.text(text || ' ', withDirty(marks, range.mark, false)),
  )
}

/** the display with the numbers replaced; null when a number is not where the XML has it */
function renumberedDisplay(
  display: FieldDisplay | null | undefined,
  edits: Array<{ at: number; from: string; to: string }>,
): FieldDisplay | null {
  if (!display || display.kind !== 'text' || typeof display.left !== 'string') return null
  let left = display.left
  let runs = display.runs
  let shift = 0
  for (const edit of edits) {
    const at = edit.at + shift
    if (left.slice(at, at + edit.from.length) !== edit.from) return null
    if (runs && runs.map((r) => r.text).join('') === left) {
      let start = 0
      const i = runs.findIndex((r) => {
        const inside = at >= start && at + edit.from.length <= start + r.text.length
        if (!inside) start += r.text.length
        return inside
      })
      runs =
        i < 0
          ? undefined
          : runs.map((r, n) =>
              n === i
                ? {
                    ...r,
                    text:
                      r.text.slice(0, at - start) +
                      edit.to +
                      r.text.slice(at - start + edit.from.length),
                  }
                : r,
            )
    } else runs = undefined
    left = left.slice(0, at) + edit.to + left.slice(at + edit.from.length)
    shift += edit.to.length - edit.from.length
  }
  const next: FieldDisplay = { ...display, left }
  if (runs) next.runs = runs
  else delete next.runs
  return next
}

export interface RenumberOptions {
  skip?: ReadonlySet<string>
  /** an explicit update: a field keeps no `w:dirty` even where its number stands */
  dropDirty?: boolean
  changedLabels?: Set<string>
}

/** numbers each label's SEQ fields in document order, with the references that quote a caption; returns how many changed */
export function renumberSeqFields(
  tr: Transform,
  blocks: Block[],
  options: RenumberOptions = {},
): number {
  const doc = tr.doc
  const byIndex = blockIndex(blocks)
  const owners: XmlOwner[] = []
  const sites: Array<XmlSite | MarkSite> = []
  const markType = doc.type.schema.marks.instrField
  const fields = markType ? collectFields(doc, markType) : []
  let nextField = 0

  const xmlSites = (xml: string, owner: number) => {
    seqFieldsOfXml(xml).forEach((field, i) => {
      const seq = parseSeqInstruction(field.instr)
      if (!seq) return
      sites.push({
        kind: 'xml',
        seq,
        current: field.result,
        dirty: field.dirtyAttr !== null,
        owner: field.spans.length > 0 ? owner : -1,
        field: i,
        at: field.spans[0]?.[0] ?? 0,
      })
    })
  }

  doc.forEach((node, pos) => {
    const end = pos + node.nodeSize
    if (node.type.name === 'docProtected') {
      const base = baseXmlOf(node, byIndex)
      if (!/SEQ\s/i.test(base.xml)) return
      const earlier = base.source ? (node.attrs.seqRenumber as SeqRenumber | null) : null
      const xml = earlier ? rewriteSeqResultsXml(base.xml, earlier.results) : base.xml
      const count = seqFieldsOfXml(xml).length
      // only a caption paragraph is rewritten: the captions of a table or a text box are counted as they stand
      const writable =
        /^<w:p[\s>]/.test(xml) &&
        (node.attrs.fieldDisplay as FieldDisplay | null)?.kind === 'text' &&
        !/<w:(?:txbxContent|drawing|pict|object)\b/.test(xml)
      if (writable)
        owners.push({
          pos,
          node,
          xml,
          source: base.source,
          count,
          results: earlier ? [...earlier.results] : new Array<string | null>(count).fill(null),
        })
      xmlSites(xml, writable ? owners.length - 1 : -1)
      return
    }
    // what the file holds inside an editable block (a caption in a table cell) is counted from its XML
    const idx = node.attrs.docxIndex as number | null | undefined
    const original = typeof idx === 'number' ? byIndex.get(idx)?.originalXml : undefined
    if (original && /SEQ\s/i.test(original)) xmlSites(original, -1)
    while (nextField < fields.length && fields[nextField].from < end) {
      const range = fields[nextField++]
      if (range.from < pos) continue
      const seq = parseSeqInstruction(range.instr)
      if (!seq) continue
      sites.push({
        kind: 'mark',
        seq,
        current: doc.textBetween(range.from, range.to, ''),
        dirty: range.mark.attrs.dirty === true,
        range,
      })
    }
  })
  if (sites.length === 0) return 0

  const skip = new Set(options.skip)
  for (const site of sites) if (site.seq.unsupported) skip.add(site.seq.id)

  const counter = new SeqCounter()
  const changedLabels = new Set<string>()
  let changed = 0
  const markJobs: Array<{ range: FieldRange; text: string }> = []
  const displayEdits = new Map<number, Array<{ at: number; from: string; to: string }>>()
  for (const site of sites) {
    const { text } = counter.next(site.seq)
    if (text === null || skip.has(site.seq.id)) continue
    const differs = text !== site.current.trim()
    if (site.kind === 'mark') {
      // a SEQ the editor holds as a mark was made in this session: its number is ours
      if (differs || site.dirty) markJobs.push({ range: site.range, text })
      if (differs) {
        changed++
        changedLabels.add(site.seq.id)
      }
      continue
    }
    // a caption counted but not written that the count disagrees with still tells on its label
    if (site.owner < 0 && differs) changedLabels.add(site.seq.id)
    if (site.owner < 0 || !(differs || (options.dropDirty && site.dirty))) continue
    const owner = owners[site.owner]
    owner.results[site.field] = text
    if (!differs) continue
    changed++
    changedLabels.add(site.seq.id)
    const edits = displayEdits.get(site.owner) ?? []
    edits.push({ at: offsetOfResult(owner.xml, site.at), from: site.current, to: text })
    displayEdits.set(site.owner, edits)
  }

  const quoted = new Map<string, { from: string; to: string }>()
  const start = tr.steps.length
  const parents = new Map<number, { names: string[]; text: string }>()
  for (const { range } of markJobs) {
    const $from = doc.resolve(range.from)
    const names = [
      ...((($from.parent.attrs.bookmarks as string[] | null) ?? []) as string[]),
      ...((($from.parent.attrs.hiddenBookmarks as string[] | null) ?? []) as string[]),
    ]
    if (names.length > 0) parents.set($from.before(), { names, text: $from.parent.textContent })
  }

  const steps: Array<{ pos: number; run: () => void }> = []
  for (const job of markJobs)
    steps.push({ pos: job.range.from, run: () => writeField(tr, job.range, job.text) })
  owners.forEach((owner, i) => {
    const rewritten = owner.results.some((r, n) => r !== null && n < owner.count)
    if (!rewritten) return
    const earlier = owner.source ? (owner.node.attrs.seqRenumber as SeqRenumber | null) : null
    const fresh = owner.results.map((r, n) => (r !== null && r !== earlier?.results[n] ? r : null))
    const xml = rewriteSeqResultsXml(owner.xml, fresh)
    if (xml === owner.xml) return
    const before = bookmarkTexts(owner.xml)
    for (const [name, to] of bookmarkTexts(xml)) {
      const from = before.get(name)
      if (from !== undefined && from !== to) quoted.set(name, { from, to })
    }
    const shown = owner.node.attrs.fieldDisplay as FieldDisplay | null
    // a text edited since keeps what it shows: the number is no longer where the XML has it
    const display = renumberedDisplay(shown, displayEdits.get(i) ?? []) ?? shown
    const { source } = owner
    steps.push({
      pos: owner.pos,
      run: () =>
        tr.setNodeMarkup(owner.pos, undefined, {
          ...owner.node.attrs,
          fieldDisplay: display,
          ...(source
            ? {
                seqRenumber: {
                  results: owner.results,
                  left: leftOfRenumbered(source, owner.results),
                } satisfies SeqRenumber,
              }
            : { genXml: xml, seqRenumber: null }),
        }),
    })
  })
  for (const step of steps.sort((a, b) => b.pos - a.pos)) step.run()

  for (const [pos, parent] of parents) {
    const node = tr.doc.nodeAt(tr.mapping.slice(start).map(pos))
    if (!node || node.textContent === parent.text) continue
    for (const name of parent.names) quoted.set(name, { from: parent.text, to: node.textContent })
  }
  for (const label of changedLabels) options.changedLabels?.add(label)
  if (quoted.size > 0) {
    const refType = tr.doc.type.schema.marks.refField
    const refs = refType ? collectFields(tr.doc, refType) : []
    for (const ref of refs.reverse()) {
      const change = quoted.get(String(ref.mark.attrs.name))
      if (!change || !plainRef(ref.mark.attrs.instr as string | null)) continue
      if (tr.doc.textBetween(ref.from, ref.to, '').trim() !== change.from.trim()) continue
      writeField(tr, ref, change.to.trim())
      changed++
    }
  }
  return changed
}

/** before a save: nothing is left for Word to recompute; not a user edit, so out of the undo history */
export function settleFieldResults(editor: Editor, blocks: Block[]): void {
  const skip = unreliableSeqLabelsOf(blocks, editor.schema)
  const tr = editor.state.tr
  renumberSeqFields(tr, blocks, { skip })
  if (tr.docChanged) editor.view.dispatch(tr.setMeta('addToHistory', false))
}

/** after a caption went in, in the same undo step */
export function renumberAfterInsert(editor: Editor, blocks: Block[]): void {
  const skip = unreliableSeqLabelsOf(blocks, editor.schema)
  const tr = editor.state.tr
  renumberSeqFields(tr, blocks, { skip })
  if (tr.docChanged) editor.view.dispatch(tr.setMeta('appendedTransaction', tr))
}
