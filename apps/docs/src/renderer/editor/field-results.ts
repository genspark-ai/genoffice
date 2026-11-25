/**
 * Results for the fields this editor computes (caption numbers, page fields): a field
 * saved with its result needs no `w:dirty`, so Word does not ask to update on open.
 */
import type { Editor } from '@tiptap/core'
import type { Mark, Node as PmNode, Schema } from '@tiptap/pm/model'
import { Transform } from '@tiptap/pm/transform'
import {
  NUMBERING_FIELD,
  STYLEREF_NOT_FOUND,
  SeqCounter,
  computeListMarkerInfos,
  computeListNumberSites,
  fieldNumberText,
  isFieldNumberFormat,
  numberingEventsOfXml,
  paragraphNumberText,
  parseSeqInstruction,
  parseStyleRefInstruction,
  rewriteSeqResultsXml,
  seqFieldsOfXml,
  setCaptionNumberingXml,
  styleRefResult,
  unreliableSeqLabels,
  type Block,
  type CaptionNumbering,
  type FieldDisplay,
  type ListMarkerInfo,
  type ListNumberSite,
  type NoteInfo,
  type NumberingDef,
  type SeqInstruction,
  type StyleInfo,
  type StyleRefInstruction,
  type StyleRefSurroundings,
  type StyledParagraph,
  type TextboxDisplay,
  type XmlSeqField,
} from '@genoffice/docx-engine'
import { blocksToPmDoc } from './convert'
import { listItemRefsOf, type ListNumberingStorage } from './extensions'
import { PAGE_FIELD_INSTR, collectFields, isFieldLocked, type FieldRange } from './field-codes'
import {
  TEXT_FORMAT_SWITCHES,
  formatFieldNumber,
  formatFieldText,
  formatSwitch,
  formatSwitchOf,
  numericPictureOf,
} from './field-switches'

/** the numbers a renumbering gave the SEQ fields of a block that keeps its place in the file */
export interface SeqRenumber {
  /** one per SEQ field of the block's XML; null where the file's own result stands */
  results: Array<string | null>
  /** the text the file's block shows with these numbers; null when that cannot be told */
  left: string | null
}

/** what tells the headings from the other paragraphs and gives them their numbers */
export interface NumberingContext {
  defs: Map<string, NumberingDef>
  styles?: Map<string, StyleInfo>
}

export function numberingContextOf(editor: Editor): NumberingContext | undefined {
  const storage = editor.storage.listNumbering as ListNumberingStorage | undefined
  return storage ? { defs: storage.defs, styles: storage.styles } : undefined
}

const levelsOfStyles = new WeakMap<Map<string, StyleInfo>, Map<string, number>>()

function headingLevelsOf(context: NumberingContext): Map<string, number> {
  if (!context.styles) return new Map()
  let levels = levelsOfStyles.get(context.styles)
  if (!levels) {
    levels = new Map()
    for (const [id, style] of context.styles)
      if (style.headingLevel) levels.set(id, style.headingLevel)
    levelsOfStyles.set(context.styles, levels)
  }
  return levels
}

function styleRefText(
  ref: StyleRefInstruction,
  paragraphs: readonly StyledParagraph[],
  before: number,
  around: StyleRefSurroundings,
): string | null {
  const text = styleRefResult(ref, paragraphs, before, around)
  return text === null || !ref.textCase || text === STYLEREF_NOT_FOUND
    ? text
    : formatFieldText(formatSwitch(ref.textCase), text)
}

/** stands in the skip set for every STYLEREF: no SEQ identifier starts with a space */
export const STYLEREF_LABEL = ' STYLEREF'
/** the ones that show a number in context or a place: a result that disagrees there says nothing of the others */
export const STYLEREF_PLACED_LABEL = ' STYLEREF placed'

const styleRefLabel = (ref: StyleRefInstruction | null): string =>
  ref && (ref.context || ref.aboveBelow) ? STYLEREF_PLACED_LABEL : STYLEREF_LABEL

/** a field whose result depends on the headings before it */
const HEADING_BOUND = /STYLEREF\s|SEQ\s[^<]*\\s\b/i

// per heading styles: which labels count right depends on where the headings are
const unreliableOf = new WeakMap<Block[], Map<unknown, Set<string>>>()

/**
 * Labels whose numbers in the file are not what counting gives, read from its XML or as
 * the editor holds it: renumbering them would change captions nobody touched.
 */
export function unreliableSeqLabelsOf(
  blocks: Block[],
  schema: Schema,
  headings?: NumberingContext,
): Set<string> {
  const byStyles = unreliableOf.get(blocks) ?? new Map<unknown, Set<string>>()
  unreliableOf.set(blocks, byStyles)
  let labels = byStyles.get(headings?.styles)
  if (!labels) {
    labels = unreliableSeqLabels(
      blocks.map((b) => b.originalXml),
      headings && headingLevelsOf(headings),
    )
    byStyles.set(headings?.styles, labels)
    const captioned = blocks.filter((b) => b.originalXml && NUMBERING_FIELD.test(b.originalXml))
    if (captioned.length > 0) {
      try {
        const pristine = new Transform(schema.nodeFromJSON(blocksToPmDoc(blocks)))
        renumberSeqFields(pristine, blocks, { skip: labels, changedLabels: labels, headings })
      } catch {
        // the file cannot be read back for the comparison: none of its labels is renumbered
        for (const b of captioned)
          for (const f of seqFieldsOfXml(b.originalXml!)) {
            labels.add(
              parseSeqInstruction(f.instr)?.id ?? styleRefLabel(parseStyleRefInstruction(f.instr)),
            )
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

/** the text of a block's runs, a non-breaking hyphen as the display has it */
export const textOfXml = (xml: string): string =>
  unescapeXml(
    [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:(noBreakHyphen|softHyphen)\s*\/>/g)]
      .map((m) => (m[2] ? (m[2] === 'noBreakHyphen' ? '\u2011' : '\u00ad') : m[1]))
      .join(''),
  )

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

interface FieldSite {
  seq: SeqInstruction | null
  ref: StyleRefInstruction | null
  /** how many of the document's styled paragraphs stand before the field */
  before: number
}

interface XmlSite extends FieldSite {
  kind: 'xml'
  current: string
  dirty: boolean
  /** index into `nodes`, -1 for a block whose fields are counted but not written */
  owner: number
  field: number
  at: number
}

interface MarkSite extends FieldSite {
  kind: 'mark'
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

const sameSwitch = (list: readonly string[], name: string): boolean =>
  list.some((s) => s.toLowerCase() === name.toLowerCase())

/** a REF whose result is the bookmark's text, in the case a `\*` asks for: `\h` leaves it so */
function plainRef(instr: string | null): boolean {
  if (!instr) return true
  const format = formatSwitchOf(instr)
  if (/\\\*\s*DBCHAR\b/i.test(instr) || (format && !sameSwitch(TEXT_FORMAT_SWITCHES, format)))
    return false
  return (
    instr
      .replace(/^\s*(?:REF\s+)?(?:"[^"]+"|[^\s\\]+)/, '')
      .replace(/\\h\b|\\\*\s+\S+/g, '')
      .trim() === ''
  )
}

function withDirty(marks: readonly Mark[], field: Mark, dirty: boolean): Mark[] {
  return marks.map((m) => (m.eq(field) ? m.type.create({ ...m.attrs, dirty }) : m))
}

/** the field shows `text` and no longer waits for an update */
export function writeField(tr: Transform, range: FieldRange, text: string): void {
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
  /** without it a number that restarts at a heading, and a STYLEREF, stay as they are */
  headings?: NumberingContext
  /** an explicit update: a field keeps no `w:dirty` even where its number stands */
  dropDirty?: boolean
  changedLabels?: Set<string>
  /** STYLEREF \\p in the language of the interface; without them such a field stays as it is */
  words?: { above: string; below: string }
}

/** only a caption paragraph is rewritten: the captions of a table or a text box are counted as they stand */
const captionParagraph = (xml: string, node: PmNode): boolean =>
  /^<w:p[\s>]/.test(xml) &&
  (node.attrs.fieldDisplay as FieldDisplay | null)?.kind === 'text' &&
  !/<w:(?:txbxContent|drawing|pict|object)\b/.test(xml)

/** a protected block's XML with the numbers of an earlier renumbering in it */
function protectedXmlOf(node: PmNode, byIndex: Map<number, Block>) {
  const base = baseXmlOf(node, byIndex)
  const earlier = base.source ? (node.attrs.seqRenumber as SeqRenumber | null) : null
  const xml = earlier ? rewriteSeqResultsXml(base.xml, earlier.results) : base.xml
  return { xml, source: base.source, earlier }
}

function headingsBound(doc: PmNode, byIndex: Map<number, Block>, fields: FieldRange[]): boolean {
  if (fields.some((f) => HEADING_BOUND.test(f.instr))) return true
  let bound = false
  doc.forEach((node) => {
    if (bound) return
    const idx = node.attrs.docxIndex as number | null | undefined
    const xml =
      node.type.name === 'docProtected'
        ? baseXmlOf(node, byIndex).xml
        : typeof idx === 'number'
          ? byIndex.get(idx)?.originalXml
          : undefined
    bound = !!xml && HEADING_BOUND.test(xml)
  })
  return bound
}

/**
 * Numbers each label's SEQ fields in document order and gives each STYLEREF the heading
 * it points at, with the references that quote a caption; returns how many changed.
 */
export function renumberSeqFields(
  tr: Transform,
  blocks: Block[],
  options: RenumberOptions = {},
): number {
  const doc = tr.doc
  const byIndex = blockIndex(blocks)
  const owners: XmlOwner[] = []
  const sites: Array<XmlSite | MarkSite | { kind: 'heading'; level: number }> = []
  const markType = doc.type.schema.marks.instrField
  const fields = markType ? collectFields(doc, markType) : []
  let nextField = 0
  const context = options.headings
  const levels =
    context && headingsBound(doc, byIndex, fields) ? headingLevelsOf(context) : undefined
  const paragraphs: StyledParagraph[] = []
  const markers = new Map<number, ListMarkerInfo | null>()
  const itemAt = new Map<number, number>()
  let numberSites: ReadonlyArray<ListNumberSite | null> = []
  let itemNodes: Array<{ pos: number; node: PmNode }> = []
  if (levels && context) {
    const { refs, nodes } = listItemRefsOf(doc)
    computeListMarkerInfos(refs, context.defs).forEach((m, i) => markers.set(nodes[i].pos, m))
    nodes.forEach((n, i) => itemAt.set(n.pos, i))
    numberSites = computeListNumberSites(refs, context.defs)
    itemNodes = nodes
  }
  const itemOf = new Map<StyledParagraph, number>()
  /** the list item a position is in, for a number relative to it */
  const itemHolding = (pos: number): number | undefined => {
    const at = itemNodes.findIndex((n) => pos >= n.pos && pos < n.pos + n.node.nodeSize)
    return at < 0 ? undefined : at
  }

  const styled = (paragraph: StyledParagraph) => {
    paragraphs.push(paragraph)
    if (paragraph.headingLevel) sites.push({ kind: 'heading', level: paragraph.headingLevel })
  }
  const styledNode = (node: PmNode, pos: number) => {
    const styleId = (node.attrs.styleId as string | null) ?? null
    const style = styleId ? context?.styles?.get(styleId) : undefined
    const own =
      node.type.name === 'docHeading' && !node.attrs.outlineOnly ? Number(node.attrs.level) : null
    const headingLevel = style?.headingLevel ?? own
    if (!styleId && !headingLevel) return
    const marker = markers.get(pos)
    const paragraph: StyledParagraph = {
      styleId,
      styleName: style?.name,
      headingLevel,
      text: node.textContent,
      marker: marker?.text,
      number: marker?.number,
    }
    const item = itemAt.get(pos)
    if (item !== undefined) itemOf.set(paragraph, item)
    styled(paragraph)
  }
  const what = (instr: string): Pick<FieldSite, 'seq' | 'ref'> | null => {
    const seq = parseSeqInstruction(instr)
    const ref = seq ? null : parseStyleRefInstruction(instr)
    return seq || ref ? { seq, ref } : null
  }
  const xmlSite = (field: XmlSeqField, index: number, owner: number) => {
    const found = what(field.instr)
    if (!found) return
    sites.push({
      kind: 'xml',
      ...found,
      before: paragraphs.length,
      current: field.result,
      dirty: field.dirtyAttr !== null,
      owner: field.spans.length > 0 && !field.locked ? owner : -1,
      field: index,
      at: field.spans[0]?.[0] ?? 0,
    })
  }

  doc.forEach((node, pos) => {
    const end = pos + node.nodeSize
    if (node.type.name === 'docProtected') {
      const { xml, source, earlier } = protectedXmlOf(node, byIndex)
      const fielded = NUMBERING_FIELD.test(xml)
      if (!fielded && !levels) return
      const writable = fielded && captionParagraph(xml, node)
      if (writable) {
        const count = seqFieldsOfXml(xml).length
        owners.push({
          pos,
          node,
          xml,
          source,
          count,
          results: earlier ? [...earlier.results] : new Array<string | null>(count).fill(null),
        })
      }
      for (const event of numberingEventsOfXml(xml, levels))
        if (event.kind === 'field')
          xmlSite(event.field, event.index, writable ? owners.length - 1 : -1)
        // what a paragraph kept as XML shows is not read here
        else
          styled({
            styleId: event.styleId,
            styleName: context?.styles?.get(event.styleId)?.name,
            headingLevel: event.level,
            text: '',
            unknown: true,
          })
      return
    }
    // what the file holds inside an editable block (a caption in a table cell) is counted from its XML
    const idx = node.attrs.docxIndex as number | null | undefined
    const original = typeof idx === 'number' ? byIndex.get(idx)?.originalXml : undefined
    if (original && NUMBERING_FIELD.test(original))
      seqFieldsOfXml(original).forEach((field, i) => xmlSite(field, i, -1))
    const inner: Array<{ pos: number; run: () => void }> = []
    if (levels) {
      if (node.isTextblock) inner.push({ pos, run: () => styledNode(node, pos) })
      else
        node.descendants((child, at) => {
          const childPos = pos + 1 + at
          if (child.isTextblock)
            inner.push({ pos: childPos, run: () => styledNode(child, childPos) })
          return !child.isTextblock
        })
    }
    while (nextField < fields.length && fields[nextField].from < end) {
      const range = fields[nextField++]
      if (range.from < pos) continue
      const found = what(range.instr)
      if (!found) continue
      inner.push({
        pos: range.from,
        run: () =>
          sites.push({
            kind: 'mark',
            ...found,
            before: paragraphs.length,
            current: doc.textBetween(range.from, range.to, ''),
            dirty: range.mark.attrs.dirty === true,
            range,
          }),
      })
    }
    for (const item of inner.sort((a, b) => a.pos - b.pos)) item.run()
  })
  if (!sites.some((site) => site.kind !== 'heading')) return 0

  const skip = new Set(options.skip)
  for (const site of sites)
    if (site.kind !== 'heading' && site.seq)
      if (site.seq.unsupported || (site.seq.headingLevel !== null && !levels)) skip.add(site.seq.id)

  const counter = new SeqCounter()
  const changedLabels = new Set<string>()
  let changed = 0
  const markJobs: Array<{ range: FieldRange; text: string }> = []
  const displayEdits = new Map<number, Array<{ at: number; from: string; to: string }>>()
  for (const site of sites) {
    if (site.kind === 'heading') {
      counter.heading(site.level)
      continue
    }
    const label = site.seq ? site.seq.id : styleRefLabel(site.ref)
    const from = site.kind === 'mark' ? itemHolding(site.range.from) : undefined
    const text = site.seq
      ? counter.next(site.seq).text
      : levels && site.ref
        ? styleRefText(site.ref, paragraphs, site.before, {
            words: options.words,
            numberOf: (paragraph, numberContext, bare) => {
              const item = itemOf.get(paragraph)
              return item === undefined
                ? null
                : paragraphNumberText(numberSites, item, numberContext, { from, bare })
            },
          })
        : null
    if (text === null || skip.has(label)) continue
    const differs = text.trim() !== site.current.trim()
    if (site.kind === 'mark') {
      if (isFieldLocked(site.range.mark)) continue
      // a field the editor holds as a mark was made in this session: its result is ours
      if (differs || site.dirty) markJobs.push({ range: site.range, text })
      if (differs) {
        changed++
        changedLabels.add(label)
      }
      continue
    }
    // a caption counted but not written that the count disagrees with still tells on its label
    if (site.owner < 0 && differs) changedLabels.add(label)
    if (site.owner < 0 || !(differs || (options.dropDirty && site.dirty))) continue
    const owner = owners[site.owner]
    owner.results[site.field] = text
    if (!differs) continue
    changed++
    changedLabels.add(label)
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
      if (!change || isFieldLocked(ref.mark) || !plainRef(ref.mark.attrs.instr as string | null))
        continue
      const shaped = (text: string) => formatFieldText(ref.instr, text.trim())
      if (tr.doc.textBetween(ref.from, ref.to, '').trim() !== shaped(change.from)) continue
      writeField(tr, ref, shaped(change.to))
      changed++
    }
  }
  return changed
}

/** one laid-out page, as the layout that drew it has it */
export interface PageFacts {
  /** its place among the laid-out pages: two positions share a page when these are equal */
  index: number
  number: number
  /** the number as the page shows it: the section's format, behind its chapter number */
  display: string
  chaptered: boolean
  sectionPages: number
}

export interface PageWords {
  above: string
  below: string
  onPage: (page: string) => string
}

/** a page layout that stands for the document as it is; there is none while a pass is due or outside the page view */
export interface PageSource {
  /** the page the character at `pos` is drawn on, or with `edge` the page that edge of the node at `pos` is on */
  pageAt: (pos: number, edge?: 'top' | 'bottom') => PageFacts | null
  pageCount: number
  /** changed since it was opened or saved: until then what the file holds stands */
  edited: boolean
  /** PAGEREF \p in the language of the interface */
  words?: PageWords
  /** the first page showing the header or footer of that kind of page in the section ending at block `sectionEnd` */
  hfPage?: (sectionEnd: number, variant: 'default' | 'first' | 'even') => PageFacts | null
  /** the page that draws the footnote holding the bookmark; null when no footnote of the file holds it */
  noteBookmarkPage?: (name: string) => PageFacts | null
}

/**
 * The footnote that holds the bookmark in the file. One whose text changed since is
 * written anew as plain text, without its bookmarks.
 */
export function footnoteOfBookmark(
  name: string,
  notes: readonly NoteInfo[],
  fromFile: readonly NoteInfo[],
): string | null {
  const held = fromFile.find((note) => note.bookmarks?.includes(name))
  if (!held) return null
  const now = notes.find((note) => note.id === held.id)
  return now && !now.created && !now.origin && now.text === held.text ? held.id : null
}

const WIDTH_FORMAT = /\\\*\s*(DBCHAR|SBCHAR)\b/gi

/** `\* DBCHAR` / `\* SBCHAR`: Arabic digits, full width or half width (ECMA-376 17.16.4.3) */
function digitWidthOf(instr: string): 'full' | 'half' | null {
  const last = [...instr.matchAll(WIDTH_FORMAT)].pop()
  return !last ? null : last[1].toUpperCase() === 'DBCHAR' ? 'full' : 'half'
}

/** a `\*` format this editor has no rule for: the result is Word's to compute */
function unreadableFormat(instr: string): boolean {
  const format = formatSwitchOf(instr.replace(WIDTH_FORMAT, ''))
  // a width beside another format or a picture: which of them Word lets win is not known here
  if (digitWidthOf(instr)) return !!format || numericPictureOf(instr) !== null
  return !!format && !isFieldNumberFormat(format) && !sameSwitch(TEXT_FORMAT_SWITCHES, format)
}

function countText(instr: string, value: number): string | null {
  if (unreadableFormat(instr)) return null
  const width = digitWidthOf(instr)
  const format = formatSwitchOf(instr)
  // a value the format has no rule for is Word's to compute, not shown as plain digits
  if (!width && format && isFieldNumberFormat(format) && numericPictureOf(instr) === null)
    return fieldNumberText(value, format)
  if (!width) return formatFieldNumber(instr, value)
  const digits = String(value)
  return width === 'half'
    ? digits
    : digits.replace(/\d/g, (d) => String.fromCharCode(0xff10 + Number(d)))
}

/** the field's own switches give its number a shape */
function shapesNumber(instr: string): boolean {
  const format = formatSwitchOf(instr)
  return (
    digitWidthOf(instr) !== null ||
    numericPictureOf(instr) !== null ||
    (!!format && isFieldNumberFormat(format))
  )
}

/** a page number as the page shows it, unless the field's switches shape it themselves */
function pageText(instr: string, page: PageFacts): string | null {
  if (unreadableFormat(instr)) return null
  if (!shapesNumber(instr)) return formatFieldText(instr, page.display)
  // a chapter number in front of a number the field formats itself: Word's to compute
  return page.chaptered ? null : countText(instr, page.number)
}

const escapeRegExp = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

interface BookmarkSite {
  pos: number
  end: number
  /** the bookmark opens where its block does, so the block's first line is its page */
  atStart: boolean
  /** where inside its paragraph the bookmark opens, while the paragraph reads as the file has it */
  inner?: number
}

function placeInParagraph(node: PmNode, pos: number, xml: string, at: number): number | null {
  if (!node.isTextblock || node.textBetween(0, node.content.size, '') !== textOfXml(xml))
    return null
  let left = textOfXml(xml.slice(0, at)).length
  let place = pos + 1
  for (let i = 0; i < node.childCount; i++) {
    const child = node.child(i)
    const length = child.isText ? child.text!.length : 0
    if (left < length) return place + left
    left -= length
    place += child.nodeSize
  }
  // behind the last character: the end of the paragraph
  return left === 0 ? place : null
}

function bookmarkSite(doc: PmNode, name: string, byIndex: Map<number, Block>): BookmarkSite | null {
  const start = new RegExp(`<w:bookmarkStart\\b[^>]*w:name="${escapeRegExp(name)}"`)
  let found: BookmarkSite | null = null
  doc.forEach((top, topPos) => {
    if (found) return
    const xml = baseXmlOf(top, byIndex).xml
    const at = xml.search(start)
    // nothing of the block's text stands before it in the file
    const leads = at >= 0 && !/<w:r[\s>]|<\/w:p>/.test(xml.slice(0, at))
    const holds = (node: PmNode, pos: number): boolean => {
      const names = [
        ...((node.attrs?.bookmarks as string[] | null) ?? []),
        ...((node.attrs?.hiddenBookmarks as string[] | null) ?? []),
      ]
      if (names.includes(name)) {
        // a bookmark the file does not have yet is written at the start of its block
        const atStart = at < 0 || (node === top && leads)
        const inner = !atStart && node === top ? placeInParagraph(node, pos, xml, at) : null
        found = { pos, end: pos + node.nodeSize, atStart, ...(inner !== null ? { inner } : {}) }
        return true
      }
      // a caption keeps its `_Ref` bookmark in its own XML
      if (node === top && node.type.name === 'docProtected' && at >= 0) {
        found = { pos, end: pos + node.nodeSize, atStart: leads }
        return true
      }
      return false
    }
    if (holds(top, topPos)) return
    top.descendants((node, pos) => !found && !holds(node, topPos + 1 + pos))
  })
  return found
}

function bookmarkPage(site: BookmarkSite, pages: PageSource): PageFacts | null {
  const first = pages.pageAt(site.pos, 'top')
  if (!first || site.atStart) return first
  const opens = site.inner === undefined ? null : pages.pageAt(site.inner)
  if (opens) return opens
  // where in the block it opens is not known: only a block on one page tells
  const last = pages.pageAt(site.pos, 'bottom')
  return last && last.index === first.index ? first : null
}

/** where a page field stands: in the text at `from`, or in a text box, on the page `page` gives */
interface PageFieldPlace {
  instr: string
  from?: number
  page: () => PageFacts | null
}

function pageFieldText(
  doc: PmNode,
  range: PageFieldPlace,
  pages: PageSource,
  byIndex: Map<number, Block>,
): string | null {
  const keyword = PAGE_FIELD_INSTR.exec(range.instr)?.[1]?.toUpperCase()
  if (keyword === 'NUMPAGES') return countText(range.instr, pages.pageCount)
  if (keyword === 'PAGE' || keyword === 'SECTIONPAGES') {
    const here = range.page()
    if (!here) return null
    return keyword === 'PAGE'
      ? pageText(range.instr, here)
      : countText(range.instr, here.sectionPages)
  }
  const name = /^\s*PAGEREF\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(range.instr)
  const site = name ? bookmarkSite(doc, name[1] ?? name[2], byIndex) : null
  const relative = /\\p\b/i.test(range.instr)
  if (name && !site) {
    // above or below a footnote is not a question of two places in the text
    const noted = relative ? null : pages.noteBookmarkPage?.(name[1] ?? name[2])
    return noted ? pageText(range.instr, noted) : null
  }
  const target = site ? bookmarkPage(site, pages) : null
  if (!site || !target) return null
  if (!relative) return pageText(range.instr, target)
  // the words say where the bookmark is from the field: a field in a text box has no such place
  if (range.from === undefined) return null
  const here = range.page()
  // a number inside the words that the field shapes itself is Word's to compute
  if (!here || !pages.words || unreadableFormat(range.instr) || shapesNumber(range.instr))
    return null
  if (here.index !== target.index)
    return formatFieldText(range.instr, pages.words.onPage(target.display))
  const inside = range.from > site.pos && range.from < site.end
  // the field stands in the bookmark's own block, and where in it the bookmark opens is not known
  if (inside && !site.atStart && site.inner === undefined) return null
  const opens = site.inner ?? site.pos
  return formatFieldText(range.instr, opens < range.from ? pages.words.above : pages.words.below)
}

/**
 * Gives PAGE, NUMPAGES, SECTIONPAGES and PAGEREF fields the page the layout shows. `tr` must
 * still hold the document `pages` was laid out from. A document as it was opened or saved is
 * left alone unless `all` asks (Update Field): our pages are not Word's to the line, and what
 * the file holds stands until something changed. Where the layout cannot tell, the field stays
 * as it is, mark included.
 */
export function refreshPageFields(
  tr: Transform,
  pages: PageSource,
  blocks: Block[] = [],
  all = false,
): number {
  const doc = tr.doc
  const markType = doc.type.schema.marks.instrField
  if (!markType || (!all && !pages.edited)) return 0
  const byIndex = blockIndex(blocks)
  const jobs: Array<{ range: FieldRange; text: string }> = []
  for (const range of collectFields(doc, markType)) {
    if (!PAGE_FIELD_INSTR.test(range.instr) || range.mark.attrs.beginXml) continue
    if (isFieldLocked(range.mark)) continue
    const waiting = range.mark.attrs.dirty === true
    const place = { instr: range.instr, from: range.from, page: () => pages.pageAt(range.from) }
    const text = pageFieldText(doc, place, pages, byIndex)
    if (text === null || text === '') continue
    const differs = text !== doc.textBetween(range.from, range.to, '').trim()
    if (differs || waiting) jobs.push({ range, text })
  }
  let changed = settleTextboxPageFields(tr, pages, byIndex)
  for (const job of jobs.reverse()) {
    if (job.text !== doc.textBetween(job.range.from, job.range.to, '').trim()) changed++
    writeField(tr, job.range, job.text)
  }
  return changed
}

type BoxRun = TextboxDisplay['paras'][number]['runs'][number]

/** a page field of a text box that still waits for its page */
const waitsInBox = (run: BoxRun, runs: readonly BoxRun[], at: number): boolean =>
  run.fldDirty === true &&
  !run.fldLock &&
  !run.fldBeginXml &&
  run.instrField !== undefined &&
  PAGE_FIELD_INSTR.test(run.instrField) &&
  // a result in several runs is not one this rewrites
  runs[at - 1]?.instrField !== run.instrField &&
  runs[at + 1]?.instrField !== run.instrField

function eachTextboxBlock(doc: PmNode, visit: (node: PmNode, pos: number) => void): void {
  doc.descendants((node, pos) => {
    if (node.type.name !== 'docProtected') return true
    if (Array.isArray(node.attrs.textboxes) && node.attrs.textboxes.length > 0) visit(node, pos)
    return false
  })
}

/**
 * The page fields put into text boxes this session take the page of the block that holds
 * their box, when that block sits on one page. Sizes stay: only attributes change.
 */
function settleTextboxPageFields(
  tr: Transform,
  pages: PageSource,
  byIndex: Map<number, Block>,
): number {
  const doc = tr.doc
  let changed = 0
  eachTextboxBlock(doc, (node, pos) => {
    let page: PageFacts | null | undefined
    const pageOfBlock = (): PageFacts | null => {
      if (page === undefined) {
        const first = pages.pageAt(pos, 'top')
        const last = first && pages.pageAt(pos, 'bottom')
        page = first && last && first.index === last.index ? first : null
      }
      return page
    }
    let touched = false
    const boxes = (node.attrs.textboxes as TextboxDisplay[]).map((box) => ({
      ...box,
      paras: box.paras.map((para) => ({
        ...para,
        runs: para.runs.map((run, at) => {
          if (!waitsInBox(run, para.runs, at)) return run
          const text = pageFieldText(
            doc,
            { instr: run.instrField!, page: pageOfBlock },
            pages,
            byIndex,
          )
          if (text === null || text === '') return run
          const { fldDirty: _settled, ...rest } = run
          touched = true
          if (text !== run.text.trim()) changed++
          return { ...rest, text }
        }),
      })),
    }))
    if (touched) tr.setNodeMarkup(pos, undefined, { ...node.attrs, textboxes: boxes })
  })
  return changed
}

/** a page field still marked for an update: its page is not in the document yet */
export function hasWaitingPageField(doc: PmNode): boolean {
  const markType = doc.type.schema.marks.instrField
  if (!markType) return false
  let boxed = false
  eachTextboxBlock(doc, (node) => {
    boxed ||= (node.attrs.textboxes as TextboxDisplay[]).some((box) =>
      box.paras.some((para) => para.runs.some((run, at) => waitsInBox(run, para.runs, at))),
    )
  })
  return (
    boxed ||
    collectFields(doc, markType).some(
      (range) =>
        range.mark.attrs.dirty === true &&
        !isFieldLocked(range.mark) &&
        !range.mark.attrs.beginXml &&
        PAGE_FIELD_INSTR.test(range.instr),
    )
  )
}

/**
 * Rewrites field runs located before `tr` took its steps. Each end is mapped inward, so a
 * field rewritten right beside one of them is not taken with it.
 */
export function replaceFieldRuns(
  tr: Transform,
  jobs: Array<{ from: number; to: number; text: string; marks: readonly Mark[] }>,
): void {
  const { schema } = tr.doc.type
  for (const job of [...jobs].sort((a, b) => b.from - a.from))
    tr.replaceWith(
      tr.mapping.map(job.from, 1),
      tr.mapping.map(job.to, -1),
      schema.text(job.text, [...job.marks]),
    )
}

/**
 * Before a save: nothing is left for Word to recompute; not a user edit, so out of the undo
 * history. `pages` is the layout standing for the document as it is, when there is one.
 */
export function settleFieldResults(
  editor: Editor,
  blocks: Block[],
  pages?: PageSource | null,
): void {
  const headings = numberingContextOf(editor)
  const skip = unreliableSeqLabelsOf(blocks, editor.schema, headings)
  const tr = editor.state.tr
  // first: the pages were laid out from the document as the view has it
  if (pages) refreshPageFields(tr, pages, blocks)
  renumberSeqFields(tr, blocks, { skip, headings })
  if (tr.docChanged) editor.view.dispatch(tr.setMeta('addToHistory', false))
}

/**
 * Gives every caption of the label the numbering Word's Caption Numbering dialog set. The
 * numbers they show stand until the next renumbering. A caption from the file becomes
 * the editor's own, as one a cross-reference bookmarked does.
 */
export function applyCaptionNumbering(
  tr: Transform,
  blocks: Block[],
  label: string,
  numbering: CaptionNumbering,
): number {
  const byIndex = blockIndex(blocks)
  let changed = 0
  tr.doc.forEach((node, pos) => {
    if (node.type.name !== 'docProtected') return
    const { xml } = protectedXmlOf(node, byIndex)
    if (!NUMBERING_FIELD.test(xml) || !captionParagraph(xml, node)) return
    const next = setCaptionNumberingXml(xml, label, numbering)
    if (!next) return
    const from = textOfXml(xml)
    const to = textOfXml(next)
    let head = 0
    while (head < from.length && head < to.length && from[head] === to[head]) head++
    let tail = 0
    while (
      tail < from.length - head &&
      tail < to.length - head &&
      from[from.length - 1 - tail] === to[to.length - 1 - tail]
    )
      tail++
    const shown = node.attrs.fieldDisplay as FieldDisplay | null
    const edit = {
      at: head,
      from: from.slice(head, from.length - tail),
      to: to.slice(head, to.length - tail),
    }
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      docxIndex: null,
      genXml: next,
      seqRenumber: null,
      fieldDisplay: renumberedDisplay(shown, [edit]) ?? shown,
    })
    changed++
  })
  return changed
}

/** after a caption went in, in the same undo step; `numbering` is what the dialog set for its label */
export function renumberAfterInsert(
  editor: Editor,
  blocks: Block[],
  numbering?: { label: string } & CaptionNumbering,
): void {
  const headings = numberingContextOf(editor)
  const skip = unreliableSeqLabelsOf(blocks, editor.schema, headings)
  const tr = editor.state.tr
  if (numbering) applyCaptionNumbering(tr, blocks, numbering.label, numbering)
  renumberSeqFields(tr, blocks, { skip, headings })
  if (tr.docChanged) editor.view.dispatch(tr.setMeta('appendedTransaction', tr))
}

/** the text a caption block shows once it stands at the selection and the document is renumbered */
export function captionPreview(
  editor: Editor,
  blocks: Block[],
  attrs: Record<string, unknown>,
  numbering?: { label: string } & CaptionNumbering,
): string {
  const xml = String(attrs.genXml)
  const type = editor.schema.nodes.docProtected
  const headings = numberingContextOf(editor)
  try {
    const { $to } = editor.state.selection
    const at = $to.depth > 0 ? $to.after(1) : $to.pos
    const tr = new Transform(editor.state.doc).insert(at, type.create(attrs))
    if (numbering) applyCaptionNumbering(tr, blocks, numbering.label, numbering)
    renumberSeqFields(tr, blocks, {
      skip: unreliableSeqLabelsOf(blocks, editor.schema, headings),
      headings,
    })
    return textOfXml(String(tr.doc.nodeAt(at)?.attrs.genXml ?? xml))
  } catch {
    return textOfXml(xml)
  }
}
