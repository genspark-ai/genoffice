import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import { NodeSelection } from '@tiptap/pm/state'
import {
  docxWithBody,
  fieldSwitchArg,
  parseDocx,
  parseFieldInstruction,
  parseSeqInstruction,
  seqFieldsOfXml,
  unlinkFieldsXml,
} from '@genoffice/docx-engine'
import type {
  Block,
  CitationStyleId,
  FieldTableLine,
  FieldTablePageRun,
  FieldTextRun,
  SourceInfo,
} from '@genoffice/docx-engine'
import {
  bibliographyFieldNodes,
  bibliographyRange,
  restyleCitations,
} from '../editor/source-citations'
import {
  collectPlanEntries,
  collectTocEntriesWithPages,
  editorStyles,
  fieldTableNodes,
  findFieldRange,
  findFieldRanges,
  findTocRange,
  headingPagesAt,
  planEntriesWithPages,
  setTableLocked,
  tableLocked,
  tableSpecAt,
  tocFieldNodes,
  TOC_INSTR,
  updateTocField,
  xmlOfNode,
  type TocRange,
} from '../editor/toc-update'
import {
  authoritiesOptionsOf,
  captionBody,
  DEFAULT_SPECS,
  figuresOptionsOf,
  indexHeadingPattern,
  indexOptionsOf,
  indexPageSep,
  tocPlan,
  type IndexTablesKind,
  type TableSpec,
} from '../editor/index-tables'
import {
  citationOf,
  citationRuns,
  citationSwitches,
  resolveCitations,
  shortCitationSwitches,
  type Citation,
} from '../editor/citations'
import { blocksToPmDoc, pmDocOptions } from '../editor/convert'
import { collectBookmarks } from '../editor/cross-reference'
import { textOfXml } from '../editor/field-results'
import { indexEntryOptions, insertFieldMarks } from '../editor/index-entry'
import { t } from '../i18n/locale'

/**
 * Puts generated blocks at the caret. With a block selected (a picture, a
 * caption just inserted) they go below it: inserting over a node selection
 * would replace the block.
 */
function insertBlocks(editor: Editor, nodes: Array<Record<string, unknown>>): void {
  const { selection } = editor.state
  if (selection instanceof NodeSelection) {
    // below the top-level block that holds the selection
    const below = selection.$to.depth > 0 ? selection.$to.after(1) : selection.to
    editor
      .chain()
      .focus()
      .insertContentAt(below, nodes as never)
      .run()
    return
  }
  editor
    .chain()
    .focus()
    .insertContent(nodes as never)
    .run()
}

/** the page each top-level block starts on, or with `bottom` the page it ends on */
type PagesOf = (positions: number[], edge?: 'top' | 'bottom') => number[] | null

const BOOKMARK_TAG = /<w:bookmark(Start|End)\b([^>]*)>/g
const xmlAttr = (attrs: string, name: string) =>
  new RegExp(`\\bw:${name}="([^"]*)"`).exec(attrs)?.[1]

/**
 * The top-level block each bookmark of the file closes in, by name. A
 * bookmark made here sits on one paragraph and is not listed.
 */
function bookmarkEnds(editor: Editor, blocks: Block[]): Map<string, number> {
  const open = new Map<string, { name: string; pos: number }>()
  const ends = new Map<string, number>()
  let before: number | null = null
  editor.state.doc.forEach((node, pos) => {
    const xml = xmlOfNode(node as never, blocks)
    for (const m of xml.includes('<w:bookmark') ? xml.matchAll(BOOKMARK_TAG) : []) {
      const id = xmlAttr(m[2], 'id')
      if (id === undefined) continue
      if (m[1] === 'Start') open.set(id, { name: xmlAttr(m[2], 'name') ?? '', pos })
      else if (open.has(id)) {
        const start = open.get(id)!
        // closed ahead of a later block's text: the bookmark took the paragraph mark before it
        const leads = start.pos !== pos && !/<w:t[\s>]/.test(xml.slice(0, m.index))
        ends.set(start.name, leads && before !== null ? before : pos)
      }
    }
    before = pos
  })
  return ends
}

/** puts a table in at the caret, or in place of the one it replaces */
function placeTable(
  editor: Editor,
  nodes: Array<Record<string, unknown>>,
  replace?: TocRange | null,
): void {
  if (!replace) return insertBlocks(editor, nodes)
  editor
    .chain()
    .focus()
    .deleteRange({ from: replace.from, to: replace.to })
    .insertContentAt(replace.from, nodes as never)
    .run()
}

/**
 * A table of contents built from the document's headings, with their page
 * numbers when known. With a spec the table follows its switches, and one
 * the selection is in is rebuilt in place.
 */
export function insertToc(
  editor: Editor,
  blocks: Block[],
  headingPages: () => number[] | null,
  spec?: TableSpec,
  pagesOf?: PagesOf,
): void {
  const replace = spec ? tableAtSelection(editor, blocks, 'toc')?.range : null
  if (replace) {
    updateTocField(editor, blocks, headingPages, spec, replace, pagesOf)
    return
  }
  const entries = spec
    ? planEntriesWithPages(
        editor,
        tocPlan(spec.instr),
        pagesOf ?? headingPagesAt(editor, headingPages),
      )
    : collectTocEntriesWithPages(editor, headingPages)
  const nodes = entries.length > 0 ? tocFieldNodes(entries, spec, !!spec, editorStyles(editor)) : []
  if (nodes.length === 0) {
    window.alert(t('ribbonTocNoHeadings'))
    return
  }
  insertBlocks(editor, nodes)
}

/**
 * The bibliography of the document's sources, under its heading. A document
 * that has one already gets that one rewritten instead of a second.
 */
export function insertBibliography(
  editor: Editor,
  blocks: Block[],
  sources: SourceInfo[],
  style: CitationStyleId,
): void {
  if (sources.length === 0) {
    window.alert(t('ribbonNoSources'))
    return
  }
  if (bibliographyRange(editor.state.doc, blocks)) {
    restyleCitations(editor, blocks, sources, style)
    return
  }
  insertBlocks(editor, [
    {
      type: 'docHeading',
      attrs: { docxIndex: null, styleId: null, aiChanged: false, level: 1 },
      content: [{ type: 'text', text: t('ribbonBibliographyHeading') }],
    },
    ...bibliographyFieldNodes(editor, sources, style),
  ])
}

const INDEX_INSTR = /<w:instrText[^>]*>\s*INDEX[\s\\]/

/** the index of every marked entry, sorted; null with none marked */
function indexFieldNodes(
  editor: Editor,
  blocks: Block[],
  pagesOf: PagesOf,
  spec: TableSpec,
): Array<Record<string, unknown>> | null {
  const { doc } = editor.state
  const topOf = (pos: number) => {
    const $pos = doc.resolve(pos)
    return $pos.depth === 0 ? pos : $pos.before(1)
  }
  const bookmarks = new Map(collectBookmarks(doc).map((b) => [b.name, topOf(b.pos)]))
  const ends = bookmarkEnds(editor, blocks)
  const marks: Array<IndexMark & { pos: number; end?: number }> = []
  doc.forEach((block, pos) => {
    block.descendants((node) => {
      if (node.type.name !== 'docXeMark' || typeof node.attrs.ta === 'string' || node.attrs.tc)
        return
      const term = String(node.attrs.term).trim()
      if (!term) return
      const { bookmark, ...options } = indexEntryOptions(node.attrs.switches as string | null)
      // a page range entry stands where its bookmark is, not where it was marked
      const at = bookmark ? bookmarks.get(bookmark) : undefined
      if (at === undefined) marks.push({ term, pos, ...options })
      else marks.push({ term, pos: at, end: Math.max(ends.get(bookmark!) ?? at, at), ...options })
    })
  })
  if (marks.length === 0) {
    window.alert(t('ribbonNoIndexEntries'))
    return null
  }
  const pages = pagesOf(marks.map((m) => m.pos))
  // a range runs to the page its bookmark's last paragraph ends on
  const ranged = marks.filter((m) => m.end !== undefined)
  const lastPages =
    ranged.length > 0
      ? pagesOf(
          ranged.map((m) => m.end!),
          'bottom',
        )
      : null
  const pageTo = new Map(ranged.map((m, i) => [m, lastPages?.[i]]))
  const options = indexOptionsOf(spec)
  const lines = indexLines(
    marks.map((mark, i) => {
      const { pos: _pos, end: _end, ...m } = mark
      if (!pages || pages.length !== marks.length) return m
      const to = pageTo.get(mark)
      return { ...m, page: pages[i], ...(to !== undefined && to > pages[i] ? { pageTo: to } : {}) }
    }),
    options.runIn,
    options.rightAlign ? undefined : indexPageSep(spec.instr),
    { pattern: indexHeadingPattern(spec.instr), accented: options.accented },
  )
  return fieldTableNodes(
    lines,
    {
      ...spec,
      ...(options.rightAlign ? {} : { pageSep: indexPageSep(spec.instr) }),
      noLink: true,
      tabPos: 4300,
    },
    t('ribbonIndexFieldLabel'),
    editorStyles(editor),
  )
}

/** one marked entry as the index reads it */
export interface IndexMark {
  term: string
  page?: number
  /** the last page of a page range entry (XE \\r), when it ends past `page` */
  pageTo?: number
  /** the text the index shows in place of this mark's page (XE \t) */
  cross?: string
  bold?: boolean
  italic?: boolean
}

interface IndexRefs {
  pages: Array<{ page: number; to: number; bold: boolean; italic: boolean }>
  crosses: string[]
}

const noRefs = (): IndexRefs => ({ pages: [], crosses: [] })

function addRef(refs: IndexRefs, mark: IndexMark): void {
  if (mark.cross) {
    if (!refs.crosses.includes(mark.cross)) refs.crosses.push(mark.cross)
    return
  }
  if (mark.page === undefined) return
  const to = mark.pageTo ?? mark.page
  const seen = refs.pages.find((p) => p.page === mark.page && p.to === to)
  if (seen) {
    seen.bold ||= !!mark.bold
    seen.italic ||= !!mark.italic
  } else refs.pages.push({ page: mark.page, to, bold: !!mark.bold, italic: !!mark.italic })
}

/** what follows an entry: its pages in order, then its cross-references */
function refRuns(refs: IndexRefs): FieldTablePageRun[] {
  const out: FieldTablePageRun[] = []
  ;[...refs.pages]
    .sort((a, b) => a.page - b.page || a.to - b.to)
    .forEach((p, i) => {
      if (i > 0) out.push({ text: ', ' })
      out.push({
        // a range reads first page, en dash, last page
        text: p.to > p.page ? `${p.page}\u2013${p.to}` : String(p.page),
        ...(p.bold ? { bold: true } : {}),
        ...(p.italic ? { italic: true } : {}),
      })
    })
  if (refs.crosses.length > 0)
    out.push({ text: `${out.length > 0 ? '. ' : ''}${refs.crosses.join('; ')}` })
  return out
}

const runsText = (runs: FieldTablePageRun[]) => runs.map((r) => r.text).join('')

const baseLetter = (letter: string) => [...letter.normalize('NFD')][0] ?? ''

/** letters in heading order: by base letter, a plain letter ahead of its accented forms */
const byLetter = (a: string, b: string): number =>
  baseLetter(a).localeCompare(baseLetter(b)) || (a < b ? -1 : a > b ? 1 : 0)

const HEADING_LETTER = /^[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}]$/u

/**
 * The letter an entry is filed under: its first letter in capitals, without
 * its accent unless accented letters get headings of their own. '' for an
 * entry that starts with a digit, a symbol or a script without letter headings.
 */
export function indexLetter(entry: string, accented: boolean): string {
  const first = [...entry.trim()][0] ?? ''
  const upper = first.toLocaleUpperCase()
  const base = accented ? upper : baseLetter(upper)
  return HEADING_LETTER.test(base) ? base : ''
}

/** the page side of a line: plain when nothing in it is bold or italic */
function pageSide(refs: IndexRefs): Pick<FieldTableLine, 'pageNo' | 'pageRuns'> {
  const runs = refRuns(refs)
  return {
    pageNo: runsText(runs),
    ...(runs.some((r) => r.bold || r.italic) ? { pageRuns: runs } : {}),
  }
}

/**
 * The lines of an index. "Main:Sub" marks a subentry: indented, it takes a
 * line of its own under the main entry; run-in, the subentries follow the
 * main entry on its line. An entry marked with a cross-reference shows that
 * text where its page would be.
 */
export function indexLines(
  marks: IndexMark[],
  runIn: boolean,
  /** what joins an entry to its pages when they are part of the text (run-in lines) */
  pageSep = ', ',
  /** INDEX \h: a line above each letter's entries, "A" in the pattern standing for the letter */
  headings: { pattern?: string; accented?: boolean } = {},
): Array<FieldTableLine & { level: number }> {
  const byText = (a: string, b: string) => a.localeCompare(b, 'zh-CN')
  const letterOf = (name: string) => indexLetter(name, !!headings.accented)
  // under headings the entries of a letter stay together: an accented letter is filed with
  // its base letter, or (\a) forms a group of its own behind it
  const byName =
    headings.pattern === undefined
      ? byText
      : (a: string, b: string) => byLetter(letterOf(a), letterOf(b)) || byText(a, b)
  const mains = new Map<string, { refs: IndexRefs; subs: Map<string, IndexRefs> }>()
  for (const mark of marks) {
    const cut = mark.term.indexOf(':')
    const main = (cut === -1 ? mark.term : mark.term.slice(0, cut)).trim()
    const sub = cut === -1 ? '' : mark.term.slice(cut + 1).trim()
    if (!main) continue
    const entry = mains.get(main) ?? { refs: noRefs(), subs: new Map<string, IndexRefs>() }
    mains.set(main, entry)
    if (!sub) addRef(entry.refs, mark)
    else {
      const refs = entry.subs.get(sub) ?? noRefs()
      entry.subs.set(sub, refs)
      addRef(refs, mark)
    }
  }
  // an entry that only points elsewhere reads "Entry. See Other"
  const named = (name: string, refs: IndexRefs) =>
    refs.pages.length === 0 && refs.crosses.length > 0
      ? `${name}. ${refs.crosses.join('; ')}`
      : name
  const out: Array<FieldTableLine & { level: number }> = []
  let letter: string | null = null
  for (const main of [...mains.keys()].sort(byName)) {
    const { refs, subs } = mains.get(main)!
    const names = [...subs.keys()].sort(byText)
    if (headings.pattern !== undefined && letterOf(main) !== letter) {
      letter = letterOf(main)
      // the entries before the first letter (digits, symbols) stand without a heading
      if (letter)
        out.push({
          // only the A stands for the letter: the rest of the pattern is literal
          text: headings.pattern
            .replace(/A/g, letter)
            .replace(/a/g, letter.toLocaleLowerCase())
            .trim(),
          level: 1,
          style: 'IndexHeading',
          noPage: true,
        })
    }
    if (runIn && names.length > 0) {
      const withPages = (name: string, r: IndexRefs) =>
        r.pages.length > 0
          ? `${name}${pageSep.replace(/\t/g, ' ')}${runsText(refRuns(r))}`
          : named(name, r)
      out.push({
        text: `${withPages(main, refs)}: ${names.map((n) => withPages(n, subs.get(n)!)).join('; ')}`,
        noPage: true,
        level: 1,
        style: 'Index1',
      })
      continue
    }
    out.push({
      text: named(main, refs),
      level: 1,
      style: 'Index1',
      indent: 0,
      ...(refs.pages.length > 0
        ? pageSide(refs)
        : { noPage: names.length > 0 || refs.crosses.length > 0 }),
    })
    for (const name of names) {
      const r = subs.get(name)!
      out.push({
        text: named(name, r),
        level: 2,
        style: 'Index2',
        indent: 240,
        ...(r.pages.length > 0 ? pageSide(r) : r.crosses.length > 0 ? { noPage: true } : {}),
      })
    }
  }
  return out
}

export function insertIndex(
  editor: Editor,
  blocks: Block[],
  pagesOf: PagesOf,
  spec?: TableSpec,
  /** rebuild the document's index of that order in place, wherever the selection is; without it, the one the selection is in */
  nth?: number,
): boolean {
  const ranges = findFieldRanges(editor.state.doc, blocks, INDEX_INSTR)
  const replace =
    nth !== undefined ? ranges[nth] : spec ? tableAtSelection(editor, blocks, 'index')?.range : null
  // the index to rebuild is gone: nothing is added in its place
  if (nth !== undefined && !replace) return false
  const nodes = indexFieldNodes(editor, blocks, pagesOf, spec ?? DEFAULT_SPECS.index)
  if (nodes) placeTable(editor, nodes, replace)
  return nodes !== null
}

/** the column count an index inserted with that spec carries */
export const indexColumnsOf = (spec: TableSpec): number => indexOptionsOf(spec).columns

/** whether the document has an entry an index can list: without one, building an index fails */
export function hasIndexEntries(editor: Editor): boolean {
  let found = false
  editor.state.doc.descendants((node) => {
    if (found) return false
    found =
      node.type.name === 'docXeMark' &&
      typeof node.attrs.ta !== 'string' &&
      !node.attrs.tc &&
      String(node.attrs.term).trim() !== ''
    return !found
  })
  return found
}

export interface BlockSpan {
  index: number
  firstBlock: number
  lastBlock: number
}

/**
 * Whether the index at top-level nodes [first, last] has its section to itself
 * at each end. A section but the last may end in a paragraph of its own that
 * only carries the break (`isBreakOnly`); any other block after the index is text
 * that shares the section.
 */
export function indexSectionEdges(
  [first, last]: [number, number],
  sections: BlockSpan[],
  isBreakOnly: (block: number) => boolean,
): { home: BlockSpan; startsIt: boolean; endsIt: boolean } | null {
  const at = (block: number) => sections.find((s) => block >= s.firstBlock && block <= s.lastBlock)
  const home = at(first)
  if (!home) return null
  const closed = home.index < sections.length - 1 && home.lastBlock === last + 1
  return {
    home,
    startsIt: first === 0 || at(first - 1)?.index !== home.index,
    endsIt: home.lastBlock <= last || (closed && isBreakOnly(last + 1)),
  }
}

/** every index of the document as its top-level nodes [first, last], in document order */
export function indexBlockRanges(editor: Editor, blocks: Block[]): Array<[number, number]> {
  const { doc } = editor.state
  return findFieldRanges(doc, blocks, INDEX_INSTR).map((range) => [
    doc.resolve(range.from).index(0),
    doc.resolve(range.to).index(0) - 1,
  ])
}

/** how many indexes start before that position: where a new one inserted there stands among them */
export const indexesBefore = (editor: Editor, blocks: Block[], pos: number): number =>
  findFieldRanges(editor.state.doc, blocks, INDEX_INSTR).filter((range) => range.from < pos).length

/** which of the document's indexes the selection is in; -1 when in none */
export const indexAtSelection = (editor: Editor, blocks: Block[]): number => {
  const { from } = editor.state.selection
  return findFieldRanges(editor.state.doc, blocks, INDEX_INSTR).findIndex(
    (range) => from >= range.from && from < range.to,
  )
}

/** rebuilds the document's index (or the one at `at`) in place from the entries marked now */
export function updateIndex(
  editor: Editor,
  blocks: Block[],
  pagesOf: PagesOf,
  at?: TocRange,
): boolean {
  const range = at ?? findFieldRange(editor.state.doc, blocks, INDEX_INSTR)
  if (!range) {
    window.alert(t('ribbonIndexNotFound'))
    return false
  }
  if (tableLocked(editor.state.doc, blocks, range)) return false
  const spec = tableSpecAt(editor.state.doc, blocks, range.from, 'INDEX') ?? DEFAULT_SPECS.index
  const nodes = indexFieldNodes(editor, blocks, pagesOf, spec)
  if (!nodes) return false
  placeTable(editor, nodes, range)
  return true
}

const FIGURES_INSTR = /<w:instrText[^>]*>\s*TOC\b[^<]*\\[ac][\s"]/

/** a caption paragraph: the label of its first SEQ field, written as a simple or a complex field */
function captionLabelOf(xml: string): string | null {
  if (!/^<w:p[\s>]/.test(xml)) return null
  for (const field of seqFieldsOfXml(xml)) {
    const seq = parseSeqInstruction(field.instr)
    if (seq) return seq.id
  }
  return null
}

export interface CaptionEntry {
  label: string
  text: string
  /** position of the caption's top-level node */
  pos: number
}

/** every caption of the document, in order: the paragraphs that number themselves with a SEQ field */
export function collectCaptions(editor: Editor, blocks: Block[]): CaptionEntry[] {
  const out: CaptionEntry[] = []
  editor.state.doc.forEach((node, pos) => {
    const xml = xmlOfNode(node as never, blocks)
    // a table's XML holds its cells' captions: only a paragraph is one itself
    const label = captionLabelOf(xml)
    if (label === null) return
    // text typed since the file was read is in the node; a generated caption only has its XML
    const text = node.textContent.trim() || textOfXml(xml).trim()
    if (text) out.push({ label, text, pos })
  })
  return out
}

/** the caption labels the document uses, in the order they first appear */
export const captionLabels = (editor: Editor, blocks: Block[]): string[] => [
  ...new Set(collectCaptions(editor, blocks).map((c) => c.label)),
]

function figuresFieldNodes(
  editor: Editor,
  blocks: Block[],
  spec: TableSpec,
  pagesOf: PagesOf,
): Array<Record<string, unknown>> | null {
  const options = figuresOptionsOf(spec)
  const plan = tocPlan(spec.instr)
  // a style or TC fields named in Options stand in for the label's captions
  const listed = options.style || options.entryId
  const captions: CaptionEntry[] = listed
    ? collectPlanEntries(editor, {
        ...plan,
        headings: false,
        outline: false,
        entryFields: !!options.entryId,
        entryLevels: null,
      }).map((e) => ({ label: '', text: e.text, pos: e.pos }))
    : collectCaptions(editor, blocks).filter((c) => c.label === options.label)
  if (captions.length === 0) {
    window.alert(t('ribbonTofNoCaptions'))
    return null
  }
  const pages = pagesOf(captions.map((c) => c.pos))
  const lines = captions.map((c, i) => ({
    text: options.includeLabel || listed ? c.text : captionBody(c.text, c.label) || c.text,
    level: 1,
    style: 'TableofFigures',
    noPage: !options.pageNumbers,
    ...(pages && pages.length === captions.length ? { pageNo: pages[i] } : {}),
  }))
  return fieldTableNodes(
    lines,
    { ...spec, pageSep: plan.pageSep, noLink: !plan.hyperlinks },
    t('ribbonItFigures'),
    editorStyles(editor),
  )
}

const figuresSpecOf = (label: string): TableSpec => ({
  // the label is quoted inside the instruction: a quote of its own would end it early
  instr: ` TOC \\h \\z \\c "${label.replace(/"/g, '')}" `,
  leader: 'dot',
})

/**
 * A table of figures over the captions of one label; the first label in use
 * when none is named. With a spec the table follows its switches, and one
 * the selection is in is rebuilt in place.
 */
export function insertTableOfFigures(
  editor: Editor,
  blocks: Block[],
  pagesOf: PagesOf,
  label?: string,
  spec?: TableSpec,
): void {
  const options = spec ? figuresOptionsOf(spec) : null
  const wanted = options
    ? options.label || options.style || options.entryId
    : label || captionLabels(editor, blocks)[0]
  if (!wanted) {
    window.alert(t('ribbonTofNoCaptions'))
    return
  }
  const replace = spec ? tableAtSelection(editor, blocks, 'figures')?.range : null
  const nodes = figuresFieldNodes(editor, blocks, spec ?? figuresSpecOf(wanted), pagesOf)
  if (nodes) placeTable(editor, nodes, replace)
}

/** rebuilds the document's first table of figures (or the one at `at`) in place, for the label it was made with */
export function updateTableOfFigures(
  editor: Editor,
  blocks: Block[],
  pagesOf: PagesOf,
  at?: TocRange,
): boolean {
  const range = at ?? findFieldRange(editor.state.doc, blocks, FIGURES_INSTR)
  if (!range) {
    window.alert(t('ribbonTofNotFound'))
    return false
  }
  if (tableLocked(editor.state.doc, blocks, range)) return false
  const spec = tableSpecAt(editor.state.doc, blocks, range.from, 'TOC')
  if (!spec) return false
  const nodes = figuresFieldNodes(editor, blocks, spec, pagesOf)
  if (!nodes) return false
  placeTable(editor, nodes, range)
  return true
}

const AUTHORITIES_INSTR = /<w:instrText[^>]*>\s*TOA[\s\\]/

/**
 * Puts a citation marker at the position: the text before it is what it
 * cites. With `mentions`, each of those positions takes the short citation
 * alone, as Word's Mark All leaves the later mentions of an authority.
 */
export function markCitation(
  editor: Editor,
  at: number,
  citation: Citation,
  mentions: number[] = [],
): void {
  insertFieldMarks(editor, [
    {
      at,
      attrs: {
        term: '',
        ta: citationSwitches(citation),
        ...(citation.runs ? { taRuns: JSON.stringify(citation.runs) } : {}),
      },
    },
    ...mentions.map((pos) => ({
      at: pos,
      attrs: { term: '', ta: shortCitationSwitches(citation.short) },
    })),
  ])
}

/** the citations marked in the document, each with the top-level block it sits in */
export function collectCitations(editor: Editor): Array<Citation & { pos: number }> {
  const out: Array<Citation & { pos: number }> = []
  editor.state.doc.forEach((block, pos) => {
    block.descendants((node) => {
      if (node.type.name !== 'docXeMark' || typeof node.attrs.ta !== 'string') return
      const runs = citationRuns(node.attrs.taRuns)
      out.push({ ...citationOf(node.attrs.ta), ...(runs ? { runs } : {}), pos })
    })
  })
  return resolveCitations(out)
}

/** the categories that have a citation, lowest first */
export const citationCategories = (editor: Editor): number[] =>
  [...new Set(collectCitations(editor).map((c) => c.category))].sort((a, b) => a - b)

/** Word lists an authority cited on this many pages or more as "passim" (TOA \p) */
const PASSIM_PAGES = 5

/**
 * The number of the sequence `id` in effect at each position: the result of
 * the last SEQ field of that identifier in or before the block ('' ahead of
 * the first one). What TOA \\s puts before a page number.
 */
function sequenceNumbersAt(
  editor: Editor,
  blocks: Block[],
  id: string,
  positions: number[],
): string[] {
  const steps: Array<{ pos: number; value: string }> = []
  let count = 0
  editor.state.doc.forEach((node, pos) => {
    for (const field of seqFieldsOfXml(xmlOfNode(node as never, blocks))) {
      if (parseSeqInstruction(field.instr)?.id.toLowerCase() !== id.toLowerCase()) continue
      count += 1
      steps.push({ pos, value: field.result.trim() || String(count) })
    }
  })
  return positions.map((at) => steps.filter((step) => step.pos <= at).at(-1)?.value ?? '')
}

function authoritiesFieldNodes(
  editor: Editor,
  blocks: Block[],
  spec: TableSpec,
  headingOf: (category: number) => string,
  pagesOf: PagesOf,
): Array<Record<string, unknown>> | null {
  const options = authoritiesOptionsOf(spec)
  const all = collectCitations(editor)
  const instr = parseFieldInstruction(spec.instr)
  // \\s: the number of that sequence stands before each page, joined by \\d (a hyphen without it)
  const sequence = fieldSwitchArg(instr, 's')
  const sequenceSep = fieldSwitchArg(instr, 'd') || '-'
  // \\l: what stands between the pages of one authority
  const pageListSep = fieldSwitchArg(instr, 'l') || ', '
  const categories = options.category
    ? [options.category]
    : [...new Set(all.map((c) => c.category))].sort((a, b) => a - b)
  const lines: Array<FieldTableLine & { level: number }> = []
  for (const category of categories) {
    const cited = all.filter((c) => c.category === category)
    if (cited.length === 0) continue
    const pages = pagesOf(cited.map((c) => c.pos))
    const numbers = sequence
      ? sequenceNumbersAt(
          editor,
          blocks,
          sequence,
          cited.map((c) => c.pos),
        )
      : null
    // pages under a sequence number stay in document order: the numbers group them
    const numbered = !!numbers?.some(Boolean)
    // a citation made in several places is one line, with every page it is on
    const byText = new Map<string, Array<{ page: number; label: string }>>()
    const formatted = new Map<string, FieldTextRun[]>()
    cited.forEach((c, i) => {
      const list = byText.get(c.long) ?? []
      if (pages && pages.length === cited.length) {
        const label = numbers?.[i] ? `${numbers[i]}${sequenceSep}${pages[i]}` : String(pages[i])
        if (!list.some((on) => on.label === label)) list.push({ page: pages[i], label })
      }
      byText.set(c.long, list)
      if (options.keepFormatting && c.runs && !formatted.has(c.long)) formatted.set(c.long, c.runs)
    })
    lines.push({ text: headingOf(category), level: 1, style: 'TOAHeading' })
    for (const [text, on] of [...byText.entries()].sort(([a], [b]) => a.localeCompare(b))) {
      const pageNo =
        options.passim && on.length >= PASSIM_PAGES
          ? 'passim'
          : (numbered ? on : on.sort((a, b) => a.page - b.page))
              .map((entry) => entry.label)
              .join(pageListSep)
      lines.push({
        text,
        level: 2,
        style: 'TableofAuthorities',
        ...(formatted.has(text) ? { textRuns: formatted.get(text) } : {}),
        ...(pageNo ? { pageNo } : {}),
      })
    }
  }
  if (lines.length === 0) {
    window.alert(t('ribbonToaNone'))
    return null
  }
  return fieldTableNodes(
    lines,
    { ...spec, noLink: true },
    t('ribbonItAuthorities'),
    editorStyles(editor),
  )
}

const authoritiesSpecOf = (category?: number): TableSpec => ({
  instr: category ? ` TOA \\h \\c "${Math.round(category)}" \\p ` : ' TOA \\h \\p ',
  leader: 'dot',
})

/**
 * A table of authorities: every category that has a citation, each under its
 * heading, or the one category named. With a spec the table follows its
 * switches, and one the selection is in is rebuilt in place.
 */
export function insertTableOfAuthorities(
  editor: Editor,
  blocks: Block[],
  pagesOf: PagesOf,
  headingOf: (category: number) => string,
  category?: number,
  spec?: TableSpec,
): void {
  const replace = spec ? tableAtSelection(editor, blocks, 'authorities')?.range : null
  const nodes = authoritiesFieldNodes(
    editor,
    blocks,
    spec ?? authoritiesSpecOf(category),
    headingOf,
    pagesOf,
  )
  if (nodes) placeTable(editor, nodes, replace)
}

/** rebuilds the document's first table of authorities (or the one at `at`) in place, for the categories it lists */
export function updateTableOfAuthorities(
  editor: Editor,
  blocks: Block[],
  pagesOf: PagesOf,
  headingOf: (category: number) => string,
  at?: TocRange,
): boolean {
  const range = at ?? findFieldRange(editor.state.doc, blocks, AUTHORITIES_INSTR)
  if (!range) {
    window.alert(t('ribbonToaNotFound'))
    return false
  }
  if (tableLocked(editor.state.doc, blocks, range)) return false
  const spec = tableSpecAt(editor.state.doc, blocks, range.from, 'TOA')
  if (!spec) return false
  const nodes = authoritiesFieldNodes(editor, blocks, spec, headingOf, pagesOf)
  if (!nodes) return false
  placeTable(editor, nodes, range)
  return true
}

export interface FieldTables {
  figures: boolean
  index: boolean
  authorities: boolean
}

interface FieldTableCache {
  originals: Map<number, string>
  nodes: WeakMap<object, FieldTables>
  docs: WeakMap<object, FieldTables>
}
const fieldTableCaches = new WeakMap<Block[], FieldTableCache>()

/**
 * Which of the tables the three Update buttons rebuild the document has. A
 * node's answer is kept: an edit makes a new document out of the same nodes.
 */
export function fieldTablesOf(
  doc: Parameters<typeof findFieldRange>[0],
  blocks: Block[],
): FieldTables {
  let cache = fieldTableCaches.get(blocks)
  if (!cache) {
    const originals = new Map<number, string>()
    for (const block of blocks)
      if (block.docxIndex != null && block.originalXml)
        originals.set(block.docxIndex, block.originalXml)
    cache = { originals, nodes: new WeakMap(), docs: new WeakMap() }
    fieldTableCaches.set(blocks, cache)
  }
  const known = cache.docs.get(doc)
  if (known) return known
  const { originals, nodes } = cache
  const found: FieldTables = { figures: false, index: false, authorities: false }
  doc.forEach((node) => {
    let own = nodes.get(node)
    if (!own) {
      const xml = node.attrs.genXml
        ? String(node.attrs.genXml)
        : (originals.get(node.attrs.docxIndex as number) ?? '')
      own = {
        figures: FIGURES_INSTR.test(xml),
        index: INDEX_INSTR.test(xml),
        authorities: AUTHORITIES_INSTR.test(xml),
      }
      nodes.set(node, own)
    }
    found.figures ||= own.figures
    found.index ||= own.index
    found.authorities ||= own.authorities
  })
  cache.docs.set(doc, found)
  return found
}

const TABLE_INSTR: Record<IndexTablesKind, RegExp> = {
  toc: TOC_INSTR,
  figures: FIGURES_INSTR,
  index: INDEX_INSTR,
  authorities: AUTHORITIES_INSTR,
}

/** the table of that kind the selection is in, with the switches it carries */
export function tableAtSelection(
  editor: Editor,
  blocks: Block[],
  kind: IndexTablesKind,
): { range: TocRange; spec: TableSpec } | null {
  const { doc, selection } = editor.state
  const ranges = findFieldRanges(doc, blocks, TABLE_INSTR[kind])
  // Word's own table of contents is a content control around the field
  const sdt = kind === 'toc' ? findTocRange(doc, blocks) : null
  if (sdt?.kind === 'sdt') ranges.push(sdt)
  const range = ranges.find((r) => selection.from >= r.from && selection.from < r.to)
  if (!range) return null
  const keyword = kind === 'index' ? 'INDEX' : kind === 'authorities' ? 'TOA' : 'TOC'
  const spec = tableSpecAt(doc, blocks, range.from, keyword)
  return spec ? { range, spec } : null
}

export interface FieldTableSite {
  kind: IndexTablesKind
  range: TocRange
}

/** every table of contents, of figures and of authorities and every index, in document order */
export function fieldTableSites(
  doc: Parameters<typeof findFieldRange>[0],
  blocks: Block[],
): FieldTableSite[] {
  const sites: FieldTableSite[] = []
  for (const kind of Object.keys(TABLE_INSTR) as IndexTablesKind[])
    for (const range of findFieldRanges(doc, blocks, TABLE_INSTR[kind])) sites.push({ kind, range })
  // Word's own table of contents is a content control around the field
  const sdt = findTocRange(doc, blocks)
  if (sdt?.kind === 'sdt') sites.push({ kind: 'toc', range: sdt })
  return sites.sort((a, b) => a.range.from - b.range.from)
}

/** the tables the selection touches, a caret inside one included */
export function fieldTablesAtSelection(editor: Editor, blocks: Block[]): FieldTableSite[] {
  const { from, to } = editor.state.selection
  return fieldTableSites(editor.state.doc, blocks).filter(
    ({ range }) => range.from < Math.max(to, from + 1) && range.to > from,
  )
}

/** the document has a locked table of that kind where Update would rebuild one */
export function updateHitsLockedTable(
  editor: Editor,
  blocks: Block[],
  kind: IndexTablesKind,
): boolean {
  const { doc } = editor.state
  const first = fieldTableSites(doc, blocks).find((site) => site.kind === kind)
  return !!first && tableLocked(doc, blocks, first.range)
}

/** Word's Lock Fields / Unlock Fields on the tables the selection touches; returns how many changed */
export function setFieldTablesLocked(editor: Editor, blocks: Block[], locked: boolean): number {
  if (!editor.isEditable) return 0
  const tr = closeHistory(editor.state.tr)
  let done = 0
  for (const { range } of fieldTablesAtSelection(editor, blocks))
    if (setTableLocked(tr, blocks, range, locked)) done++
  if (done > 0) editor.view.dispatch(tr)
  return done
}

export interface FieldTableUpdateHost {
  blocks: Block[]
  pagesOf: PagesOf
  headingPages: () => number[] | null
  citationHeading: (category: number) => string
}

/**
 * Update Field on the tables the selection touches: each is rebuilt as its own
 * Update button rebuilds it; a locked one stays. The last first, so the others
 * keep their places.
 */
export function updateFieldTablesAtSelection(
  editor: Editor,
  host: FieldTableUpdateHost,
): { updated: number; locked: number } {
  const { blocks, pagesOf } = host
  const count = { updated: 0, locked: 0 }
  if (!editor.isEditable) return count
  for (const { kind, range } of fieldTablesAtSelection(editor, blocks).reverse()) {
    if (tableLocked(editor.state.doc, blocks, range)) {
      count.locked++
      continue
    }
    const done =
      kind === 'toc'
        ? updateTocField(editor, blocks, host.headingPages, undefined, range, pagesOf)
        : kind === 'index'
          ? updateIndex(editor, blocks, pagesOf, range)
          : kind === 'figures'
            ? updateTableOfFigures(editor, blocks, pagesOf, range)
            : updateTableOfAuthorities(editor, blocks, pagesOf, host.citationHeading, range)
    if (done) count.updated++
  }
  return count
}

type JsonNode = { type: string; attrs?: Record<string, unknown>; content?: JsonNode[] }

/** a node read from another parse of this file: it is new content here, written from what it holds */
const asNewContent = (node: JsonNode): JsonNode => ({
  ...node,
  ...(node.attrs && 'docxIndex' in node.attrs ? { attrs: { ...node.attrs, docxIndex: null } } : {}),
  ...(node.content ? { content: node.content.map(asNewContent) } : {}),
})

/**
 * Word's Unlink Fields on the tables the selection touches: the field goes and
 * its lines stay, read again from their XML. A line that reads as a paragraph
 * (an index entry) can be edited from then on; one that is drawn as a table
 * line whatever it holds (a paragraph in a TOC style) keeps its place as it is
 * drawn, without the field. `bytes` is the file the document was read from.
 */
export async function unlinkFieldTablesAtSelection(
  editor: Editor,
  blocks: Block[],
  bytes: Uint8Array | null | undefined,
): Promise<number> {
  if (!editor.isEditable) return 0
  const sites = fieldTablesAtSelection(editor, blocks)
  if (sites.length === 0) return 0
  const { doc } = editor.state
  type Job = { size: number; attrs: Record<string, unknown>; xml: string; fielded: boolean }
  const jobs = new Map<number, Job>()
  doc.forEach((node, pos) => {
    if (node.type.name !== 'docProtected') return
    if (!sites.some(({ range }) => pos >= range.from && pos < range.to)) return
    const xml = xmlOfNode(node as never, blocks)
    const plain = unlinkFieldsXml(xml)
    // a line between the field's two ends holds no field character and is a line of it all the same
    jobs.set(pos, { size: node.nodeSize, attrs: node.attrs, xml: plain, fielded: plain !== xml })
  })
  const list = [...jobs.entries()].sort((a, b) => a[0] - b[0])
  let fresh: JsonNode[] | null = null
  try {
    const body = list.map(([, job]) => job.xml).join('')
    // a paragraph that closes a section is not a line of the table to read again
    const pkg = bytes && !body.includes('<w:sectPr') ? await docxWithBody(bytes, body) : null
    const parsed = pkg ? await parseDocx(pkg) : null
    // the section properties that close the body are not a line
    const lines = parsed?.blocks.filter((block) => !block.hidden) ?? []
    const content = parsed
      ? ((blocksToPmDoc(lines, undefined, pmDocOptions(parsed)).content ?? []) as JsonNode[])
      : []
    if (content.length === list.length) fresh = content.map(asNewContent)
  } catch {
    fresh = null
  }
  // edited while the lines were read: their places no longer stand
  if (editor.isDestroyed || editor.state.doc !== doc) return 0
  const tr = closeHistory(editor.state.tr)
  for (let i = list.length - 1; i >= 0; i--) {
    const [pos, job] = list[i]
    const node = fresh?.[i]
    if (node && node.type !== 'docProtected')
      tr.replaceWith(pos, pos + job.size, editor.schema.nodeFromJSON(node))
    else if (job.fielded)
      tr.setNodeMarkup(pos, undefined, { ...job.attrs, docxIndex: null, genXml: job.xml })
  }
  if (!tr.docChanged) return 0
  editor.view.dispatch(tr)
  return sites.length
}

/** a table spec as a command payload */
export const tableSpecPayload = (spec: TableSpec): string => JSON.stringify(spec)

/** the table spec a command payload carries; undefined for any other payload */
export function tableSpecOfPayload(payload: string | undefined): TableSpec | undefined {
  if (!payload?.startsWith('{')) return undefined
  try {
    const spec = JSON.parse(payload) as Partial<TableSpec>
    return typeof spec.instr === 'string' && typeof spec.leader === 'string'
      ? { instr: spec.instr, leader: spec.leader }
      : undefined
  } catch {
    return undefined
  }
}

/** the command that inserts each kind of table; its payload is the table's spec */
export const INDEX_TABLES_COMMAND = {
  toc: 'insert-toc',
  figures: 'insert-table-of-figures',
  index: 'insert-index',
  authorities: 'toa-insert',
} as const satisfies Record<IndexTablesKind, string>

/** the spec of the table of each kind the selection is in */
export function indexTablesAtSelection(
  editor: Editor | null,
  blocks: Block[],
): Partial<Record<IndexTablesKind, TableSpec>> {
  const out: Partial<Record<IndexTablesKind, TableSpec>> = {}
  if (!editor) return out
  for (const kind of Object.keys(INDEX_TABLES_COMMAND) as IndexTablesKind[]) {
    const found = tableAtSelection(editor, blocks, kind)
    if (found) out[kind] = found.spec
  }
  return out
}
