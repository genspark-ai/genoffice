import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  fieldInstructionOfXml,
  generateFieldTableXml,
  generateTocFieldXml,
  rewriteSeqResultsXml,
  type Block,
  type FieldDisplay,
  type FieldTableLine,
  type StyleInfo,
  type TocEntry,
} from '@genoffice/docx-engine'
import { collectHeadings } from './headings'
import { documentStyleId } from './table-formats'
import { tableLeaderOf, tocEntryField, tocPlan, type TableSpec, type TocPlan } from './index-tables'
import { t } from '../i18n/locale'

/**
 * Table-of-contents detection + regeneration, shared by the
 * References ribbon tab and the TOC context menu.
 *
 * Two shapes exist in the wild:
 * - a TOC *field* (what our built-in "Table of Contents" inserts): one
 *   w:fldChar begin…end run across the entry paragraphs, opened by a
 *   `TOC \o "1-N" \h \z \u` instruction;
 * - Word's automatic TOC: a w:sdt gallery control ("Table of Contents")
 *   whose content is a TOCHeading paragraph followed by TOC1–TOC9 entry
 *   paragraphs (PAGEREF fields). It carries no TOC instruction, so the
 *   field scan below cannot see it — the sdt scan handles that shape.
 */

export interface TocRange {
  from: number
  to: number
  keepPageBreak: boolean
  kind: 'field' | 'sdt'
}

export function collectTocEntries(editor: Editor): TocEntry[] {
  return collectHeadings(editor.state.doc, editor.storage.listNumbering?.styles).map(
    ({ level, text }) => ({ level, text }),
  )
}

/** Heading entries + real page numbers (headingPages and collectTocEntries share document order) */
export function collectTocEntriesWithPages(
  editor: Editor,
  headingPages?: () => number[] | null,
): TocEntry[] {
  const entries = collectTocEntries(editor)
  const pages = headingPages?.()
  if (pages && pages.length === entries.length) {
    return entries.map((e, i) => ({ ...e, pageNo: pages[i] }))
  }
  return entries
}

export type StyleLookup = ReadonlyMap<string, Pick<StyleInfo, 'name' | 'display'>>

/** the document's styles as the editor holds them, pending definitions included */
export const editorStyles = (editor: Editor): StyleLookup | undefined =>
  editor.storage.listNumbering?.styles as StyleLookup | undefined

export interface PlannedEntry extends TocEntry {
  /** position of the top-level block the entry stands in */
  pos: number
}

/**
 * The entries a TOC instruction lists, in document order: the built-in
 * headings of its levels, paragraphs with an outline level of their own (\u),
 * paragraphs in the styles it names (\t) and TC fields (\f).
 */
export function collectPlanEntries(editor: Editor, plan: TocPlan): PlannedEntry[] {
  const { doc } = editor.state
  const styles = editorStyles(editor)
  const headings = new Map(
    collectHeadings(doc, editor.storage.listNumbering?.styles).map((h) => [h.pos, h]),
  )
  const levelOfStyle = new Map(plan.styles.map((s) => [s.name.toLowerCase(), s.level]))
  const out: PlannedEntry[] = []
  doc.forEach((node, pos) => {
    const heading = headings.get(pos)
    const styleId = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : ''
    const styleName = styleId
      ? (styles?.get(styleId)?.name ?? styleId)
      : node.type.name === 'docHeading' && node.attrs.outlineOnly !== true
        ? `heading ${Number(node.attrs.level) || 1}`
        : ''
    let level = node.isTextblock
      ? (levelOfStyle.get(styleName.toLowerCase()) ?? levelOfStyle.get(styleId.toLowerCase()))
      : undefined
    if (level === undefined && heading) {
      // a heading has an outline level through its style, so \\u lists it too
      const listed = plan.outline || (plan.headings && node.attrs.outlineOnly !== true)
      if (listed && heading.level >= plan.from && heading.level <= plan.to) level = heading.level
    }
    if (level !== undefined && node.textContent.trim())
      out.push({ text: node.textContent, level, pos })
    if (!plan.entryFields) return
    node.descendants((child) => {
      if (child.type.name !== 'docXeMark' || typeof child.attrs.tc !== 'string') return
      const field = tocEntryField(child.attrs.tc)
      if (!field.text || (field.id ?? '').toLowerCase() !== (plan.entryId ?? '').toLowerCase())
        return
      const [from, to] = plan.entryLevels ?? [1, 9]
      if (field.level >= from && field.level <= to)
        out.push({ text: field.text, level: field.level, pos })
    })
  })
  return out
}

/** the planned entries with the page each one is on, when the layout can say */
export function planEntriesWithPages(
  editor: Editor,
  plan: TocPlan,
  pagesOf?: (positions: number[]) => number[] | null,
): PlannedEntry[] {
  const entries = collectPlanEntries(editor, plan)
  const pages = pagesOf?.(entries.map((e) => e.pos))
  return pages && pages.length === entries.length
    ? entries.map((e, i) => ({ ...e, pageNo: pages[i] }))
    : entries
}

/** PM nodes for the lines of a generated TOC / TOA / INDEX field (one docProtected per line) */
export function fieldTableNodes(
  lines: Array<FieldTableLine & { level: number }>,
  table: TableSpec & { pageSep?: string; noLink?: boolean; tabPos?: number },
  label: string,
  /** the document's styles: a line's own indent is drawn on top of what its style carries */
  styles?: StyleLookup,
): Array<Record<string, unknown>> {
  // a built-in style goes by the id this document stores it under
  const styled = lines.map((line) =>
    line.style ? { ...line, style: documentStyleId(line.style, styles) } : line,
  )
  return generateFieldTableXml(styled, table).map((xml, i) => {
    const line = styled[i]
    const styleLeft = (line.style && styles?.get(line.style)?.display?.indentLeftTwips) || 0
    return {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label,
        genXml: xml,
        styleId: line.style ?? null,
        fieldDisplay: {
          kind: 'tocLine',
          left: line.text,
          ...(line.textRuns?.length ? { leftRuns: line.textRuns } : {}),
          right: line.noPage || line.pageNo === undefined ? '' : String(line.pageNo),
          level: line.level,
          leader: table.leader,
          ...(line.indent !== undefined
            ? { indentLeftTwips: Math.max(0, line.indent - styleLeft) }
            : {}),
          ...(line.noPage ? { noPage: true } : {}),
          ...(line.pageRuns?.length && !line.noPage ? { rightRuns: line.pageRuns } : {}),
          ...(table.pageSep !== undefined && !line.noPage ? { pageSep: table.pageSep } : {}),
          ...(table.noLink ? { noLink: true } : {}),
        } satisfies FieldDisplay,
      },
    }
  })
}

/**
 * PM nodes for a freshly generated TOC field (one docProtected per line).
 * With a spec the entries follow its switches: the levels listed, page
 * numbers, their separator, links.
 */
export function tocFieldNodes(
  entries: TocEntry[],
  spec?: TableSpec,
  /** the entries are the instruction's own (collectPlanEntries): none is filtered out again */
  planned = false,
  styles?: StyleLookup,
): Array<Record<string, unknown>> {
  if (spec) {
    const plan = tocPlan(spec.instr)
    const lines = entries
      .filter((e) => planned || (e.level >= plan.from && e.level <= plan.to))
      .map((e) => ({
        text: e.text,
        pageNo: e.pageNo,
        level: e.level,
        style: `TOC${Math.min(Math.max(e.level, 1), 9)}`,
        noPage: plan.noPage(e.level),
      }))
    return fieldTableNodes(
      lines,
      { ...spec, pageSep: plan.pageSep, noLink: !plan.hyperlinks },
      t('ribbonTocFieldLabel'),
      styles,
    )
  }
  return generateTocFieldXml(entries).map((xml, i) => {
    const own = `TOC${Math.min(Math.max(entries[i].level, 1), 9)}`
    const styleId = documentStyleId(own, styles)
    return {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonTocFieldLabel'),
        genXml:
          styleId === own
            ? xml
            : xml.replace(`<w:pStyle w:val="${own}"/>`, `<w:pStyle w:val="${styleId}"/>`),
        styleId,
        fieldDisplay: {
          kind: 'tocLine',
          left: entries[i].text,
          right: entries[i].pageNo !== undefined ? String(entries[i].pageNo) : '',
          level: entries[i].level,
        },
      },
    }
  })
}

/** heading pages as a lookup by position, for a caller that only measures headings */
export function headingPagesAt(
  editor: Editor,
  headingPages?: () => number[] | null,
): ((positions: number[]) => number[] | null) | undefined {
  if (!headingPages) return undefined
  return (positions) => {
    const pages = headingPages()
    const headings = collectHeadings(editor.state.doc, editor.storage.listNumbering?.styles)
    if (!pages || pages.length !== headings.length) return null
    const byPos = new Map(headings.map((h, i) => [h.pos, pages[i]]))
    const out = positions.map((pos) => byPos.get(pos))
    return out.every((page) => page !== undefined) ? (out as number[]) : null
  }
}

/** the instruction and tab leader of the table that starts at `from`; null when it has no instruction there */
export function tableSpecAt(
  doc: PmNode,
  blocks: Block[],
  from: number,
  keyword: 'TOC' | 'INDEX' | 'TOA',
): TableSpec | null {
  const first = doc.nodeAt(from)
  if (!first) return null
  const instr = fieldInstructionOfXml(xmlOfNode(first as never, blocks))
  // an entry of Word's own table can open with its PAGEREF instead
  if (!new RegExp(`^\\s*${keyword}(?:\\s|\\\\|$)`, 'i').test(instr)) return null
  // the entry's own right tab says it; a tab the style carries reaches here through the parsed line
  const xml = xmlOfNode(first as never, blocks)
  const own = /<w:tab\b(?=[^>]*w:val="right")[^>]*>/.exec(xml)?.[0]
  const display = first.attrs.fieldDisplay as FieldDisplay | null
  const leader = own
    ? ((/w:leader="([^"]*)"/.exec(own)?.[1] ?? 'none') as FieldDisplay['leader'])
    : display?.leader
  return { instr, leader: tableLeaderOf(leader) }
}

const PAGE_BREAK_PARAGRAPH_XML = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

/** OOXML of a top-level PM node: editor-generated fragment or original slice */
export function xmlOfNode(node: { attrs: Record<string, unknown> }, blocks: Block[]): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const idx = node.attrs.docxIndex
  if (idx === null || idx === undefined) return ''
  const xml = blocks.find((b) => b.docxIndex === idx)?.originalXml ?? ''
  // captions renumbered since the file was read show their new numbers
  const renumbered = node.attrs.seqRenumber as { results: Array<string | null> } | null | undefined
  return renumbered ? rewriteSeqResultsXml(xml, renumbered.results) : xml
}

/** the w:sdt opening bytes a top-level node carries (live attr, else the original block) */
function sdtOpenOf(
  node: { attrs: Record<string, unknown> },
  blocks: Block[],
): {
  open: string
  group: number | undefined
} {
  const raw = node.attrs.sdtShell
  if (typeof raw === 'string' && raw) {
    try {
      const shell = JSON.parse(raw) as { openXml?: string; group?: number }
      if (shell.openXml) return { open: shell.openXml, group: shell.group }
    } catch {
      // fall through to the block lookup
    }
  }
  const idx = node.attrs.docxIndex
  const block =
    idx === null || idx === undefined ? undefined : blocks.find((b) => b.docxIndex === idx)
  return { open: block?.sdtShell?.openXml ?? '', group: block?.sdtShell?.group }
}

// a table of figures is a TOC field too, told apart by its \c switch
export const TOC_INSTR = /<w:instrText[^>]*>\s*TOC[\s\\](?![^<]*\\[ac][\s"])/

/** Word automatic TOC: one w:sdt gallery control whose members share a sdt group */
function findSdtTocRange(doc: PmNode, blocks: Block[]): TocRange | null {
  const top: Array<{
    offset: number
    size: number
    xml: string
    open: string
    group: number | undefined
  }> = []
  doc.forEach((node, offset) => {
    const { open, group } = sdtOpenOf(node as never, blocks)
    top.push({ offset, size: node.nodeSize, xml: xmlOfNode(node as never, blocks), open, group })
  })
  const gi = top.findIndex((n) => TOC_GALLERY.test(n.open))
  if (gi === -1) return null
  const group = top[gi].group
  let last = gi
  if (group !== undefined) {
    top.forEach((n, i) => {
      if (n.group === group) last = Math.max(last, i)
    })
  }
  // keep a leading "Table of Contents" heading paragraph (Word updates the
  // entries under it, it does not drop the heading)
  const keepHeading = TOC_HEADING_STYLE.test(top[gi].xml)
  const from = keepHeading ? top[gi].offset + top[gi].size : top[gi].offset
  const to = top[last].offset + top[last].size
  if (from >= to) return null
  return {
    from,
    to,
    keepPageBreak: PAGE_BREAK.test(top[last].xml),
    kind: 'sdt',
  }
}

const TOC_GALLERY = /docPartGallery[^>]*Table of Contents/
const TOC_HEADING_STYLE = /w:pStyle\s+w:val\s*=\s*["']TOCHeading["']/
const PAGE_BREAK = /<w:br\s[^>]*w:type="page"/

/**
 * The replaced region of the document's first table of contents, or null
 * when the document has none (field shape first, then Word's sdt shape).
 */
export function findTocRange(doc: PmNode, blocks: Block[]): TocRange | null {
  const field = findFieldRange(doc, blocks, TOC_INSTR)
  if (field) return field
  return findSdtTocRange(doc, blocks)
}

/** The region of the first field with that instruction (field begin ... matching
 * end, tracked by fldChar depth across top-level blocks). */
export function findFieldRange(doc: PmNode, blocks: Block[], instr: RegExp): TocRange | null {
  return findFieldRanges(doc, blocks, instr)[0] ?? null
}

/** every field with that instruction, in document order */
export function findFieldRanges(doc: PmNode, blocks: Block[], instr: RegExp): TocRange[] {
  const out: TocRange[] = []
  let from = -1
  let depth = 0
  doc.forEach((node, offset) => {
    const xml = xmlOfNode(node as never, blocks)
    if (from === -1) {
      if (!instr.test(xml)) return
      from = offset
      depth = 0
    }
    depth += (xml.match(/w:fldCharType="begin"/g) ?? []).length
    depth -= (xml.match(/w:fldCharType="end"/g) ?? []).length
    if (depth <= 0) {
      out.push({
        from,
        to: offset + node.nodeSize,
        keepPageBreak: PAGE_BREAK.test(xml),
        kind: 'field',
      })
      from = -1
    }
  })
  return out
}

/**
 * Regenerate the document's table of contents from the current headings,
 * replacing the existing TOC (field or Word sdt shape) in place.
 * Returns false — after an alert, as the References tab always did — when
 * the document has no headings or no table of contents.
 */
export function updateTocField(
  editor: Editor,
  blocks: Block[],
  headingPages?: () => number[] | null,
  /** the instruction and leader to rebuild with; the table's own when not given */
  spec?: TableSpec,
  /** the table to rebuild; the document's first when not given */
  at?: TocRange,
  /** pages of any blocks: what an entry that is not a heading (a styled paragraph, a TC field) needs */
  pagesOf?: (positions: number[]) => number[] | null,
): boolean {
  const range = at ?? findTocRange(editor.state.doc, blocks)
  const own = spec ?? (range && tableSpecAt(editor.state.doc, blocks, range.from, 'TOC')) ?? null
  const entries = own
    ? planEntriesWithPages(
        editor,
        tocPlan(own.instr),
        pagesOf ?? headingPagesAt(editor, headingPages),
      )
    : collectTocEntriesWithPages(editor, headingPages)
  if (entries.length === 0) {
    window.alert(t('ribbonTocNoHeadings'))
    return false
  }
  if (!range) {
    window.alert(t('ribbonTocNotFound'))
    return false
  }

  const nodes = tocFieldNodes(entries, own ?? undefined, !!own, editorStyles(editor))
  if (nodes.length === 0) {
    window.alert(t('ribbonTocNoHeadings'))
    return false
  }
  if (range.keepPageBreak) {
    nodes.push({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonPageBreak'),
        genXml: PAGE_BREAK_PARAGRAPH_XML,
        fieldDisplay: { kind: 'pageBreak' },
      },
    })
  }

  editor
    .chain()
    .focus()
    .deleteRange({ from: range.from, to: range.to })
    .insertContentAt(range.from, nodes as never)
    .run()
  return true
}
