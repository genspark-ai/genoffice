import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { generateTocFieldXml, type Block, type TocEntry } from '@genoffice/docx-engine'
import { collectHeadings } from './headings'
import { t } from '../i18n/locale'

/**
 * Table-of-contents (目录) detection + regeneration, shared by the
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

/** PM nodes for a freshly generated TOC field (one docProtected per line) */
export function tocFieldNodes(entries: TocEntry[]): Array<Record<string, unknown>> {
  return generateTocFieldXml(entries).map((xml, i) => ({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonTocFieldLabel'),
      genXml: xml,
      fieldDisplay: {
        kind: 'tocLine',
        left: entries[i].text,
        right: entries[i].pageNo !== undefined ? String(entries[i].pageNo) : '',
        level: entries[i].level,
      },
    },
  }))
}

const PAGE_BREAK_PARAGRAPH_XML = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

/** OOXML of a top-level PM node: editor-generated fragment or original slice */
export function xmlOfNode(node: { attrs: Record<string, unknown> }, blocks: Block[]): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const idx = node.attrs.docxIndex
  if (idx === null || idx === undefined) return ''
  return blocks.find((b) => b.docxIndex === idx)?.originalXml ?? ''
}

/** the w:sdt opening bytes a top-level node carries (live attr, else the original block) */
function sdtOpenOf(node: { attrs: Record<string, unknown> }, blocks: Block[]): {
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

const TOC_INSTR = /<w:instrText[^>]*>\s*TOC[\s\\]/

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
  const field = findFieldTocRange(doc, blocks)
  if (field) return field
  return findSdtTocRange(doc, blocks)
}

/** Find the TOC field region (field begin ... matching end, tracked by fldChar
 * depth across top-level blocks). */
function findFieldTocRange(doc: PmNode, blocks: Block[]): TocRange | null {
  let from = -1
  let to = -1
  let keepPageBreak = false
  let depth = 0
  let found = false
  doc.forEach((node, offset) => {
    if (found) return
    const xml = xmlOfNode(node as never, blocks)
    if (from === -1) {
      if (TOC_INSTR.test(xml)) {
        from = offset
        depth = 0
      } else {
        return
      }
    }
    depth += (xml.match(/w:fldCharType="begin"/g) ?? []).length
    depth -= (xml.match(/w:fldCharType="end"/g) ?? []).length
    if (depth <= 0) {
      to = offset + node.nodeSize
      keepPageBreak = PAGE_BREAK.test(xml)
      found = true
    }
  })
  if (from === -1 || !found) return null
  return { from, to, keepPageBreak, kind: 'field' }
}

/** whether a document position sits inside the document's table of contents */
export function tocAtPos(doc: PmNode, blocks: Block[], pos: number): boolean {
  const range = findTocRange(doc, blocks)
  return range !== null && pos >= range.from && pos <= range.to
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
): boolean {
  const entries = collectTocEntriesWithPages(editor, headingPages)
  if (entries.length === 0) {
    window.alert(t('ribbonTocNoHeadings'))
    return false
  }

  const range = findTocRange(editor.state.doc, blocks)
  if (!range) {
    window.alert(t('ribbonTocNotFound'))
    return false
  }

  const nodes = tocFieldNodes(entries)
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
