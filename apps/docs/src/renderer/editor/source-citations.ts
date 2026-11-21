import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import {
  bibliographyEntries,
  bibliographyFieldXml,
  citationFieldOf,
  citationInstruction,
  formatCitation,
  type Block,
  type CitationStyleId,
  type SourceInfo,
} from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import { findFieldRange, type TocRange } from './toc-update'

/** the style a document's citations take until one is picked */
export const DEFAULT_CITATION_STYLE: CitationStyleId = 'apa'

const CITATION_INSTR = /^\s*CITATION\s/i
const BIBLIOGRAPHY_INSTR = /<w:instrText[^>]*>\s*BIBLIOGRAPHY\b/

export interface CitationField {
  from: number
  to: number
  instr: string
  text: string
}

/** every CITATION field of the document, in order; a result split over runs is one field */
export function citationFields(doc: PmNode): CitationField[] {
  const out: CitationField[] = []
  const markType = doc.type.schema.marks.instrField
  if (!markType) return out
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return
    let offset = pos + 1
    let open: (CitationField & { id: unknown }) | null = null
    node.forEach((child) => {
      const mark = child.isText ? child.marks.find((m) => m.type === markType) : undefined
      const instr = mark ? String(mark.attrs.instr ?? '') : ''
      if (mark && CITATION_INSTR.test(instr)) {
        if (open && open.instr === instr && open.id === mark.attrs.fieldId) {
          open.to = offset + child.nodeSize
          open.text += child.text ?? ''
        } else {
          open = {
            from: offset,
            to: offset + child.nodeSize,
            instr,
            text: child.text ?? '',
            id: mark.attrs.fieldId,
          }
          out.push(open)
        }
      } else open = null
      offset += child.nodeSize
    })
    return false
  })
  return out.map(({ from, to, instr, text }) => ({ from, to, instr, text }))
}

/** the sources the text cites, in the order it first names them */
export function citedOrder(doc: PmNode): string[] {
  const tags = citationFields(doc).flatMap((field) => citationFieldOf(field.instr)?.tags ?? [])
  return [...new Set(tags)]
}

/** the document's bibliography field, when it has one */
export const bibliographyRange = (doc: PmNode, blocks: Block[]): TocRange | null =>
  findFieldRange(doc, blocks, BIBLIOGRAPHY_INSTR)

function bibliographyNodes(
  sources: SourceInfo[],
  style: CitationStyleId,
  order: string[],
): Array<Record<string, unknown>> {
  const entries = bibliographyEntries(sources, style, order)
  return bibliographyFieldXml(entries).map((xml, i) => ({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonBibliography'),
      genXml: xml,
      fieldDisplay: {
        kind: 'text',
        left: entries[i].map((seg) => seg.text).join(''),
        runs: entries[i].map((seg) => ({
          text: seg.text,
          ...(seg.italic ? { italic: true } : {}),
        })),
        indentLeftTwips: 720,
        indentFirstLineTwips: -720,
      },
    },
  }))
}

/** rewrites, inside the transaction, every citation and the bibliography of its document in the style */
function restyleInto(
  tr: Transaction,
  blocks: Block[],
  sources: SourceInfo[],
  style: CitationStyleId,
): void {
  const { schema } = tr.doc.type
  const markType = schema.marks.instrField
  const before = tr.doc
  const order = citedOrder(before)
  for (const field of citationFields(before).reverse()) {
    const text = formatCitation(field.instr, sources, style, order)
    if (text === null || text === field.text) continue
    const marks = (before.nodeAt(field.from)?.marks ?? []).map((m) =>
      m.type === markType ? markType.create({ ...m.attrs, fieldPart: 'single' }) : m,
    )
    tr.replaceWith(field.from, field.to, schema.text(text, marks))
  }
  const range = sources.length > 0 ? bibliographyRange(tr.doc, blocks) : null
  if (range)
    tr.replaceWith(
      range.from,
      range.to,
      bibliographyNodes(sources, style, order).map((json) => schema.nodeFromJSON(json)),
    )
}

/**
 * Writes every citation and the bibliography in the style: what Word does
 * when its Style: box changes. One step of the undo history.
 */
export function restyleCitations(
  editor: Editor,
  blocks: Block[],
  sources: SourceInfo[],
  style: CitationStyleId,
): void {
  const tr = editor.state.tr
  restyleInto(tr, blocks, sources, style)
  if (tr.docChanged) editor.view.dispatch(tr)
}

/** a number no field of the document carries yet: it keeps two fields side by side apart */
export function nextFieldId(doc: PmNode): number {
  let max = 0
  doc.descendants((node) => {
    for (const mark of node.marks)
      if (mark.type.name === 'instrField') max = Math.max(max, Number(mark.attrs.fieldId) || 0)
  })
  return max + 1
}

/**
 * Puts a citation of the source at the caret, as a CITATION field showing it
 * in the style. The citations a numbered style renumbers and the bibliography
 * that takes the source change in the same step, so one Undo takes all of it back.
 */
export function insertCitationField(
  editor: Editor,
  blocks: Block[],
  source: SourceInfo,
  sources: SourceInfo[],
  style: CitationStyleId,
): void {
  const { state } = editor
  const markType = state.schema.marks.instrField
  const instr = citationInstruction(source.tag)
  const field = markType.create({ instr, fieldId: nextFieldId(state.doc), fieldPart: 'single' })
  const marks = [...state.selection.$from.marks().filter((m) => m.type !== markType), field]
  const tr = state.tr.replaceSelectionWith(
    state.schema.text(formatCitation(instr, sources, style) ?? `(${source.tag})`, marks),
    false,
  )
  restyleInto(tr, blocks, sources, style)
  editor.view.dispatch(tr.scrollIntoView())
  editor.commands.focus()
}

/** the bibliography of the document's sources at the caret, or in place of the one it has */
export function bibliographyFieldNodes(
  editor: Editor,
  sources: SourceInfo[],
  style: CitationStyleId,
): Array<Record<string, unknown>> {
  return bibliographyNodes(sources, style, citedOrder(editor.state.doc))
}
