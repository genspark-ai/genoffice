import type { Node as PmNode } from '@tiptap/pm/model'
import type { StyleInfo } from '@genoffice/docx-engine'

export interface HeadingRef {
  text: string
  level: number
  /** top-level position of the heading node */
  pos: number
}

/** Paragraph-style lookup that resolves the heading level a style confers (never written back on save) */
export type HeadingStyles = ReadonlyMap<
  string,
  Pick<StyleInfo, 'headingLevel' | 'headingOutlineOff'>
>

/** Node types that can be a heading without being a docHeading */
const HEADING_CAPABLE: Record<string, true> = { docParagraph: true, docListItem: true }

/**
 * Single heading predicate shared by TOC, nav pane and TOC page backfill (document order).
 *
 * A paragraph is a heading as a docHeading node, or through its paragraph style:
 * Word's numbered headings keep their numbering in the style (w:pPr/w:numPr → the
 * multilevel list definition), so the parser classifies those paragraphs as list
 * items — Word still lists them in its navigation pane and the TOC, and so does
 * this predicate when the caller supplies the document's styles.
 */
export function collectHeadings(doc: PmNode, styles?: HeadingStyles): HeadingRef[] {
  const out: HeadingRef[] = []
  doc.forEach((node, offset) => {
    if (node.type.name === 'docHeading') {
      if (node.textContent.trim())
        out.push({ text: node.textContent, level: Number(node.attrs.level) || 1, pos: offset })
      return
    }
    if (!styles || HEADING_CAPABLE[node.type.name] !== true) return
    const styleId = typeof node.attrs?.styleId === 'string' ? node.attrs.styleId : ''
    const style = styleId ? styles.get(styleId) : undefined
    // w:outlineLvl 9 on the style means body text even when a basedOn ancestor is a heading
    if (!style?.headingLevel || style.headingOutlineOff) return
    if (!node.textContent.trim()) return
    out.push({ text: node.textContent, level: style.headingLevel, pos: offset })
  })
  return out
}
