import type { Node as PmNode } from '@tiptap/pm/model'

export interface HeadingRef {
  text: string
  level: number
  /** top-level position of the docHeading node */
  pos: number
}

/** Single heading predicate shared by TOC, nav pane and TOC page backfill (document order) */
export function collectHeadings(doc: PmNode): HeadingRef[] {
  const out: HeadingRef[] = []
  doc.forEach((node, offset) => {
    if (node.type.name === 'docHeading' && node.textContent.trim()) {
      out.push({ text: node.textContent, level: Number(node.attrs.level) || 1, pos: offset })
    }
  })
  return out
}
