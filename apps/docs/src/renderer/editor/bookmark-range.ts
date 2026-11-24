import type { Node as PmNode } from '@tiptap/pm/model'

export const RANGE_BOOKMARK_MARK = 'docRangeBookmark'

/** the text a bookmark was made from, while any of it is left */
export function bookmarkRange(doc: PmNode, name: string): { from: number; to: number } | null {
  const wanted = name.toUpperCase()
  let from = -1
  let to = -1
  doc.descendants((node, pos) => {
    if (!node.isInline) return true
    const marked = node.marks.some(
      (m) => m.type.name === RANGE_BOOKMARK_MARK && String(m.attrs.name).toUpperCase() === wanted,
    )
    if (marked) {
      if (from < 0) from = pos
      to = pos + node.nodeSize
    }
    return false
  })
  return from < 0 ? null : { from, to }
}
