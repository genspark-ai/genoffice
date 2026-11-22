import type { Node as PmNode } from '@tiptap/pm/model'

/** the first footnote or endnote reference at or after `from`; null when there is none */
export function nextNoteRefPos(
  doc: PmNode,
  from: number,
  kind: 'footnote' | 'endnote',
): number | null {
  let found: number | null = null
  doc.nodesBetween(Math.min(from, doc.content.size), doc.content.size, (node, pos) => {
    if (found !== null) return false
    if (pos >= from && node.type.name === 'docNoteRef' && node.attrs.kind === kind) found = pos
    return found === null
  })
  return found
}

export interface NoteRef {
  kind: 'footnote' | 'endnote'
  id: string
}

/** the first note reference at or after `from`, or else the document's first; null when it has none */
export function noteRefNear(doc: PmNode, from: number): NoteRef | null {
  const refs: Array<NoteRef & { pos: number }> = []
  doc.descendants((node, pos) => {
    if (node.type.name !== 'docNoteRef') return true
    refs.push({ kind: node.attrs.kind, id: String(node.attrs.id), pos })
    return false
  })
  const ref = refs.find((r) => r.pos >= from) ?? refs[0]
  return ref ? { kind: ref.kind, id: ref.id } : null
}
