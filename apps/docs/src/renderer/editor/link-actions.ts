/**
 * Word's hyperlink context menu: Edit / Open / Copy / Remove Hyperlink act on
 * the link run under the pointer, not on the selection.
 */
import type { Editor } from '@tiptap/core'
import type { Mark, Node as PmNode } from '@tiptap/pm/model'
import type { EditorState } from '@tiptap/pm/state'

export interface LinkRange {
  from: number
  to: number
  href: string
  tooltip: string | null
}

/** the whole link run around `pos` (either side of a boundary counts), null when not on a link */
export function linkRangeAt(state: EditorState, pos: number): LinkRange | null {
  const markType = state.schema.marks.link
  if (!markType || pos < 0 || pos > state.doc.content.size) return null
  const $pos = state.doc.resolve(pos)
  const parent = $pos.parent
  if (!parent.isTextblock) return null
  const findMark = (node: PmNode | null | undefined) =>
    node?.isText ? node.marks.find((m) => m.type === markType) : undefined
  const after = parent.childAfter($pos.parentOffset)
  let mark = findMark(after.node)
  let index = after.index
  if (!mark && $pos.parentOffset > 0) {
    const before = parent.childBefore($pos.parentOffset)
    mark = findMark(before.node)
    index = before.index
  }
  if (!mark) return null
  return rangeOfMark(parent, $pos.start(), index, mark)
}

function rangeOfMark(parent: PmNode, base: number, index: number, mark: Mark): LinkRange {
  let startIndex = index
  let endIndex = index + 1
  while (startIndex > 0 && mark.isInSet(parent.child(startIndex - 1).marks)) startIndex--
  while (endIndex < parent.childCount && mark.isInSet(parent.child(endIndex).marks)) endIndex++
  let from = base
  for (let i = 0; i < startIndex; i++) from += parent.child(i).nodeSize
  let to = from
  for (let i = startIndex; i < endIndex; i++) to += parent.child(i).nodeSize
  return {
    from,
    to,
    href: typeof mark.attrs.href === 'string' ? mark.attrs.href : '',
    tooltip: typeof mark.attrs.tooltip === 'string' ? mark.attrs.tooltip : null,
  }
}

/** Remove Hyperlink: the text stays, only the mark goes; the selection is left alone */
export function removeLink(editor: Editor, range: { from: number; to: number }): boolean {
  const markType = editor.state.schema.marks.link
  if (!markType || !editor.isEditable) return false
  editor.view.dispatch(editor.state.tr.removeMark(range.from, range.to, markType))
  return true
}

export type LinkTarget = { kind: 'external'; url: string } | { kind: 'anchor'; name: string } | null

/** what Open Hyperlink does with an address: browser for http(s), in-document jump for #bookmark */
export function linkTarget(href: string): LinkTarget {
  const text = href.trim()
  if (!text) return null
  if (text.startsWith('#')) return text.length > 1 ? { kind: 'anchor', name: text.slice(1) } : null
  if (/^https?:\/\//i.test(text)) return { kind: 'external', url: text }
  // bare domains from imported documents ("www.example.com")
  if (/^www\.[^\s/]+\.[a-z]{2,}(?:[/?#].*)?$/i.test(text))
    return { kind: 'external', url: `https://${text}` }
  return null
}
