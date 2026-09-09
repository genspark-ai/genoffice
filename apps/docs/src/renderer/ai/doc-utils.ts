import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

/** ProseMirror positions of a top-level child index range */
export function blockRangePositions(
  editor: Editor,
  startIndex: number,
  endIndex: number,
): { from: number; to: number } {
  const doc = editor.state.doc
  let from = 0
  let to = 0
  let index = 0
  doc.forEach((node, offset) => {
    if (index === startIndex) from = offset
    if (index === endIndex) to = offset + node.nodeSize
    index++
  })
  return { from, to }
}

// ---- tracked deletions (pending revisions are not current content) ----

const hasDelMark = (node: ProseMirrorNode) => node.marks.some((m) => m.type.name === 'del')

/** block text as it reads once pending tracked deletions are applied (textContent minus del runs) */
export function liveText(node: ProseMirrorNode): string {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return ''
  let out = ''
  const walk = (child: ProseMirrorNode): void => {
    if (hasDelMark(child)) return
    if (child.isText) out += child.text ?? ''
    else if (child.isLeaf) out += child.type.spec.leafText?.(child) ?? ''
    else child.forEach(walk)
  }
  node.forEach(walk)
  return out
}

/**
 * The whole block is a pending deletion revision (struck through in the editor).
 * Checked structurally, not by text length: atom leaves (inline formulas, ruby)
 * carry no textContent, so a live formula must still count as live content.
 */
export function isTrackedDeleted(node: ProseMirrorNode): boolean {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return true
  let hasContent = false
  let hasLive = false
  node.descendants((child) => {
    if (child.isText || (child.isInline && child.isLeaf)) {
      hasContent = true
      if (!hasDelMark(child)) hasLive = true
    }
  })
  return hasContent && !hasLive
}
