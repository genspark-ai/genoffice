import type { Editor } from '@tiptap/core'

export interface OutlineItem {
  /** heading level, 1-6 */
  level: number
  /** collapsed single-line heading text */
  text: string
  /** document position before the heading node (resolves to a cursor inside it) */
  pos: number
}

/**
 * Headings in document order for the outline sidebar (PDF bookmark parity:
 * {@link apps/pdf/src/renderer/OutlinePanel.tsx}). Empty headings are skipped
 * so the pane never shows blank rows.
 */
export function collectOutline(editor: Editor): OutlineItem[] {
  const items: OutlineItem[] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'heading') return
    const text = node.textContent.replace(/\s+/g, ' ').trim()
    if (!text) return
    const level = node.attrs.level
    items.push({
      level: typeof level === 'number' ? Math.min(Math.max(Math.round(level), 1), 6) : 1,
      text,
      pos,
    })
  })
  return items
}
