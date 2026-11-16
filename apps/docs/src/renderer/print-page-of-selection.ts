/**
 * Where the print dialog's preview should start.
 *
 * The editor stamps each top-level block with data-idx (its docxIndex), and the
 * pagination preview clones those blocks into .pv-page sheets. Matching the two
 * gives the page the caret is on, so opening the print sheet shows the page the
 * reader was looking at instead of always opening on page 1.
 */
import type { Editor } from '@tiptap/core'

/** data-idx of the block holding the selection, or null when it cannot be read. */
export function blockIdAtSelection(editor: Editor | null): string | null {
  if (!editor) return null
  try {
    const { node } = editor.view.domAtPos(editor.state.selection.from)
    const el = node.nodeType === 1 ? (node as HTMLElement) : node.parentElement
    return el?.closest('[data-idx]')?.getAttribute('data-idx') ?? null
  } catch {
    // A selection at a boundary can have no DOM range yet.
    return null
  }
}

/** data-idx values present in each preview sheet, in sheet order. */
export function pageBlockIds(pages: readonly HTMLElement[]): string[][] {
  return pages.map((page) =>
    Array.from(page.querySelectorAll<HTMLElement>('[data-idx]'))
      .map((el) => el.dataset.idx)
      .filter((id): id is string => Boolean(id)),
  )
}

/** First sheet holding the block. A block split across sheets maps to its first. */
export function pageOfBlockId(
  ids: readonly (readonly string[])[],
  blockId: string | null,
): number | null {
  if (!blockId) return null
  const i = ids.findIndex((list) => list.includes(blockId))
  return i >= 0 ? i : null
}

/** 0-based preview page for the caret; falls back to the first page. */
export function initialPrintPageIndex(
  ids: readonly (readonly string[])[],
  blockId: string | null,
): number {
  return pageOfBlockId(ids, blockId) ?? 0
}
