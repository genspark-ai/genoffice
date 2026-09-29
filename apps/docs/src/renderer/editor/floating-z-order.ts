import type { Editor } from '@tiptap/core'

// Stacking order among overlapping floating anchors. z-order only has a
// visible effect on floating (front/behind) anchors, so the callers enable it
// there; a bring-forward on an inline one also floats it (Word parity).

/** Rank of the selected anchor (docProtected): imageZOrder, absent = 0. */
function currentZOrder(editor: Editor): number {
  return Number(editor.getAttributes('docProtected').imageZOrder ?? 0)
}

/** z-order of every floating anchor in the document plus the selected anchor's own rank (an inline selection is not in the document list, so its rank is seeded here); Word's to-front/to-back are document-global */
function floatingZOrders(editor: Editor, rank: number): number[] {
  const zs: number[] = [rank]
  editor.state.doc.descendants((n) => {
    if (
      n.type.name === 'docProtected' &&
      (n.attrs.imageWrap === 'front' || n.attrs.imageWrap === 'behind')
    )
      zs.push(Number(n.attrs.imageZOrder ?? 0))
  })
  return zs
}

/** Apply a rank; an inline anchor has no paint order, so it floats in front to make the reorder meaningful. */
function setZOrder(editor: Editor, z: number): void {
  const attrs: Record<string, unknown> = { imageZOrder: z }
  const wrap = editor.getAttributes('docProtected').imageWrap
  if (wrap !== 'front' && wrap !== 'behind') attrs.imageWrap = 'front'
  editor.chain().focus().updateAttributes('docProtected', attrs).run()
}

export function bringToFront(editor: Editor): void {
  setZOrder(editor, Math.max(...floatingZOrders(editor, currentZOrder(editor))) + 1)
}

export function bringForward(editor: Editor): void {
  setZOrder(editor, currentZOrder(editor) + 1)
}

export function sendBackward(editor: Editor): void {
  setZOrder(editor, currentZOrder(editor) - 1)
}

export function sendToBack(editor: Editor): void {
  setZOrder(editor, Math.min(...floatingZOrders(editor, currentZOrder(editor))) - 1)
}
