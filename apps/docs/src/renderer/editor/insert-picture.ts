import type { Editor } from '@tiptap/core'
import type { ResolvedPos } from '@tiptap/pm/model'

export interface NewPicture {
  dataUrl: string
  base64: string
  mime: string
  widthPx: number
  heightPx: number
  label: string
}

function cellDepthOf($pos: ResolvedPos): number {
  for (let depth = $pos.depth; depth > 0; depth--) {
    const name = $pos.node(depth).type.name
    if (name === 'docTableCell' || name === 'docTableHeader') return depth
  }
  return -1
}

/** usable width of the cell's box, from its live DOM (0 when not rendered) */
function cellContentWidthPx(editor: Editor, cellPos: number): number {
  const dom = editor.view.nodeDOM(cellPos)
  if (!(dom instanceof HTMLElement)) return 0
  const style = getComputedStyle(dom)
  const pad = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight)
  return Math.floor(dom.clientWidth - (Number.isFinite(pad) ? pad : 0))
}

/**
 * Insert a picture at the caret. In the body it is a protected block (Word's
 * inline picture paragraph); inside a table cell, where that block is not
 * allowed, it becomes an inline picture run scaled to the cell width.
 */
export function insertPictureAtSelection(editor: Editor, picture: NewPicture): boolean {
  const { $from } = editor.state.selection
  const cellDepth = cellDepthOf($from)
  if (cellDepth > 0) {
    const cellWidth = cellContentWidthPx(editor, $from.before(cellDepth))
    const scale = cellWidth > 0 ? Math.min(1, cellWidth / picture.widthPx) : 1
    return editor
      .chain()
      .focus()
      .insertContent({
        type: 'docInlineImage',
        attrs: {
          dataUrl: picture.dataUrl,
          widthPx: Math.max(1, Math.round(picture.widthPx * scale)),
          heightPx: Math.max(1, Math.round(picture.heightPx * scale)),
        },
      })
      .run()
  }
  return editor
    .chain()
    .focus()
    .insertContent({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'image',
        label: picture.label,
        imageDataUrl: picture.dataUrl,
        imageWidthPx: picture.widthPx,
        imageHeightPx: picture.heightPx,
        genImage: {
          base64: picture.base64,
          mime: picture.mime,
          widthPx: picture.widthPx,
          heightPx: picture.heightPx,
        },
      },
    })
    .run()
}
