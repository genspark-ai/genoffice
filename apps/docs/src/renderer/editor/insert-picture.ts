import type { Editor } from '@tiptap/core'
import type { ResolvedPos } from '@tiptap/pm/model'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import { heldPictureXml } from './convert'
import { pictureVectorAttr } from './picture-vectors'

export interface NewPicture {
  dataUrl: string
  /** the picture as an SVG data URL: what the page draws while the bitmap is this one */
  vectorUrl?: string
  base64: string
  mime: string
  widthPx: number
  heightPx: number
  label: string
  altText?: string
}

function cellDepthOf($pos: ResolvedPos): number {
  for (let depth = $pos.depth; depth > 0; depth--) {
    const name = $pos.node(depth).type.name
    if (name === 'docTableCell' || name === 'docTableHeader' || name === 'hfCell') return depth
  }
  return -1
}

const contentWidthPx = (dom: unknown): number => {
  if (!(dom instanceof HTMLElement)) return 0
  const style = getComputedStyle(dom)
  const pad = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight)
  return Math.floor(dom.clientWidth - (Number.isFinite(pad) ? pad : 0))
}

/** usable width of the cell's box, from its live DOM (0 when not rendered) */
const cellContentWidthPx = (editor: Editor, cellPos: number): number =>
  contentWidthPx(editor.view.nodeDOM(cellPos))

/** the pictures as picture runs of the caret's paragraph, no wider than `maxWidth` px (0 = any width) */
function insertPictureRuns(
  editor: Editor,
  pictures: NewPicture[],
  maxWidth: number,
  named = false,
  /** the editor of a header, footer or text box: each run carries its drawing to the save */
  held = false,
): boolean {
  return (
    editor
      .chain()
      // a header's strip is already on screen; the default focus scroll jumps the page
      .focus(undefined, { scrollIntoView: !held })
      .insertContent(
        pictures.map((picture) => {
          const scale = maxWidth > 0 ? Math.min(1, maxWidth / picture.widthPx) : 1
          const widthPx = Math.max(1, Math.round(picture.widthPx * scale))
          const heightPx = Math.max(1, Math.round(picture.heightPx * scale))
          return {
            type: 'docInlineImage',
            attrs: {
              dataUrl: picture.dataUrl,
              vector: pictureVectorAttr(picture.dataUrl, picture.vectorUrl),
              ...(named && picture.altText ? { altText: picture.altText } : {}),
              ...(held ? { xml: heldPictureXml({ ...picture, widthPx, heightPx }) } : {}),
              widthPx,
              heightPx,
            },
          }
        }),
      )
      .run()
  )
}

/**
 * Pictures that sit in the line like a character (Word's icons): all of them
 * go into the caret's paragraph, side by side, in one undo step. Where there
 * is no paragraph to take them (a selected object) each gets one of its own.
 */
export function insertPicturesInLine(editor: Editor, pictures: NewPicture[]): boolean {
  if (pictures.length === 0) return false
  const { $from } = editor.state.selection
  const held = !editor.schema.nodes.docProtected
  if (!$from.parent.isTextblock) return !held && insertPictures(editor, pictures, true)
  const cellDepth = cellDepthOf($from)
  const maxWidth =
    cellDepth > 0
      ? cellContentWidthPx(editor, $from.before(cellDepth))
      : held
        ? contentWidthPx(editor.view.dom)
        : 0
  return insertPictureRuns(editor, pictures, maxWidth, true, held)
}

/**
 * Insert pictures at the caret, one after another, in one transaction. In the
 * body each is a protected block (Word's inline picture paragraph); inside a
 * table cell, where that block is not allowed, they become inline picture
 * runs scaled to the cell width.
 */
function insertPictures(editor: Editor, pictures: NewPicture[], settle: boolean): boolean {
  if (pictures.length === 0) return false
  const { $from } = editor.state.selection
  const cellDepth = cellDepthOf($from)
  if (cellDepth > 0)
    return insertPictureRuns(editor, pictures, cellContentWidthPx(editor, $from.before(cellDepth)))
  return editor
    .chain()
    .focus()
    .insertContent(
      pictures.map((picture) => ({
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'image',
          label: picture.label,
          imageDataUrl: picture.dataUrl,
          imageVector: pictureVectorAttr(picture.dataUrl, picture.vectorUrl),
          imageWidthPx: picture.widthPx,
          imageHeightPx: picture.heightPx,
          ...(picture.altText ? { imageAlt: picture.altText } : {}),
          genImage: {
            base64: picture.base64,
            mime: picture.mime,
            widthPx: picture.widthPx,
            heightPx: picture.heightPx,
          },
        },
      })),
    )
    .command(({ tr }) => {
      if (settle) settleCaret(tr)
      return true
    })
    .run()
}

export function insertPictureAtSelection(editor: Editor, picture: NewPicture): boolean {
  return insertPictures(editor, [picture], false)
}

/**
 * A picture inserted into an empty document is the only node, with a node
 * selection on it: there is no text position to type at, and the next
 * keystroke or insert replaces the picture. A paragraph follows it and takes
 * the caret, as in Word. A mid-paragraph insert already leaves the caret in
 * the split-off rest of the paragraph.
 */
function settleCaret(tr: Transaction): void {
  const { doc } = tr
  const $after = doc.resolve(Math.min(tr.selection.to, doc.content.size))
  if ($after.parent.isTextblock) return
  const paragraph = doc.type.schema.nodes.docParagraph
  if ($after.nodeAfter?.isTextblock !== true && paragraph) tr.insert($after.pos, paragraph.create())
  tr.setSelection(TextSelection.near(tr.doc.resolve($after.pos + 1)))
}

export function settleCaretAfterPicture(editor: Editor): void {
  editor
    .chain()
    .command(({ tr }) => {
      settleCaret(tr)
      return true
    })
    .run()
}
