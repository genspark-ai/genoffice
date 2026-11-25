import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection } from '@tiptap/pm/state'
import { NEW_PICTURE_SVG_SUFFIX, retargetImageBlip } from '@genoffice/docx-engine'
import { getActiveSubEditor } from './active-editor'
import { heldPictureToken } from './convert'
import { originalPictureSize, pictureDensityOf } from './picture-density'
import { pictureVectorAttr } from './picture-vectors'

export async function imageSizeOf(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight })
    img.onerror = () => reject(new Error('Failed to read image'))
    img.src = dataUrl
  })
}

// enough of the file for its header segments (an Exif block is at most 64 KB)
const HEAD_BYTES = 192 * 1024

async function headBytesOf(url: string): Promise<Uint8Array> {
  const data = /^data:[^;,]+;base64,/.exec(url)
  if (!data) return new Uint8Array(await (await fetch(url)).arrayBuffer()).subarray(0, HEAD_BYTES)
  const head = atob(url.slice(data[0].length, data[0].length + (HEAD_BYTES / 3) * 4))
  return Uint8Array.from(head, (ch) => ch.charCodeAt(0))
}

/** Word's original size of a picture in px: its pixels at the density the file states */
export async function originalPictureSizeOf(
  url: string,
): Promise<{ width: number; height: number }> {
  const pixels = await imageSizeOf(url)
  const density = await headBytesOf(url).then(pictureDensityOf, () => null)
  return originalPictureSize(pixels, density)
}

/** the selected picture that sits in a line of text (a single-layer run picture) */
export function selectedRunPicture(editor: Editor): PmNode | null {
  const { selection } = editor.state
  if (!(selection instanceof NodeSelection)) return null
  const { node } = selection
  return node.type.name === 'docInlineImage' && !node.attrs.rule && !node.attrs.layers ? node : null
}

/** Change Picture has a picture to work on: a picture block, or a picture in a line of text */
export const canChangePicture = (editor: Editor): boolean =>
  editor.getAttributes('docProtected')?.blockType === 'image' || selectedRunPicture(editor) !== null

// the editor of a text box, header or footer in which a right-click last selected a picture
let pointed: Editor | null = null

/**
 * The header, footer or text box whose selected picture Change Picture acts on: the one
 * being edited, or the text box a right-click picked a picture in while that box is still
 * the object selected in the document.
 */
export function subPictureEditor(editor: Editor): Editor | null {
  const active = getActiveSubEditor()
  if (active?.isEditable && selectedRunPicture(active)) return active
  if (!pointed || pointed.isDestroyed || !selectedRunPicture(pointed)) return null
  const { selection } = editor.state
  const dom = selection instanceof NodeSelection ? editor.view.nodeDOM(selection.from) : null
  return dom?.contains(pointed.view.dom) ? pointed : null
}

/** the editor whose selected picture Change Picture acts on: a text box or strip, else the document */
export function changePictureEditor(editor: Editor): Editor | null {
  return subPictureEditor(editor) ?? (canChangePicture(editor) ? editor : null)
}

/**
 * A right-click on a picture in a header, footer or text box selects it in the editor
 * that holds it, as a right-click on a picture of the document does; false when it is not on one.
 */
export function selectSubEditorPicture(editor: Editor, target: EventTarget | null): boolean {
  const img = (target as Element | null)?.closest?.('img.doc-inline-img')
  const root = img?.closest('.ProseMirror') as (Element & { editor?: Editor }) | null | undefined
  const sub = root?.editor
  if (!img || !sub || sub === editor || sub.isDestroyed) return false
  let at = -1
  sub.state.doc.descendants((node, pos) => {
    // the picture's own element, or the span a floating one is drawn in
    if (at === -1 && node.type.name === 'docInlineImage' && sub.view.nodeDOM(pos)?.contains(img))
      at = pos
    return at === -1
  })
  if (at === -1) return false
  sub.view.dispatch(sub.state.tr.setSelection(NodeSelection.create(sub.state.doc, at)))
  pointed = sub
  return true
}

/**
 * A picture in a line of text takes new bytes where it stands: its drawing
 * stays (place in the line, wrapping, rotation), pointed at a placeholder the
 * save turns into the new media part. Width holds, height follows the new picture.
 */
async function replaceRunPictureBytes(
  editor: Editor,
  picture: PmNode,
  dataUrl: string,
  vectorUrl?: string,
): Promise<void> {
  const natural = await imageSizeOf(dataUrl)
  const { selection } = editor.state
  if (!(selection instanceof NodeSelection) || selection.node !== picture) return
  const widthPx = Number(picture.attrs.widthPx) || Math.min(natural.width, 620)
  const heightPx = Math.max(1, Math.round((widthPx * natural.height) / natural.width))
  const vector = pictureVectorAttr(dataUrl, vectorUrl)
  const token = heldPictureToken()
  const xml = String(picture.attrs.xml ?? '')
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(selection.from, undefined, {
      ...picture.attrs,
      dataUrl,
      vector,
      widthPx,
      heightPx,
      // bytes that no longer come from the linked file: the drawing drops the link with them
      link: null,
      // a picture not saved yet has no drawing: the save writes one from its bytes
      xml:
        xml && retargetImageBlip(xml, token, vector ? token + NEW_PICTURE_SVG_SUFFIX : undefined),
    }),
  )
  // a text box hands its content to the document when it is asked to, as before a save
  if (editor.view.dom.closest('.doc-textbox'))
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
}

/**
 * Replace the selected image's bytes (shared by Change Picture / remove background / crop).
 * Original images (docxIndex set) swap bytes in place via the imageReplace patch: the
 * drawing XML survives, so wrap/position/docxIndex — and with them the Position gallery —
 * keep working. Images not yet saved (genImage) just update their pending payload.
 * Display size keeps the current width; height adapts to the new image's aspect ratio.
 */
export async function replaceSelectedPictureBytes(
  editor: Editor,
  dataUrl: string,
  /** the size to show the new bytes at, when they are a part of the old picture */
  shown?: { widthPx: number; heightPx: number },
  /** the SVG the new bytes were drawn from, shown and saved in their place */
  vectorUrl?: string,
): Promise<void> {
  // a text box that is selected as an object is not typed in, yet its picture can be changed
  if (!editor.isEditable && editor !== pointed) return
  const m = /^data:(image\/(?:png|jpeg|gif));base64,(.*)$/s.exec(dataUrl)
  if (!m) return
  const inLine = selectedRunPicture(editor)
  if (inLine) return replaceRunPictureBytes(editor, inLine, dataUrl, vectorUrl).catch(() => {})
  const attrs = editor.getAttributes('docProtected')
  if (attrs?.blockType !== 'image') return
  try {
    const natural = await imageSizeOf(dataUrl)
    const currentW = Number(attrs.imageWidthPx) || Math.min(natural.width, 620)
    const w = Math.max(1, Math.round(shown?.widthPx ?? currentW))
    const h = Math.max(
      1,
      Math.round(shown?.heightPx ?? (currentW * natural.height) / natural.width),
    )
    const isOriginal = attrs.docxIndex !== null && attrs.docxIndex !== undefined
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', {
        imageDataUrl: dataUrl,
        imageVector: pictureVectorAttr(dataUrl, vectorUrl),
        imageWidthPx: w,
        imageHeightPx: h,
        // The new bytes are the full picture (crop/cutout bake destructively) and
        // the replace pipeline strips a:srcRect on save — drop a Word-authored
        // crop/fill window or it would keep clipping the new image until reload
        imageCrop: null,
        imageFillRect: null,
        // bytes that no longer come from the linked file: the save drops the link with them
        imageLink: null,
        ...(isOriginal
          ? { imageReplace: { base64: m[2], mime: m[1] } }
          : { genImage: { base64: m[2], mime: m[1], widthPx: w, heightPx: h } }),
      })
      .run()
  } catch {
    /* image decode failed: keep the original untouched */
  }
}
