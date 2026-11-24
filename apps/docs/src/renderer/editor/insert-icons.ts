import type { Editor } from '@tiptap/core'
import { ICON_INSERT_PX, iconSvg, loadIconLibrary } from '../icon-library'
import { getActiveSubEditor } from './active-editor'
import { insertPicturesInLine, type NewPicture } from './insert-picture'
import { rememberPictureVector } from './picture-vectors'
import { rasteriseSvg, svgDataUrl } from './svg-raster'

const COMMIT_EVENT = 'ai-docs-commit-tables'

const prepared = new Map<string, NewPicture>()

/**
 * Draws the fallback bitmaps of the named library icons. Done while the dialog
 * is still up, so that inserting them afterwards waits for nothing and no
 * typing or caret move can come between two icons.
 */
export async function prepareLibraryIcons(names: string[]): Promise<void> {
  const icons = await loadIconLibrary()
  await Promise.all(
    names.map(async (name) => {
      const icon = icons.find((candidate) => candidate.name === name)
      if (!icon || prepared.has(name)) return
      const svg = iconSvg(icon)
      const dataUrl = await rasteriseSvg(svg, ICON_INSERT_PX, ICON_INSERT_PX)
      const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
      rememberPictureVector(base64, btoa(svg), icon.label)
      prepared.set(name, {
        dataUrl,
        vectorUrl: svgDataUrl(svg),
        base64,
        mime: 'image/png',
        widthPx: ICON_INSERT_PX,
        heightPx: ICON_INSERT_PX,
        label: icon.label,
        altText: icon.label,
      })
    }),
  )
}

/** a prepared icon, for Change Picture > From Icons */
export const preparedIcon = (name: string): NewPicture | undefined => prepared.get(name)

/**
 * Puts the prepared icons into the caret's paragraph, in the order given, as
 * one undo step. The caret may be in a header, footer or text box: the editor
 * that holds it takes them.
 */
export function insertPreparedIcons(editor: Editor, names: string[]): boolean {
  const sub = getActiveSubEditor()
  // a click in the dialog ends a text box's editing; the box still is where the icons go
  if (!editor.isEditable || !names.every((name) => prepared.has(name))) return false
  const inserted = insertPicturesInLine(
    sub ?? editor,
    names.map((name) => prepared.get(name)!),
  )
  // the box hands its text to the document, so the icons are in the undo history at once
  if (inserted && sub && !sub.schema.nodes.hfField) window.dispatchEvent(new Event(COMMIT_EVENT))
  return inserted
}

/** the same for icons nobody prepared: drawn first, then inserted together */
export async function insertLibraryIcons(editor: Editor, names: string[]): Promise<void> {
  await prepareLibraryIcons(names)
  insertPreparedIcons(
    editor,
    names.filter((name) => prepared.has(name)),
  )
}
