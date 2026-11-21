import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { runPictureAltOf } from '@genoffice/docx-engine'
import { autoCropDataUrl } from './auto-crop'
import { replaceSelectedPictureBytes } from './picture-bytes'

const EMU_PER_PX = 9525

/** what Size and Position shows of the selected picture */
export interface PictureGeometry {
  widthPx: number
  heightPx: number
  /** degrees clockwise, 0-359 */
  rotDeg: number
  /** a picture in a line of text keeps the turn its file gives it */
  canRotate: boolean
  /** a floating picture placed by offsets: where it is, px from what its anchor measures from */
  offsetPx: { x: number; y: number } | null
  /** a picture alone in its paragraph can stand at a preset place of the text area */
  canPreset: boolean
  /** the preset place it stands at; in the result, the one to move it to */
  preset: PicturePreset | null
  /** it floats: text wraps around it, or it lies over or under the text */
  floating: boolean
  /** the picture's bytes, for the size it was made at */
  dataUrl: string
}

export interface PicturePreset {
  h: 'left' | 'center' | 'right'
  v: 'top' | 'center' | 'bottom'
}

const PRESET_H = ['left', 'center', 'right'] as const
const PRESET_V = ['top', 'center', 'bottom'] as const

interface Selected {
  pos: number
  inline: boolean
  attrs: Record<string, unknown>
}

function selectedPicture(editor: Editor): Selected | null {
  const sel = editor.state.selection
  if (!(sel instanceof NodeSelection)) return null
  const { node } = sel
  if (node.type.name === 'docInlineImage') return { pos: sel.from, inline: true, attrs: node.attrs }
  if (node.type.name === 'docProtected' && node.attrs.blockType === 'image')
    return { pos: sel.from, inline: false, attrs: node.attrs }
  return null
}

export function selectedPictureGeometry(editor: Editor): PictureGeometry | null {
  const picked = selectedPicture(editor)
  if (!picked) return null
  const a = picked.attrs
  const widthPx = Number(picked.inline ? a.widthPx : a.imageWidthPx)
  const heightPx = Number(picked.inline ? a.heightPx : a.imageHeightPx)
  if (!(widthPx > 0) || !(heightPx > 0)) return null
  const x = picked.inline ? a.offsetXEmu : a.imageOffsetXEmu
  const y = picked.inline ? a.offsetYEmu : a.imageOffsetYEmu
  const wrap = picked.inline ? a.wrap : a.imageWrap
  const presetH = PRESET_H.find((k) => k === a.imagePosH)
  const presetV = PRESET_V.find((k) => k === a.imagePosV)
  return {
    widthPx,
    heightPx,
    rotDeg: Number(picked.inline ? a.rotDeg : a.imageRotDeg) || 0,
    canRotate: !picked.inline,
    offsetPx:
      wrap && x != null && y != null
        ? { x: Number(x) / EMU_PER_PX, y: Number(y) / EMU_PER_PX }
        : null,
    canPreset: !picked.inline,
    preset: !picked.inline && presetH && presetV ? { h: presetH, v: presetV } : null,
    floating: !!wrap,
    dataUrl: String((picked.inline ? a.dataUrl : a.imageDataUrl) ?? ''),
  }
}

/** Size and Position's OK: the picture takes the size, turn and place in one undo step. */
export function applyPictureGeometry(editor: Editor, next: PictureGeometry): boolean {
  const picked = selectedPicture(editor)
  if (!picked || !editor.isEditable) return false
  const widthPx = Math.max(1, Math.round(next.widthPx))
  const heightPx = Math.max(1, Math.round(next.heightPx))
  const rot = ((Math.round(next.rotDeg) % 360) + 360) % 360 || null
  const offset = next.offsetPx
    ? {
        x: Math.round(next.offsetPx.x * EMU_PER_PX),
        y: Math.round(next.offsetPx.y * EMU_PER_PX),
      }
    : null
  const preset = picked.inline ? null : next.preset
  const attrs = picked.inline
    ? {
        ...picked.attrs,
        widthPx,
        heightPx,
        ...(offset ? { offsetXEmu: offset.x, offsetYEmu: offset.y } : {}),
      }
    : {
        ...picked.attrs,
        imageWidthPx: widthPx,
        imageHeightPx: heightPx,
        imageRotDeg: rot,
        ...(preset
          ? {
              // the same place the Position gallery gives
              imageWrap: preset.h === 'right' ? 'square-right' : 'square-left',
              imagePosH: preset.h,
              imagePosV: preset.v,
              // what the save writes the pair against, so the page shows it there before a reopen
              imageRelH: 'margin',
              imageRelV: 'margin',
              imageOffsetXEmu: null,
              imageOffsetYEmu: null,
            }
          : offset
            ? { imageOffsetXEmu: offset.x, imageOffsetYEmu: offset.y }
            : {}),
      }
  const { state, view } = editor
  const tr = state.tr.setNodeMarkup(picked.pos, undefined, attrs)
  view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, picked.pos)))
  return true
}

/** the alternative text of the selected picture; null when no picture that can hold one is selected */
export function selectedPictureAlt(editor: Editor): string | null {
  const picked = selectedPicture(editor)
  if (!picked) return null
  if (!picked.inline) return String(picked.attrs.imageAlt ?? '')
  if (picked.attrs.rule) return null
  if (picked.attrs.altText != null) return String(picked.attrs.altText)
  const xml = String(picked.attrs.xml ?? '')
  // a picture inserted in this session gets its drawing when the file is saved
  return xml ? runPictureAltOf(xml) : ''
}

export function setSelectedPictureAlt(editor: Editor, alt: string): boolean {
  const picked = selectedPicture(editor)
  const current = selectedPictureAlt(editor)
  if (!picked || current === null || !editor.isEditable) return false
  const text = alt.trim()
  if (text === current) return false
  if (!picked.inline)
    return editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageAlt: text || null })
      .run()
  const { state, view } = editor
  const tr = state.tr.setNodeMarkup(picked.pos, undefined, { ...picked.attrs, altText: text })
  view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, picked.pos)))
  return true
}

/**
 * Word's Auto Crop: the uniform border around the selected picture goes, and
 * what is left keeps the scale it was shown at. False when there is no border.
 */
export async function autoCropSelectedPicture(editor: Editor): Promise<boolean> {
  const picked = selectedPicture(editor)
  if (!picked || picked.inline || !editor.isEditable) return false
  const src = String(picked.attrs.imageDataUrl ?? '')
  if (!src) return false
  let cropped: Awaited<ReturnType<typeof autoCropDataUrl>>
  try {
    cropped = await autoCropDataUrl(src)
  } catch {
    return false
  }
  if (!cropped) return false
  // the decode took a moment: only the picture that was asked about is cropped
  const now = selectedPicture(editor)
  if (!now || now.pos !== picked.pos || now.attrs.imageDataUrl !== picked.attrs.imageDataUrl)
    return true
  const shownW = Number(picked.attrs.imageWidthPx) || cropped.width
  const shownH = Number(picked.attrs.imageHeightPx) || cropped.height
  await replaceSelectedPictureBytes(editor, cropped.dataUrl, {
    widthPx: (shownW * cropped.box.w) / cropped.width,
    heightPx: (shownH * cropped.box.h) / cropped.height,
  })
  return true
}
