import { Fragment } from 'react'
import type { CSSProperties, ReactNode } from 'react'
import { pdfRectToCss, pdfToView } from './annotations'
import type { PageGeom } from './annotations'
import { colorSegments, decodeStyle, encodeStyle, runsToColors } from './color-runs'
import type { CharStyle } from './color-runs'
import type { DocFontStyle } from './doc-font'
import { EDIT_FONTS } from '../shared/ipc'
import type { TextEditInput, TextEditValidation, TextInsertInput } from '../shared/ipc'

export const EDIT_FONT_BY_ID = new Map<string, (typeof EDIT_FONTS)[number]>(
  EDIT_FONTS.map((f) => [f.id, f]),
)

let measureCtx: CanvasRenderingContext2D | null = null
/** Width of text in the given CSS font (shared hidden canvas) */
export function measureTextWidth(text: string, font: string): number {
  measureCtx ??= document.createElement('canvas').getContext('2d')
  if (!measureCtx) return 0
  measureCtx.font = font
  return measureCtx.measureText(text).width
}

export const rgbToHex = (c: readonly [number, number, number]): string =>
  `#${c
    .map((v) =>
      Math.round(v * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`
export const hexToRgb = (hex: string): [number, number, number] => [
  parseInt(hex.slice(1, 3), 16) / 255,
  parseInt(hex.slice(3, 5), 16) / 255,
  parseInt(hex.slice(5, 7), 16) / 255,
]

/** Text-edit colors travel as 0-255 RGB (PDFium fill color), unlike markups' 0-1 floats */
export const hexTo255 = (hex: string): [number, number, number] => [
  parseInt(hex.slice(1, 3), 16),
  parseInt(hex.slice(3, 5), 16),
  parseInt(hex.slice(5, 7), 16),
]
export const rgb255ToHex = (c: readonly [number, number, number]): string =>
  `#${c.map((v) => v.toString(16).padStart(2, '0')).join('')}`

/** Committed IPC style runs → encoded-key runs over newText (the draft/preview form) */
export const styleRunsToKeyRuns = (
  runs: NonNullable<TextEditInput['styleRuns']>,
): { start: number; end: number; color: string }[] =>
  runs.map((r) => ({
    start: r.start,
    end: r.end,
    color: encodeStyle({
      color: r.color ? rgb255ToHex(r.color) : undefined,
      font: r.font,
      size: r.size,
      bold: r.bold,
      italic: r.italic,
    }),
  }))

/** CSS of one styled segment in the editor mirror / pending preview. Explicit on/off
    overrides the inherited draft-level weight/slant; size scales like the host text. */
export const styleSegCss = (s: CharStyle, scale: number): CSSProperties => ({
  ...(s.color ? { color: s.color } : {}),
  ...(s.font ? { fontFamily: EDIT_FONT_BY_ID.get(s.font)?.css } : {}),
  ...(s.size !== undefined ? { fontSize: s.size * scale * 0.92 } : {}),
  ...(s.bold !== undefined ? { fontWeight: s.bold ? 700 : 400 } : {}),
  ...(s.italic !== undefined ? { fontStyle: s.italic ? 'italic' : 'normal' } : {}),
})

/** Encoded-key runs → the IPC styleRuns the engine consumes */
export const keyRunsToStyleRuns = (
  runs: { start: number; end: number; color: string }[],
): NonNullable<TextEditInput['styleRuns']> =>
  runs.map((r) => {
    const s = decodeStyle(r.color)
    return {
      start: r.start,
      end: r.end,
      color: s.color ? hexTo255(s.color) : undefined,
      font: s.font,
      size: s.size,
      bold: s.bold,
      italic: s.italic,
    }
  })

export interface LocalTextEdit {
  id: string
  input: TextEditInput
  /** Matched run's ink bounds from validation (PDF user space). The edit rect is a pdf.js
      layout box; glyph ink can poke out of it, so the preview covers this instead */
  cover?: [number, number, number, number]
  /** The run's base ink (display-only, from the draft's probe): the pending preview
      shows the document's real color when the edit doesn't change it */
  baseInk?: string
  /** Local look-alike of the run's own font (display-only, from the PostScript
      name): the pending preview reads like the document. Never sent to the engine. */
  baseFont?: DocFontStyle
  /** Accumulated block-move delta (PDF user space). Renderer metadata only: the preview
      and hover box draw at rect + moveBy while input.rect stays at the original position
      (it is the save-time match key). The engine-side position rides in input.translate
      (pure moves) or the shifted input.origin (moved rebuild edits). */
  moveBy?: [number, number]
}

/** Stable per-render key for a clustered block's rect (same idea as imageRectKey) */
export const blockRectKey = (r: readonly number[]): string => r.map((v) => v.toFixed(2)).join(',')

export const shiftRect = (
  r: readonly [number, number, number, number],
  d: readonly [number, number],
): [number, number, number, number] => [r[0] + d[0], r[1] + d[1], r[2] + d[0], r[3] + d[1]]

export interface LocalTextInsert {
  id: string
  input: TextInsertInput
}

/** Area a pending edit must blank: the edit rect grown to the validated ink bounds */
export const unionCover = (
  rect: readonly [number, number, number, number],
  cover: readonly [number, number, number, number] | undefined,
): [number, number, number, number] =>
  cover
    ? [
        Math.min(rect[0], cover[0]),
        Math.min(rect[1], cover[1]),
        Math.max(rect[2], cover[2]),
        Math.max(rect[3], cover[3]),
      ]
    : [rect[0], rect[1], rect[2], rect[3]]

/** Expand a CSS box by p px on every side (antialiasing bleeds past exact ink bounds) */
export const inflateCss = (
  b: { left: number; top: number; width: number; height: number },
  p: number,
) => ({
  left: b.left - p,
  top: b.top - p,
  width: b.width + 2 * p,
  height: b.height + 2 * p,
})

/** Pending text-insert preview style; shared by the canvas overlay and the thumbnail mirror */
export const textInsertPreviewStyle = (
  insert: LocalTextInsert,
  geom: PageGeom,
  scale: number,
): CSSProperties => {
  const [vx, vy] = pdfToView(geom, insert.input.origin[0], insert.input.origin[1])
  const align = insert.input.align ?? 'left'
  const style: CSSProperties = {
    left: vx * scale,
    top: (vy - insert.input.fontSize) * scale,
    fontSize: insert.input.fontSize * scale * 0.92,
    lineHeight: insert.input.lineLeading ? `${insert.input.lineLeading * scale}px` : 1.2,
    color: `rgb(${insert.input.color.join(', ')})`,
    whiteSpace: 'pre',
    transform:
      align === 'center' ? 'translateX(-50%)' : align === 'right' ? 'translateX(-100%)' : undefined,
    textAlign: align,
  }
  if (insert.input.font) style.fontFamily = EDIT_FONT_BY_ID.get(insert.input.font)?.css
  if (insert.input.bold) style.fontWeight = 700
  if (insert.input.italic) style.fontStyle = 'italic'
  return style
}

/** Pending text-edit preview style + original-run cover; shared with the thumbnail mirror */
export const textEditPreviewParts = (
  te: LocalTextEdit,
  geom: PageGeom,
  scale: number,
): { style: CSSProperties; coverStyle: CSSProperties | null } => {
  const fs = (te.input.newFontSize ?? te.input.fontSize) * scale * 0.92
  const lineCount = te.input.newText.split('\n').length
  const leadPx = te.input.lineLeading ? te.input.lineLeading * scale : fs * 1.2
  const style: CSSProperties = {
    // A moved block previews at its new position; input.rect stays at the original
    // (it is the save-time match key, and the cover below must hide the original ink)
    ...pdfRectToCss(geom, te.moveBy ? shiftRect(te.input.rect, te.moveBy) : te.input.rect, scale),
    fontSize: fs,
    ...(te.input.lineLeading ? { lineHeight: `${leadPx}px` } : {}),
  }
  if (te.input.newColor) {
    style.color = `rgb(${te.input.newColor.join(', ')})`
  } else if (te.baseInk) {
    style.color = te.baseInk
  }
  const baseF = te.input.newFont ? undefined : te.baseFont
  if (te.input.newFont) {
    style.fontFamily = EDIT_FONT_BY_ID.get(te.input.newFont)?.css
  } else if (baseF) {
    // Look-alike of the document's own face
    style.fontFamily = baseF.css
  }
  if (te.input.newBold) style.fontWeight = 700
  else if (baseF?.weight) style.fontWeight = baseF.weight
  if (te.input.newItalic) style.fontStyle = 'italic'
  else if (baseF?.italic) style.fontStyle = 'italic'
  if (lineCount > 1) {
    // Grow below the original rect, same leading the engine writes
    // (block edits carry the paragraph's own leading)
    style.height = lineCount * leadPx
    style.lineHeight = te.input.lineLeading ? `${leadPx}px` : 1.2
    style.alignItems = 'flex-start'
  } else if (typeof style.height === 'number' && fs + 2 > style.height) {
    // Enlarged single line: the engine keeps the run's baseline, so glyphs grow
    // upward past the original rect — extend the box up and bottom-align, or the
    // preview clips the taller glyphs (overflow: hidden)
    const grown = Math.ceil(fs) + 2
    style.top = (style.top as number) - (grown - style.height)
    style.height = grown
    style.alignItems = 'flex-end'
  }
  // The rebuilt run grows right past the original rect when the replacement is
  // longer; the preview must too, or the extra characters look cut off until
  // the save (overflow: hidden)
  const previewFont = `${te.input.newItalic || baseF?.italic ? 'italic ' : ''}${
    te.input.newBold ? 'bold ' : baseF?.weight ? `${baseF.weight} ` : ''
  }${fs}px ${
    (te.input.newFont && EDIT_FONT_BY_ID.get(te.input.newFont)?.css) ||
    baseF?.css ||
    getComputedStyle(document.body).fontFamily
  }`
  const widest = Math.max(
    ...te.input.newText.split('\n').map((l) => measureTextWidth(l, previewFont)),
  )
  if (typeof style.width === 'number' && widest > style.width) {
    style.width = widest + 2
  }
  if (te.input.align) {
    // The preview is a flex container and its text is one shrink-to-fit anonymous
    // item: textAlign only aligns lines within that item, justifyContent moves
    // the item itself off main-start
    style.textAlign = te.input.align
    if (te.input.align === 'center') style.justifyContent = 'center'
    if (te.input.align === 'right') style.justifyContent = 'flex-end'
  }
  const coverStyle = te.cover
    ? inflateCss(pdfRectToCss(geom, unionCover(te.input.rect, te.cover), scale), 1.5)
    : null
  return { style, coverStyle }
}

/** Styled-run children of a text-edit preview; shared with the thumbnail mirror */
export const textEditPreviewContent = (te: LocalTextEdit, scale: number): ReactNode =>
  (te.input.styleRuns ?? te.input.colorRuns)?.length ? (
    // One wrapper span = one flex item: the preview is a row flex container, and
    // bare segments would become separate items laid out horizontally, breaking
    // '\n' stacking in multi-line previews
    <span>
      {colorSegments(
        te.input.newText,
        runsToColors(
          te.input.newText.length,
          styleRunsToKeyRuns(te.input.styleRuns ?? te.input.colorRuns ?? []),
        ),
      ).map((seg, i) => {
        if (!seg.color) return <Fragment key={i}>{seg.text}</Fragment>
        const s = decodeStyle(seg.color)
        return (
          <span key={i} style={styleSegCss(s, scale)}>
            {seg.text}
          </span>
        )
      })}
    </span>
  ) : (
    te.input.newText
  )

/** Editor state for the floating text-edit box; editId set when re-opening a pending edit */
export interface TextDraft {
  origIdx: number
  rect: [number, number, number, number]
  oldText: string
  fontSize: number
  value: string
  /** Style overrides; undefined = keep the run's original size/color */
  size?: number
  /** CSS hex like '#d32f2f' */
  color?: string
  /** Selection-level styles, one encoded key per code unit of value (see encodeStyle);
      '' = base (the draft-level overrides ?? original). undefined/all-'' = uniform
      draft (the pre-existing whole-run behavior). */
  charStyles?: string[]
  /** Colors the document already draws the run with (async, from the open probe),
      pre-seeded into charStyles as color-only keys. A commit whose styles still equal
      these carries no *change* — they only ride along so a rebuild repaints them. */
  seedStyleRuns?: { start: number; end: number; color: string }[]
  /** The run's base ink in the document (async, from the open probe). Display-only:
      the editor/preview text shows the real color; never committed as a change. */
  seedInk?: string
  /** Local look-alike of the run's own font (async, from the dominant run's
      PostScript name). Display + reflow measurement only; never committed. */
  seedFont?: DocFontStyle
  /** EDIT_FONTS id; undefined = automatic rebuild font */
  font?: string
  /** Style toggles; true = on, undefined = off (resolved via font variants at save) */
  bold?: true
  italic?: true
  editId?: string
  /** Further pending edits folded into this block draft (besides editId); the
      commit replaces editId and removes these — they would overlap the block
      edit at save otherwise */
  foldedIds?: string[]
  /** The value the editor opened with when pending edits were folded in: an
      unmodified commit must keep those edits instead of converting them */
  foldBase?: string
  /** charStyles the fold seeded (styles of the folded line edits); an unmodified
      commit compares against these, not against empty */
  foldStyles?: string[]
  /** Ink bounds of the run being edited (async, from a dry-run validate) */
  cover?: [number, number, number, number]
  /** Present for paragraph (block) edits: the geometry the commit reflows into.
      lineHeight is the block's original leading at the original font size;
      bottomPt is the block's bottom edge in firstBaseline's (possibly moved)
      frame — the overflow guard measures growth against it. */
  block?: {
    leftPt: number
    firstBaseline: number
    widthPt: number
    lineHeight: number
    align: 'left' | 'center' | 'right'
    bottomPt: number
  }
  /** Carried from a reopened moved edit: the floating editor draws at rect + moveBy
      (where the preview sits) while rect itself stays the save-time match key */
  moveBy?: [number, number]
}

/** Seed a fresh draft's colors from the open probe's report of what the document
    already draws: the run's base ink (display-only) plus earlier saved selection
    colors. Selection colors only while the draft is pristine — the runs are offsets
    into oldText, and once typing starts they no longer align (the engine-side
    rebuild still preserves the colors on save). */
export const seedDraftColors = (d: TextDraft, v: TextEditValidation): TextDraft => {
  let next = d
  // Near-white ink would vanish on the editor's white background; keep default ink
  if (v.baseColor && !next.color && !next.seedInk) {
    const [r, g, b] = v.baseColor
    if ((0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 <= 0.85)
      next = { ...next, seedInk: rgb255ToHex(v.baseColor) }
  }
  if (!v.colorRuns || v.colorRuns.length === 0) return next
  if (next.charStyles || next.seedStyleRuns || next.value !== next.oldText) return next
  const keyRuns = v.colorRuns.map((r) => ({
    start: r.start,
    end: r.end,
    color: encodeStyle({ color: rgb255ToHex(r.color) }),
  }))
  return {
    ...next,
    charStyles: runsToColors(next.oldText.length, keyRuns),
    seedStyleRuns: keyRuns,
  }
}
