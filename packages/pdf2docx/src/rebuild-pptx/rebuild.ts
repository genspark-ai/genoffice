/**
 * IR → pptx deck assembly (P25). Every PDF page becomes one slide of the same
 * size; every block lands at its measured coordinates — the geometry IS the
 * layout (the docx canvas path's lesson, without the flow compromises).
 *
 * spTree paint order (z): background color plate → full-page background
 * render → background panels (by source z) → decor shadings → behind-floats →
 * reading-order blocks (inline images / text boxes / tables) → footnotes.
 */
import { rectHeight, rectWidth, type Rect } from '../geometry'
import type { ImageBlock, IrPage, TextBlock } from '../ir'
import {
  addElement,
  addPicture,
  appendRawElements,
  buildTableGridXml,
  createBlankPptx,
  insertBlankSlide,
  openPptx,
  savePptx,
  setSlideSize,
  type OpenedPptx,
  type Slide,
} from '../../../pptx-engine/src/index'
import { tableGridOptions, type PageMapper } from './table'
import { textBlockParagraph } from './text'

const EMU_PER_PT = 12700

/**
 * Wrap headroom on text-box width (P19 frame lesson): substituted fonts
 * render a hair wider, and one wrapped line breaks the measured rhythm. The
 * pad extends AWAY from the anchored edge so visible ink never drifts.
 */
const WRAP_PAD_RATIO = 0.04
const WRAP_PAD_MIN_PT = 4

/** default fallback slide size (pt): 4:3 letter-ish, only for empty PDFs */
const FALLBACK_W_PT = 720
const FALLBACK_H_PT = 540

/** uniform page→slide mapper; non-uniform page sizes letterbox-center */
function pageMapper(page: IrPage, deckWPt: number, deckHPt: number): PageMapper {
  const s = Math.min(deckWPt / Math.max(1, page.widthPt), deckHPt / Math.max(1, page.heightPt))
  const ox = (deckWPt - page.widthPt * s) / 2
  const oy = (deckHPt - page.heightPt * s) / 2
  return {
    scale: s,
    len: (pt) => Math.round(pt * s * EMU_PER_PT),
    rect: (box: Rect) => ({
      x: Math.round((ox + box.x0 * s) * EMU_PER_PT),
      y: Math.round((oy + (page.heightPt - box.y1) * s) * EMU_PER_PT),
      cx: Math.max(1, Math.round(rectWidth(box) * s * EMU_PER_PT)),
      cy: Math.max(1, Math.round(rectHeight(box) * s * EMU_PER_PT)),
    }),
  }
}

const extOf = (mime: string): string => (mime === 'image/jpeg' ? 'jpg' : 'png')

function addImageAt(
  opened: OpenedPptx,
  slide: Slide,
  img: { data: Uint8Array; mime: string },
  offset: { x: number; y: number; cx: number; cy: number },
): void {
  addPicture(opened, slide, { bytes: img.data, ext: extOf(img.mime), offset })
}

/** full-page raster (scanned/degraded pages, background renders) */
function addFullPageImage(
  opened: OpenedPptx,
  slide: Slide,
  img: { data: Uint8Array; mime: string },
  deckWPt: number,
  deckHPt: number,
): void {
  addImageAt(opened, slide, img, {
    x: 0,
    y: 0,
    cx: Math.round(deckWPt * EMU_PER_PT),
    cy: Math.round(deckHPt * EMU_PER_PT),
  })
}

function addTextBlock(slide: Slide, block: TextBlock, m: PageMapper): void {
  if (block.lines.length > 0) {
    const widthPt = Math.max(1, rectWidth(block.box)) * m.scale
    const padPt = Math.max(WRAP_PAD_MIN_PT, widthPt * WRAP_PAD_RATIO)
    const offset = m.rect(block.box)
    const padEmu = Math.round(padPt * EMU_PER_PT)
    if (block.align === 'center') offset.x -= Math.round(padEmu / 2)
    else if (block.align === 'right') offset.x -= padEmu
    offset.cx += padEmu
    addElement(slide, {
      kind: 'textbox',
      offset,
      paragraphs: [textBlockParagraph(block, m.scale)],
      bodyPr: { wrap: 'square', anchor: 't', insetsEmu: { l: 0, t: 0, r: 0, b: 0 } },
    })
  }
  if (block.border) addDecorBorder(slide, block, m)
}

/** decorative rule attached to a paragraph (P7 pBdr) → a thin filled rect */
function addDecorBorder(slide: Slide, block: TextBlock, m: PageMapper): void {
  const b = block.border!
  const box = block.box
  let rect: Rect
  if (b.side === 'left') {
    const x1 = box.x0 - b.spacePt
    rect = { x0: x1 - b.widthPt, x1, y0: box.y0, y1: box.y1 }
  } else {
    const x0 = box.x0 + (b.indentLeftPt ?? 0)
    const x1 = box.x1 - (b.indentRightPt ?? 0)
    const y = b.side === 'top' ? box.y1 + b.spacePt : box.y0 - b.spacePt
    rect = { x0, x1: Math.max(x0 + 1, x1), y0: y - b.widthPt / 2, y1: y + b.widthPt / 2 }
  }
  addElement(slide, { kind: 'rect', offset: m.rect(rect), fillColor: `#${b.color}` })
}

function emitPage(
  opened: OpenedPptx,
  slideIndex: number,
  page: IrPage,
  deckWPt: number,
  deckHPt: number,
): void {
  // appendRawElements reparses the slide model, so never cache the reference
  const slideAt = (): Slide => opened.deck.slides[slideIndex]!
  const m = pageMapper(page, deckWPt, deckHPt)

  // scanned/degraded page: its bitmap is the whole slide
  if (page.scanned || page.degraded) {
    if (page.render) addFullPageImage(opened, slideAt(), page.render, deckWPt, deckHPt)
    return
  }

  if (page.bgColor) {
    addElement(slideAt(), {
      kind: 'rect',
      offset: {
        x: 0,
        y: 0,
        cx: Math.round(deckWPt * EMU_PER_PT),
        cy: Math.round(deckHPt * EMU_PER_PT),
      },
      fillColor: `#${page.bgColor}`,
    })
  }
  if (page.bgRender) addFullPageImage(opened, slideAt(), page.bgRender, deckWPt, deckHPt)

  // one underlay pool in source paint order (z): background panels, decor
  // shadings, behind-floats, and the leftover fill slabs the flow rebuild
  // drops (dark-text card plates like deck-food's red banner live ONLY here —
  // the light-text backdrop pass never lifts them into bgPanels)
  const isBehind = (b: ImageBlock): boolean => b.float?.wrap === 'behind'
  type Underlay = { z: number; ord: number; paint: () => void }
  const underlays: Underlay[] = []
  let ord = 0
  for (const panel of page.bgPanels ?? []) {
    underlays.push({
      z: panel.z ?? 0,
      ord: ord++,
      paint: () => addImageAt(opened, slideAt(), panel, m.rect(panel.box)),
    })
  }
  for (const decor of page.decorImages ?? []) {
    underlays.push({
      z: decor.z ?? 0,
      ord: ord++,
      paint: () => addImageAt(opened, slideAt(), decor, m.rect(decor.box)),
    })
  }
  for (const img of page.blocks.filter((b): b is ImageBlock => b.kind === 'image' && isBehind(b))) {
    underlays.push({
      z: img.z ?? 0,
      ord: ord++,
      paint: () => addImageAt(opened, slideAt(), img, m.rect(img.box)),
    })
  }
  for (const fill of page.shapes?.fills ?? []) {
    // real source ink the flow path leaves behind; translucent washes keep
    // their alpha (an opaque slab where the source painted a scrim reads as
    // a black bar). Consumed fills (cell shading/highlights) repaint the
    // same color under their construct — harmless by construction.
    const alphaHex =
      fill.alpha !== undefined ? Math.round(fill.alpha).toString(16).padStart(2, '0') : ''
    underlays.push({
      z: fill.z ?? 0,
      ord: ord++,
      paint: () =>
        addElement(slideAt(), {
          kind: 'rect',
          offset: m.rect(fill.box),
          fillColor: `#${fill.color}${alphaHex}`,
        }),
    })
  }
  underlays.sort((a, b) => a.z - b.z || a.ord - b.ord)
  for (const u of underlays) u.paint()

  for (const block of page.blocks) {
    if (block.kind === 'image') {
      if (!isBehind(block)) addImageAt(opened, slideAt(), block, m.rect(block.box))
    } else if (block.kind === 'table') {
      const xml = buildTableGridXml(slideAt(), tableGridOptions(block, m))
      appendRawElements(opened, slideIndex, [xml])
    } else {
      addTextBlock(slideAt(), block, m)
    }
  }

  for (const note of page.footnotes ?? []) {
    for (const block of note.blocks) addTextBlock(slideAt(), block, m)
  }
}

/** Rebuild the analyzed pages as a pptx deck (one slide per page). */
export async function rebuildPptx(pages: IrPage[]): Promise<Uint8Array> {
  const opened = await openPptx(await createBlankPptx())
  const deckWPt = pages[0]?.widthPt ?? FALLBACK_W_PT
  const deckHPt = pages[0]?.heightPt ?? FALLBACK_H_PT
  setSlideSize(opened, Math.round(deckWPt * EMU_PER_PT), Math.round(deckHPt * EMU_PER_PT))
  // the blank deck ships slide 1; every further page appends after the last
  for (let i = 1; i < pages.length; i++) insertBlankSlide(opened, i - 1)
  for (const [i, page] of pages.entries()) emitPage(opened, i, page, deckWPt, deckHPt)
  return savePptx(opened)
}
