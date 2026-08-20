/**
 * Floating-image classification (P3): an image clearly detached from the text
 * flow — text running BESIDE it (wrapped / sitting in a column gutter) or ON
 * TOP of it (a backdrop) — becomes an anchored float; everything else keeps
 * P1's inline-image treatment. Never drops an image either way.
 */
import { intersectArea, overlapRatio, rectArea, verticalOverlapRatio } from '../geometry'
import type { Rect } from '../geometry'
import type { ImageBlock } from '../ir'
import type { LineUnit } from './units'

/** units beside the image (vertically overlapping, horizontally clear) needed to float */
const SIDE_UNITS_MIN = 2
/** units mostly inside the image box needed to float it as a backdrop */
const OVER_UNITS_MIN = 2
/** a unit counts as beside when its share inside the image is at most this */
const SIDE_X_TOL_PT = 2
/**
 * images at most this big (pt) are marker icons (list bullets, link favicons);
 * inside a text band they must not flow inline — each would push the whole
 * band down by its height and stack negative gaps (P5)
 */
const TINY_ICON_MAX_PT = 20
/**
 * corner badges (P9 C): an image at most this share of the page sharing its
 * vertical band with even ONE text unit (slide titles run alone on their
 * line) is furniture, not a figure — inline it would inject its full height
 * into the flow on every page and stack clamped negative gaps
 */
const BADGE_MAX_AREA_RATIO = 0.02
/**
 * an image mostly inside another image's box (P9 C) is an absolutely
 * positioned overlay (ring decor over a photo, badge on a banner) — flowing
 * both inline stacks their heights and inflates the page
 */
const IMAGE_OVERLAP_FLOAT_MIN = 0.6
/**
 * an image whose visible part covers this share of the page is a wallpaper
 * photo (P11 A) — always a behindDoc backdrop, whatever the text does; inline
 * it would consume a whole page of flow and push the body onto extra pages
 */
const PAGE_COVER_FLOAT_MIN = 0.9

// ── text-shadow glyph rasters (P11 B) ──
/** the image is only modestly larger than its line (blur padding) */
const SHADOW_MAX_AREA_RATIO = 3.5
/** the line spans most of the image width */
const SHADOW_MIN_WIDTH_SHARE = 0.65
/** centers agree horizontally (share of image width) … */
const SHADOW_MAX_CX_OFFSET = 0.15
/** … and vertically (share of image height) */
const SHADOW_MAX_CY_OFFSET = 0.3
/** only display text gets exported shadow rasters */
const SHADOW_MIN_FONT_PT = 18

/**
 * Drop glyph-raster shadow images (P11 B): slide exporters bake a display
 * line's soft shadow into a bitmap of the SAME text at the SAME spot, with
 * the real chars drawn on top. Rebuilt, the bitmap pins exactly while the
 * flow text lands a little off — the pair reads as a ghosted double strike.
 * The bitmap twin (not the text) is the disposable copy. Judged on the
 * display chars that actually sit INSIDE the image (one visual line can span
 * two side-by-side shadow rasters, so unit boxes are the wrong frame).
 */
export function suppressTextShadowImages(
  images: readonly ImageBlock[],
  units: readonly LineUnit[],
): ImageBlock[] {
  // hidden chars (w:vanish, e.g. a re-hidden OCR layer) paint nothing — an
  // image over them duplicates nothing and must stay (P29 B)
  const chars = units.flatMap((u) => u.chars).filter((c) => c.text.trim() !== '' && !c.invisible)
  return images.filter((img) => {
    const w = img.box.x1 - img.box.x0
    const h = img.box.y1 - img.box.y0
    if (w <= 0 || h <= 0) return true
    const inside = chars.filter((c) => {
      const cx = (c.box.x0 + c.box.x1) / 2
      const cy = (c.box.y0 + c.box.y1) / 2
      return cx >= img.box.x0 && cx <= img.box.x1 && cy >= img.box.y0 && cy <= img.box.y1
    })
    if (inside.length < 2) return true
    const union = {
      x0: Math.min(...inside.map((c) => c.box.x0)),
      y0: Math.min(...inside.map((c) => c.box.y0)),
      x1: Math.max(...inside.map((c) => c.box.x1)),
      y1: Math.max(...inside.map((c) => c.box.y1)),
    }
    if (rectArea(img.box) > rectArea(union) * SHADOW_MAX_AREA_RATIO) return true
    if (union.x1 - union.x0 < w * SHADOW_MIN_WIDTH_SHARE) return true
    if (
      Math.abs((union.x0 + union.x1) / 2 - (img.box.x0 + img.box.x1) / 2) >
      w * SHADOW_MAX_CX_OFFSET
    )
      return true
    if (
      Math.abs((union.y0 + union.y1) / 2 - (img.box.y0 + img.box.y1) / 2) >
      h * SHADOW_MAX_CY_OFFSET
    )
      return true
    const sizes = inside.map((c) => c.fontSize).sort((a, b) => a - b)
    return (sizes[sizes.length >> 1] ?? 0) < SHADOW_MIN_FONT_PT
  })
}

export interface ClassifiedImages {
  /** float field set; excluded from column detection and the flow chain */
  floats: ImageBlock[]
  inline: ImageBlock[]
}

/** invisible micro text (kept for fidelity, P11 C) is not wrap evidence */
/** an image fully inside this share of the page's top/bottom edge is decor (P16 F) */
const PAGE_EDGE_BAND_RATIO = 0.12
/** edge-band decor tolerates this many furniture units (page number, footer
 * line) overlapping or beside it — more text means real content (P20) */
const EDGE_BAND_FURNITURE_MAX = 2

const UNIT_MIN_FONT_PT = 2.5
const isVisibleUnit = (u: LineUnit): boolean => u.chars.some((c) => c.fontSize >= UNIT_MIN_FONT_PT)

/** Split page images into floats and inline; floats get their placement filled. */
export function classifyFloatImages(
  images: readonly ImageBlock[],
  allUnits: readonly LineUnit[],
  /** page area (pt²) — enables the corner-badge rule (P9 C) when provided */
  pageAreaPt2?: number,
  /** page rect — enables the wallpaper page-cover rule (P11 A) when provided */
  pageBox?: Rect,
): ClassifiedImages {
  const floats: ImageBlock[] = []
  const inline: ImageBlock[] = []
  const units = allUnits.filter(isVisibleUnit)
  const bodyLeft = units.length > 0 ? Math.min(...units.map((u) => u.box.x0)) : 0
  const bodyRight = units.length > 0 ? Math.max(...units.map((u) => u.box.x1)) : 0

  for (const img of images) {
    const over = units.filter((u) => overlapRatio(u.box, img.box) >= 0.8)
    const beside = units.filter(
      (u) =>
        verticalOverlapRatio(u.box, img.box) >= 0.5 &&
        (u.box.x1 <= img.box.x0 + SIDE_X_TOL_PT || u.box.x0 >= img.box.x1 - SIDE_X_TOL_PT),
    )

    if (
      pageBox !== undefined &&
      intersectArea(img.box, pageBox) >= rectArea(pageBox) * PAGE_COVER_FLOAT_MIN
    ) {
      // full-bleed wallpaper photo (single image or a clipped over-scan tile)
      img.float = { wrap: 'behind', xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
      floats.push(img)
    } else if (over.length >= OVER_UNITS_MIN) {
      img.float = { wrap: 'behind', xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
      floats.push(img)
    } else if (
      images.some(
        (other) => other !== img && overlapRatio(img.box, other.box) >= IMAGE_OVERLAP_FLOAT_MIN,
      )
    ) {
      // overlay stack (decor built from layered images) — before the square
      // rule (P11 D): square floats reserve column band space, and a stack of
      // five layered hexagons beside a text zone reserved five image heights
      img.float = { wrap: 'behind', xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
      floats.push(img)
    } else if (beside.length >= SIDE_UNITS_MIN) {
      // image on the side where the text is NOT
      const leftText = beside.filter((u) => u.box.x1 <= img.box.x0 + SIDE_X_TOL_PT).length
      const rightText = beside.length - leftText
      const imgCenter = (img.box.x0 + img.box.x1) / 2
      const wrap =
        leftText > rightText
          ? 'square-right'
          : rightText > leftText
            ? 'square-left'
            : imgCenter > (bodyLeft + bodyRight) / 2
              ? 'square-right'
              : 'square-left'
      img.float = { wrap, xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
      floats.push(img)
    } else if (
      beside.length >= 1 &&
      over.length === 0 &&
      pageAreaPt2 !== undefined &&
      rectArea(img.box) <= pageAreaPt2 * BADGE_MAX_AREA_RATIO
    ) {
      // corner badge / page logo sharing a title's band: behind-anchored so it
      // has zero effect on the text flow
      img.float = { wrap: 'behind', xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
      floats.push(img)
    } else if (
      pageBox !== undefined &&
      over.length <= EDGE_BAND_FURNITURE_MAX &&
      beside.length <= EDGE_BAND_FURNITURE_MAX &&
      (img.box.y1 <= pageBox.y0 + PAGE_EDGE_BAND_RATIO * (pageBox.y1 - pageBox.y0) ||
        img.box.y0 >= pageBox.y1 - PAGE_EDGE_BAND_RATIO * (pageBox.y1 - pageBox.y0))
    ) {
      // page-edge decor band (P16 F): footer bars / header ribbons live
      // entirely inside the page's top or bottom margin band — flowed inline
      // they ride wherever the body text ends. A page number / footer line
      // sharing the band is furniture, not wrap evidence (P20): a full-width
      // footer graphic contains the centred page number and stayed inline,
      // and being the page's leftmost block it dragged the derived margin to
      // the page edge, shifting every unindented paragraph left
      img.float = { wrap: 'behind', xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
      floats.push(img)
    } else if (
      img.box.x1 - img.box.x0 <= TINY_ICON_MAX_PT &&
      img.box.y1 - img.box.y0 <= TINY_ICON_MAX_PT &&
      units.length > 0 &&
      img.box.y1 <= Math.max(...units.map((u) => u.box.y1)) &&
      img.box.y0 >= Math.min(...units.map((u) => u.box.y0))
    ) {
      // marker icon (list bullet, link favicon, timeline dot) inside the
      // page's text span — anchor it behind the text
      img.float = { wrap: 'behind', xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
      floats.push(img)
    } else {
      inline.push(img)
    }
  }

  // gallery rows (P9 C): images laid out side by side share a vertical band —
  // flowing them inline stacks their heights (a 2×4 photo grid became 8 page
  // heights). Every row member pins behind at its measured position instead.
  const sideBySide = (a: ImageBlock, b: ImageBlock): boolean =>
    verticalOverlapRatio(a.box, b.box) >= 0.5 &&
    (a.box.x1 <= b.box.x0 + SIDE_X_TOL_PT || a.box.x0 >= b.box.x1 - SIDE_X_TOL_PT)
  const snapshot = [...inline]
  for (const img of snapshot) {
    if (!snapshot.some((other) => other !== img && sideBySide(img, other))) continue
    img.float = { wrap: 'behind', xOffsetPt: Math.max(0, img.box.x0 - bodyLeft) }
    floats.push(img)
    inline.splice(inline.indexOf(img), 1)
  }
  return { floats, inline }
}
