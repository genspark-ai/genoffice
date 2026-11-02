// Reserved header/footer heights per section and variant: one probe path for
// the canvas slicer (single and multi-section) and the preview.
import type { HeaderFooter, HfImage, SectionInfo } from '@genoffice/docx-engine'

import { hfReservedHeightPx, hfStripGeom, type HfStripGeom } from './editor/hf-dom'
import { hfPartWidthPx, type HfVariant } from './pagination-hf'
import type { SectionHfHeights } from './pagination-types'

export type HfKind = 'header' | 'footer'

/** the strip a page lays out, plus the section it was picked from (Link to Previous) */
export interface HfStripSource {
  value: HeaderFooter | null
  images?: HfImage[]
  owner?: number
}

export function hfVariantHeightsPx(
  resolve: (kind: HfKind, variant: HfVariant) => HfStripSource,
  widthOf: (kind: HfKind, variant: HfVariant, owner?: number) => number,
  geom: HfStripGeom,
  firstOn: boolean,
  evenOn: boolean,
): SectionHfHeights {
  const px = (kind: HfKind, variant: HfVariant) => {
    const part = resolve(kind, variant)
    return hfReservedHeightPx(
      kind,
      part.value,
      widthOf(kind, variant, part.owner),
      part.images,
      geom,
    )
  }
  return {
    headerPx: px('header', 'default'),
    footerPx: px('footer', 'default'),
    ...(firstOn
      ? { firstHeaderPx: px('header', 'first'), firstFooterPx: px('footer', 'first') }
      : {}),
    ...(evenOn ? { evenHeaderPx: px('header', 'even'), evenFooterPx: px('footer', 'even') } : {}),
  }
}

/** per-section heights for sectionGeoms; a lone section follows the live titlePg toggle */
export function sectionHfHeights(
  secs: SectionInfo[],
  resolve: (sectionIndex: number, kind: HfKind, variant: HfVariant) => HfStripSource,
  opts: { titlePg: boolean; evenOddHf: boolean },
): SectionHfHeights[] {
  return secs.map((s, i) =>
    hfVariantHeightsPx(
      (kind, variant) => resolve(i, kind, variant),
      (kind, variant, owner) => hfPartWidthPx(secs, i, kind, variant, owner),
      hfStripGeom(s.settings),
      secs.length === 1 ? opts.titlePg : !!s.titlePg,
      opts.evenOddHf,
    ),
  )
}
