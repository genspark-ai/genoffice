// Section geometry: page box, margins, columns, doc-grid pitch and the
// per-block column / width / vertical-alignment specs derived from sections.
import type { SectionInfo, SectionSettings } from '@genoffice/docx-engine'

import type {
  BlockBox,
  ColumnBlockPlacement,
  PageSlice,
  SectionGeom,
  SectionHfHeights,
} from './pagination-types'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** Physical section page box shared by editor and pagination preview. */
export function sectionPageBox(set: SectionSettings): {
  width: number
  height: number
  contentWidth: number
  headerDist: number
  footerDist: number
} {
  return {
    width: twipsToPx(set.pageWidth),
    height: twipsToPx(set.pageHeight),
    contentWidth: twipsToPx(set.pageWidth - set.marginLeft - set.marginRight),
    headerDist: twipsToPx(set.headerDist ?? 720),
    footerDist: twipsToPx(set.footerDist ?? 720),
  }
}

/** Body top = max(marginTop, headerDist + header height) */
export function effectiveTopPx(set: SectionSettings, headerPx: number): number {
  const dist = twipsToPx(set.headerDist ?? 720)
  return Math.max(twipsToPx(set.marginTop), headerPx > 0 ? dist + headerPx : 0)
}

/** Body bottom margin = max(marginBottom, footerDist + footer height) */
export function effectiveBottomPx(set: SectionSettings, footerPx: number): number {
  const dist = twipsToPx(set.footerDist ?? 720)
  return Math.max(twipsToPx(set.marginBottom), footerPx > 0 ? dist + footerPx : 0)
}

/** Typed line-grid pitch (pt) of one section (w:docGrid lines/linesAndChars), else null */
export function sectionGridPitchPt(s: SectionInfo): number | null {
  const g = s.settings.docGrid
  if (!g || (g.type !== 'lines' && g.type !== 'linesAndChars') || !g.linePitch) return null
  return g.linePitch / 20
}

/**
 * Uniform typed line grid (w:docGrid type lines/linesAndChars): the pitch in
 * points when EVERY section declares the same typed pitch, else null. Mixed
 * docs deliver per-section pitches through sectionGridPitchSpecs instead.
 * The value feeds .doc-page { --doc-grid-pitch } which line-height round(up)
 * expressions consume.
 */
export function docGridPitchPt(sections: SectionInfo[]): number | null {
  if (sections.length === 0) return null
  let pitch: number | null = null
  for (const s of sections) {
    const p = sectionGridPitchPt(s)
    if (p === null) return null
    if (pitch === null) pitch = p
    else if (pitch !== p) return null
  }
  return pitch
}

/**
 * Character-grid letter-spacing delta (pt) of one section. Word probes
 * (2026-09-02, MS Mincho/Arial, sizes 10.5/12pt): with w:docGrid
 * type="linesAndChars" and w:charSpace=N, EVERY character in the section
 * (east-asian, latin, halfwidth) advances by its natural width + N/4096 pt —
 * a uniform letter-spacing, positive or negative. Types lines/default ignore
 * charSpace; snapToChars snaps to a Normal-style-based pitch instead (not
 * modeled). Table cells and snapToGrid=0 paragraphs are spaced too; headers
 * and footers are not.
 */
export function sectionCharSpacePt(s: SectionInfo): number {
  const g = s.settings.docGrid
  if (!g || g.type !== 'linesAndChars' || !g.charSpace) return 0
  return g.charSpace / 4096
}

/**
 * Uniform character-grid delta (pt) when every section declares the same
 * nonzero value, else null. Feeds .doc-page { --doc-char-space }; mixed docs
 * deliver per-section values through sectionCharSpaceSpecs instead.
 */
export function docCharSpacePt(sections: SectionInfo[]): number | null {
  if (sections.length === 0) return null
  const v = sectionCharSpacePt(sections[0])
  for (const s of sections) if (sectionCharSpacePt(s) !== v) return null
  return v !== 0 ? v : null
}

/** Column count of a section (w:cols w:num; covers equal and explicit-width columns) */
export function sectionColumns(s: SectionInfo): number {
  return Math.max(1, s.settings.columns ?? 1)
}

/** Section column geometry (px). Equal-width columns divide evenly per w:cols
 *  w:space (default 720 twips); w:equalWidth="0" reads the explicit w:col
 *  width/space list (falling back to even division when the list is absent). */
export function sectionColGeom(s: SectionInfo): {
  cols: number
  /** first column's width — the uniform width when equalWidth */
  colWidthPx: number
  gapPx: number
  equalWidth: boolean
  /** per-column widths (length cols) */
  widths: number[]
  /** gap after column k (length cols-1) */
  gaps: number[]
} {
  const set = s.settings
  const contentW = twipsToPx(set.pageWidth - set.marginLeft - set.marginRight)
  const cols = sectionColumns(s)
  const gapPx = twipsToPx(set.colSpace ?? 720)
  const even = cols > 1 ? (contentW - gapPx * (cols - 1)) / cols : contentW
  let equalWidth = true
  let widths = Array.from({ length: cols }, () => even)
  let gaps = Array.from({ length: Math.max(cols - 1, 0) }, () => gapPx)
  if (cols > 1 && /<w:cols[^>]*w:equalWidth="0"/.test(s.sectPrXml ?? '')) {
    const colsXml = /<w:cols\b[^>]*>([\s\S]*?)<\/w:cols>/.exec(s.sectPrXml ?? '')?.[1] ?? ''
    const list = Array.from(colsXml.matchAll(/<w:col\b([^>]*)\/?>/g), (m) => ({
      w: twipsToPx(parseInt(/w:w="(\d+)"/.exec(m[1])?.[1] ?? '0', 10)),
      space: twipsToPx(parseInt(/w:space="(\d+)"/.exec(m[1])?.[1] ?? '0', 10)),
    }))
    if (list.length === cols && list.every((c) => c.w > 0)) {
      equalWidth = false
      widths = list.map((c) => c.w)
      gaps = list.slice(0, -1).map((c) => c.space)
    }
  }
  return { cols, gapPx, colWidthPx: widths[0] ?? contentW, equalWidth, widths, gaps }
}

/** RTL section (sectPr w:bidi): columns fill right-to-left (visual order only; engine indices stay logical) */
export function sectionBidi(s: SectionInfo): boolean {
  return /<w:bidi(?:\s*\/>|\s+w:val="(?:1|true|on)")/.test(s.sectPrXml ?? '')
}

/**
 * Mixed-column canvas placements: for every block on a regioned page, the
 * column width plus a per-column constant translate mapping its stacked
 * single-flow position into the column slot. dy = region top − the column
 * start's offset from the page start (negative for later columns/regions:
 * they pull up over the vacated stacked space). Blocks are placed whole by
 * their top; floated/eless blocks are skipped.
 */
export function columnLayoutSpecs(
  blocks: BlockBox[],
  slices: PageSlice[],
  sections: SectionInfo[],
): ColumnBlockPlacement[] {
  const specs: ColumnBlockPlacement[] = []
  if (blocks.length === 0) return specs
  let bi = 0
  for (const slice of slices) {
    if (!slice.regions) {
      while (bi < blocks.length && blocks[bi].top < slice.end - 0.5) bi++
      continue
    }
    for (const region of slice.regions) {
      const sec = sections[Math.max(0, Math.min(region.section, sections.length - 1))]
      if (!sec) continue
      const geom = sectionColGeom(sec)
      const rtl = geom.cols > 1 && sectionBidi(sec)
      // left edge of each column (LTR): cumulative widths + gaps
      const xs: number[] = []
      let x = 0
      for (let c = 0; c < geom.cols; c++) {
        xs.push(x)
        x += geom.widths[c] + (geom.gaps[c] ?? 0)
      }
      const totalW = x
      for (let c = 0; c < region.columns.length; c++) {
        const col = region.columns[c]
        const w = geom.widths[Math.min(c, geom.widths.length - 1)]
        const dx = geom.cols > 1 ? (rtl ? totalW - (xs[c] ?? 0) - w : (xs[c] ?? 0)) : 0
        const dy = region.top - (col.start - slice.start)
        const widthPx = geom.cols > 1 ? w : undefined
        if (widthPx === undefined && Math.abs(dx) < 0.01 && Math.abs(dy) < 0.01) {
          // untouched single-column region at its natural place: no decorations
          while (bi < blocks.length && blocks[bi].top < col.end - 0.5) bi++
          continue
        }
        while (bi < blocks.length && blocks[bi].top < col.end - 0.5) {
          const b = blocks[bi]
          bi++
          if (!b.el || b.floated) continue
          if (b.top < col.start - 0.5) continue
          specs.push({ el: b.el, ...(widthPx !== undefined ? { widthPx } : {}), dx, dy })
        }
      }
    }
  }
  return specs
}

const isTableBlock = (el: HTMLElement) =>
  el.tagName === 'TABLE' || el.getAttribute('data-doc-protected') === 'table'

/**
 * Per-block wrap widths and horizontal placement for documents whose sections
 * disagree on content width or side margins, e.g. a landscape section in a
 * portrait document, or a full-bleed cover section (w:pgMar left="0") ahead of
 * body sections with real margins. The canvas lays the whole flow on one page
 * width and pads it by the FIRST section's margins (--page-pad), so a section
 * that disagrees needs both its own wrap width and an x offset back onto its
 * own left margin — otherwise every later page keeps the cover's margin-less
 * text column. Every block gets an explicit width (not just the differing
 * sections'): preview clones render into per-section wrap widths, so any
 * container-relative block would reflow there. Tables keep their inline min()
 * width and get the section geometry via --doc-content-w / --doc-margin-* instead.
 */
export function sectionWidthSpecs(
  blocks: BlockBox[],
  sections: SectionInfo[],
  geoms: SectionGeom[],
): ColumnBlockPlacement[] {
  const canvasW = geoms[0]?.contentWidth
  // the canvas pads by sections[0] (App.tsx's canvasSection): placement offsets
  // are relative to that inset, and the first section sits on it unshifted
  const canvasInsetLeftPx = sections[0]?.settings ? twipsToPx(sections[0].settings.marginLeft) : 0
  const differsAt = (i: number): boolean => {
    const g = geoms[i]
    const set = sections[i]?.settings
    if (!g || !set) return false
    const widthDiffers =
      g.contentWidth !== undefined && Math.abs(g.contentWidth - (canvasW ?? 0)) > 0.5
    const insetDiffers = Math.abs(twipsToPx(set.marginLeft) - canvasInsetLeftPx) > 0.5
    return widthDiffers || insetDiffers
  }
  if (canvasW === undefined || !geoms.some((_, i) => differsAt(i))) return []
  const specs: ColumnBlockPlacement[] = []
  for (const b of blocks) {
    if (!b.el || b.floated) continue
    const si = Math.max(0, Math.min(b.section ?? 0, geoms.length - 1))
    const w = geoms[si]?.contentWidth
    if (w === undefined) continue
    const set = sections[si]?.settings
    const spec: ColumnBlockPlacement = {
      el: b.el,
      // shift the block from the canvas' (first section's) text column onto its
      // own section's left margin; dx composes with column/vAlign translates
      dx: (set ? twipsToPx(set.marginLeft) : 0) - canvasInsetLeftPx,
      dy: 0,
      contentWPx: w,
      marginLeftPx: set ? twipsToPx(set.marginLeft) : 0,
      marginRightPx: set ? twipsToPx(set.marginRight) : 0,
    }
    // blocks with an own inline width (tables, textboxes) size themselves;
    // paragraphs get the section width minus their indent margins
    if (!isTableBlock(b.el) && !b.el.style.width) {
      const cs = getComputedStyle(b.el)
      spec.widthPx = w - (parseFloat(cs.marginLeft) || 0) - (parseFloat(cs.marginRight) || 0)
    }
    specs.push(spec)
  }
  return specs
}

/**
 * Per-block docGrid pitch for documents whose sections disagree on the typed
 * line grid (mixed pitches, or typed + untyped): each block carries its own
 * section's pitch through the layout decoration channel. Uniform docs get no
 * specs (the single .doc-page injection stays). Blocks of untyped sections opt
 * out like snapToGrid=0 (via the channel's own doc-grid-nosnap class, keeping
 * auto line multiples alive); blocks carrying their own doc-nosnap keep their
 * inline opt-out untouched (the classes are distinct, so the skip is stable
 * across remeasure passes).
 */
export function sectionGridPitchSpecs(
  blocks: BlockBox[],
  sections: SectionInfo[],
): ColumnBlockPlacement[] {
  const pitches = sections.map(sectionGridPitchPt)
  const typed = pitches.filter((p): p is number => p != null)
  if (
    typed.length === 0 ||
    (typed.length === pitches.length && typed.every((p) => p === typed[0]))
  ) {
    return []
  }
  const specs: ColumnBlockPlacement[] = []
  for (const b of blocks) {
    if (!b.el || b.el.classList.contains('doc-nosnap')) continue
    const si = Math.max(0, Math.min(b.section ?? 0, pitches.length - 1))
    specs.push({ el: b.el, dx: 0, dy: 0, gridPitchPt: pitches[si] ?? 0 })
  }
  return specs
}

/**
 * Per-block character-grid letter-spacing for documents whose sections
 * disagree on the docGrid charSpace delta (e.g. prod-sas 043: section 1 typed
 * lines only, section 2 linesAndChars charSpace=-820). Uniform docs get no
 * specs (the single .doc-page --doc-char-space injection stays). Blocks of
 * sections without an effective charSpace simply carry no spec: with no
 * uniform injection present there is nothing to override.
 */
export function sectionCharSpaceSpecs(
  blocks: BlockBox[],
  sections: SectionInfo[],
): ColumnBlockPlacement[] {
  const deltas = sections.map(sectionCharSpacePt)
  if (!deltas.some((d) => d !== 0) || deltas.every((d) => d === deltas[0])) return []
  const specs: ColumnBlockPlacement[] = []
  for (const b of blocks) {
    if (!b.el) continue
    const si = Math.max(0, Math.min(b.section ?? 0, deltas.length - 1))
    const d = deltas[si] ?? 0
    if (d !== 0) specs.push({ el: b.el, dx: 0, dy: 0, charSpacePt: d })
  }
  return specs
}

/**
 * sectPr w:vAlign (center/bottom) page shifts: blocks of a vertically aligned
 * page translate down into the page's free space — purely visual, through the
 * same decoration channel as column placement (the page band already spans the
 * full page height, so shifted blocks stay inside their page). 'both'
 * (justified) keeps top alignment, as does a page where a block crosses either
 * boundary (shifting one half of a split block would tear it) and multi-column
 * pages.
 */
export function vAlignShiftSpecs(
  blocks: BlockBox[],
  slices: PageSlice[],
  sections: SectionInfo[],
  geoms: SectionGeom[],
): ColumnBlockPlacement[] {
  const specs: ColumnBlockPlacement[] = []
  if (!sections.some((s) => s.settings.vAlign === 'center' || s.settings.vAlign === 'bottom'))
    return specs
  for (const slice of slices) {
    const va = sections[slice.section]?.settings.vAlign
    if (va !== 'center' && va !== 'bottom') continue
    if (slice.regions) continue
    const colH = geoms[slice.section]?.contentHeight ?? 0
    const free = colH - (slice.end - slice.start)
    if (free < 1) continue
    const dy = va === 'center' ? free / 2 : free
    const page: ColumnBlockPlacement[] = []
    let whole = true
    for (const b of blocks) {
      if (b.top + b.height <= slice.start + 0.5) continue
      if (b.top >= slice.end - 0.5) break
      if (b.top < slice.start - 0.5 || b.top + b.height > slice.end + 2) {
        whole = false
        break
      }
      if (!b.el || b.floated) continue
      page.push({ el: b.el, dx: 0, dy })
    }
    if (whole) specs.push(...page)
  }
  return specs
}

/** SectionInfo[] → pagination geometry
 *  - continuous with unchanged page geometry: no forced break (content flows on the same page)
 *  - continuous with changed page geometry (width/height change, e.g. landscape → portrait): forced break
 *  - nextPage/evenPage/oddPage: forced break
 *  - with hfHeights, oversized headers/footers squeeze body capacity
 */
export function sectionGeoms(
  sections: SectionInfo[],
  hfHeights?: SectionHfHeights[],
): SectionGeom[] {
  return sections.map((s, i) => {
    const cols = sectionColumns(s)
    let forceBreak = false
    let colBreakStart = false
    if (i > 0) {
      // nextColumn with the same multi-column count advances one column (Word,
      // tdf#135343 c14/c15); a changed count or a single-column layout acts like a
      // page break (tdf#135343 c12v3, n#750255)
      colBreakStart =
        s.startType === 'nextColumn' && cols > 1 && sectionColumns(sections[i - 1]) === cols
      const asContinuous = s.startType === 'continuous' || colBreakStart
      if (!asContinuous) {
        forceBreak = true
      } else {
        // continuous section: force a page break if the page size differs from the previous section (e.g. landscape → portrait)
        const prev = sections[i - 1].settings
        const cur = s.settings
        if (prev.pageWidth !== cur.pageWidth || prev.pageHeight !== cur.pageHeight) {
          forceBreak = true
        }
      }
    }
    const set = s.settings
    const hf = hfHeights?.[i]
    const firstContentHeight =
      hf?.firstHeaderPx !== undefined || hf?.firstFooterPx !== undefined
        ? twipsToPx(set.pageHeight) -
          effectiveTopPx(set, hf.firstHeaderPx ?? 0) -
          effectiveBottomPx(set, hf.firstFooterPx ?? 0)
        : undefined
    return {
      contentHeight:
        twipsToPx(set.pageHeight) -
        effectiveTopPx(set, hf?.headerPx ?? 0) -
        effectiveBottomPx(set, hf?.footerPx ?? 0),
      ...(firstContentHeight !== undefined ? { firstContentHeight } : {}),
      contentWidth: twipsToPx(set.pageWidth - set.marginLeft - set.marginRight),
      topPx: effectiveTopPx(set, hf?.headerPx ?? 0),
      forceBreak,
      startType: s.startType,
      // colWidths only for explicit-width columns: the narrower-column gate is
      // meaningless when every column is the same width
      ...(cols > 1
        ? {
            cols,
            ...(sectionColGeom(s).equalWidth ? {} : { colWidths: sectionColGeom(s).widths }),
          }
        : {}),
      ...(colBreakStart && !forceBreak ? { colBreakStart: true } : {}),
    }
  })
}

/**
 * Live section list: when a non-final section's break paragraph (the block at
 * lastBlockIndex) has been deleted from the canvas, that section merges into the
 * next (content before a section break takes the following section's
 * page setup). This is derived and doesn't mutate the authoritative sections state,
 * so undoing the deletion restores naturally; readSections rebuilds after saving.
 */
export function liveSections(
  sections: SectionInfo[],
  blocks: BlockBox[],
  /** boundaries in the DOM but not in the block list (zero-height section-break chips) */
  extraPresent?: Set<number>,
  /** break paragraphs whose mark is a tracked deletion: Word's markup view drops the break */
  trackedDeleted?: Set<number>,
): SectionInfo[] {
  if (sections.length <= 1) return sections
  const present = new Set<number>(extraPresent)
  for (const b of blocks) if (b.docxIndex !== undefined) present.add(b.docxIndex)
  if (trackedDeleted) for (const i of trackedDeleted) present.delete(i)
  const out: SectionInfo[] = []
  let carryFirst: number | null = null
  let changed = false
  sections.forEach((s, i) => {
    const first = carryFirst ?? s.firstBlockIndex
    carryFirst = null
    if (i < sections.length - 1 && !present.has(s.lastBlockIndex)) {
      changed = true
      carryFirst = first
      return
    }
    out.push(first === s.firstBlockIndex ? s : { ...s, firstBlockIndex: first })
  })
  return changed ? out : sections
}

/** Tag each block's owning section by the sections' block ranges (lastBlockIndex); new blocks without docxIndex inherit from the previous block */
export function assignSections(blocks: BlockBox[], sections: SectionInfo[]): void {
  const ends = sections.map((s) => s.lastBlockIndex)
  let prev = 0
  for (const block of blocks) {
    let s = prev
    if (block.docxIndex !== undefined) {
      const i = ends.findIndex((end) => block.docxIndex! <= end)
      s = i >= 0 ? i : ends.length - 1
    }
    block.section = s
    prev = s
  }
}
