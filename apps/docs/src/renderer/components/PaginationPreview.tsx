import { useEffect, useMemo, useState } from 'react'
import type {
  HeaderFooter,
  HfImage,
  HfPartInfo,
  SectionInfo,
  SectionSettings,
} from '@genoffice/docx-engine'
import {
  appendEndnotesBlock,
  assignSections,
  effectiveBottomPx,
  effectiveHfRefs,
  effectiveTopPx,
  formatPageNumber,
  liveSections,
  measureBlocks,
  pageNumbers,
  sectionColGeom,
  sectionFirstPages,
  sectionGeoms,
  sectionPageBox,
  sliceWithLineSplit,
  type BlockBox,
  type BlockMetaOf,
  type PageNoteItem,
  type PageSlice,
  type SectionHfHeights,
} from '../pagination'
import { estimateHfHeight, FOOTNOTE_SEPARATOR_H } from '../line-metrics'
import { toRoman } from '../note-format'
import { useI18n } from '../i18n/locale'
import { HeaderFooterArea } from './HeaderFooterArea'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** Snapshot of one top-level canvas block for pruned per-page clones (virtual gapless coordinates, layout px) */
export interface CloneChild {
  html: string
  vTop: number
  vBottom: number
  /** CSS margins (layout px): spacer heights must exclude them to keep flow positions exact */
  mt: number
  mb: number
  /** zero-height marker (hidden bookmarks etc.): always kept, never worth pruning */
  zero: boolean
}

/**
 * Per-page full-document clones cost pages × doc DOM; past this budget (top-level
 * blocks × pages) a 300+-page document OOMs the renderer during preview/export
 * ("Promise was collected"), so pages switch to pruned clones: blocks outside the
 * page window collapse into fixed-height spacers.
 */
const CLONE_PRUNE_BUDGET = 150_000
/** window slack around a page (px): keeps neighbours whose floats/overflow bleed into the page */
const CLONE_PRUNE_PAD = 2000

/** pruned clone for one page window: blocks intersecting [from-pad, to+pad] verbatim, pruned runs as spacers (exported for tests) */
export function prunedCloneHtml(kids: CloneChild[], from: number, to: number): string {
  const lo = from - CLONE_PRUNE_PAD
  const hi = to + CLONE_PRUNE_PAD
  const parts: string[] = []
  let lastKept: CloneChild | null = null
  let pruned = false
  for (const c of kids) {
    if (!c.zero && (c.vBottom <= lo || c.vTop >= hi)) {
      pruned = true
      continue
    }
    if (pruned) {
      // spacer replaces the pruned run; its height re-derives the next block's
      // border-box top from the previous kept block's margin edge (spacers
      // suppress margin collapse, so both adjacent margins apply in full)
      const base = lastKept ? lastKept.vBottom + lastKept.mb : 0
      const h = Math.max(0, c.vTop - c.mt - base)
      parts.push(
        `<div class="pv-prune-spacer" style="margin:0;border:0;padding:0;height:${h}px"></div>`,
      )
      pruned = false
    }
    parts.push(c.html)
    // zero markers advance the flow position too (a spacer may have been
    // emitted before them); skipping them leaves `base` stale and the next
    // spacer re-adds the distance already consumed, shifting content down
    lastKept = c
  }
  return parts.join('')
}

export interface HfSet {
  header: HeaderFooter | null
  footer: HeaderFooter | null
  headerFirst: HeaderFooter | null
  footerFirst: HeaderFooter | null
  headerEven: HeaderFooter | null
  footerEven: HeaderFooter | null
  titlePg: boolean
  evenOddHf: boolean
  /** images in each variant part (logos etc., display-only) */
  images?: Partial<
    Record<
      'header' | 'footer' | 'headerFirst' | 'footerFirst' | 'headerEven' | 'footerEven',
      HfImage[]
    >
  >
}

/**
 * Pagination preview: a read-only snapshot of real page slicing over the canvas's continuous
 * flow. Each page = a full content clone + overflow clipping + negative-margin offset; the
 * clone is fixed at the canvas content width (line breaks from measurement must not change),
 * and paper size/margins render per each page's section (mixed portrait/landscape across
 * sections is real). Headers/footers render per page by Word variant rules (first page /
 * odd-even), with real page numbers.
 */
export function PaginationPreview({
  section,
  sections,
  hfParts,
  colFlow,
  zoom,
  hf,
  watermark,
  pageColor,
  blockMetaOf,
  pageFootnotesOf,
  endnoteItems,
  sectionHfOverride,
  onExportPdf,
  onClose,
}: {
  /** Canvas geometry (final section): for the measurement origin / clone width */
  section: SectionSettings
  /** All sections: for per-page paper geometry (empty array = single section per `section`) */
  sections: SectionInfo[]
  /** rId → header/footer parts (multi-section picks by each section's references) */
  hfParts: Record<string, HfPartInfo>
  /** Canvas column-flow geometry (non-null when the canvas column CSS is active): shared by the measuring state / clone wrap width */
  colFlow: { cols: number; colWidthPx: number; gapPx: number } | null
  zoom: number
  hf: HfSet
  watermark: string | null
  pageColor: string | null
  /** docxIndex → parse-layer pagination constraints (keepNext/widow/table-row flags) */
  blockMetaOf?: BlockMetaOf
  /** Per-page footnote collection (referencing page → entry list), for page-bottom rendering */
  pageFootnotesOf?: (blocks: BlockBox[], slices: PageSlice[]) => PageNoteItem[][]
  /** Endnote entries (placed together at the document end, take part in slicing, may continue across pages) */
  endnoteItems?: PageNoteItem[]
  /** Multi-section: unsaved per-section header/footer edit overrides (default variant) */
  sectionHfOverride?: (sectionIndex: number, kind: 'header' | 'footer') => HeaderFooter | null
  onExportPdf: () => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [slices, setSlices] = useState<PageSlice[]>([])
  const [pageNotes, setPageNotes] = useState<PageNoteItem[][]>([])
  /** Top Y of the endnote area (virtual coordinates); null = no endnotes */
  const [endnotesTop, setEndnotesTop] = useState<number | null>(null)
  const [html, setHtml] = useState('')
  /** non-null = pruned-clone mode (large documents): per-page windows instead of full clones */
  const [cloneKids, setCloneKids] = useState<CloneChild[] | null>(null)
  /** Live section list: a section whose break block was deleted (unsaved) merges into the next, matching the canvas */
  const [secs, setSecs] = useState<SectionInfo[]>(sections)

  const canvasContentW = twipsToPx(section.pageWidth - section.marginLeft - section.marginRight)
  // clone wrap width = measurement width: a columned canvas measures single-flow at column width; the clone must match to reproduce line breaks
  const wrapW = colFlow?.colWidthPx ?? canvasContentW
  // canvas content-area top = effective top margin after header push-down (matches --page-pad)
  const canvasMTop = effectiveTopPx(
    section,
    estimateHfHeight(hf.header, canvasContentW, hf.images?.header),
  )
  /** Settings of the page's section (single-section documents fall back to the canvas geometry) */
  const settingsOf = (slice: PageSlice): SectionSettings =>
    secs[Math.min(slice.section, secs.length - 1)]?.settings ?? section

  useEffect(() => {
    const pm = document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
    if (!pm) return
    const factor = zoom / 100
    // switch the columned canvas to the single-flow measuring state (CSS columns off, width = column width), matching engine column-flow coordinates
    if (colFlow) pm.classList.add('measuring-columns')
    try {
      const origin = pm.getBoundingClientRect().top + canvasMTop * factor
      const { blocks, totalHeight } = measureBlocks(pm, origin, factor)
      const live = liveSections(sections, blocks)
      setSecs(live)
      if (live.length > 0) assignSections(blocks, live)
      const withEndnotes = appendEndnotesBlock(
        blocks,
        totalHeight,
        endnoteItems ?? [],
        FOOTNOTE_SEPARATOR_H,
      )
      const flowH = withEndnotes?.totalHeight ?? totalHeight
      setEndnotesTop(withEndnotes?.top ?? null)
      let computed: PageSlice[]
      if (live.length > 0) {
        // each section's default-variant header/footer estimated heights → body push-down (matching the canvas)
        const refs = effectiveHfRefs(live)
        const hfHs: SectionHfHeights[] = live.map((s, i) => {
          const set = s.settings
          const w = twipsToPx(set.pageWidth - set.marginLeft - set.marginRight)
          const pick = (kind: 'header' | 'footer'): HeaderFooter | null => {
            if (i === live.length - 1) return kind === 'header' ? hf.header : hf.footer
            const ov = sectionHfOverride?.(i, kind)
            if (ov) return ov
            const rId = refs[i]?.[kind]?.default
            const part = rId ? hfParts[rId] : undefined
            return part
              ? { text: part.text, pageNumber: part.hasPageNumber, paras: part.paras }
              : null
          }
          const imagesOf = (kind: 'header' | 'footer') => {
            const rId = refs[i]?.[kind]?.default
            const fromPart = rId ? hfParts[rId]?.images : undefined
            if (fromPart?.length) return fromPart
            return i === live.length - 1 ? hf.images?.[kind] : undefined
          }
          return {
            headerPx: estimateHfHeight(pick('header'), w, imagesOf('header')),
            footerPx: estimateHfHeight(pick('footer'), w, imagesOf('footer')),
          }
        })
        const geoms = sectionGeoms(live, hfHs)
        // when the canvas column CSS is inactive, measure as full-width single flow; the geometry drops column flow to match
        if (!colFlow) for (const g of geoms) if (g.cols) g.cols = undefined
        computed = sliceWithLineSplit(blocks, geoms, flowH, factor, blockMetaOf)
      } else {
        const contentH =
          twipsToPx(section.pageHeight) -
          effectiveTopPx(section, estimateHfHeight(hf.header, canvasContentW, hf.images?.header)) -
          effectiveBottomPx(section, estimateHfHeight(hf.footer, canvasContentW, hf.images?.footer))
        computed = sliceWithLineSplit(
          blocks,
          [
            {
              contentHeight: contentH,
              forceBreak: false,
              ...(colFlow ? { cols: colFlow.cols } : {}),
            },
          ],
          flowH,
          factor,
          blockMetaOf,
        )
      }
      setSlices(computed)
      setPageNotes(pageFootnotesOf ? pageFootnotesOf(blocks, computed) : [])
      // Per-page full clones explode on large documents (pages × doc DOM →
      // renderer OOM / "Promise was collected" during printToPDF). Past the
      // budget, snapshot per-block geometry and render pruned windows instead.
      const kidEls = Array.from(pm.children) as HTMLElement[]
      if (computed.length * kidEls.length >= CLONE_PRUNE_BUDGET) {
        const metas: CloneChild[] = []
        let gapAccum = 0
        for (const el of kidEls) {
          const rect = el.getBoundingClientRect()
          if (el.classList.contains('page-gap')) {
            gapAccum += rect.height
            continue
          }
          let innerGap = 0
          for (const g of el.querySelectorAll('.page-gap-inline'))
            innerGap += g.getBoundingClientRect().height
          const cs = window.getComputedStyle(el)
          const vTop = (rect.top - origin - gapAccum) / factor
          const h = (rect.height - innerGap) / factor
          gapAccum += innerGap
          metas.push({
            html: el.outerHTML,
            vTop,
            vBottom: vTop + h,
            mt: parseFloat(cs.marginTop) || 0,
            mb: parseFloat(cs.marginBottom) || 0,
            zero: rect.height <= 0,
          })
        }
        setCloneKids(metas)
        setHtml('')
      } else {
        setCloneKids(null)
        setHtml(pm.innerHTML)
      }
    } finally {
      if (colFlow) pm.classList.remove('measuring-columns')
    }
    // snapshot: measure once on open; deps intentionally empty
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onClose()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose])

  const multiSection = secs.length > 1
  const effRefs = useMemo(() => effectiveHfRefs(secs), [secs])
  // single-section also uses pageNumbers: pgNumType w:start renumbering applies to single-section documents too
  const nums = useMemo(
    () => (secs.length > 0 ? pageNumbers(slices, secs) : slices.map((_, i) => i + 1)),
    [slices, secs],
  )
  const firsts = useMemo(() => sectionFirstPages(slices), [slices])

  // line positions of endnote entries in virtual coordinates (matching appendEndnotesBlock's line boxes)
  const endnoteRows = useMemo(() => {
    if (endnotesTop === null || !endnoteItems || endnoteItems.length === 0) return []
    let off = endnotesTop
    return endnoteItems.map((item, i) => {
      const height = (i === 0 ? FOOTNOTE_SEPARATOR_H : 0) + item.height
      const row = { item, top: off, height, withSeparator: i === 0 }
      off += height
      return row
    })
  }, [endnotesTop, endnoteItems])

  const toHf = (rId: string | undefined): HeaderFooter | null => {
    const part = rId ? hfParts[rId] : undefined
    if (!part) return null
    return {
      text: part.text,
      pageNumber: part.hasPageNumber,
      paras: part.paras.length > 0 ? part.paras : undefined,
    }
  }

  /** Single section: reuse the editing state (unsaved header edits are visible); multi-section: pick parts by each section's references */
  const hfFor = (
    i: number,
  ): {
    header: HeaderFooter | null
    footer: HeaderFooter | null
    headerImages?: HfImage[]
    footerImages?: HfImage[]
  } => {
    const pageNo = nums[i]
    if (!multiSection) {
      if (hf.titlePg && i === 0) {
        return {
          header: hf.headerFirst,
          footer: hf.footerFirst,
          headerImages: hf.images?.headerFirst,
          footerImages: hf.images?.footerFirst,
        }
      }
      if (hf.evenOddHf && pageNo % 2 === 0) {
        return {
          header: hf.headerEven,
          footer: hf.footerEven,
          headerImages: hf.images?.headerEven,
          footerImages: hf.images?.footerEven,
        }
      }
      return {
        header: hf.header,
        footer: hf.footer,
        headerImages: hf.images?.header,
        footerImages: hf.images?.footer,
      }
    }
    const slice = slices[i]
    const sec = secs[Math.min(slice.section, secs.length - 1)]
    const refs = effRefs[Math.min(slice.section, effRefs.length - 1)]
    const variant =
      sec.titlePg && firsts[i] ? 'first' : hf.evenOddHf && pageNo % 2 === 0 ? 'even' : 'default'
    // unsaved per-section header/footer edits take priority over document parts (default variant)
    const ovHeader = variant === 'default' ? sectionHfOverride?.(slice.section, 'header') : null
    const ovFooter = variant === 'default' ? sectionHfOverride?.(slice.section, 'footer') : null
    const headerRId = refs.header[variant]
    const footerRId = refs.footer[variant]
    return {
      header: ovHeader ?? toHf(headerRId),
      footer: ovFooter ?? toHf(footerRId),
      headerImages: headerRId ? hfParts[headerRId]?.images : undefined,
      footerImages: footerRId ? hfParts[footerRId]?.images : undefined,
    }
  }

  return (
    <div className="pagination-preview">
      <div className="pv-toolbar">
        <span className="pv-title">{t('appPaginationPreview')}</span>
        <span className="pv-count">{t('appTotalPagesN', { n: slices.length })}</span>
        <span className="pv-hint">{t('appPvHint')}</span>
        <button className="pv-close" title={t('appPvExportTip')} onClick={onExportPdf}>
          {t('appExportPdf')}
        </button>
        <button className="pv-close" onClick={onClose}>
          {t('appClose')}
        </button>
      </div>
      <div className="pv-scroll">
        {slices.map((slice, i) => {
          const parts = hfFor(i)
          const s = settingsOf(slice)
          const pageBox = sectionPageBox(s)
          const pageW = pageBox.width
          const pageH = pageBox.height
          const secContentW = pageBox.contentWidth
          // effective margins after this page's variant header/footer push-down (an over-tall header pushes the body down)
          const mTop = effectiveTopPx(s, estimateHfHeight(parts.header, secContentW))
          const mBottom = effectiveBottomPx(s, estimateHfHeight(parts.footer, secContentW))
          const contentH = pageH - mTop - mBottom
          // page vertical alignment (sectPr w:vAlign): content of non-full pages shifts down as a whole
          const usedH = Math.min(slice.end - slice.start, contentH)
          const vSpare = Math.max(0, contentH - usedH)
          const vOffset = s.vAlign === 'center' ? vSpare / 2 : s.vAlign === 'bottom' ? vSpare : 0
          // page numbers display in the owning section's number format (w:pgNumType w:fmt)
          const pageNoText = formatPageNumber(
            nums[i],
            secs[Math.min(slice.section, secs.length - 1)]?.pageNumberFmt,
          )
          return (
            <div
              key={i}
              className="pv-page"
              style={
                {
                  width: pageW,
                  height: pageH,
                  '--pv-page-h': `${pageH}px`,
                  '--page-w': `${pageW}px`,
                  '--page-h': `${pageH}px`,
                  '--header-dist': `${pageBox.headerDist}px`,
                  '--footer-dist': `${pageBox.footerDist}px`,
                  '--pv-mr': `${twipsToPx(s.marginRight)}px`,
                  padding: `${mTop}px ${twipsToPx(s.marginRight)}px ${mBottom}px ${twipsToPx(s.marginLeft)}px`,
                  background: pageColor ? `#${pageColor}` : undefined,
                } as React.CSSProperties
              }
            >
              {watermark && (
                <div className="page-watermark" aria-hidden="true">
                  {watermark}
                </div>
              )}
              {parts.header && (
                <HeaderFooterArea
                  kind="header"
                  value={parts.header}
                  images={parts.headerImages}
                  readOnly
                  onCommit={() => {}}
                  pageNo={pageNoText}
                  pageTotal={slices.length}
                />
              )}
              {slice.repeatHeader && !slice.regions && (
                // tblHeader repeated headers: a broken table's page first renders a clone of the source table's header rows
                // (the engine already reserved repeatHeader.height on this page)
                <div className="pv-clip" style={{ height: slice.repeatHeader.height }}>
                  <div
                    className="pv-offset"
                    style={{ marginTop: -slice.repeatHeader.top, width: wrapW }}
                  >
                    <div
                      className="doc-page pv-content"
                      dangerouslySetInnerHTML={{
                        __html: cloneKids
                          ? prunedCloneHtml(
                              cloneKids,
                              slice.repeatHeader.top,
                              slice.repeatHeader.top + slice.repeatHeader.height,
                            )
                          : html,
                      }}
                    />
                  </div>
                </div>
              )}
              {slice.regions ? (
                // column flow: regions stack vertically; within a region, columns are narrow-clipped side by side (column-leading repeated headers follow their column)
                slice.regions.map((region, ri) => {
                  const rSec = secs[Math.min(region.section, secs.length - 1)]
                  const rg = rSec
                    ? sectionColGeom(rSec)
                    : (colFlow ?? { cols: 1, colWidthPx: canvasContentW, gapPx: 0 })
                  const extent =
                    ri + 1 < slice.regions!.length
                      ? slice.regions![ri + 1].top - region.top
                      : undefined
                  const multi = rg.cols > 1
                  return (
                    <div
                      key={ri}
                      className="pv-region"
                      style={{ gap: rg.gapPx, ...(extent !== undefined ? { height: extent } : {}) }}
                    >
                      {region.columns.map((col, ci) => (
                        <div
                          key={ci}
                          className="pv-col"
                          style={{ width: multi ? rg.colWidthPx : undefined }}
                        >
                          {col.repeatHeader && (
                            <div className="pv-clip" style={{ height: col.repeatHeader.height }}>
                              <div
                                className="pv-offset"
                                style={{ marginTop: -col.repeatHeader.top, width: wrapW }}
                              >
                                <div
                                  className="doc-page pv-content"
                                  dangerouslySetInnerHTML={{
                                    __html: cloneKids
                                      ? prunedCloneHtml(
                                          cloneKids,
                                          col.repeatHeader.top,
                                          col.repeatHeader.top + col.repeatHeader.height,
                                        )
                                      : html,
                                  }}
                                />
                              </div>
                            </div>
                          )}
                          <div
                            className="pv-clip"
                            style={{
                              height: Math.min(
                                col.end - col.start,
                                region.height - (col.repeatHeader?.height ?? 0),
                              ),
                            }}
                          >
                            <div
                              className="pv-offset"
                              style={{ marginTop: -col.start, width: wrapW }}
                            >
                              <div
                                className="doc-page pv-content"
                                dangerouslySetInnerHTML={{
                                  __html: cloneKids
                                    ? prunedCloneHtml(cloneKids, col.start, col.end)
                                    : html,
                                }}
                              />
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )
                })
              ) : (
                <div
                  className="pv-clip"
                  style={{
                    // last page opens to full capacity: slice bounds can drift a few
                    // lines short of the clone's real height (page-crossing tables),
                    // silently dropping the document tail from export/print;
                    // past the real content bottom the window is empty anyway
                    height:
                      i === slices.length - 1 && vOffset <= 0.5
                        ? contentH - (slice.repeatHeader?.height ?? 0)
                        : Math.min(
                            slice.end - slice.start,
                            contentH - (slice.repeatHeader?.height ?? 0),
                          ),
                    ...(vOffset > 0.5 ? { marginTop: vOffset } : {}),
                  }}
                >
                  {/* the offset lives on a separate wrapper: print rules zero out .doc-page's margin;
                      width is fixed to the measured wrap width (columned canvas = column width), so wider landscape-section paper doesn't reflow */}
                  <div className="pv-offset" style={{ marginTop: -slice.start, width: wrapW }}>
                    <div
                      className="doc-page pv-content"
                      dangerouslySetInnerHTML={{
                        __html: cloneKids
                          ? prunedCloneHtml(
                              cloneKids,
                              slice.start,
                              // last page opens its clip to full capacity; the window must cover it
                              i === slices.length - 1 ? slice.start + contentH : slice.end,
                            )
                          : html,
                      }}
                    />
                  </div>
                </div>
              )}
              {(pageNotes[i]?.length ?? 0) > 0 && (
                // page-bottom footnotes (Word behavior: placed at the bottom of the page's content area, separator on top)
                <div
                  className="pv-footnotes"
                  style={{
                    left: twipsToPx(s.marginLeft),
                    width: pageW - twipsToPx(s.marginLeft) - twipsToPx(s.marginRight),
                    bottom: twipsToPx(s.marginBottom),
                    height:
                      pageNotes[i]!.reduce((sum, n) => sum + n.height, 0) + FOOTNOTE_SEPARATOR_H,
                  }}
                >
                  {pageNotes[i]!.map((n) => (
                    // entries get fixed heights from the estimates, strictly matching the engine's reservation
                    <div key={n.id} className="pv-footnote" style={{ height: n.height }}>
                      <sup>{n.no}</sup>
                      {n.richParas
                        ? n.richParas.map((para, pi) => (
                            <span key={pi}>
                              {pi > 0 && <br />}
                              {para.map((run, ri) => (
                                <span
                                  key={ri}
                                  style={{
                                    fontWeight: run.bold ? 600 : undefined,
                                    fontStyle: run.italic ? 'italic' : undefined,
                                    textDecoration:
                                      [run.underline && 'underline', run.strike && 'line-through']
                                        .filter(Boolean)
                                        .join(' ') || undefined,
                                    color: run.color ? `#${run.color}` : undefined,
                                    fontSize: run.sizeHalfPoints
                                      ? `${run.sizeHalfPoints / 2}pt`
                                      : undefined,
                                  }}
                                >
                                  {run.text}
                                </span>
                              ))}
                            </span>
                          ))
                        : n.text}
                    </div>
                  ))}
                </div>
              )}
              {(() => {
                // endnotes: immediately after the body's end, placed on pages per the slices, may continue across pages
                const rows = endnoteRows.filter(
                  (r) => r.top >= slice.start - 0.5 && r.top < slice.end - 0.5,
                )
                if (rows.length === 0) return null
                return (
                  <div
                    className={`pv-endnotes${rows[0].withSeparator ? ' with-separator' : ''}`}
                    style={{
                      left: twipsToPx(s.marginLeft),
                      width: pageW - twipsToPx(s.marginLeft) - twipsToPx(s.marginRight),
                      top: mTop + (slice.repeatHeader?.height ?? 0) + (rows[0].top - slice.start),
                    }}
                  >
                    {rows.map(({ item: n, height, withSeparator }) => (
                      <div key={n.id} className="pv-footnote" style={{ height }}>
                        {withSeparator && <div className="pv-endnote-separator" />}
                        <sup>{toRoman(n.no)}</sup>
                        {n.richParas
                          ? n.richParas.map((para, pi) => (
                              <span key={pi}>
                                {pi > 0 && <br />}
                                {para.map((run, ri) => (
                                  <span
                                    key={ri}
                                    style={{
                                      fontWeight: run.bold ? 600 : undefined,
                                      fontStyle: run.italic ? 'italic' : undefined,
                                      textDecoration:
                                        [run.underline && 'underline', run.strike && 'line-through']
                                          .filter(Boolean)
                                          .join(' ') || undefined,
                                      color: run.color ? `#${run.color}` : undefined,
                                      fontSize: run.sizeHalfPoints
                                        ? `${run.sizeHalfPoints / 2}pt`
                                        : undefined,
                                    }}
                                  >
                                    {run.text}
                                  </span>
                                ))}
                              </span>
                            ))
                          : n.text}
                      </div>
                    ))}
                  </div>
                )
              })()}
              {parts.footer && (
                <HeaderFooterArea
                  kind="footer"
                  value={parts.footer}
                  images={parts.footerImages}
                  readOnly
                  onCommit={() => {}}
                  pageNo={pageNoText}
                  pageTotal={slices.length}
                />
              )}
              <div className="pv-pageno">{i + 1}</div>
            </div>
          )
        })}
      </div>
    </div>
  )
}
