import { useEffect, useMemo, useRef, useState } from 'react'
import type {
  CommentInfo,
  HeaderFooter,
  HfImage,
  HfPartInfo,
  SectionInfo,
  SectionSettings,
} from '@genoffice/docx-engine'
import {
  pageTextEnds,
  effectiveBottomPx,
  effectiveHfRefs,
  effectiveTopPx,
  noteAreaPlacement,
  type HfSet,
  hfPartWidthPx,
  type HfVariant,
  type PageHf,
  pageHfStrips,
  singleVariantGeomOf,
  anchorBoxLift,
  anchorShiftPx,
  sectionBidi,
  sectionColGeom,
  sectionPageCounts,
  sectionVertical,
  blockInlineExtraPx,
  sectionPageBox,
  rowSplitCss,
  seamWindow,
  hoistWindow,
  pageCutDiffs,
  pageOwners,
  applyBlockMeta,
  type PageCut,
  type BlockMetaOf,
  type FloatBox,
  type PageNoteItem,
  type PageSlice,
  type SectionGeom,
  type SectionHfHeights,
  type SliceOutputs,
} from '../pagination'
import { applyFloatPullUps, sliceFlow, type PassContext } from '../pagination-pass'
import type { PageLayout } from '../editor/goto'
import { hfVariantHeightsPx, sectionHfHeights } from '../hf-heights'
import { pageMargins } from '../page-margins'
import { noteSeparatorPx, noteSeparatorRulePx } from '../line-metrics'
import { cssFontFamily } from '../text-layout/font-rules'
import { syncPreviewLineNumbers } from '../editor/line-numbers'
import type { EditorView } from '@tiptap/pm/view'
import { revWhen, type AnchorBlock } from '../editor/margin-annotations'
import { pageBorderArtStripStyle, pageBorderStyleOf } from '../editor/pagination-gaps'
import { textColorValue } from '../editor/text-color'
import { textOutlineCssValue } from '../editor/text-outline'
import { noteMarkText } from '../note-format'
import { noteRowBox, noteRunClass } from '../editor/note-dom'
import { useI18n } from '../i18n/locale'
import {
  hfFloatPagePos,
  hfImageHangsOnPara,
  hfFloatTransform,
  hfStripGeom,
  hfWashoutFilter,
  wordArtSvgMarkup,
} from '../editor/hf-dom'
import { HeaderFooterArea } from './HeaderFooterArea'

import {
  anchoredBoxesOf,
  canvasChildWidths,
  canvasContentLeft,
  canvasEditorEl,
  canvasFlowOrigin,
  cloneChildrenOf,
  hoistMarkedEls,
  isLaidOutAbsoluteBox,
  previewScrollEl,
  publishPreviewProbes,
  shapeBoxExtent,
  bodyWindowHeight,
  canvasWrapWidth,
  cloneBlockHtml,
  cloneChildGeom,
  collectCellPageBoxes,
  hoistPaperClamp,
  hoistPlan,
  hoistSlotOf,
  hoistWindowCss,
  leadLeakCss,
  measureCommentSpots,
  measureRevisionSpots,
  pageRelXPictureHtml,
  pinnedCloneCss,
  prunedCloneHtml,
  shouldPruneClones,
  verticalPageCss,
  type CellPageBox,
  type CloneChild,
  type CommentSpot,
  type PageRelXPicture,
  type RevSpot,
} from '../preview-canvas-snapshot'

export * from '../preview-canvas-snapshot'

const twipsToPx = (twips: number) => (twips / 1440) * 96
/** separator row CSS variables of a notes area (styles.css sizes the rule and its spacing from them) */
const separatorVars = (kind: 'footnote' | 'endnote') =>
  ({
    '--note-sep-h': `${noteSeparatorPx(kind)}px`,
    '--note-sep-rule': `${noteSeparatorRulePx(kind)}px`,
  }) as React.CSSProperties

/**
 * Word's print-markup geometry (measured against Word for Mac PDF output):
 * the sheet is laid out on a virtual page widened by a markup strip, the whole
 * thing is scaled uniformly to fit the paper width and centered vertically,
 * and balloons keep their anchor's unscaled Y. On A4 that comes out to
 * k ≈ 0.745 with the gray strip running from the content's right edge.
 */
const MARKUP_EXTRA_W = 272
const MARKUP_BAND_GUTTER = 8
const BUBBLE_ENTRY = 40
const BUBBLE_RIGHT_PAD = 24
const BUBBLE_STACK_GAP = 8
const BUBBLE_FONT_PX = 12
const BUBBLE_LINE_H = 16
const BUBBLE_PAD_H = 12
/** revision balloons (.pv-rev-bubble): 11px/14px text, 2px padding, 1px border */
const REV_FONT_PX = 11
const REV_LINE_H = 14
const REV_PAD_H = 6
const REV_STACK_GAP = 3
const REV_TEXT_MAX = 300

type BubbleMetrics = { fontPx: number; lineH: number; padH: number }
const COMMENT_METRICS: BubbleMetrics = {
  fontPx: BUBBLE_FONT_PX,
  lineH: BUBBLE_LINE_H,
  padH: BUBBLE_PAD_H,
}
const REV_METRICS: BubbleMetrics = { fontPx: REV_FONT_PX, lineH: REV_LINE_H, padH: REV_PAD_H }

/** width-weighted wrap estimate: bubbles are absolutely stacked before layout */
function estimateBubbleHeight(text: string, innerW: number, m = COMMENT_METRICS): number {
  let w = 0
  let lines = 1
  for (const ch of text) {
    const wide = (ch.codePointAt(0) ?? 0) > 0x2e7f
    const cw = ch === '\n' ? Infinity : wide ? m.fontPx : m.fontPx * 0.55
    if (w + cw > innerW) {
      lines += 1
      w = ch === '\n' ? 0 : cw
    } else {
      w += cw
    }
  }
  return lines * m.lineH + m.padH
}

/** one balloon to stack on a page */
export type BalloonItem = {
  key: string
  /** anchor line top on the page (page coordinates) */
  anchorTop: number
  /** stacking position (see mergeBalloonLists) */
  seq: number
  endX: number
  height: number
  /** height once the page is crowded (one line, text clipped); absent = never compacts */
  compactHeight?: number
  /** always prints; non-sticky balloons are the ones Word routes to its overflow pane */
  sticky: boolean
}

/**
 * One stacking sequence from two lists that are each already in Word's order
 * (comments: document order; revisions: anchor order): merge by `orderTop`
 * (a monotonic stacking Y, so the lists' own order is kept), same-line ties
 * left to right, and number the result. Comparing positions across the two
 * lists would otherwise be arbitrary on a shared line.
 */
export function mergeBalloonLists<T extends { orderTop: number; endX: number }>(
  a: T[],
  b: T[],
): Array<T & { seq: number }> {
  const out: Array<T & { seq: number }> = []
  let i = 0
  let j = 0
  while (i < a.length || j < b.length) {
    const x = a[i]
    const y = b[j]
    const takeA =
      !y || (!!x && (x.orderTop < y.orderTop || (x.orderTop === y.orderTop && x.endX <= y.endX)))
    const it = takeA ? a[i++] : b[j++]
    out.push({ ...it, seq: out.length })
  }
  return out
}

/**
 * Word's balloon stack: each balloon sits at its anchor or below the previous
 * one, then the stack is pulled up from the bottom so it ends inside the
 * content band (balloons may float above their anchor, leaders slant down).
 * A stack that still overflows compacts its compactable balloons, then drops
 * non-sticky ones from the end.
 */
export function stackBalloons<T extends BalloonItem>(
  items: T[],
  bandTop: number,
  bandBottom: number,
  gapOf: (it: T) => number,
): Array<T & { top: number; compact: boolean }> {
  let sorted = [...items].sort((a, b) => a.seq - b.seq)
  const layout = (compact: boolean) => {
    const hs = sorted.map((it) =>
      compact && it.compactHeight != null ? it.compactHeight : it.height,
    )
    const tops: number[] = []
    let prevBottom = -Infinity
    sorted.forEach((it, i) => {
      const top = Math.max(it.anchorTop, prevBottom + gapOf(it))
      tops.push(top)
      prevBottom = top + hs[i]
    })
    let limit = bandBottom
    for (let i = sorted.length - 1; i >= 0; i--) {
      if (tops[i] + hs[i] > limit) tops[i] = limit - hs[i]
      limit = tops[i] - gapOf(sorted[i])
    }
    return { tops, hs, fits: sorted.length === 0 || tops[0] >= bandTop }
  }
  let compact = false
  let res = layout(false)
  if (!res.fits) {
    compact = true
    res = layout(true)
  }
  while (!res.fits) {
    const drop = sorted.map((it) => it.sticky).lastIndexOf(false)
    if (drop < 0) break
    sorted = sorted.filter((_, i) => i !== drop)
    res = layout(true)
  }
  // only sticky balloons left and still too tall: keep them stacked from the band top
  const shift = res.fits ? 0 : bandTop - res.tops[0]
  return sorted.map((it, i) => ({
    ...it,
    top: res.tops[i] + shift,
    compact: compact && it.compactHeight != null,
  }))
}

/**
 * Pagination preview: a read-only snapshot of real page slicing over the canvas's continuous
 * flow. Each page = a full content clone + overflow clipping + negative-margin offset; the
 * clone is fixed at the canvas content width (line breaks from measurement must not change),
 * and paper size/margins render per each page's section (mixed portrait/landscape across
 * sections is real). Headers/footers render per page by Word variant rules (first page /
 * odd-even), with real page numbers.
 */
/** anchored header/footer picture on a preview page; an a:srcRect crop becomes
 *  an overflow-hidden window over the scaled image (mirrors hfImgNode) */
function FloatHfImg({
  img,
  pos,
}: {
  img: HfImage
  pos: ReturnType<typeof hfFloatPagePos>
}): React.JSX.Element {
  const base: React.CSSProperties = {
    left: pos.x,
    top: pos.y,
    transform: hfFloatTransform(img, pos),
    ...(img.widthPx ? { width: img.widthPx } : {}),
    ...(img.heightPx ? { height: img.heightPx } : {}),
    ...(img.washout ? { filter: hfWashoutFilter(img.washout) } : {}),
    ...(img.behind ? {} : { zIndex: 1 }),
  }
  if (img.wordArt) {
    return (
      <span
        className="pv-watermark-img"
        aria-hidden="true"
        style={base}
        dangerouslySetInnerHTML={{ __html: wordArtSvgMarkup(img) }}
      />
    )
  }
  const c = img.crop
  if (c && img.widthPx && img.heightPx) {
    const span = (a: number, b: number) => Math.max(0.01, 1 - a - b)
    const sw = img.widthPx / span(c.l, c.r)
    const sh = img.heightPx / span(c.t, c.b)
    return (
      <span className="pv-watermark-img" aria-hidden="true" style={{ ...base, overflow: 'hidden' }}>
        <img
          src={img.dataUrl}
          alt=""
          style={{
            position: 'absolute',
            left: -c.l * sw,
            top: -c.t * sh,
            width: sw,
            height: sh,
            maxWidth: 'none',
          }}
        />
      </span>
    )
  }
  return (
    <img className="pv-watermark-img" src={img.dataUrl} alt="" aria-hidden="true" style={base} />
  )
}

export function PaginationPreview({
  section,
  canvasTop,
  sections,
  legacyBands,
  mirrorMargins = false,
  delSectBreaks,
  hfParts,
  colFlow,
  colMode,
  hf,
  watermarkDirty,
  pendingWatermark,
  blockMetaOf,
  layout,
  layoutRev,
  footnotesBeneathText,
  endnoteItems,
  resolveHf,
  clearPageGaps,
  comments,
  anchorBlocks,
  revisionView,
  balloonKinds,
  onExportPdf,
  onClose,
  suppressEscape,
  behindPrintDialog,
}: {
  /** Canvas geometry (final section): clone width, single-section fallbacks */
  section: SectionSettings
  /** canvas content-area top (px): the first section's effective top margin, the paper's padding-top */
  canvasTop: number
  /** All sections: for per-page paper geometry (empty array = single section per `section`) */
  sections: SectionInfo[]
  /** settings.xml w:mirrorMargins: even pages swap the side margins */
  mirrorMargins?: boolean
  /** section-break paragraphs whose mark is a tracked deletion (no break in markup views) */
  delSectBreaks?: Set<number>
  /** rId → header/footer parts (multi-section picks by each section's references) */
  hfParts: Record<string, HfPartInfo>
  /** Canvas column-flow geometry (non-null when the canvas column CSS is active): shared by the measuring state / clone wrap width */
  colFlow: { cols: number; colWidthPx: number; gapPx: number } | null
  /** canvas column mode: 'uniform' = whole-page CSS multicol, 'mixed' = per-block layout decorations */
  colMode: 'none' | 'uniform' | 'mixed'
  hf: HfSet
  /** compatibilityMode < 15 header band rule (hfStripGeom) */
  legacyBands?: boolean
  /** unsaved Design → Watermark edit: hide the parsed watermark shape, draw `pendingWatermark` instead */
  watermarkDirty?: boolean
  /** the unsaved watermark as the header image it will parse back as (text or picture) */
  pendingWatermark?: HfImage | null
  /** docxIndex → parse-layer pagination constraints (keepNext/widow/table-row flags) */
  blockMetaOf?: BlockMetaOf
  /** the canvas pass's layout (after any pending pass, or a new one when `fresh`): the PageCut list the sheets print */
  layout?: (fresh?: boolean) => PageLayout | null
  /** bumped after every canvas pass: the sheets rebuild on the new cuts */
  layoutRev?: number
  /** Per-page footnote collection (referencing page → entry list), for page-bottom rendering */
  /** w:footnotePr w:pos="beneathText": the note area follows the last body line instead of sitting at the page bottom */
  footnotesBeneathText?: boolean
  /** Endnote entries (placed together at the document end, take part in slicing, may continue across pages) */
  endnoteItems?: PageNoteItem[]
  /** Multi-section: the strip a section shows (pending edits, Link to Previous, inheritance) */
  resolveHf?: (
    sections: SectionInfo[],
    sectionIndex: number,
    kind: 'header' | 'footer',
    variant: 'default' | 'first' | 'even',
  ) => { value: HeaderFooter | null; images?: HfImage[]; owner?: number }
  /**
   * Clears the canvas page-gap decorations before the snapshot measure. In-table
   * gap/repeated-header widgets are extra <tr>s that consume rowspan slots, so a
   * vMerge-heavy table measures with collapsed columns (exploding row heights)
   * while they are present; the canvas rebuilds them on its next debounced
   * remeasure after the snapshot.
   */
  clearPageGaps?: () => void
  /** Comment threads: open threads print as Word-style margin balloons (scaled sheet + markup strip) */
  comments?: CommentInfo[]
  /** parsed blocks: anchors for comment ranges that never produced a text mark (cross-paragraph / image / table ranges) */
  anchorBlocks?: AnchorBlock[]
  /** editor view while the canvas is in balloon mode (print view, All Markup): deletions / format changes print as margin balloons */
  revisionView?: EditorView | null
  /** revision balloon kinds of the current Show Markup → Balloons mode */
  balloonKinds?: ReadonlySet<RevSpot['kind']>
  onExportPdf: () => void
  onClose: () => void
  /** While true (e.g. the print dialog is stacked on top), Escape must not close the preview */
  suppressEscape?: boolean
  /** the print dialog is open: the preview is only its hidden print source */
  behindPrintDialog?: boolean
}) {
  const { t } = useI18n()
  const [slices, setSlices] = useState<PageCut[]>([])
  const [splitCss, setSplitCss] = useState('')
  const [hoistWinCss, setHoistWinCss] = useState('')
  const [vertCss, setVertCss] = useState('')
  const [leakCss, setLeakCss] = useState('')
  const [pageNotes, setPageNotes] = useState<PageNoteItem[][]>([])
  /** per page: data-idx of the table whose header the page repeats (probe hook) */
  const [headerTables, setHeaderTables] = useState<Array<string | undefined>>([])
  /** probe hook: every pass block's data-idx and the top the clone is expected to paint it at */
  const [blockTops, setBlockTops] = useState<Array<[string, number]>>([])
  /** per-page flow-coordinate bottom of the body text (beneathText footnote anchor) */
  const [textEnds, setTextEnds] = useState<number[]>([])
  /** Top Y of the endnote area (virtual coordinates); null = no endnotes */
  const [endnotesTop, setEndnotesTop] = useState<number | null>(null)
  const [html, setHtml] = useState('')
  /** non-null = pruned-clone mode (large documents): per-page windows instead of full clones */
  const [cloneKids, setCloneKids] = useState<CloneChild[] | null>(null)
  const [cellPageBoxes, setCellPageBoxes] = useState<CellPageBox[]>([])
  const [relXPictures, setRelXPictures] = useState<PageRelXPicture[]>([])
  /** Live section list: a section whose break block was deleted (unsaved) merges into the next, matching the canvas */
  const [secs, setSecs] = useState<SectionInfo[]>(sections)
  /** open comment threads' anchors (clone flow coordinates), in document order */
  const [commentSpots, setCommentSpots] = useState<CommentSpot[]>([])
  const [revSpots, setRevSpots] = useState<RevSpot[]>([])

  const canvasContentW = twipsToPx(section.pageWidth - section.marginLeft - section.marginRight)
  /** the canvas content width as Blink laid it out (snapshot), null until measured or when no block spans it */
  const [snappedWrapW, setSnappedWrapW] = useState<number | null>(null)
  /** the canvas editor's lang (hyphenation documents): Blink breaks lines per language */
  const [cloneLang, setCloneLang] = useState<string | undefined>(undefined)
  const snappedPass = useRef<PageLayout | null>(null)
  /** Settings of the page's section (single-section documents fall back to the canvas geometry) */
  const settingsOf = (slice: PageSlice): SectionSettings =>
    secs[Math.min(slice.section, secs.length - 1)]?.settings ?? section
  /** Clone wrap width = the section's measurement width (columned canvas = column width);
   *  differing-width sections wrap at their own content width (per-block width decorations ride the clone) */
  const wrapWOf = (sectionIdx: number): number => {
    if (colFlow) return colFlow.colWidthPx
    const w = sectionPageBox(
      secs[Math.min(sectionIdx, Math.max(secs.length - 1, 0))]?.settings ?? section,
    ).contentWidth
    return snappedWrapW !== null && Math.abs(w - canvasContentW) < 0.01 ? snappedWrapW : w
  }

  useEffect(() => {
    const pm = canvasEditorEl()
    if (!pm) return
    // before the gaps go: a pending canvas pass runs now and rebuilds its
    // widgets, which the clones below must not carry
    let pass = layout?.()
    // a pass flushed by the previous run re-enters here with its sheets built
    if (pass && pass === snappedPass.current) return
    // ProseMirror redraws the nodes a snapshot stamped: a pass snapshotted
    // before (preview closed and reopened) holds detached elements
    const detached = (el: Element | null | undefined) => el != null && !el.isConnected
    if (
      pass?.flow?.blocks.some((b) => detached(b.el)) ||
      pass?.flow?.floats.some((f) => detached(f.el))
    )
      pass = layout?.(true)
    clearPageGaps?.()
    // the sheets print the canvas pass's cuts and nothing else: with no pass
    // (no document) there is nothing to print; an empty document still has one cut
    if (!pass?.flow || pass.cuts.length === 0) {
      snappedPass.current = null
      setSlices([])
      return
    }
    // Measure at zoom 1: CSS zoom rounds every box to device pixels, so dividing
    // zoomed rects by the factor drifts from the zoom-1 clones the pages render
    // (a long table accumulates rows of error — repeated headers overprint the
    // first data row and page cuts land mid-row). Neutralizing the canvas zoom
    // for the snapshot makes measurement and clone layout share one geometry.
    const zoomHost = pm.closest<HTMLElement>('.doc-zoom')
    const savedZoom = zoomHost?.style.zoom ?? ''
    if (zoomHost) zoomHost.style.zoom = '1'
    const factor = 1
    // switch the columned canvas to the single-flow measuring state (uniform: CSS columns
    // off, width = column width; mixed: block translates off), matching engine column-flow
    // coordinates. vAlign documents carry the same visual translates on the canvas
    // (vAlignShiftSpecs) and the preview applies its own vOffset, so they must be
    // neutralized here too or the shifted rects double-apply.
    const measureNeutralize =
      colMode !== 'none' ||
      section.vAlign === 'center' ||
      section.vAlign === 'bottom' ||
      sections.some(
        (s) =>
          s.settings.vAlign === 'center' ||
          s.settings.vAlign === 'bottom' ||
          sectionVertical(s.settings),
      )
    if (measureNeutralize) pm.classList.add('measuring-columns')
    try {
      const origin = canvasFlowOrigin(pm, canvasTop, factor)
      // the sheets position against the very flow the pass sliced (its blocks,
      // floating boxes, sections, endnotes, spill, flow height). The flow is
      // reused as is: preparing it again would re-fold the pass's live blocks
      // (foldSectionStartSpaceBefore is not idempotent) and re-pull its boxes.
      const flow = pass.flow
      const { blocks, floats } = flow
      // the shadow re-slices a twin of the pass flow with the preview's own
      // header/footer resolver and counts where it disagrees (previewCutMismatch,
      // asserted 0 by the invariants e2e)
      const shadow =
        // page-cut: allow — the shadow re-slice is a dev / probe twin, never the printed result
        import.meta.env.DEV ||
        (window as unknown as { __previewShadow?: boolean }).__previewShadow === true
      /** each section's header/footer variant heights → body push-down (matching
       *  the canvas); the shadow only */
      const previewHfHeights = (live: SectionInfo[]): SectionHfHeights[] => {
        const refs = effectiveHfRefs(live)
        const partOf = (i: number, kind: 'header' | 'footer', variant: HfVariant) => {
          const rId = refs[i]?.[kind]?.[variant]
          return rId ? hfParts[rId] : undefined
        }
        // page-cut: allow — shadow resolver
        return sectionHfHeights(
          live,
          resolveHf
            ? (i, kind, variant) => resolveHf(live, i, kind, variant)
            : (i, kind, variant) => {
                const part = partOf(i, kind, variant)
                const value = part
                  ? { text: part.text, pageNumber: part.hasPageNumber, paras: part.paras }
                  : null
                if (variant !== 'default') return { value, images: part?.images }
                // the parsed set carries the last section's default strips
                const last = i === live.length - 1
                return {
                  value: last ? (kind === 'header' ? hf.header : hf.footer) : value,
                  images: part?.images?.length ? part.images : last ? hf.images?.[kind] : undefined,
                }
              },
          { titlePg: hf.titlePg, evenOddHf: hf.evenOddHf, legacyBands },
        )
      }
      const singleHfPx: SectionHfHeights = shadow
        ? // page-cut: allow — shadow resolver
          hfVariantHeightsPx(
            (kind, variant) => {
              const strips = pageHfStrips(variant, undefined, hf)
              return kind === 'header'
                ? { value: strips.header, images: strips.headerImages }
                : { value: strips.footer, images: strips.footerImages }
            },
            () => canvasContentW,
            hfStripGeom(section, legacyBands),
            hf.titlePg,
            hf.evenOddHf,
          )
        : pass.hfHeights[0]
      const singleGeom: SectionGeom = shadow
        ? {
            contentHeight:
              twipsToPx(section.pageHeight) -
              effectiveTopPx(section, singleHfPx.headerPx) -
              effectiveBottomPx(section, singleHfPx.footerPx),
            forceBreak: false,
            ...singleVariantGeomOf(section, singleHfPx),
            ...(sections[0]?.pageNumberStart !== undefined
              ? { pageNumberStart: sections[0].pageNumberStart }
              : {}),
            ...(colFlow ? { cols: colFlow.cols } : {}),
          }
        : pass.geoms[0]
      const splitOut: SliceOutputs = { rowSplits: [], floatSplits: [] }
      /** the canvas pass's pipeline (pagination-pass) for the shadow twin */
      const passCtx: PassContext = {
        sections,
        section,
        delSectBreaks,
        endnoteItems: endnoteItems ?? [],
        endnoteSeparatorPx: noteSeparatorPx('endnote'),
        hfHeightsOf: (live) => (shadow ? previewHfHeights(live) : pass.hfHeights),
        singleHfPx,
        singleGeom,
        evenOddHf: hf.evenOddHf,
        titlePg: hf.titlePg,
        // when the canvas column layout is inactive, measure as full-width single flow
        dropCols: colMode === 'none',
        blockMetaOf,
        ...(sections.some((s) => sectionVertical(s.settings))
          ? { inlineExtraOf: blockInlineExtraPx }
          : {}),
        factor,
        out: splitOut,
      }
      const live = flow.secList ?? []
      setSecs(live)
      setEndnotesTop(flow.endnotesTop)
      // the pass's cuts are printed against this snapshot's blocks, so the
      // snapshot carries the same parse-layer facts the slicer saw (footnote
      // reservations and bands, keep flags) in every build: a DEV-only shadow
      // used to apply them and hid that production never did (split
      // paragraphs printed every footnote on their first fragment's page)
      applyBlockMeta(blocks, blockMetaOf, factor)
      const { cuts: computed, geoms: liveGeoms, hfHeights: liveHfHs, rowSplits } = pass
      let previewCutDiag: string[] = []
      const shadowed = shadow
      if (shadow) {
        // the shadow slices a twin: it must not swap the pass's spill block,
        // and its boxes start where the pass measured them (sliceFlow pulls
        // them up itself)
        const twinFloats = floats.map((f) => ({ ...f }))
        applyFloatPullUps(
          twinFloats,
          pass.floatPullUps.map((p) => ({ ...p, dyPx: -p.dyPx })),
        )
        const twin = { ...flow, blocks: [...blocks], floats: twinFloats }
        previewCutDiag = pageCutDiffs(computed, sliceFlow(twin, passCtx).cuts)
      }
      const lastIdx = computed.length - 1
      // the cuts name the page every floated / anchor-shifted block and every
      // floating box lands on; the sheets never look a page up by Y
      const owners = pageOwners(computed)
      const pageOfBlock = (i: number) => owners.blocks.get(i) ?? 0
      const pageOfBox = (i: number) => owners.boxes.get(i) ?? 0
      const boxIndex = new Map(floats.map((f, i) => [f, i] as const))
      const boxPage = (f: FloatBox) => pageOfBox(boxIndex.get(f) ?? -1)
      /** flow-coordinate bottom of page i's clip window (the last page opens to full capacity unless vertically aligned) */
      const winEndOf = (s: PageCut, i: number) => {
        const vAlign =
          live.length > 0
            ? live[Math.min(s.section, live.length - 1)]?.settings.vAlign
            : section.vAlign
        const full = i === lastIdx && vAlign !== 'center' && vAlign !== 'bottom'
        return full ? s.start + s.capacityPx : Math.min(s.end, s.start + s.capacityPx)
      }
      // anchor-shifted textbox wrappers paint off their flow slot: the owning
      // page hides the ride-along copies and its boxes are shifted back inside
      // that page's window (Word keeps a box on its anchor's page)
      blocks.forEach((b, bi) => {
        const el = b.el
        if (!b.anchorShifted || !el) return
        const pg = pageOfBlock(bi)
        el.dataset.pinPage = String(pg)
        const slice = computed[pg]
        // column windows and repeated table headers reshape the clip: no lift there
        if (slice.regions || slice.repeatHeader) {
          el.style.removeProperty('--pv-anchor-lift')
          return
        }
        const paintTop = b.top + (b.leadFoldPx ?? 0) + anchorShiftPx(el)
        const { top, bottom } = shapeBoxExtent(el, paintTop, factor)
        const lift =
          top <= bottom ? anchorBoxLift(top, bottom, slice.start, winEndOf(slice, pg)) : 0
        if (Math.abs(lift) > 0.5) el.style.setProperty('--pv-anchor-lift', `${lift}px`)
        else el.style.removeProperty('--pv-anchor-lift')
      })
      // stamp each page-pinned box's owning page on its canvas wrapper before
      // cloning so the per-page CSS rules can hide the copies on other pages
      for (const f of floats) {
        if (f.pinned) {
          const wrap = f.el.closest<HTMLElement>('.doc-protected-pagepinned')
          if (wrap) wrap.dataset.pinPage = String(boxPage(f))
        } else if (f.pageRelV) {
          // page-relative V boxes ride every clone too (absolute boxes ignore
          // the pv-clip overflow): nearby pages showed visible duplicates —
          // a cover band re-painted over the TOC two pages later (real_run2/69).
          // Stamped per box: sibling boxes of one anchor can own different pages.
          f.el.dataset.pinPage = String(boxPage(f))
        }
      }
      // wrappers whose every box is page-relative V get un-positioned in the
      // preview (like pinned covers): the boxes escape the pv-clip — sized to
      // the page's flow content, it cut a page-bottom cover band clean off —
      // and resolve against the page box at their page-relative offsets
      for (const f of floats) {
        if (!f.pageRelV) continue
        const wrap = f.el.closest<HTMLElement>(
          '.doc-protected-floating, .doc-img-float, .doc-anchor-origin',
        )
        if (!wrap) continue
        const boxes = anchoredBoxesOf(wrap)
        // static siblings (an inline drawing sharing the paragraph) neither
        // need the shared containing block nor ride the page-margin translate:
        // only absolutely positioned siblings gate the un-positioning; zero-height
        // ones (degenerate guide lines) are invisible and not measured as floats.
        // A single shape's float slot sits on the chrome rotator above the box.
        const absBoxes = boxes.filter(isLaidOutAbsoluteBox)
        if (absBoxes.length > 0 && absBoxes.every((b) => b.dataset.pageRelV === '1')) {
          wrap.dataset.pvPagerel = '1'
        } else {
          delete wrap.dataset.pvPagerel
        }
      }
      // paragraph-anchored floats spilling past their page's flow window: the
      // pv-clip cuts them on the owning page and the next page's window
      // repaints the spilled part at its top (all windows share one flow
      // clone). Un-position the wrapper so the boxes escape the clip and
      // re-pin to the page box at the anchor's content-area Y; the stamped
      // owning page lets other clones hide the ride-along copies.
      for (const el of hoistMarkedEls(pm)) {
        delete el.dataset.pvHoist
        delete el.dataset.pvHoistSlot
        delete el.dataset.pinPage
        delete el.dataset.pinWin
        el.style.removeProperty('--pv-hoist-dy')
      }
      const hoistWins = new Map<number, Set<string>>()
      // the hoist translate uses each page's own top margin (a titlePg first page
      // renders the first-variant header), so the paper clamp must too
      const firstOfSection = (pg: number) =>
        (pg === 0 || computed[pg - 1].section !== computed[pg].section) &&
        !computed[pg].continuedSection
      const hoistPageTopPx = (pg: number, section: number): number => {
        const first = liveHfHs[section]?.firstHeaderPx
        return firstOfSection(pg) && first !== undefined && live[section]
          ? effectiveTopPx(live[section].settings, first)
          : (liveGeoms[section]?.topPx ?? 0)
      }
      // run-anchored pictures with a page-relative X: the anchor paragraph
      // keeps its position (and lead shift) in the clone; the picture itself
      // draws once per owning page as a page overlay, every clone copy hidden
      const relXPics: PageRelXPicture[] = []
      for (const f of floats) {
        if (!f.el.matches('.doc-inline-img-anchor > img[data-page-rel-x]') || f.pageRelV) continue
        delete f.el.dataset.pvOverlaid
        const pg = boxPage(f)
        const slice = computed[pg]
        if (!slice || slice.repeatHeader) continue
        const win = hoistWindow(slice, f.anchorTop)
        if (!win) continue
        f.el.dataset.pvOverlaid = '1'
        relXPics.push({
          page: pg,
          offX: parseFloat(f.el.dataset.pageRelXPx ?? '') || 0,
          fromPaperEdge: f.el.dataset.pageRelXFrom === 'page',
          topY: win.dy + f.top - f.anchorTop,
          behind: f.el.classList.contains('doc-inline-img--behind'),
          html: pageRelXPictureHtml(f.el),
        })
      }
      setRelXPictures(relXPics)
      const wrapFloats = new Map<HTMLElement, FloatBox[]>()
      for (const f of floats) {
        const wrap = f.el.closest<HTMLElement>('.doc-protected-floating, .doc-img-float')
        if (!wrap) continue
        const list = wrapFloats.get(wrap)
        if (list) list.push(f)
        else wrapFloats.set(wrap, [f])
      }
      for (const [wrap, all] of wrapFloats) {
        // pinned / page-relative wrappers already position against the page
        if (wrap.dataset.pvPagerel === '1' || wrap.classList.contains('doc-protected-pagepinned'))
          continue
        const plan = hoistPlan(all)
        if (!plan) continue
        const { boxes: wfs, mixed } = plan
        // the hoist CSS re-pins against the page box, so each slot (left px,
        // left:50%, right:0) gets its own column-to-page correction
        const slot = hoistSlotOf(wfs[0])
        if (!slot || !wfs.every((f) => hoistSlotOf(f) === slot)) continue
        const pg = boxPage(wfs[0])
        const slice = computed[pg]
        if (!slice || slice.repeatHeader) continue
        const win = hoistWindow(slice, wfs[0].anchorTop)
        if (!win) continue
        // in a multi-column region only page-relative X boxes re-pin to the
        // page (the column clone shifts and clips them); column-relative boxes
        // stay with their column, whose left the page has no way to restore
        const pageRelX = wfs.every((f) => f.el.dataset.pageRelX === '1')
        if (win.multiCol && !pageRelX) continue
        const spills = wfs.some((f) => f.top + f.height > win.end + 1 || f.top < win.start - 1)
        if (!spills && !win.multiCol && !mixed) continue
        wrap.dataset.pvHoist = '1'
        wrap.dataset.pvHoistSlot = slot
        wrap.dataset.pinPage = String(pg)
        // a regioned page clones the flow once per column window; the escaped
        // boxes would paint once per window, each copy over the earlier
        // windows' text, so only the anchor's own window keeps them
        const winKey = win.start.toFixed(1)
        wrap.dataset.pinWin = winKey
        ;(hoistWins.get(pg) ?? hoistWins.set(pg, new Set()).get(pg)!).add(winKey)
        if (mixed) wrap.dataset.pvPagerel = '1'
        const dy = hoistPaperClamp(
          win.dy,
          hoistPageTopPx(pg, slice.section),
          wfs.map((f) => f.top - f.anchorTop),
        )
        wrap.style.setProperty('--pv-hoist-dy', `${dy}px`)
      }
      setSplitCss(rowSplitCss(rowSplits, blocks, computed))
      setHoistWinCss(hoistWindowCss(hoistWins))
      setVertCss(
        liveGeoms.some((g) => g.vertical) ? verticalPageCss(blocks, computed, live, liveGeoms) : '',
      )
      setLeakCss(leadLeakCss(blocks, computed))
      setSlices(computed)
      setBlockTops(
        blocks.flatMap((b) =>
          b.el?.dataset.idx !== undefined
            ? [
                [b.el.dataset.idx, cloneChildGeom(b, { vTop: b.top, h: b.height }).vTop] as [
                  string,
                  number,
                ],
              ]
            : [],
        ),
      )
      setHeaderTables(
        computed.map((c) =>
          c.repeatHeader && (c.lead.kind === 'row' || c.lead.kind === 'cell')
            ? blocks[c.lead.block]?.el?.dataset.idx
            : undefined,
        ),
      )
      // read-only snapshot for e2e: the cuts the sheets print and the shadow's verdict
      ;(window as unknown as Record<string, unknown>).__previewDebug = {
        slices: computed,
        shadowed,
        passGen: pass.passGen,
        previewCutMismatch: previewCutDiag.length,
        previewCutDiag,
        // the snapshot's block geometry, to locate where a shadow disagreement starts
        blocks: shadowed
          ? blocks.map((b) => ({
              top: b.top,
              height: b.height,
              spaceBeforePx: b.spaceBeforePx,
              spaceAfterPx: b.spaceAfterPx,
              idx: b.el?.dataset.idx,
            }))
          : undefined,
      }
      setPageNotes(pass.pageNotes)
      setTextEnds(footnotesBeneathText ? pageTextEnds(blocks, computed) : [])
      // comment anchors, measured in the same neutralized geometry as the blocks;
      // content-relative X: the pm element is the padded .doc-page, but the
      // preview clones strip that padding and sit inside the sheet's own
      const pmContentLeft = canvasContentLeft(pm)
      setCommentSpots(
        comments && comments.length > 0
          ? measureCommentSpots(pm, comments, anchorBlocks, origin, factor, pmContentLeft)
          : [],
      )
      setRevSpots(
        revisionView
          ? measureRevisionSpots(
              revisionView,
              pm,
              anchorBlocks,
              origin,
              factor,
              pmContentLeft,
              balloonKinds,
            )
          : [],
      )
      // Per-page full clones explode on large documents (pages × doc DOM →
      // renderer OOM / "Promise was collected" during printToPDF). Past the
      // budget, snapshot per-block geometry and render pruned windows instead.
      const kidEls = Array.from(pm.children) as HTMLElement[]
      setCloneLang(pm.getAttribute('lang') ?? undefined)
      setSnappedWrapW(canvasWrapWidth(canvasChildWidths(kidEls, factor), canvasContentW))
      setCellPageBoxes(collectCellPageBoxes(pm))
      if (shouldPruneClones(computed.length, kidEls.length, pm.getElementsByTagName('*').length)) {
        const metas = cloneChildrenOf(kidEls, blocks, origin, factor)
        setCloneKids(metas)
        setHtml('')
      } else {
        setCloneKids(null)
        setHtml(Array.from(pm.children, (c) => cloneBlockHtml(c as HTMLElement)).join(''))
      }
      snappedPass.current = pass
    } finally {
      if (measureNeutralize) pm.classList.remove('measuring-columns')
      if (zoomHost) zoomHost.style.zoom = savedZoom
    }
    // snapshot: measure once on open; deps intentionally empty
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [layoutRev])

  useEffect(() => {
    // stopPropagation cannot shield this window-level listener from the print
    // dialog's own Escape handler (same target, same phase), so the dialog
    // suppresses it via prop while it is stacked on top
    if (suppressEscape) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onClose()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose, suppressEscape])

  // line numbers (w:lnNumType) ride the cloned pages: measured after the sheets mount;
  // markup spots rescale the sheets, so they re-measure too
  useEffect(() => {
    const root = previewScrollEl()
    if (root) syncPreviewLineNumbers(root, secs, blockMetaOf)
  }, [secs, blockMetaOf, slices, commentSpots, revSpots])

  // probe: where the clone's table lands inside each repeated-header window
  // (expected at its top) and how tall its header row is against the reserved
  // height; read after mount, published for the export dump
  useEffect(() => {
    publishPreviewProbes(blockTops)
  }, [slices, headerTables, blockTops, commentSpots, revSpots])

  const multiSection = secs.length > 1
  const commentById = useMemo(() => new Map((comments ?? []).map((c) => [c.id, c])), [comments])
  const effRefs = useMemo(() => effectiveHfRefs(secs), [secs])
  const nums = useMemo(() => slices.map((s) => s.pageNo), [slices])
  const sectionPages = useMemo(() => sectionPageCounts(slices), [slices])

  // line positions of endnote entries in virtual coordinates (matching appendEndnotesBlock's line boxes)
  const endnoteSepPx = noteSeparatorPx('endnote')
  const endnoteRows = useMemo(() => {
    if (endnotesTop === null || !endnoteItems || endnoteItems.length === 0) return []
    let off = endnotesTop
    return endnoteItems.map((item, i) => {
      const height = (i === 0 ? endnoteSepPx : 0) + item.height
      const row = { item, top: off, height, withSeparator: i === 0 }
      off += height
      return row
    })
  }, [endnotesTop, endnoteItems, endnoteSepPx])

  const toHf = (rId: string | undefined): HeaderFooter | null => {
    const part = rId ? hfParts[rId] : undefined
    if (!part) return null
    return {
      text: part.text,
      pageNumber: part.hasPageNumber,
      paras: part.paras.length > 0 ? part.paras : undefined,
    }
  }

  /** Single section: the resolver's pending state (strip edits land there, not in `hf`); multi-section: pick by each section's references */
  const hfFor = (i: number): PageHf => {
    const slice = slices[i]
    // the cut names the variant the slicer reserved the strip heights for
    const variant = slice.hf.variant
    if (!multiSection)
      return pageHfStrips(variant, resolveHf && ((kind) => resolveHf(secs, 0, kind, variant)), hf)
    const sec = secs[Math.min(slice.section, secs.length - 1)]
    const refs = effRefs[Math.min(slice.section, effRefs.length - 1)]
    const ownW = twipsToPx(
      sec.settings.pageWidth - sec.settings.marginLeft - sec.settings.marginRight,
    )
    const inheritedW = (kind: 'header' | 'footer') => {
      const owner = resolveHf?.(secs, slice.section, kind, variant).owner
      const w = hfPartWidthPx(secs, slice.section, kind, variant, owner)
      return Math.abs(w - ownW) > 0.5 ? w : undefined
    }
    const widths = { headerColumnPx: inheritedW('header'), footerColumnPx: inheritedW('footer') }
    if (resolveHf) {
      const h = resolveHf(secs, slice.section, 'header', variant)
      const f = resolveHf(secs, slice.section, 'footer', variant)
      return {
        header: h.value,
        footer: f.value,
        headerImages: h.images,
        footerImages: f.images,
        ...widths,
      }
    }
    const headerRId = refs.header[variant]
    const footerRId = refs.footer[variant]
    return {
      header: toHf(headerRId),
      footer: toHf(footerRId),
      headerImages: headerRId ? hfParts[headerRId]?.images : undefined,
      footerImages: footerRId ? hfParts[footerRId]?.images : undefined,
      ...widths,
    }
  }

  return (
    <div
      className={behindPrintDialog ? 'pagination-preview pv-behind-dialog' : 'pagination-preview'}
    >
      <div className="pv-toolbar">
        <span className="pv-title">{t('appPaginationPreview')}</span>
        <span className="pv-count">{t('appTotalPagesN', { n: slices.length })}</span>
        <span className="pv-hint">{t('appPvHint')}</span>
        <button className="pv-close" data-tip={t('appPvExportTip')} onClick={onExportPdf}>
          {t('appExportPdf')}
        </button>
        <button className="pv-close" onClick={onClose}>
          {t('appClose')}
        </button>
      </div>
      <style>{pinnedCloneCss(slices.length)}</style>
      {splitCss && <style>{splitCss}</style>}
      {hoistWinCss && <style>{hoistWinCss}</style>}
      {vertCss && <style>{vertCss}</style>}
      {leakCss && <style>{leakCss}</style>}
      {/* aria-hidden: the cloned page stack is a visual print preview; exposing
          its full-document DOM to the accessibility tree overflows Blink's AX
          update queue on long documents and crashes the renderer */}
      <div className="pv-scroll" aria-hidden="true">
        {slices.map((slice, i) => {
          const parts = hfFor(i)
          const s = settingsOf(slice)
          // this page's side margins: a mirrored document swaps them on even pages
          const sideMargins = pageMargins(s, nums[i], mirrorMargins)
          const mL = twipsToPx(sideMargins.left)
          const mR = twipsToPx(sideMargins.right)
          const pageBox = sectionPageBox(s)
          const pageW = pageBox.width
          const pageH = pageBox.height
          const secContentW = pageBox.contentWidth
          // effective margins after this page's variant header/footer push-down (an over-tall
          // header pushes the body down): the cut's reserved heights, as the slicer used them
          const mTop = effectiveTopPx(s, slice.hf.headerPx)
          const footerReservedPx = slice.hf.footerPx
          const mBottom = effectiveBottomPx(s, footerReservedPx)
          const contentH = pageH - mTop - mBottom
          const notes = pageNotes[i] ?? []
          const notesH =
            notes.length > 0
              ? notes.reduce((sum, n) => sum + n.height, 0) + noteSeparatorPx('footnote')
              : 0
          // page vertical alignment (sectPr w:vAlign): content of non-full pages shifts down as a whole
          const usedH = Math.min(slice.end - slice.start, contentH)
          const vSpare = Math.max(0, contentH - usedH)
          // (vertical-text pages: the slice span is the sideways fill, not leftover height)
          const vOffset = sectionVertical(s)
            ? 0
            : s.vAlign === 'center'
              ? vSpare / 2
              : s.vAlign === 'bottom'
                ? vSpare
                : 0
          const noteArea = noteAreaPlacement(notesH, textEnds[i], slice, {
            pageH,
            mTop,
            mBottom,
            headerH: slice.repeatHeader?.height ?? 0,
            vOffset,
          })
          // the window opens up into the top margin: fully on the first page (nothing
          // precedes the flow there, so a side-wrapped picture anchored above its
          // paragraph paints into the margin like Word), else by a lifted table's hang
          const openTop = i === 0 ? mTop : Math.min(slice.liftTop ?? 0, mTop)
          // a cut table's incoming edge: the first window on the page grows one pixel
          // up into the top margin so the whole straddling border shows in place;
          // a page opening on a table's bottom edge drops that pixel instead (the
          // previous page's window keeps the whole bottom border)
          const seam = seamWindow(slice, slices[i + 1])
          const seamLift = Math.max(seam.lift, 0)
          const bodyLift = slice.repeatHeader ? 0 : seam.lift
          // page numbers display in the owning section's number format (w:pgNumType w:fmt)
          const pageNoText = slice.display
          // page border (w:pgBorders): drawn per sheet; w:display counts pages within the section
          const pageBorder = pageBorderStyleOf(s)
          const firstOfSection = i === 0 || slices[i - 1].section !== slice.section
          const drawPageBorder =
            pageBorder &&
            !(pageBorder.display === 'firstPage' && !firstOfSection) &&
            !(pageBorder.display === 'notFirstPage' && firstOfSection)
          // Word print-markup: uniform scale to make room for the markup strip,
          // scaled sheet centered vertically (see MARKUP_EXTRA_W)
          const markupOn = commentSpots.length > 0 || revSpots.length > 0
          const markupK = pageW / (pageW + MARKUP_EXTRA_W)
          const markupOffY = (pageH - pageH * markupK) / 2
          const markupTransform = {
            transform: `translateY(${markupOffY}px) scale(${markupK})`,
            transformOrigin: '0 0',
          } as const
          return (
            <div
              key={i}
              className="pv-page"
              data-pv-page={i}
              data-pv-section={slice.section}
              // inert: cloned pages carry natively focusable nodes (links,
              // contenteditable cells); keyboard focus must not enter the
              // aria-hidden subtree. Kept off .pv-scroll so it stays scrollable.
              inert
              style={
                {
                  width: pageW,
                  height: pageH,
                  '--pv-page-h': `${pageH}px`,
                  '--page-w': `${pageW}px`,
                  '--page-h': `${pageH}px`,
                  '--section-content-w': `${secContentW}px`,
                  '--header-dist': `${pageBox.headerDist}px`,
                  '--footer-dist': `${pageBox.footerDist}px`,
                  '--pv-mr': `${mR}px`,
                  '--pv-ml': `${mL}px`,
                  '--pv-mt': `${mTop}px`,
                  '--pv-mt-page': `${twipsToPx(s.marginTop)}px`,
                } as React.CSSProperties
              }
            >
              <div
                className="pv-sheet"
                style={{
                  padding: `${mTop}px ${mR}px ${mBottom}px ${mL}px`,
                  ...(markupOn ? markupTransform : {}),
                }}
              >
                {drawPageBorder && (
                  <div
                    className={`pv-page-border${pageBorder.zOrder === 'back' ? ' pv-page-border-back' : ''}`}
                    aria-hidden="true"
                    style={{
                      top: pageBorder.sides.top?.insetPx ?? 0,
                      right: pageBorder.sides.right?.insetPx ?? 0,
                      bottom: pageBorder.sides.bottom?.insetPx ?? 0,
                      left: pageBorder.sides.left?.insetPx ?? 0,
                      borderTop: pageBorder.sides.top?.css,
                      borderRight: pageBorder.sides.right?.css,
                      borderBottom: pageBorder.sides.bottom?.css,
                      borderLeft: pageBorder.sides.left?.css,
                      boxShadow: pageBorder.boxShadow,
                    }}
                  >
                    {(['top', 'right', 'bottom', 'left'] as const).map((name) => {
                      const art = pageBorder.sides[name]?.art
                      return art ? (
                        <div
                          key={name}
                          className="page-border-art"
                          style={pageBorderArtStripStyle(name, art) as React.CSSProperties}
                        />
                      ) : null
                    })}
                  </div>
                )}
                {[
                  ...(watermarkDirty && pendingWatermark ? [pendingWatermark] : []),
                  ...(parts.headerImages ?? []).filter(
                    (img) =>
                      img.floating &&
                      !hfImageHangsOnPara(img) &&
                      !(watermarkDirty && (img.wordArt || img.watermark)),
                  ),
                ].map((img, k) => {
                  // picture watermark (anchored image in the header): drawn once
                  // per page behind the body (negative z-index; .pv-page isolates)
                  const pos = hfFloatPagePos(img, {
                    pageW,
                    pageH,
                    marginLeft: mL,
                    marginRight: mR,
                    marginTop: mTop,
                    marginBottom: mBottom,
                    headerDist: pageBox.headerDist,
                    sectMarginTop: twipsToPx(s.marginTop),
                    sectMarginBottom: twipsToPx(s.marginBottom),
                  })
                  return <FloatHfImg key={`wm${k}`} img={img} pos={pos} />
                })}
                {cellPageBoxes
                  .filter((b) => b.page === i)
                  .map((b, k) => (
                    <div
                      key={`cpb${k}`}
                      className="pv-cell-page-box"
                      aria-hidden="true"
                      style={{
                        left: b.fromPaperEdge ? b.offX : mL + b.offX,
                        top: b.fromPage ? b.offY : mTop + b.offY,
                        // the isolated page clone renders after the overlay:
                        // front boxes need a level above it, behind ones below
                        zIndex: b.behind ? -1 : 1,
                      }}
                      dangerouslySetInnerHTML={{ __html: b.html }}
                    />
                  ))}
                {relXPictures
                  .filter((b) => b.page === i)
                  .map((b, k) => (
                    <div
                      key={`rxp${k}`}
                      className="pv-cell-page-box"
                      aria-hidden="true"
                      style={{
                        left: b.fromPaperEdge ? b.offX : mL + b.offX,
                        // paragraph-relative Y rides the body window's vAlign shift;
                        // Word keeps the picture on the paper (hoistPaperClamp)
                        top: Math.max(0, mTop + (slice.regions ? 0 : vOffset) + b.topY),
                        zIndex: b.behind ? -1 : 1,
                      }}
                      dangerouslySetInnerHTML={{ __html: b.html }}
                    />
                  ))}
                {(parts.footerImages ?? [])
                  .filter((img) => img.floating && !hfImageHangsOnPara(img))
                  .map((img, k) => {
                    // anchored footer picture (seal beside the address block):
                    // paragraph-relative offsets measure from the footer strip top
                    const pos = hfFloatPagePos(img, {
                      pageW,
                      pageH,
                      marginLeft: mL,
                      marginRight: mR,
                      marginTop: mTop,
                      marginBottom: mBottom,
                      headerDist: pageBox.headerDist,
                      sectMarginTop: twipsToPx(s.marginTop),
                      sectMarginBottom: twipsToPx(s.marginBottom),
                      paraOriginY: pageH - pageBox.footerDist - footerReservedPx,
                    })
                    return <FloatHfImg key={`fwm${k}`} img={img} pos={pos} />
                  })}
                {/* image-only parts (logo headers) have a null text value but must still render */}
                {(parts.header || parts.headerImages?.some((img) => !img.floating)) && (
                  <HeaderFooterArea
                    kind="header"
                    value={parts.header ?? { text: '' }}
                    images={parts.headerImages?.filter(
                      (img) => !img.floating || hfImageHangsOnPara(img),
                    )}
                    readOnly
                    onCommit={() => {}}
                    pageNo={pageNoText}
                    pageTotal={slices.length}
                    sectionPages={sectionPages[i]}
                    style={
                      parts.headerColumnPx !== undefined
                        ? { width: parts.headerColumnPx }
                        : undefined
                    }
                    boxGeom={{ ...hfStripGeom(s, legacyBands), columnW: parts.headerColumnPx }}
                  />
                )}
                {slice.repeatHeader && !slice.regions && (
                  // tblHeader repeated headers: a broken table's page first renders a clone of the source table's header rows
                  // (the engine already reserved repeatHeader.height on this page)
                  <div
                    className="pv-clip pv-repeat-header"
                    data-pv-table={headerTables[i]}
                    data-pv-reserved={slice.repeatHeader.height}
                    data-pv-lift={seamLift}
                    style={{ height: slice.repeatHeader.height + seamLift, marginTop: -seamLift }}
                  >
                    <div
                      className="pv-offset"
                      style={{
                        marginTop: -(slice.repeatHeader.top - seamLift),
                        width: wrapWOf(slice.section),
                      }}
                    >
                      <div
                        className="doc-page pv-content"
                        lang={cloneLang}
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
                    const rtl = multi && rSec != null && sectionBidi(rSec)
                    const geo = rg as Partial<{ widths: number[]; gaps: number[] }> & typeof rg
                    // per-column width/gap (w:equalWidth="0" lists differ per column);
                    // gaps ride the columns as margins so unequal spaces work too
                    const widthOf = (ci: number) =>
                      multi ? (geo.widths?.[ci] ?? rg.colWidthPx) : undefined
                    const gapAfter = (ci: number) =>
                      multi && ci < region.columns.length - 1 ? (geo.gaps?.[ci] ?? rg.gapPx) : 0
                    return (
                      <div
                        key={ri}
                        className="pv-region"
                        style={{
                          ...(extent !== undefined ? { height: extent } : {}),
                          // RTL section (w:bidi): columns fill right-to-left
                          ...(rtl ? { flexDirection: 'row-reverse' as const } : {}),
                        }}
                      >
                        {region.columns.map((col, ci) => {
                          // the document's very last column window opens to the region's
                          // full capacity: slice bounds can drift a few lines short of the
                          // clone's real height (same allowance as the single-flow branch),
                          // silently cutting the document tail mid-line from export/print
                          const tailWindow =
                            i === slices.length - 1 &&
                            ri === slice.regions!.length - 1 &&
                            ci === region.columns.length - 1 &&
                            vOffset <= 0.5
                          return (
                            <div
                              key={ci}
                              className="pv-col"
                              style={{
                                width: widthOf(ci),
                                ...(rtl
                                  ? { marginLeft: gapAfter(ci) }
                                  : { marginRight: gapAfter(ci) }),
                              }}
                            >
                              {col.repeatHeader && (
                                <div
                                  className="pv-clip"
                                  style={{ height: col.repeatHeader.height }}
                                >
                                  <div
                                    className="pv-offset"
                                    style={{
                                      marginTop: -col.repeatHeader.top,
                                      width: wrapWOf(region.section),
                                    }}
                                  >
                                    <div
                                      className="doc-page pv-content"
                                      lang={cloneLang}
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
                                  height: bodyWindowHeight(col, region.height, {
                                    full: tailWindow,
                                    hasNotes: notes.length > 0,
                                    seamExtend: 0,
                                  }),
                                }}
                              >
                                <div
                                  className="pv-offset"
                                  style={{ marginTop: -col.start, width: wrapWOf(region.section) }}
                                >
                                  <div
                                    className="doc-page pv-content"
                                    lang={cloneLang}
                                    data-pv-body=""
                                    data-pv-win={col.start.toFixed(1)}
                                    dangerouslySetInnerHTML={{
                                      __html: cloneKids
                                        ? prunedCloneHtml(
                                            cloneKids,
                                            col.start,
                                            tailWindow ? col.start + region.height : col.end,
                                          )
                                        : html,
                                    }}
                                  />
                                </div>
                              </div>
                            </div>
                          )
                        })}
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
                        openTop +
                        bodyLift +
                        bodyWindowHeight(slice, contentH, {
                          full: (i === slices.length - 1 && vOffset <= 0.5) || !!sectionVertical(s),
                          hasNotes: notes.length > 0,
                          seamExtend: seam.extend,
                        }),
                      ...(vOffset > 0.5 || openTop > 0 || bodyLift
                        ? { marginTop: (vOffset > 0.5 ? vOffset : 0) - openTop - bodyLift }
                        : {}),
                      ...(openTop > 0 ? { paddingTop: openTop } : {}),
                    }}
                  >
                    {/* the offset lives on a separate wrapper: print rules zero out .doc-page's margin;
                      width is fixed to the section's wrap width so the clone never reflows against the paper */}
                    <div
                      className="pv-offset"
                      style={{
                        marginTop: -(slice.start - bodyLift),
                        width: wrapWOf(slice.section),
                      }}
                    >
                      <div
                        className="doc-page pv-content"
                        lang={cloneLang}
                        data-pv-body=""
                        data-pv-win={slice.start.toFixed(1)}
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
                {notes.length > 0 && (
                  <div
                    className="pv-footnotes"
                    style={{
                      left: mL,
                      width: pageW - mL - mR,
                      height: notesH,
                      ...separatorVars('footnote'),
                      ...(noteArea.top === null
                        ? { bottom: twipsToPx(s.marginBottom) }
                        : { top: noteArea.top }),
                    }}
                  >
                    {notes.map((n) => (
                      // entries get reserved heights (DOM-measured at these exact styles);
                      // min-height lets a residual long entry spill into the bottom margin
                      // instead of overprinting the next entry
                      <div
                        key={n.id}
                        className={['pv-footnote', ...noteRowBox(n).classes].join(' ')}
                        style={{ minHeight: n.height, ...noteRowBox(n).style }}
                      >
                        {!n.noRefMark && <sup>{noteMarkText('footnote', n.no)}</sup>}
                        {n.richParas
                          ? n.richParas.map((para, pi) => (
                              <span key={pi}>
                                {pi > 0 && <br />}
                                {para.map((run, ri) => (
                                  <span
                                    key={ri}
                                    className={noteRunClass(run)}
                                    style={{
                                      fontWeight: run.bold ? 600 : undefined,
                                      fontStyle: run.italic ? 'italic' : undefined,
                                      textDecoration:
                                        [run.underline && 'underline', run.strike && 'line-through']
                                          .filter(Boolean)
                                          .join(' ') || undefined,
                                      color: run.color ? textColorValue(run.color) : undefined,
                                      fontSize: run.sizeHalfPoints
                                        ? `${run.sizeHalfPoints / 2}pt`
                                        : undefined,
                                      fontFamily: run.fontAscii
                                        ? cssFontFamily(run.fontAscii)
                                        : undefined,
                                      WebkitTextStroke: run.textOutline
                                        ? textOutlineCssValue(run.textOutline)
                                        : undefined,
                                      textTransform: run.caps === 'all' ? 'uppercase' : undefined,
                                      fontVariantCaps:
                                        run.caps === 'small' ? 'small-caps' : undefined,
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
                        left: mL,
                        width: pageW - mL - mR,
                        top:
                          mTop +
                          (slice.regions ? 0 : vOffset) +
                          (slice.repeatHeader?.height ?? 0) +
                          (rows[0].top - slice.start) +
                          noteArea.endnoteShift,
                      }}
                    >
                      {rows.map(({ item: n, height, withSeparator }) => (
                        <div
                          key={n.id}
                          className={['pv-footnote', ...noteRowBox(n).classes].join(' ')}
                          style={{ minHeight: height, ...noteRowBox(n).style }}
                        >
                          {withSeparator && (
                            <div
                              className="pv-endnote-separator"
                              style={separatorVars('endnote')}
                            />
                          )}
                          {!n.noRefMark && <sup>{noteMarkText('endnote', n.no)}</sup>}
                          {n.richParas
                            ? n.richParas.map((para, pi) => (
                                <span key={pi}>
                                  {pi > 0 && <br />}
                                  {para.map((run, ri) => (
                                    <span
                                      key={ri}
                                      className={noteRunClass(run)}
                                      style={{
                                        fontWeight: run.bold ? 600 : undefined,
                                        fontStyle: run.italic ? 'italic' : undefined,
                                        textDecoration:
                                          [
                                            run.underline && 'underline',
                                            run.strike && 'line-through',
                                          ]
                                            .filter(Boolean)
                                            .join(' ') || undefined,
                                        color: run.color ? textColorValue(run.color) : undefined,
                                        fontSize: run.sizeHalfPoints
                                          ? `${run.sizeHalfPoints / 2}pt`
                                          : undefined,
                                        fontFamily: run.fontAscii
                                          ? cssFontFamily(run.fontAscii)
                                          : undefined,
                                        WebkitTextStroke: run.textOutline
                                          ? textOutlineCssValue(run.textOutline)
                                          : undefined,
                                        textTransform: run.caps === 'all' ? 'uppercase' : undefined,
                                        fontVariantCaps:
                                          run.caps === 'small' ? 'small-caps' : undefined,
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
                {(parts.footer || parts.footerImages?.some((img) => !img.floating)) && (
                  <HeaderFooterArea
                    kind="footer"
                    value={parts.footer ?? { text: '' }}
                    images={parts.footerImages?.filter(
                      (img) => !img.floating || hfImageHangsOnPara(img),
                    )}
                    readOnly
                    onCommit={() => {}}
                    pageNo={pageNoText}
                    pageTotal={slices.length}
                    sectionPages={sectionPages[i]}
                    style={
                      parts.footerColumnPx !== undefined
                        ? { width: parts.footerColumnPx }
                        : undefined
                    }
                    boxGeom={{ ...hfStripGeom(s, legacyBands), columnW: parts.footerColumnPx }}
                  />
                )}
                <div className="pv-pageno">{i + 1}</div>
              </div>
              {markupOn &&
                (() => {
                  const contentRight = pageW - mR
                  const bubbleLeft = contentRight + BUBBLE_ENTRY
                  const bubbleW = pageW + MARKUP_EXTRA_W - BUBBLE_RIGHT_PAD - bubbleLeft
                  const headerH = slice.repeatHeader?.height ?? 0
                  // the last page opens its clip to full capacity (slice bounds can
                  // land short); its tail anchors must still balloon
                  const sliceEnd =
                    i === slices.length - 1 && vOffset <= 0.5
                      ? Math.max(slice.end, slice.start + contentH - headerH)
                      : slice.end
                  const inPageTopOf = (sp: { top: number }): number | null => {
                    if (!slice.regions) {
                      if (sp.top < slice.start - 0.5 || sp.top >= sliceEnd - 0.5) return null
                      return mTop + vOffset + headerH + (sp.top - slice.start)
                    }
                    const base = slice.regions[0]?.top ?? slice.start
                    for (const region of slice.regions) {
                      for (const col of region.columns) {
                        // the document's last column window opens to the region's full
                        // capacity (see the pv-clip tail allowance): match its anchors too
                        const colEnd =
                          i === slices.length - 1 &&
                          region === slice.regions.at(-1) &&
                          col === region.columns.at(-1) &&
                          vOffset <= 0.5
                            ? Math.max(
                                col.end,
                                col.start + region.height - (col.repeatHeader?.height ?? 0),
                              )
                            : col.end
                        if (sp.top >= col.start - 0.5 && sp.top < colEnd - 0.5)
                          return (
                            mTop +
                            (region.top - base) +
                            (col.repeatHeader?.height ?? 0) +
                            (sp.top - col.start)
                          )
                      }
                    }
                    return null
                  }
                  type Item = Omit<BalloonItem, 'seq'> & {
                    kind: 'comment' | 'del' | 'fmt'
                    label: string
                    text: string
                    /** revision balloons: author line with the relative time (Word's header) */
                    head?: { author: string; when: string }
                    endY: number
                    /** merge key: comments stay in document order even where anchors invert (next column) */
                    orderTop: number
                  }
                  const items: Item[] = []
                  let orderFloor = -Infinity
                  for (const sp of commentSpots) {
                    const anchorTop = inPageTopOf(sp)
                    if (anchorTop === null) continue
                    const root = commentById.get(sp.id)
                    const replies = (comments ?? []).filter((r) => r.parentId === sp.id)
                    const text = [root?.text ?? '', ...replies.map((r) => `${r.author}: ${r.text}`)]
                      .filter(Boolean)
                      .join('\n')
                    // Word labels by initials + running number: "Commented [KICC1]"
                    const label = `${t('appPvCommented')} [${root?.initials ?? ''}${sp.no}]: `
                    orderFloor = Math.max(orderFloor, anchorTop)
                    items.push({
                      key: `c${sp.id}`,
                      kind: 'comment',
                      label,
                      text,
                      anchorTop,
                      orderTop: orderFloor,
                      endX: sp.endX,
                      endY: sp.endY,
                      // the label prefix wraps with the text, so it counts toward the height
                      height: estimateBubbleHeight(label + text, bubbleW - 16),
                      sticky: true,
                    })
                  }
                  const revItems: Item[] = []
                  for (const sp of revSpots) {
                    const anchorTop = inPageTopOf(sp)
                    if (anchorTop === null) continue
                    const label = `${t(sp.kind === 'del' ? 'editorRevDeleted' : 'editorRevFormatted')}: `
                    const text =
                      sp.text.length > REV_TEXT_MAX ? `${sp.text.slice(0, REV_TEXT_MAX)}…` : sp.text
                    revItems.push({
                      key: sp.key,
                      kind: sp.kind,
                      label,
                      text,
                      head: {
                        author: sp.author || t('editorUnknownAuthor'),
                        when: revWhen(sp.date),
                      },
                      anchorTop,
                      orderTop: anchorTop,
                      endX: sp.endX,
                      endY: sp.endY,
                      height:
                        REV_LINE_H + estimateBubbleHeight(label + text, bubbleW - 12, REV_METRICS),
                      compactHeight: 2 * REV_LINE_H + REV_PAD_H,
                      sticky: false,
                    })
                  }
                  const placed = stackBalloons(
                    mergeBalloonLists(items, revItems),
                    mTop,
                    mTop + contentH,
                    (it) => (it.kind === 'comment' ? BUBBLE_STACK_GAP : REV_STACK_GAP),
                  )
                  return (
                    <div
                      className="pv-markup"
                      style={{
                        width: pageW + MARKUP_EXTRA_W,
                        height: pageH,
                        ...markupTransform,
                      }}
                    >
                      <div
                        className="pv-markup-band"
                        style={{
                          left: contentRight + MARKUP_BAND_GUTTER,
                          width: pageW + MARKUP_EXTRA_W - contentRight - MARKUP_BAND_GUTTER * 2,
                        }}
                      />
                      {placed.length > 0 && (
                        <svg
                          className="pv-comment-leaders"
                          width={pageW + MARKUP_EXTRA_W}
                          height={pageH}
                        >
                          {placed.map((p) => {
                            // leader runs from the range end; a range that ends on
                            // another page falls back to the anchor's first line
                            const endsHere =
                              !slice.regions &&
                              p.endY >= slice.start - 0.5 &&
                              p.endY < sliceEnd + 0.5
                            const ex = endsHere ? mL + p.endX : contentRight
                            const ey = endsHere
                              ? mTop + vOffset + headerH + (p.endY - slice.start)
                              : p.anchorTop + BUBBLE_LINE_H - 4
                            return (
                              <path
                                key={p.key}
                                className={p.kind === 'comment' ? undefined : 'pv-leader-rev'}
                                d={`M ${ex} ${ey} L ${contentRight + BUBBLE_ENTRY / 2} ${ey} L ${bubbleLeft} ${p.top + 10}`}
                              />
                            )
                          })}
                        </svg>
                      )}
                      {placed.map((p) => (
                        <div
                          key={p.key}
                          className={
                            p.kind === 'comment'
                              ? 'pv-comment-bubble'
                              : `pv-comment-bubble pv-rev-bubble${p.compact ? ' pv-rev-bubble-compact' : ''}`
                          }
                          style={{ left: bubbleLeft, top: p.top, width: bubbleW }}
                        >
                          {p.head && (
                            <div className="pv-rev-head">
                              <span className="pv-comment-label">{p.head.author}</span>
                              <span className="pv-rev-when">{p.head.when}</span>
                            </div>
                          )}
                          <div className="pv-rev-body">
                            <span className="pv-comment-label">{p.label}</span>
                            {p.text}
                          </div>
                        </div>
                      ))}
                    </div>
                  )
                })()}
            </div>
          )
        })}
      </div>
    </div>
  )
}
