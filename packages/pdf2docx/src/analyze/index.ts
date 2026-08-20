/**
 * Analysis layer — pure geometry, zero PDFium dependency. Turns an extracted
 * page (chars + images + shapes + flags) into the IR page the rebuild layer
 * consumes: shape normalization → lattice tables → stream tables → floats →
 * XY-Cut sections/columns → paragraph flow → before_space chain.
 *
 * Plain pages (one section, one column, no stream tables, no floats) go
 * through the exact P1/P2 assembly path so their output stays byte-stable.
 */
import type { ExtractedPage } from '../extract'
import type { Rect } from '../geometry'
import { intersectArea, median, overlapRatio, rectArea, rectUnionAll } from '../geometry'
import type {
  ImageBlock,
  IrPage,
  PageBlock,
  PageColumn,
  PageSection,
  Stroke,
  TextBlock,
} from '../ir'
import { groupIntoBlocks, bodyContextOf } from './blocks'
import {
  analyzeChars,
  calibrateFontSizes,
  dedupeDoubleDrawnChars,
  normalizeRegionalFontArtifacts,
} from './chars'
import { pageConfidence } from './confidence'
import { applyDecorBorders } from './decor'
import type { LayoutSection, SectionElement } from './columns'
import { detectSections, mergeTwinSections } from './columns'
import { classifyFloatImages, suppressTextShadowImages } from './floats'
import { detectFootnotes } from './footnotes'
import { extractEmptyFrames } from './frames'
import { detectFormTables } from './form'
import { clusterCombiningMarks, groupIntoLines } from './lines'
import { detectListBlocks } from './lists'
import { mergeSideBySidePanels } from './panels'
import { encodeRgbaPng } from '../extract/png'
import { normalizeArabicForms } from './rtl'
import {
  extractBackgroundPanels,
  extractFullBleedTiles,
  extractPageBackground,
  extractTextBackdrops,
  hexLuminance,
  isNearWhite,
  normalizeShapes,
} from './shapes'
import { applySpacingChain } from './spacing'
import { detectStreamTables } from './stream'
import { applyTextShapeStyles } from './styling'
import { detectTables, solveLatticeGrids } from './table'
import {
  detectTocBlocks,
  detectTocRows,
  hasDotLeaderRun,
  LEADER_PAGE_MIN_LINES,
  LEADER_PAGE_MIN_SHARE,
} from './toc'
import { clusterUnitRows, splitIntoUnits } from './units'
import { detectRuleSeparatedZones } from './zones'

export { analyzeChars, dedupeDoubleDrawnChars, normalizeRegionalFontArtifacts } from './chars'
export { clusterCombiningMarks, groupIntoLines, type RawLine } from './lines'
export {
  groupIntoWords,
  isLetterSpacedLine,
  spaceGapThreshold,
  medianCharGap,
  type Word,
  type WordOptions,
} from './words'
export { buildSpans, isBoldChar } from './spans'
export { groupIntoBlocks, bodyContextOf, type BodyContext } from './blocks'
export {
  extractBackgroundPanels,
  extractFullBleedTiles,
  extractPageBackground,
  extractTextBackdrops,
  hexLuminance,
  isNearWhite,
  normalizeShapes,
  rectOfSubpath,
} from './shapes'
export {
  detectTables,
  detectCellVAlign,
  detectCellHAlign,
  groupStrokes,
  solveGrid,
  layoutCells,
  trimGhostEdgeColumns,
  type TableGrid,
  type CellLayout,
  type DetectedTables,
} from './table'
export { firstStrongDir, lineHasRtl, normalizeArabicForms, reorderVisualToLogical } from './rtl'
export {
  splitIntoUnits,
  clusterUnitRows,
  columnGapThreshold,
  type LineUnit,
  type UnitRow,
} from './units'
export { detectStreamTables, type DetectedStreamTables } from './stream'
export {
  detectSections,
  type SectionElement,
  type LayoutColumn,
  type LayoutSection,
} from './columns'
export { applySpacingChain } from './spacing'
export { classifyFloatImages, suppressTextShadowImages, type ClassifiedImages } from './floats'
export { applyTextShapeStyles, type StyledChars } from './styling'
export { detectFormTables, detectCheckboxSquares, type DetectedFormTables } from './form'
export { detectFootnotes, type DetectedFootnotes } from './footnotes'
export { extractEmptyFrames, type EmptyFrame } from './frames'
export { detectFurniture, type FurniturePage, type FurnitureResult } from './furniture'
export { applyDecorBorders, type DecorResult } from './decor'
export { detectListBlocks, parseListMarker } from './lists'
export { detectTocBlocks, detectTocRows } from './toc'
export { detectVectorRegions } from './vector'
export { pageConfidence, PAGE_CONFIDENCE_MIN, type ConfidenceSignals } from './confidence'
export {
  classifyPage,
  classifyPages,
  computeCanvasPrior,
  isNewsletterStrongPage,
  isSlideProducer,
  isSlideSizedPage,
  type CanvasDocPrior,
  type DocMeta,
  type PageLayoutClass,
} from './canvas'

/**
 * a region panel is one flat color — a stretched 2×2 bitmap carries it. A
 * translucent source fill keeps its alpha (P11 A): a 70%-black scrim over a
 * photo backdrop must not become an opaque slab hiding the photo.
 */
function solidPanelImage(box: Rect, color: string, alpha = 255, z?: number): ImageBlock {
  const r = parseInt(color.slice(0, 2), 16)
  const g = parseInt(color.slice(2, 4), 16)
  const b = parseInt(color.slice(4, 6), 16)
  const rgba = new Uint8Array(2 * 2 * 4)
  for (let i = 0; i < rgba.length; i += 4) {
    rgba[i] = r
    rgba[i + 1] = g
    rgba[i + 2] = b
    rgba[i + 3] = alpha
  }
  return {
    kind: 'image',
    box,
    data: encodeRgbaPng(rgba, 2, 2),
    mime: 'image/png',
    pixelWidth: 2,
    pixelHeight: 2,
    float: { wrap: 'behind', xOffsetPt: box.x0 },
    ...(z !== undefined ? { z } : {}),
  }
}

/** hollow border bitmap for an empty stroke frame (P16 K) — alpha inside */
function frameImage(box: Rect, color: string, widthPt: number): ImageBlock {
  const scale = 2
  const w = Math.max(2, Math.round((box.x1 - box.x0) * scale))
  const h = Math.max(2, Math.round((box.y1 - box.y0) * scale))
  const bw = Math.max(1, Math.round(widthPt * scale))
  const r = parseInt(color.slice(0, 2), 16)
  const g = parseInt(color.slice(2, 4), 16)
  const b = parseInt(color.slice(4, 6), 16)
  const rgba = new Uint8Array(w * h * 4)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const edge = x < bw || x >= w - bw || y < bw || y >= h - bw
      if (!edge) continue
      const o = (y * w + x) * 4
      rgba[o] = r
      rgba[o + 1] = g
      rgba[o + 2] = b
      rgba[o + 3] = 255
    }
  }
  return {
    kind: 'image',
    box,
    data: encodeRgbaPng(rgba, w, h),
    mime: 'image/png',
    pixelWidth: w,
    pixelHeight: h,
    float: { wrap: 'behind', xOffsetPt: box.x0 },
  }
}

/** a gutter stroke must cover this share of its section's height to be the column rule */
const COL_SEP_MIN_SECTION_COVER = 0.6

/** interleave non-text blocks into the text flow by vertical position (top edge) */
function mergeBlocks(textBlocks: PageBlock[], floating: PageBlock[]): PageBlock[] {
  const merged: PageBlock[] = [...textBlocks]
  for (const block of floating) {
    // insert before the first block whose top edge is below this one's
    const at = merged.findIndex((b) => b.box.y1 < block.box.y1)
    if (at === -1) merged.push(block)
    else merged.splice(at, 0, block)
  }
  return merged
}

/** wrap reading-order blocks as the single-column page structure */
function singleSectionOf(blocks: PageBlock[]): PageSection[] {
  if (blocks.length === 0) return []
  const firstText = blocks.find((b): b is TextBlock => b.kind === 'text')
  const box = rectUnionAll(blocks.map((b) => b.box))
  return [
    {
      box,
      columns: [{ box, blocks }],
      gutterWidthsPt: [],
      dir: firstText?.dir ?? 'ltr',
    },
  ]
}

/** assemble one layout column's elements into reading-order blocks */
function assembleColumn(
  column: LayoutSection['columns'][number],
  singleColumn: boolean,
  pageWidthPt: number,
  listSeq: { next: number },
  landscape: boolean,
  pageBodyLeftX0?: number,
): PageColumn {
  // row-major unit order (P14 A): same-row units whose baselines differ by a
  // hair (a 21pt number badge beside a 14pt card title) must flatten left →
  // right, or the line grouper reads the x regression as a line break and
  // stacks the badge under its own title
  const units = clusterUnitRows(column.elements.filter((e) => e.unit).map((e) => e.unit!)).flatMap(
    (row) => row.units,
  )
  const lines = analyzeChars(units.flatMap((u) => u.chars))
  // single-column sections keep the page-level mirrored right edge (see
  // analyzePage below); real columns are judged against their own extent
  const body = {
    bodyLeft: column.box.x0,
    bodyRight: singleColumn ? Math.max(column.box.x1, pageWidthPt - column.box.x0) : column.box.x1,
  }
  const textBlocks = detectTocBlocks(
    detectListBlocks(
      groupIntoBlocks(lines, body, { pinOpenLeadedBreaks: landscape }),
      listSeq,
      pageBodyLeftX0,
    ),
  )
  const others = column.elements.filter((e) => e.block).map((e) => e.block!)
  return { box: column.box, blocks: mergeBlocks([...textBlocks], others) }
}

/** insert a float into the section/column its anchor position falls in */
function placeFloat(sections: PageSection[], float: ImageBlock): void {
  const cy = (float.box.y0 + float.box.y1) / 2
  const target =
    sections.find((s) => cy <= s.box.y1 && cy >= s.box.y0) ?? sections[sections.length - 1]
  if (!target) return
  const cx = (float.box.x0 + float.box.x1) / 2
  const overlapW = (c: PageColumn): number =>
    Math.min(c.box.x1, float.box.x1) - Math.max(c.box.x0, float.box.x0)
  let column = target.columns.find((c) => cx >= c.box.x0 && cx <= c.box.x1)
  if (!column) {
    column = [...target.columns].sort((a, b) => overlapW(b) - overlapW(a))[0]!
  }
  const at = column.blocks.findIndex((b) => b.box.y1 < float.box.y1)
  if (at === -1) column.blocks.push(float)
  else column.blocks.splice(at, 0, float)
}

const flattenSections = (sections: PageSection[]): PageBlock[] =>
  sections.flatMap((s) => s.columns.flatMap((c) => c.blocks))

// ── card regions (P20) ──
/** a text block belongs to a plate when ~all of its box sits inside */
const CARD_MEMBER_COVER_MIN = 0.9
/** a text block only partially on the plate poisons the card (a paragraph cannot split) */
const CARD_PARTIAL_COVER_LO = 0.1
/** any table/image overlapping the plate this much keeps it a plain panel
 * (the parked shading attempt collected form-gov table banners — never again) */
const CARD_FOREIGN_OVERLAP_MAX = 0.3
/** the P10 B light-text judgment, verbatim: only a DARK plate under LIGHT
 * text converts. Light plates with dark text (form fill-in areas, highlight
 * washes) render fine as panels — misalignment there is nearly invisible,
 * and form-gov's pale form regions must stay byte-identical. */
const CARD_MAX_FILL_LUMA = 0.75
const CARD_LIGHT_TEXT_LUMA = 0.85
const CARD_LIGHT_SHARE_MIN = 0.6
/** a card is a multi-line GROUP; single-line label banners (form-gov's green
 * section strips, P16 I label plates) hug their text and render fine as
 * panels — converting them changed a settled sentinel for no gain */
const CARD_MIN_MEMBER_LINES = 2

/** one card candidate: the plate fill and its already-minted bgPanels image */
export interface CardCandidate {
  img: ImageBlock
  box: Rect
  color: string
  rounded: boolean
  z?: number
}

/** length-weighted share of member text that is light (P10 B bar) */
function memberLightShare(members: readonly TextBlock[]): number | undefined {
  let total = 0
  let light = 0
  for (const b of members)
    for (const line of b.lines)
      for (const span of line.spans) {
        const len = span.text.trim().length
        if (len === 0) continue
        total += len
        if (hexLuminance(span.color) >= CARD_LIGHT_TEXT_LUMA) light += len
      }
  return total > 0 ? light / total : undefined
}

/**
 * Card regions (P20): a backdrop plate whose content is exactly a contiguous
 * run of text blocks in one column becomes IrPage.cards[*] — the flow rebuild
 * turns the group into one paragraph-anchored text box, so plate and text
 * ride the flow together (pinned at absolute coordinates, one reflowed line
 * above slides the plate off its own text — the byte-deck sample's FINAL
 * TAKEAWAY overlap).
 * Anything else on the plate (tables, images, straddling or decorated text,
 * low contrast) keeps the behindDoc panel exactly as today.
 * Candidates are visited small-to-large so a chip nested on a bigger card
 * claims its text first and the outer plate then declines (stays a panel).
 */
export function detectCardRegions(page: IrPage, candidates: readonly CardCandidate[]): void {
  // page flow order (flattened reading order); floats never join a card
  const flow = page.blocks.filter((b) => !(b.kind === 'image' && b.float !== undefined))
  for (const cand of [...candidates].sort((a, b) => rectArea(a.box) - rectArea(b.box))) {
    const members: TextBlock[] = []
    let poisoned = false
    for (const b of flow) {
      const ratio = overlapRatio(b.box, cand.box)
      if (b.kind === 'text') {
        if (b.cardId === undefined && ratio >= CARD_MEMBER_COVER_MIN) {
          // decorated / TOC / footnote-bearing members need flow machinery
          // the text box does not reproduce — keep the panel
          if (
            b.border !== undefined ||
            b.tocEntry !== undefined ||
            b.lines.some((l) => l.spans.some((sp) => sp.noteRef !== undefined))
          ) {
            poisoned = true
          }
          members.push(b)
        } else if (ratio > CARD_PARTIAL_COVER_LO) {
          poisoned = true
        }
      } else if (ratio >= CARD_FOREIGN_OVERLAP_MAX) {
        poisoned = true
      }
    }
    if (poisoned || members.length === 0) continue
    if (members.reduce((n, m) => n + m.lines.length, 0) < CARD_MIN_MEMBER_LINES) continue
    // contiguous run in reading order: a stranger between two members would
    // have to flow through the box (members may span a section boundary —
    // detectSections cuts bands freely across a tall plate)
    const idx = members.map((m) => flow.indexOf(m)).sort((a, b) => a - b)
    if (idx[0]! < 0 || idx[idx.length - 1]! - idx[0]! !== idx.length - 1) continue
    if (hexLuminance(cand.color) > CARD_MAX_FILL_LUMA) continue
    const lightShare = memberLightShare(members)
    if (lightShare === undefined || lightShare < CARD_LIGHT_SHARE_MIN) continue
    const cards = (page.cards ??= [])
    const cardId = cards.length
    cards.push({
      box: cand.box,
      color: cand.color,
      ...(cand.rounded ? { rounded: true } : {}),
      ...(cand.z !== undefined ? { z: cand.z } : {}),
    })
    cand.img.cardId = cardId
    for (const m of members) m.cardId = cardId
  }
}

export function analyzePage(extracted: ExtractedPage): IrPage {
  const base: IrPage = {
    index: extracted.index,
    widthPt: extracted.widthPt,
    heightPt: extracted.heightPt,
    rotation: extracted.rotation,
    blocks: [],
    degraded: extracted.degraded,
    scanned: extracted.scanned,
    hasStructTree: extracted.hasStructTree,
    degradedReason: extracted.degradedReason,
    render: extracted.render,
  }
  // fallback pages carry only their full-page render
  if (extracted.scanned || extracted.degraded) return base

  // P11 B: ghost twins of double-drawn chars go first — they'd double-weight
  // the calibration medians and rebuild as visible double strikes
  dedupeDoubleDrawnChars(extracted.chars)

  // P8: bogus declared font sizes → metric-derived effective sizes, before
  // anything downstream (spans, styling, tables, line heights) reads them
  calibrateFontSizes(extracted.chars)

  // P16 D: TC-variant families on strictly-simplified text are export-chain
  // substitution artifacts — rewrite before spans capture the family
  normalizeRegionalFontArtifacts(extracted.chars)

  const shapes = normalizeShapes(extracted.paths)
  base.shapes = shapes

  // a fill covering ~the whole page is the page wash, not content — record it
  // for the document background and keep it out of the shape pool so it never
  // reads as cell shading / highlight / vector art
  const bg = extractPageBackground(shapes.fills, extracted.widthPt, extracted.heightPt)
  if (bg !== undefined && !isNearWhite(bg)) base.bgColor = bg

  // full-bleed tile group (P22 B): a text-free page painted edge-to-edge by a
  // few big slabs pins them as behindDoc floats at their absolute coordinates
  // — left in the pool their hairline edges mint a ghost lattice table that
  // gets margin-clamped and shrunk
  const tiles = extractFullBleedTiles(
    shapes,
    extracted.chars.length,
    extracted.widthPt,
    extracted.heightPt,
  )
  if (tiles.length > 0) {
    base.bgPanels = [
      ...(base.bgPanels ?? []),
      ...tiles.map((t) => solidPanelImage(t.box, t.color, t.alpha, t.z)),
    ]
  }

  // region background panels (P10 B): bottom-z edge-flush color slabs under
  // body text (cover spines, chapter banners) leave the fill pool and pin to
  // the page as behindDoc floats — their text keeps flowing normally
  const panels = extractBackgroundPanels(
    shapes.fills,
    extracted.chars.map((c) => c.box),
    extracted.widthPt,
    extracted.heightPt,
  )
  if (panels.length > 0) {
    base.bgPanels = panels.map((p) => solidPanelImage(p.box, p.color, p.alpha, p.z))
  }

  // P9 B: pre-rendered background stack rides through to the rebuild layer,
  // which pins it behind the page's text as a full-page behindDoc float
  if (extracted.bgRender) base.bgRender = extracted.bgRender

  // P19: decor shadings ride through untouched — only the canvas rebuild
  // path pins them (flow output stays byte-identical)
  if (extracted.decorImages?.length) {
    base.decorImages = extracted.decorImages.map((img) => ({
      kind: 'image' as const,
      box: img.box,
      data: img.data,
      mime: img.mime,
      pixelWidth: img.pixelWidth,
      pixelHeight: img.pixelHeight,
      float: { wrap: 'behind' as const, xOffsetPt: img.box.x0 },
      ...(img.z !== undefined ? { z: img.z } : {}),
    }))
  }

  // P6: footnotes lift off the page bottom before anything else sees the
  // chars; their body anchors collapse into noteRef markers
  const noted = detectFootnotes(
    extracted.chars,
    shapes,
    extracted.index,
    extracted.widthPt,
    extracted.heightPt,
  )
  if (noted.footnotes.length > 0) base.footnotes = noted.footnotes

  // P4 style mapping runs before table detection so table-cell text is styled
  // too; lattice grid regions are pre-solved so their borders / cell shading
  // never read as underline / highlight
  // form-XObject strokes get only the strict ≥3×3 lattice chance (P14 C/P27):
  // slide design dividers arrive wrapped in forms, and a lone cross of them
  // must not mint a bordered 2×2 "table" around the content it separates —
  // but whole tables wrapped in forms (twotables) do solve
  const gridBoxes: Rect[] = solveLatticeGrids(shapes.strokes).map((g) => g.box)

  // light-text backdrops (P10 B): card fills whose text is near-white leave
  // the fill pool as behindDoc floats — dropped, that text is invisible on
  // the paper. After gridBoxes so lattice cell shading stays w:shd.
  const backdrops = extractTextBackdrops(
    shapes.fills,
    noted.bodyChars,
    gridBoxes,
    extracted.widthPt,
    extracted.heightPt,
    shapes.curvedFills ?? [],
    // white capsules need something beneath them to contrast with (P12 C)
    base.bgRender !== undefined || base.bgColor !== undefined || (base.bgPanels?.length ?? 0) > 0,
  )
  const cardCandidates: CardCandidate[] = []
  if (backdrops.length > 0) {
    const curved = new Set(shapes.curvedFills ?? [])
    const imgs = backdrops.map((p) => solidPanelImage(p.box, p.color, p.alpha, p.z))
    // opaque plates are card candidates (P20); translucent scrims never are
    imgs.forEach((img, i) => {
      const p = backdrops[i]!
      if (p.alpha === undefined) {
        cardCandidates.push({
          img,
          box: p.box,
          color: p.color,
          rounded: curved.has(p),
          ...(p.z !== undefined ? { z: p.z } : {}),
        })
      }
    })
    base.bgPanels = [...(base.bgPanels ?? []), ...imgs]
  }

  // empty stroked rectangles (quote/answer boxes, P16 K): whole frames pin
  // behind the text; left in the pool their horizontal edges degrade into
  // two bare paragraph rules and the vertical edges vanish
  const frames = extractEmptyFrames(shapes.strokes, noted.bodyChars, gridBoxes)
  if (frames.length > 0) {
    base.bgPanels = [
      ...(base.bgPanels ?? []),
      ...frames.map((f) => frameImage(f.box, f.color, f.widthPt)),
    ]
  }

  const styled = applyTextShapeStyles(
    noted.bodyChars,
    shapes,
    gridBoxes,
    extracted.widthPt * extracted.heightPt,
  )
  const { tables: latticeTables, remainingChars } = detectTables(
    shapes,
    styled.chars,
    extracted.heightPt,
    extracted.widthPt,
  )

  // P3 layout passes work on line units (visual lines split at column gaps)
  const prepped = clusterCombiningMarks(normalizeArabicForms(remainingChars))
  const rawLines = groupIntoLines(prepped)
  const units = splitIntoUnits(rawLines)
  // checkbox form rows (P6) — checked before the stream pass: their single
  // checkbox-led rows are exactly what the stream detector rejects by design
  const latticeBoxes = latticeTables.map((t) => t.box)
  const { tables: formTables, remainingUnits: unitsAfterForms } = detectFormTables(
    units,
    shapes,
    latticeBoxes,
  )
  // leaderless TOC rows (P6) — before the stream pass, whose aligned-columns
  // heuristics would otherwise read a contents page as a table
  const { blocks: tocRowBlocks, remainingUnits: unitsAfterToc } = detectTocRows(unitsAfterForms)
  // rule-separated side-by-side zones (P22 A) — before the stream pass so a
  // court caption's two stacks never dissolve into interleaved rows
  const {
    tables: zoneTables,
    remainingUnits: unitsAfterZones,
    consumedStrokes: zoneStrokes,
  } = detectRuleSeparatedZones(
    unitsAfterToc,
    shapes,
    [...latticeBoxes, ...formTables.map((t) => t.box)],
    { pageWidthPt: extracted.widthPt, pageHeightPt: extracted.heightPt },
  )
  const { tables: streamTables, remainingUnits } = detectStreamTables(
    unitsAfterZones,
    shapes,
    [...latticeBoxes, ...formTables.map((t) => t.box), ...zoneTables.map((t) => t.box)],
    // slides place a data table BESIDE prose absolutely — their interleaved
    // rows need the vast-valley region pass (P16 E)
    { slideRegions: extracted.widthPt > extracted.heightPt },
  )
  const imageBlocks: ImageBlock[] = extracted.images.map((img) => ({
    kind: 'image' as const,
    box: img.box,
    data: img.data,
    mime: img.mime,
    pixelWidth: img.pixelWidth,
    pixelHeight: img.pixelHeight,
    ...(img.z !== undefined ? { z: img.z } : {}),
  }))
  const { floats, inline } = classifyFloatImages(
    suppressTextShadowImages(imageBlocks, remainingUnits),
    remainingUnits,
    extracted.widthPt * extracted.heightPt,
    { x0: 0, y0: 0, x1: extracted.widthPt, y1: extracted.heightPt },
  )
  // slides have no text flow (P11 D): landscape pages pin every image at its
  // absolute position behind the text — an inline doodle or corner logo would
  // stack its height into a flow the source never had and spill the page
  if (extracted.widthPt > extracted.heightPt) {
    for (const img of inline) {
      img.float = { wrap: 'behind', xOffsetPt: img.box.x0 }
      floats.push(img)
    }
    inline.length = 0
  }
  // side-by-side lattice panels → one grid (P28); after the other table
  // detectors so their exclusion zones keep the original panel boxes
  const { tables: panelTables, notes: panelNotes } = mergeSideBySidePanels(latticeTables)
  const tables = [...panelTables, ...formTables, ...zoneTables, ...streamTables]

  // cell-text overlay images (P29): some generators paint a cell's text a
  // second time as a small raster pinned over the cell — the flow table then
  // shows the text AND a page-anchored ghost of it slightly offset. An image
  // riding a text-bearing cell's line area duplicates it; drop the raster.
  // Empty-cell overlays (label art) are unique content and stay.
  const CELL_OVERLAY_MIN_SHARE = 0.35
  const isCellTextOverlay = (img: ImageBlock): boolean => {
    const cx = (img.box.x0 + img.box.x1) / 2
    const cy = (img.box.y0 + img.box.y1) / 2
    const imgArea = Math.max(1, rectArea(img.box))
    for (const t of tables) {
      if (t.confidence !== undefined) continue
      if (cx < t.box.x0 || cx > t.box.x1 || cy < t.box.y0 || cy > t.box.y1) continue
      for (const row of t.rows) {
        for (const cell of row) {
          if (cell.vMerge === 'continue') continue
          const b = cell.box
          if (cx < b.x0 || cx > b.x1 || cy < b.y0 || cy > b.y1) continue
          let overlap = 0
          for (const block of cell.blocks) {
            for (const line of block.lines) {
              overlap += intersectArea(img.box, line.box)
            }
          }
          return overlap >= CELL_OVERLAY_MIN_SHARE * imgArea
        }
      }
    }
    return false
  }
  for (const list of [floats, inline]) {
    const keep = list.filter((img) => !isCellTextOverlay(img))
    if (keep.length !== list.length) list.splice(0, list.length, ...keep)
  }

  const elements: SectionElement[] = [
    ...remainingUnits.map((u) => ({ box: u.box, unit: u })),
    ...tables.map((t) => ({ box: t.box, block: t })),
    ...tocRowBlocks.map((b) => ({ box: b.box, block: b })),
    ...inline.map((img) => ({ box: img.box, block: img })),
  ]
  // dot-leader index pages (P22 D): the leader runs keep every entry row one
  // full-width unit, so no gutter survives any slab and the column sweep
  // shatters the page into dozens of alternating sections — LibreOffice then
  // renders every transition with its own balancing and the page multiplies.
  // Whole-page single column reproduces the rows (dots are real glyphs).
  const leaderLineCount = rawLines.filter((l) => hasDotLeaderRun(l.chars)).length
  const leaderIndexPage =
    extracted.widthPt < extracted.heightPt &&
    leaderLineCount >= LEADER_PAGE_MIN_LINES &&
    leaderLineCount >= LEADER_PAGE_MIN_SHARE * Math.max(1, rawLines.length)
  let layout: LayoutSection[]
  if (leaderIndexPage && elements.length > 0) {
    const pageBox = rectUnionAll(elements.map((e) => e.box))
    layout = [
      {
        box: pageBox,
        columns: [{ box: pageBox, elements: [...elements] }],
        gutters: [],
        dir: 'ltr',
      },
    ]
  } else {
    layout = detectSections(elements, extracted.heightPt, extracted.widthPt)
    // torn multi-column runs re-join on portrait flow pages (P22 D);
    // landscape/slide sections place content absolutely and stay apart
    if (extracted.widthPt < extracted.heightPt) layout = mergeTwinSections(layout)
  }

  // form tables consume UNITS, so the plain path's remainingChars would
  // duplicate their text — pages with them assemble from units instead
  const plainPage =
    layout.every((s) => s.columns.length === 1) &&
    streamTables.length === 0 &&
    zoneTables.length === 0 &&
    formTables.length === 0 &&
    tocRowBlocks.length === 0 &&
    floats.length === 0

  // page-unique sequence ids for ordered-list runs (rebuild maps them to numIds)
  const listSeq = { next: 0 }
  /** vertical strokes consumed as w:cols separators (P14 C) */
  const sepStrokes = new Set<Stroke>()
  let sections: PageSection[]
  if (plainPage) {
    // P1/P2 assembly with visual ordering (P5): content-stream order breaks on
    // absolutely-positioned overlay layouts (decor text drawn last merged
    // whole-page paragraph blocks and produced huge negative gaps), so lines
    // sort top→down like the multi-column path does
    const lines = analyzeChars(remainingChars).sort(
      (a, b) => b.baseline - a.baseline || a.box.x0 - b.box.x0,
    )
    let body: ReturnType<typeof bodyContextOf> | undefined
    if (lines.length > 0) {
      const content = bodyContextOf(lines)
      // right edge widened to mirror the left margin — otherwise the page's
      // single widest line (often a centered title) defines bodyRight and
      // reads as "right-aligned" against itself
      body = {
        bodyLeft: content.bodyLeft,
        bodyRight: Math.max(content.bodyRight, extracted.widthPt - content.bodyLeft),
      }
    }
    const textBlocks: TextBlock[] = detectTocBlocks(
      detectListBlocks(
        groupIntoBlocks(lines, body, {
          pinOpenLeadedBreaks: extracted.widthPt > extracted.heightPt,
        }),
        listSeq,
      ),
    )
    sections = singleSectionOf(mergeBlocks([...textBlocks], [...inline, ...tables]))
  } else {
    // page-level left edge for the weak-bullet indent evidence (P20): slide
    // layouts pin a dash sub-bullet group into a section of its own, so the
    // column has no plain neighbours to judge the indent against
    const pageBodyLeftX0 = median(layout.flatMap((ls) => ls.columns.map((c) => c.box.x0)))
    sections = layout.map((ls) => {
      const orderedColumns = ls.dir === 'rtl' ? [...ls.columns].reverse() : ls.columns
      return {
        box: ls.box,
        columns: orderedColumns.map((c) =>
          assembleColumn(
            c,
            ls.columns.length === 1,
            extracted.widthPt,
            listSeq,
            extracted.widthPt > extracted.heightPt,
            pageBodyLeftX0,
          ),
        ),
        gutterWidthsPt: ls.gutters.map((g) => g.hi - g.lo),
        dir: ls.dir,
      }
    })
    if (sections.length === 0 && floats.length > 0) {
      // image-only page (wallpaper cover rasterized to one float, P11 D):
      // with no text there is no layout section to host the float — a bare
      // single-column section keeps the page and its behindDoc anchor alive
      sections = singleSectionOf([...floats])
    } else {
      for (const float of floats) placeFloat(sections, float)
    }
  }
  if (sections.length === 0 && (base.bgPanels?.length ?? 0) > 0) {
    // panel-only page (full-bleed tile group, P22 B): no flow content at all,
    // but the behindDoc panels still need a section whose holder paragraphs
    // anchor them — otherwise the page ships blank
    const pageBox: Rect = { x0: 0, y0: 0, x1: extracted.widthPt, y1: extracted.heightPt }
    sections = [
      { box: pageBox, columns: [{ box: pageBox, blocks: [] }], gutterWidthsPt: [], dir: 'ltr' },
    ]
  }
  {
    // drawn column separators (P14 C): a vertical stroke living inside a
    // section's gutter for most of the section height is the divider between
    // the columns — carry it as w:cols w:sep and keep it away from the decor
    // pass (which would count it as an ignored vertical stray)
    for (const [i, ls] of layout.entries()) {
      const section = sections[i]
      if (!section || ls.gutters.length === 0) continue
      const sectionH = ls.box.y1 - ls.box.y0
      for (const stroke of shapes.strokes) {
        if (stroke.orientation !== 'v' || sepStrokes.has(stroke)) continue
        const x = (stroke.box.x0 + stroke.box.x1) / 2
        if (!ls.gutters.some((g) => x >= g.lo && x <= g.hi)) continue
        const overlap = Math.min(stroke.box.y1, ls.box.y1) - Math.max(stroke.box.y0, ls.box.y0)
        if (overlap < COL_SEP_MIN_SECTION_COVER * sectionH) continue
        section.colSep = true
        sepStrokes.add(stroke)
      }
    }
  }

  // P7: leftover horizontal rules become paragraph borders / thin bar
  // paragraphs before the spacing chain measures the flow
  const decor = applyDecorBorders(
    sections,
    shapes.strokes,
    [...gridBoxes, ...tables.map((t) => t.box)],
    sepStrokes.size > 0 || zoneStrokes.size > 0
      ? new Set([...styled.consumedStrokes, ...sepStrokes, ...zoneStrokes])
      : styled.consumedStrokes,
  )
  if (decor.ignoredVertical + decor.droppedBars > 0) {
    base.ignoredVerticalDecor = decor.ignoredVertical + decor.droppedBars
  }

  const warnings = [...styled.warnings, ...panelNotes, ...applySpacingChain(sections)]
  if (warnings.length > 0) base.warnings = warnings
  base.sections = sections
  base.blocks = flattenSections(sections)

  // card regions (P20): portrait flow pages only — landscape/slide pages pin
  // their text absolutely too, so plate and text cannot drift apart there
  if (extracted.widthPt < extracted.heightPt && cardCandidates.length > 0) {
    detectCardRegions(base, cardCandidates)
  }
  base.confidence = pageConfidence({
    badUnicodeRatio: extracted.badUnicodeRatio,
    streamTableConfidences: [...formTables, ...zoneTables, ...streamTables]
      .map((t) => t.confidence)
      .filter((c): c is number => c !== undefined),
    warningCount: warnings.length,
  })
  return base
}
