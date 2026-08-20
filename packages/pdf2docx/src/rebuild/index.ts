/**
 * Rebuild layer: IR pages → docx-engine SaveBlock[] → .docx bytes, following
 * the exportDocxBytes() pattern from apps/markdown (buildBlankDocx → parseDocx
 * → saveDocx). Unit conventions per docx-engine: lengths in twips (pt × 20),
 * font sizes in half-points, image display sizes in CSS px (pt × 96/72).
 */
// relative import (not the @genoffice/docx-engine package name): in a git
// worktree node_modules is a symlink into the main checkout, and a bare
// specifier would silently run ANOTHER checkout's docx-engine at runtime
// (tsx/node); tsconfig paths + the vitest alias only cover types and tests
import {
  buildAnchoredTextboxParagraphXml,
  buildBlankDocx,
  generateTableModelXml,
  parseDocx,
  saveDocx,
} from '../../../docx-engine/src/index'
import type {
  CustomNumberingLevel,
  GeneratedBlock,
  HeaderFooter,
  HfParagraph,
  NewImage,
  NoteInfo,
  NoteRun,
  ParaFormat,
  Run,
  SaveBlock,
  SaveOptions,
  SectionSettings,
  TableCell,
  TableModel,
  TableParagraph,
  TextboxContentParagraph,
} from '../../../docx-engine/src/index'
import type { FurnitureHf } from '../analyze/furniture'
import { detectCellHAlign } from '../analyze/table'
import { preflightFitBlock } from './fit'
import { applyOutputFontSubstitutions } from './fontmap'
import { rectWidth, rectHeight, rectCenterX, rectUnionAll, approxEq } from '../geometry'
import type {
  CardRegion,
  FloatPlacement,
  ImageBlock,
  IrPage,
  Line,
  PageBlock,
  PageRender,
  PageSection,
  Span,
  TableBlock,
  TextBlock,
} from '../ir'
import type { UnicodeScript } from '../script'
import { isEastAsianScript, isNoSpaceScript } from '../script'

const PT_TO_TWIPS = 20
const PT_TO_PX = 96 / 72

/**
 * page margins derived from content are clamped to this range (points).
 * The floor is small on purpose: tight-margin sources (landscape
 * newsletters) must keep their usable area — clamping the margin ABOVE the
 * measured one shrinks the page below its own ink extent and guarantees
 * overflow. Page numbers/headers that used to distort the measurement are
 * gone (P6 furniture pass).
 */
const MARGIN_MIN_PT = 12
const MARGIN_MAX_PT = 108
const MARGIN_DEFAULT_PT = 72
/**
 * The BOTTOM margin gives the measured value back this much slack (with a
 * floor). Every rebuilt page starts on an explicit page break, so pagination
 * only ever depends on the bottom margin when a page's content slightly
 * overflows (twips rounding, atLeast table rows growing a hair, substituted
 * font metrics) — the slack absorbs that without moving any fitting content.
 */
const BOTTOM_SLACK_PT = 24
const BOTTOM_MARGIN_FLOOR_PT = 20
/**
 * The page vertical budget (P8) stops this far above the slack-reduced
 * bottom margin. Larger than BOTTOM_SLACK on purpose: a budget-clamped page
 * fills flush, and substituted fonts wrapping a long paragraph one or two
 * extra lines costs ~2 line heights that only this headroom can absorb —
 * whitespace on such a page shrinks a little more instead of spilling a
 * near-empty extra page.
 */
const BUDGET_SLACK_PT = 48
/** separator rule + its leading Word adds above a page's footnote area (P12 D) */
const FOOTNOTE_SEPARATOR_PT = 12
/**
 * Whole-page compression floor (P9 C): a page whose measured block heights
 * exceed its budget shrinks proportionally, but never below this — beyond it
 * the text is unreadable and the page was going to spill regardless.
 */
const MIN_HEIGHT_SCALE = 0.55
/**
 * Compressed pages squeeze this much further below the budget: they fill
 * flush by construction, so substituted-font wrap growth (a title gaining a
 * line) would otherwise eat the whole BUDGET_SLACK and still spill.
 */
const COMPRESS_SAFETY_PT = 36
/**
 * Multi-line text at/above this font size (pt) reserves one extra line pitch
 * in the page budget: display-size titles re-wrap one line taller under a
 * substituted font, and at 30pt+ a single gained line alone out-eats the
 * whole BUDGET_SLACK.
 */
const WRAP_RISK_FONT_PT = 28

const B64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'

export function bytesToBase64(bytes: Uint8Array): string {
  // hosted runtimes (node / electron main) expose Buffer — much faster than JS
  const B = (
    globalThis as { Buffer?: { from(b: Uint8Array): { toString(enc: 'base64'): string } } }
  ).Buffer
  if (B) return B.from(bytes).toString('base64')
  const parts: string[] = []
  for (let i = 0; i < bytes.length; i += 3) {
    const b0 = bytes[i]!
    const b1 = bytes[i + 1]
    const b2 = bytes[i + 2]
    parts.push(
      B64_ALPHABET[b0 >> 2]!,
      B64_ALPHABET[((b0 & 3) << 4) | ((b1 ?? 0) >> 4)]!,
      b1 === undefined ? '=' : B64_ALPHABET[((b1 & 15) << 2) | ((b2 ?? 0) >> 6)]!,
      b2 === undefined || b1 === undefined ? '=' : B64_ALPHABET[b2 & 63]!,
    )
  }
  return parts.join('')
}

interface PageGeometry {
  section: SectionSettings
  /** usable text column width in points */
  contentWidthPt: number
  marginLeftPt: number
  marginRightPt: number
  marginTopPt: number
  /** the slack-reduced bottom margin actually written to the section (pt) */
  marginBottomPt: number
}

function clampMargin(pt: number, maxPt = MARGIN_MAX_PT): number {
  // no blocks at all → nothing measured → the conventional default;
  // content TOUCHING the page edge is a real measurement → minimal margin
  if (!Number.isFinite(pt)) return MARGIN_DEFAULT_PT
  if (pt <= 0) return MARGIN_MIN_PT
  return Math.min(maxPt, Math.max(MARGIN_MIN_PT, pt))
}

/** One section from page 1's geometry; margins hug the content across all pages. */
function computeGeometry(pages: IrPage[]): PageGeometry {
  const first = pages[0]
  const widthPt = first?.widthPt || 612
  const heightPt = first?.heightPt || 792

  let left = Infinity
  let right = Infinity
  let top = Infinity
  let bottom = Infinity
  /** page-edge decor floats (footer graphics, header ribbons) are pinned at
   * absolute page coordinates — letting them hug the HORIZONTAL margins drags
   * the text column to the page edge and every unindented paragraph with it
   * (P20). Vertical hug keeps them: the band is part of the vertical layout. */
  const isEdgeBandFloat = (block: PageBlock, page: IrPage): boolean =>
    block.kind === 'image' &&
    block.float !== undefined &&
    (block.box.y1 <= 0.12 * page.heightPt || block.box.y0 >= 0.88 * page.heightPt)
  for (const page of pages) {
    for (const block of page.blocks) {
      if (!isEdgeBandFloat(block, page)) {
        left = Math.min(left, block.box.x0)
        right = Math.min(right, page.widthPt - block.box.x1)
      }
      top = Math.min(top, page.heightPt - block.box.y1)
      bottom = Math.min(bottom, block.box.y0)
    }
  }
  // horizontal margins hug the measured content even when deep (P16 C): a
  // right-half column layout (Chromium print) has a true ~278pt left margin,
  // and capping it at 1.5in shifts every paragraph 170pt left. The cap only
  // guards against degenerate content (≥10% of the width must remain).
  const hMaxPt = Math.max(MARGIN_MAX_PT, widthPt * 0.48)
  let marginLeftPt = clampMargin(left, hMaxPt)
  let marginRightPt = clampMargin(right, hMaxPt)
  // a deep-hugged margin reproduces a narrow source column EXACTLY, and
  // substituted fonts wrap a hair wider — the opposite margin funds the
  // wrap headroom or every page spills its wrapped tail lines
  const WRAP_HEADROOM_PT = 24
  if (marginLeftPt > MARGIN_MAX_PT) {
    marginRightPt = Math.max(MARGIN_MIN_PT, marginRightPt - WRAP_HEADROOM_PT)
  } else if (marginRightPt > MARGIN_MAX_PT) {
    marginLeftPt = Math.max(MARGIN_MIN_PT, marginLeftPt - WRAP_HEADROOM_PT)
  }
  const marginBottomPt = Math.max(BOTTOM_MARGIN_FLOOR_PT, clampMargin(bottom) - BOTTOM_SLACK_PT)
  const section: SectionSettings = {
    pageWidth: Math.round(widthPt * PT_TO_TWIPS),
    pageHeight: Math.round(heightPt * PT_TO_TWIPS),
    orientation: widthPt > heightPt ? 'landscape' : 'portrait',
    marginTop: Math.round(clampMargin(top) * PT_TO_TWIPS),
    marginRight: Math.round(marginRightPt * PT_TO_TWIPS),
    marginBottom: Math.round(marginBottomPt * PT_TO_TWIPS),
    marginLeft: Math.round(marginLeftPt * PT_TO_TWIPS),
    pageBorder: false,
    columns: 1,
  }
  return {
    section,
    contentWidthPt: Math.max(72, widthPt - marginLeftPt - marginRightPt),
    marginLeftPt,
    marginRightPt,
    marginTopPt: clampMargin(top),
    marginBottomPt,
  }
}

/** default complex-script fonts when the PDF declares none (plan §5.3) */
const RTL_FALLBACK_CS_FONT: Partial<Record<UnicodeScript, string>> = {
  arabic: 'Traditional Arabic',
  hebrew: 'David',
}

/**
 * Word's w:highlight palette (docx-engine models named highlights only, no
 * run-level w:shd) — PDF fill colors map to the nearest entry.
 */
const HIGHLIGHT_PALETTE: Array<{ name: string; rgb: [number, number, number] }> = [
  { name: 'black', rgb: [0, 0, 0] },
  { name: 'blue', rgb: [0, 0, 255] },
  { name: 'cyan', rgb: [0, 255, 255] },
  { name: 'green', rgb: [0, 255, 0] },
  { name: 'magenta', rgb: [255, 0, 255] },
  { name: 'red', rgb: [255, 0, 0] },
  { name: 'yellow', rgb: [255, 255, 0] },
  { name: 'white', rgb: [255, 255, 255] },
  { name: 'darkBlue', rgb: [0, 0, 139] },
  { name: 'darkCyan', rgb: [0, 139, 139] },
  { name: 'darkGreen', rgb: [0, 100, 0] },
  { name: 'darkMagenta', rgb: [139, 0, 139] },
  { name: 'darkRed', rgb: [139, 0, 0] },
  { name: 'darkYellow', rgb: [128, 128, 0] },
  { name: 'darkGray', rgb: [128, 128, 128] },
  { name: 'lightGray', rgb: [192, 192, 192] },
]

/** hex RRGGBB → nearest named w:highlight; white maps to "no highlight" */
export function nearestHighlight(hex: string): string | undefined {
  if (!/^[0-9a-f]{6}$/i.test(hex)) return undefined
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  let best: string | undefined
  let bestDist = Infinity
  for (const { name, rgb } of HIGHLIGHT_PALETTE) {
    const dist = (r - rgb[0]) ** 2 + (g - rgb[1]) ** 2 + (b - rgb[2]) ** 2
    if (dist < bestDist) {
      bestDist = dist
      best = name
    }
  }
  return best === 'white' ? undefined : best
}

function runFromSpan(span: Span): Run {
  // footnote anchor (P6): a bare reference-marker run — Word renders the
  // number itself, the span's own text stays empty
  if (span.noteRef !== undefined) {
    return { text: span.text, noteRef: { kind: 'footnote', id: span.noteRef } }
  }
  const run: Run = { text: span.text }
  if (span.bold) run.bold = true
  if (span.italic) run.italic = true
  if (span.underline) run.underline = true
  if (span.strike) run.strike = true
  if (span.highlight) {
    const named = nearestHighlight(span.highlight)
    if (named) run.highlight = named
  }
  if (span.color && span.color !== '000000') run.color = span.color
  const halfPoints = Math.round(span.fontSize * 2)
  if (halfPoints > 0) run.sizeHalfPoints = halfPoints
  if (span.fontFamily) {
    // CJK-family scripts fill the w:eastAsia slot (docx-engine Run.font);
    // everything else declares only the Latin slots (fontAscii)
    if (isEastAsianScript(span.script)) run.font = span.fontFamily
    else run.fontAscii = span.fontFamily
  }
  if (span.dir === 'rtl') {
    run.rtl = true
    run.fontCs = span.fontFamily || RTL_FALLBACK_CS_FONT[span.script] || 'Traditional Arabic'
  }
  // character compression (P5): w:w / w:spacing are outside docx-engine's run
  // model — they ride in rawRPr, whose unmanaged children survive generation
  const compression: string[] = []
  // invisible source text (PDF Tr 3/7, Word's hidden formatting marks like
  // section-break labels): w:vanish keeps it present but unseen, as in the source (P20)
  if (span.invisible) compression.push('<w:vanish/>')
  if (span.charSpacingPt !== undefined) {
    compression.push(`<w:spacing w:val="${Math.round(span.charSpacingPt * 20)}"/>`)
  }
  if (span.charScale !== undefined) {
    compression.push(`<w:w w:val="${Math.round(span.charScale * 100)}"/>`)
  }
  if (compression.length > 0) run.rawRPr = `<w:rPr>${compression.join('')}</w:rPr>`
  return run
}

const sameRunStyle = (a: Run, b: Run): boolean =>
  a.bold === b.bold &&
  a.italic === b.italic &&
  a.underline === b.underline &&
  a.strike === b.strike &&
  a.highlight === b.highlight &&
  a.color === b.color &&
  a.sizeHalfPoints === b.sizeHalfPoints &&
  a.font === b.font &&
  a.fontAscii === b.fontAscii &&
  a.fontCs === b.fontCs &&
  a.rtl === b.rtl &&
  a.rawRPr === b.rawRPr

function pushRun(runs: Run[], run: Run): void {
  const last = runs[runs.length - 1]
  // reference-marker runs are atomic — merging text into them would lose it
  if (last && !last.noteRef && !run.noteRef && sameRunStyle(last, run)) last.text += run.text
  else runs.push(run)
}

/** trailing hyphen glyphs dropped when PDFium marked end-of-line hyphenation */
const TRAILING_HYPHEN = /[-­‐]$/

/** Flatten a text block's lines into one paragraph's runs (dehyphenate + script-aware joins). */
function paragraphRuns(block: TextBlock): Run[] {
  const runs: Run[] = []
  for (const [i, line] of block.lines.entries()) {
    if (i > 0) {
      const prevLine = block.lines[i - 1]!
      const last = runs[runs.length - 1]
      const lastSpan = prevLine.spans[prevLine.spans.length - 1]
      const nextSpan = line.spans[0]
      if (line.hardBreakBefore) {
        // intentional intra-paragraph break (P7): '\n' becomes <w:br/>
        if (last && !last.noteRef) last.text += '\n'
        else runs.push({ text: '\n' })
      } else if (prevLine.endsWithHyphen && last) {
        last.text = last.text.replace(TRAILING_HYPHEN, '')
      } else if (
        last &&
        lastSpan &&
        nextSpan &&
        !isNoSpaceScript(lastSpan.script) &&
        !isNoSpaceScript(nextSpan.script)
      ) {
        last.text += ' '
      }
    }
    for (const span of line.spans) pushRun(runs, runFromSpan(span))
  }
  return runs
}

// ── explicit paragraph spacing (P5) ──
// Every rebuilt paragraph writes its own w:spacing. Anything left implicit
// falls through to the blank template's docDefaults (w:after=120,
// w:line=276 lineRule=auto), which inflated rebuilt documents by 50–100%
// in page count: +6pt after every paragraph plus 1.15× line height.

/** floor for measured line heights (twips, 2pt) — guards degenerate ink boxes */
const MIN_LINE_TWIPS = 40
/**
 * leading factor over the tallest font size — the exact line box must clear
 * the glyphs' natural extent (ascender+descender ≈ 1.15–1.2 em), or tall runs
 * clip/overlap the line below. The measured ink-box height is only the inked
 * pixels (glyphs never fill the full em), so it badly understates the pitch
 * of a SINGLE-line block (headings/titles — the main clipping victims).
 * Multi-line blocks keep their measured pitch untouched: their box already
 * includes the inter-line leading, and flooring them to the tallest span
 * (dense CJK reports mix 1.5× emphasis spans into body text) measurably
 * doubled page counts. 1.16 tracks the blank template's ~1.15×.
 */
const LINE_LEADING = 1.16
/** near-zero exact height for empty utility paragraphs (break carriers, spacers) */
const TIGHT_LINE_TWIPS = 20

/**
 * Tallest font size (pt) of a text block. The exact line box must clear the
 * TALLEST glyph on the line, not the most common one — a mixed-size line
 * (e.g. a spaced-out 22pt chapter title riding over 14pt body, or one large
 * lead capital) still clips the big run if the box is floored to the dominant
 * size. Max keeps every glyph inside the exact box; the before_space chain
 * still owns inter-block whitespace so this does not re-inflate pages.
 */
function maxFontSizePt(block: TextBlock): number {
  let max = 0
  for (const line of block.lines) {
    for (const span of line.spans) {
      if (span.fontSize > max) max = span.fontSize
    }
  }
  return max
}

/**
 * Paragraph format for a block. Alignment stays VISUAL (docx-engine converts
 * to Word's logical jc for bidi paragraphs); each direction's default edge is
 * omitted — left for LTR, right for RTL. The measured first-line indent is
 * left-edge based, meaningless for RTL, so it is dropped there.
 *
 * Spacing is explicit: the measured line pitch (block ink height spread over
 * its lines) lands in w:line with lineRule=exact — the paragraph then occupies
 * exactly its measured extent no matter what font the viewer substitutes
 * (atLeast would re-inflate whenever the substituted font's natural line
 * height exceeds the measured pitch). Inter-block whitespace is all carried
 * by the before_space chain, so w:after is pinned to 0.
 */
function blockFormat(block: TextBlock, heightScale = 1): ParaFormat {
  // measured pitch: block ink height spread over its lines (understates the
  // true leading for short blocks — ink never fills the em box). An exact
  // baseline-span pitch was tried for pinned verse stacks (P22 E) and
  // reverted: hard-break stacks from the slide/list machinery drifted
  // whole form/deck pages for a ~0.001 poems gain.
  const measuredPitchTwips = Math.round(
    (rectHeight(block.box) / Math.max(1, block.lines.length)) * PT_TO_TWIPS,
  )
  // font-based floor, SINGLE-line blocks only: a lone line's ink box is all
  // the pitch information there is, and it understates the em — exact then
  // clips the glyphs' ascenders/descenders (title/heading truncation). A
  // multi-line box already spans its inter-line leading, so its measured
  // pitch is trusted as-is — flooring it to the tallest span re-inflates
  // dense mixed-size text into extra pages.
  const fontFloorTwips =
    block.lines.length <= 1 ? Math.round(maxFontSizePt(block) * LINE_LEADING * PT_TO_TWIPS) : 0
  // whole-page compression (P9 C) applies AFTER the floors: on a page whose
  // block heights alone exceed the budget every pitch shrinks by one factor
  const lineTwips = Math.max(MIN_LINE_TWIPS, measuredPitchTwips, fontFloorTwips)
  const format: ParaFormat = {
    lineRule: 'exact',
    lineRawTwips:
      heightScale < 1 ? Math.max(TIGHT_LINE_TWIPS, Math.round(lineTwips * heightScale)) : lineTwips,
    spaceAfter: 0,
  }
  if (block.dir === 'rtl') {
    format.bidi = true
    if (block.align !== 'right') format.align = block.align
  } else {
    if (block.align !== 'left') format.align = block.align
    if (block.firstLineIndentPt > 0) {
      format.indentFirstLine = Math.round(block.firstLineIndentPt * PT_TO_TWIPS)
    }
  }
  return format
}

// ── list numbering (P4) ──

/** bullet numbering definition shipped by docx-engine's blank template */
const BLANK_BULLET_NUM_ID = '1'
/** decimal abstractNum of the blank template (numbering.xml: 0 = bullet, 1 = decimal) */
const DECIMAL_ABSTRACT_NUM_ID = '1'
/** numIds allocated for detected ordered lists start here (well clear of the blank's 1/2) */
const FIRST_ALLOCATED_NUM_ID = 100
/** numbering depth of the blank template's abstract definitions */
const NUM_LEVELS = 5

/** collects the numbering parts one conversion needs (docx-engine SaveOptions.numbering) */
class ListNumbering {
  restartNums: NonNullable<NonNullable<SaveOptions['numbering']>['restartNums']> = []
  newDefs: NonNullable<NonNullable<SaveOptions['numbering']>['newDefs']> = []
  private byRun = new Map<string, string>()
  private nextId = FIRST_ALLOCATED_NUM_ID

  /** docx numId for a text block's list annotation (allocating on first sight) */
  numIdFor(page: number, list: NonNullable<TextBlock['list']>): string {
    if (list.kind === 'bullet') return BLANK_BULLET_NUM_ID
    const key = `${page}:${list.seqId ?? 0}`
    let numId = this.byRun.get(key)
    if (numId !== undefined) return numId
    numId = String(this.nextId++)
    this.byRun.set(key, numId)
    const start = list.start ?? 1
    if (list.style === 'multi') {
      // outline numbers ("3.1.15."): lvlText composes all ordinals; unused
      // higher levels display their start override, so the literal PDF value
      // reproduces. Indent mirrors the measured layout: number at the margin,
      // text hanging at 720 twips (36pt) like the source templates.
      const levels: CustomNumberingLevel[] = Array.from({ length: NUM_LEVELS }, (_, i) => ({
        numFmt: 'decimal',
        lvlText: `${Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.')}.`,
        indentLeft: 720,
        hanging: 720,
        start: list.startValues?.[i] ?? 1,
      }))
      this.newDefs.push({ numId, kind: 'ordered', levels })
      return numId
    }
    if (list.style === 'paren' || list.style === 'parens') {
      // "%1)" / "(%1)" markers need their own abstractNum (custom lvlText)
      const levels: CustomNumberingLevel[] = Array.from({ length: NUM_LEVELS }, (_, i) => ({
        numFmt: 'decimal',
        lvlText:
          i === list.level ? (list.style === 'paren' ? `%${i + 1})` : `(%${i + 1})`) : `%${i + 1}.`,
        indentLeft: 720 * (i + 1),
        hanging: 360,
        start: i === list.level ? start : 1,
      }))
      this.newDefs.push({ numId, kind: 'ordered', levels })
    } else {
      // plain "N." lists restart the blank decimal abstract at the PDF's first ordinal
      const startOverrides: Record<number, number> = {}
      for (let i = 0; i < NUM_LEVELS; i++) startOverrides[i] = i === list.level ? start : 1
      this.restartNums.push({ numId, abstractNumId: DECIMAL_ABSTRACT_NUM_ID, startOverrides })
    }
    return numId
  }

  toSaveOptions(): SaveOptions['numbering'] | undefined {
    if (this.restartNums.length === 0 && this.newDefs.length === 0) return undefined
    return {
      ...(this.newDefs.length > 0 ? { newDefs: this.newDefs } : {}),
      ...(this.restartNums.length > 0 ? { restartNums: this.restartNums } : {}),
    }
  }
}

function textBlockToSave(
  block: TextBlock,
  pageBreakBefore: boolean,
  spacingBeforePt: number,
  pageIndex: number,
  numbering: ListNumbering,
  heightScale = 1,
): SaveBlock {
  const format = blockFormat(block, heightScale)
  if (pageBreakBefore) format.pageBreakBefore = true
  // before_space chain (P3): analyze-computed inter-block spacing plus the
  // page-top leading, already clamped to the page's vertical budget (P8) by
  // the assembly loop
  let spacingPt = spacingBeforePt
  // decorative rule (P7): stray horizontal line mapped onto w:pBdr; 'top' =
  // the line sits above the text = w:top. Paragraph borders CONSUME vertical
  // space (LibreOffice adds w:space + the line width to the paragraph box),
  // so both are paid out of the measured spacing-before — net page height
  // stays put, or every ruled page creeps into overflow. The text gap is
  // clamped to what the spacing can fund.
  if (block.border && block.border.side === 'left') {
    // vertical accent bar (P14 C): drawn left of the text, consumes no
    // vertical flow — the spacing budget below stays untouched
    format.borders = 'l'
    format.borderStyle = {
      color: block.border.color,
      szEighths: Math.max(2, Math.round(block.border.widthPt * 8)),
      spacePt: block.border.spacePt,
    }
  } else if (block.border) {
    const b = block.border
    // +1pt beyond gap + line width: LibreOffice's border box rounds up, and
    // rebuilt pages run flush — a sub-pt overshoot per rule still overflows
    const BORDER_FUND_SLACK_PT = 1
    // a page-break paragraph's top border renders its w:space ABOVE the text
    // at the page top (nothing above funds it — LibreOffice collapses that
    // space after a section break but honours it after w:pageBreakBefore),
    // so the whole page shifts down by the gap: hug the rule to the text
    const gapPt =
      pageBreakBefore && b.side === 'top'
        ? 0
        : Math.max(0, Math.min(b.spacePt, spacingPt - b.widthPt - BORDER_FUND_SLACK_PT))
    format.borders = b.side === 'top' ? 't' : 'b'
    format.borderStyle = {
      color: b.color,
      szEighths: Math.max(2, Math.round(b.widthPt * 8)),
      spacePt: gapPt,
    }
    spacingPt = Math.max(0, spacingPt - gapPt - b.widthPt - BORDER_FUND_SLACK_PT)
    // standalone bar (no text): shrink the carrier line to near-zero and pay
    // for what still exceeds the stroke's own measured ink height — dozens of
    // rules on a designed page must not accumulate into an overflow
    if (block.lines.length === 0) {
      format.lineRawTwips = TIGHT_LINE_TWIPS
      const ownPt = TIGHT_LINE_TWIPS / PT_TO_TWIPS + b.widthPt
      spacingPt = Math.max(0, spacingPt - Math.max(0, ownPt - rectHeight(block.box)))
    }
    if (b.indentRightPt !== undefined && b.indentRightPt > 0) {
      format.indentRight = Math.round(b.indentRightPt * PT_TO_TWIPS)
    }
    if (b.indentLeftPt !== undefined && b.indentLeftPt > 0) {
      format.indentLeft = Math.round(b.indentLeftPt * PT_TO_TWIPS)
    }
  }
  if (spacingPt > 0) format.spaceBefore = Math.round(spacingPt * PT_TO_TWIPS)
  const generated: GeneratedBlock = { type: 'paragraph', runs: paragraphRuns(block) }
  if (block.list) {
    // real docx list item: the marker regenerates from the numbering part,
    // and the numbering level's own indent applies
    generated.type = 'listItem'
    generated.list = {
      kind: block.list.kind,
      numId: numbering.numIdFor(pageIndex, block.list),
      ilvl: block.list.level,
    }
    delete format.indentFirstLine
  }
  generated.format = format
  return { kind: 'generated', block: generated }
}

const escXml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/** minimal literal rPr XML for hand-built paragraphs (TOC entries, leader rulings) */
function rawRPrXml(run: Run): string {
  const props: string[] = []
  const fonts: string[] = []
  if (run.fontAscii)
    fonts.push(`w:ascii="${escXml(run.fontAscii)}" w:hAnsi="${escXml(run.fontAscii)}"`)
  if (run.font) fonts.push(`w:eastAsia="${escXml(run.font)}"`)
  if (fonts.length > 0) props.push(`<w:rFonts ${fonts.join(' ')}/>`)
  if (run.bold) props.push('<w:b/>')
  if (run.italic) props.push('<w:i/>')
  if (run.color) props.push(`<w:color w:val="${escXml(run.color)}"/>`)
  if (run.sizeHalfPoints) {
    props.push(`<w:sz w:val="${run.sizeHalfPoints}"/><w:szCs w:val="${run.sizeHalfPoints}"/>`)
  }
  return props.length > 0 ? `<w:rPr>${props.join('')}</w:rPr>` : ''
}

/** minimal literal run XML for hand-built paragraphs (TOC entries) */
function rawRunXml(run: Run): string {
  return `<w:r>${rawRPrXml(run)}<w:t xml:space="preserve">${escXml(run.text)}</w:t></w:r>`
}

// ── full-line leader rulings (P17) ──
// Answer rulings ('.....', '………', '____') re-render a hair wider under the
// substituted font: the last glyph wraps onto its own line and every ruling
// grows the page one pitch (a gov form page carries 6-9 of them). A right
// tab with a leader ends exactly at the measured x1 and cannot wrap.
const LEADER_STYLES: Record<string, string> = {
  '.': 'dot',
  '…': 'dot',
  '‥': 'dot',
  '·': 'middleDot',
  _: 'underscore',
}
/** leader style when the block is one line of a single repeated leader glyph */
function leaderStyleOf(block: TextBlock): string | null {
  if (block.lines.length !== 1 || block.tocEntry || block.list || block.border) return null
  if (block.dir === 'rtl') return null
  const spans = block.lines[0]!.spans
  if (spans.some((s) => s.noteRef !== undefined)) return null
  const compact = spans
    .map((s) => s.text)
    .join('')
    .replace(/[\s ]+/g, '')
  if (compact.length < 8) return null
  const style = LEADER_STYLES[compact[0]!]
  if (style === undefined) return null
  for (const ch of compact) if (ch !== compact[0]) return null
  return style
}

/** a leader ruling as an empty paragraph with a leader tab out to measured x1 */
function leaderToSave(
  block: TextBlock,
  leader: string,
  colBasePt: number,
  pageBreakBefore: boolean,
  spacingPt: number,
  heightScale = 1,
): SaveBlock {
  const format = blockFormat(block, heightScale)
  const indentPt = Math.max(0, block.box.x0 - colBasePt)
  const tabPosTwips = Math.max(
    Math.round(indentPt * PT_TO_TWIPS) + 36,
    Math.round((block.box.x1 - colBasePt) * PT_TO_TWIPS),
  )
  const pPr =
    '<w:pPr>' +
    (pageBreakBefore ? '<w:pageBreakBefore/>' : '') +
    `<w:tabs><w:tab w:val="right" w:leader="${leader}" w:pos="${tabPosTwips}"/></w:tabs>` +
    `<w:spacing w:before="${Math.round(spacingPt * PT_TO_TWIPS)}" w:after="0"` +
    ` w:line="${format.lineRawTwips}" w:lineRule="exact"/>` +
    (indentPt >= 1 ? `<w:ind w:left="${Math.round(indentPt * PT_TO_TWIPS)}"/>` : '') +
    '</w:pPr>'
  const styleRun = runFromSpan(block.lines[0]!.spans[0]!)
  return { kind: 'xml', xml: `<w:p>${pPr}<w:r>${rawRPrXml(styleRun)}<w:tab/></w:r></w:p>` }
}

/**
 * A TOC dot-leader entry (P6): TOC pStyle + measured indent, the literal
 * dots replaced by a right-aligned dot-leader tab, page number after it.
 * Indent and tab measure from the COLUMN base, and the tab ends at the
 * entry's measured right edge clamped into the column (P23): the old
 * full-content-width tab overshot multi-column sections, double-counted the
 * column x as indent, and LO wrapped each entry word-by-word (childAttachments
 * src4 checklist).
 */
function tocToSave(
  block: TextBlock,
  colBasePt: number,
  colWidthPt: number,
  pageBreakBefore: boolean,
  spacingPt: number,
  heightScale = 1,
): SaveBlock {
  const toc = block.tocEntry!
  const format = blockFormat(block, heightScale)
  const indentPt = Math.max(0, block.box.x0 - colBasePt)
  const tabPos = Math.min(
    Math.round(colWidthPt * PT_TO_TWIPS),
    Math.max(
      Math.round(indentPt * PT_TO_TWIPS) + 36,
      Math.round((block.box.x1 - colBasePt) * PT_TO_TWIPS),
    ),
  )
  const pPr =
    `<w:pPr><w:pStyle w:val="TOC${Math.min(toc.level, 9)}"/>` +
    (pageBreakBefore ? '<w:pageBreakBefore/>' : '') +
    `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${tabPos}"/></w:tabs>` +
    `<w:spacing w:before="${Math.round(spacingPt * PT_TO_TWIPS)}" w:after="0"` +
    ` w:line="${format.lineRawTwips}" w:lineRule="exact"/>` +
    (indentPt >= 1 ? `<w:ind w:left="${Math.round(indentPt * PT_TO_TWIPS)}"/>` : '') +
    '</w:pPr>'
  const runs = paragraphRuns(block)
  const titleXml = runs.map(rawRunXml).join('')
  const numberXml =
    '<w:r><w:tab/></w:r>' +
    rawRunXml({ ...(runs[runs.length - 1] ?? { text: '' }), text: toc.pageNumber })
  return { kind: 'xml', xml: `<w:p>${pPr}${titleXml}${numberXml}</w:p>` }
}

/** cell content → docx-engine rich cell paragraphs (RTL/alignment per block) */
function cellParagraphs(
  cell: TableBlock['rows'][number][number],
  padTrailingSpace: boolean,
  heightScale = 1,
): TableParagraph[] {
  const paras = cell.blocks.map((block) => {
    const format = blockFormat(block, heightScale)
    // zone cells (P22 A) measure their stack's internal INK gaps; other
    // detectors never set spacingBeforePt on cell blocks. The paragraph's
    // line box is taller than its ink (leading) — deduct the excess so the
    // cumulative positions stay put
    if (block.spacingBeforePt !== undefined && block.spacingBeforePt > 0) {
      const inkTwips = Math.round(rectHeight(block.box) * PT_TO_TWIPS)
      const excess = Math.max(0, (format.lineRawTwips ?? inkTwips) - inkTwips)
      const spaceBefore = Math.round(block.spacingBeforePt * PT_TO_TWIPS) - excess
      if (spaceBefore > 0) format.spaceBefore = spaceBefore
    }
    return { ...format, runs: paragraphRuns(block) }
  })
  // FORM cells end in a space (invisible in rendering): Word's own
  // field-result cells carry one, and plain-text extractors that concatenate
  // cells would otherwise fuse the last word with the next cell's first.
  // Plain tables stay verbatim — their sources have no such space.
  if (padTrailingSpace) {
    const lastRun = paras[paras.length - 1]?.runs.at(-1)
    if (lastRun && lastRun.text !== '' && !/\s$/.test(lastRun.text)) lastRun.text += ' '
  }
  return paras
}

/** measured grid row height: the smallest single-row cell decides (merged cells span rows) */
function rowHeightTwips(row: TableBlock['rows'][number]): number | null {
  const heights = row.filter((c) => c.vMerge === undefined).map((c) => rectHeight(c.box))
  if (heights.length === 0) return null
  return Math.max(1, Math.round(Math.min(...heights) * PT_TO_TWIPS))
}

/** stream/form (borderless) tables carry explicit none-borders */
const NO_BORDER = { style: 'none' } as const
const NO_BORDERS = {
  top: NO_BORDER,
  left: NO_BORDER,
  bottom: NO_BORDER,
  right: NO_BORDER,
  insideH: NO_BORDER,
  insideV: NO_BORDER,
}
/**
 * near-zero cell padding (twips): measured cell text already sits inside the
 * grid box, and Word's default 108-twip side margins would shrink the text
 * column and wrap lines the PDF kept whole — rows then grow past their
 * measured height and pages overflow
 */
const CELL_MAR_TWIPS = { left: 15, right: 15 }

/** lattice grids drawn in a non-black ink keep their color — the default
 * single borders render black and repaint a light-ruled table (deck-gov's
 * white rulings between zebra fills read as one dark slab) */
function colorBorders(color: string): NonNullable<TableModel['borders']> {
  const line = { style: 'single', szEighths: 4, color }
  return { top: line, left: line, bottom: line, right: line, insideH: line, insideV: line }
}

/** LibreOffice draws a bordered row at trHeight PLUS the horizontal border
 * width (~0.5pt single line) — deducted from measured lattice row heights */
const BORDER_EAT_TWIPS = 10

/** Word's hard page-dimension ceiling (22in per side) */
const WORD_MAX_PAGE_PT = 22 * 72

/** slack before the set-solid squeeze fires (twips) — measurement noise on
 * small tables must not trigger it */
const TABLE_SQUEEZE_TOL_TWIPS = 60

/**
 * Set-solid table squeeze (P24 B): dense tabular reports (USGS discharge
 * matrices) print rows at a pitch BELOW fontSize × LINE_LEADING, so the
 * single-line font floor grows every cell paragraph past the source row and
 * a flush-full page spills its tail rows (merge-PDFBOX-4417: +1.5pt × 46
 * rows ≈ 3 spilled lines). When the emitted rows sum past the table's own
 * measured box, scale the whole table back to its box — the table's internal
 * geometry is trusted over the per-line floor, mirroring the multi-line
 * "measured pitch is trusted" rule. vMerged grids are skipped (a spanning
 * cell's stack legitimately exceeds its home row), as are zone tables
 * (sepRule — they rebuild exact line geometry with their own spacing math).
 */
function tableSqueezeFactor(block: TableBlock, heightScale: number): number {
  if (block.sepRule !== undefined) return 1
  if (block.rows.some((row) => row.some((c) => c.vMerge !== undefined))) return 1
  const boxTwips = rectHeight(block.box) * PT_TO_TWIPS * heightScale
  let emitTwips = 0
  for (const row of block.rows) {
    let tallest = (rowHeightTwips(row) ?? 0) * heightScale
    for (const cell of row) {
      let cellTwips = 0
      for (const b of cell.blocks) {
        cellTwips +=
          (blockFormat(b, heightScale).lineRawTwips ?? MIN_LINE_TWIPS) * Math.max(1, b.lines.length)
      }
      tallest = Math.max(tallest, cellTwips)
    }
    emitTwips += tallest
  }
  if (emitTwips <= boxTwips + Math.max(TABLE_SQUEEZE_TOL_TWIPS, boxTwips * 0.02)) return 1
  return Math.max(boxTwips / emitTwips, MIN_HEIGHT_SCALE)
}

/** IR table → docx-engine TableModel → self-contained w:tbl fragment */
function tableToSave(
  block: TableBlock,
  heightScale = 1,
  indentPt = 0,
  shavePt = 0,
  floatPos?: NonNullable<TableModel['floatPos']>,
): SaveBlock {
  // set-solid squeeze (P24 B) composes with whole-page compression
  const rowScale = heightScale * tableSqueezeFactor(block, heightScale)
  const model: TableModel = {
    cellMarTwips: CELL_MAR_TWIPS,
    // canvas pages (P19): the table floats at its measured page position
    ...(floatPos !== undefined ? { floatPos } : {}),
    // measured left offset (P17): an inset table (boxed form on a designed
    // page) otherwise snaps to the margin and the whole panel shifts left
    ...(indentPt >= 2 ? { indentTwips: Math.round(indentPt * PT_TO_TWIPS) } : {}),
    // confidence is set by the borderless detectors only; lattice tables
    // (real drawn borders) keep the default single-line borders unless the
    // grid was drawn in a non-black ink
    ...(block.confidence !== undefined
      ? {
          borders:
            // rule-separated zone (P22 A): the drawn separator between the
            // side-by-side stacks survives as the inside-vertical border
            block.sepRule !== undefined
              ? { ...NO_BORDERS, insideV: { style: block.sepRule, szEighths: 4 } }
              : NO_BORDERS,
        }
      : block.borderColor !== undefined
        ? { borders: colorBorders(block.borderColor) }
        : {}),
    rows: block.rows.map((row) => {
      let startCol = 0
      return row.map((cell) => {
        // width preflight (P21 A): a cell paragraph wraps at the cell's grid
        // width minus margins — one substituted-font wrap grows the row and a
        // dense full-page table spills (cellWidth needs the SPANNED columns)
        const cellWidthPt = block.colWidthsPt
          .slice(startCol, startCol + Math.max(1, cell.gridSpan))
          .reduce((a, b) => a + b, 0)
        startCol += Math.max(1, cell.gridSpan)
        const cellAvailPt = cellWidthPt - (CELL_MAR_TWIPS.left + CELL_MAR_TWIPS.right) / PT_TO_TWIPS
        for (const cellBlock of cell.blocks) {
          // zone cells rebuild exact line geometry — no wrap slack (P22 A)
          preflightFitBlock(cellBlock, cellAvailPt, undefined, {
            strict: block.sepRule !== undefined,
          })
        }
        const richParas = cellParagraphs(cell, block.form === true, rowScale)
        // lattice cells: infer horizontal alignment against the real cell box —
        // the block analyser's frame is the text's own extent and a lone short
        // line carries no alignment information without the cell around it
        if (block.confidence === undefined) {
          richParas.forEach((para, i) => {
            const cellBlock = cell.blocks[i]
            if (para.align === undefined && cellBlock !== undefined) {
              const align = detectCellHAlign(cell.box, cellBlock.lines)
              if (align) para.align = align
            }
          })
        }
        const tc: TableCell = {
          paras:
            richParas.length > 0 ? richParas.map((p) => p.runs.map((r) => r.text).join('')) : [''],
        }
        // EMPTY cells still hold one paragraph — give it a near-zero explicit
        // height, or the docDefaults (1.15× line + w:after) grow every empty
        // row far past its measured trHeight (form templates are full of them)
        tc.richParas =
          richParas.length > 0
            ? richParas
            : [{ spaceAfter: 0, lineRule: 'exact', lineRawTwips: TIGHT_LINE_TWIPS, runs: [] }]
        if (cell.gridSpan > 1) tc.colSpan = cell.gridSpan
        if (cell.vMerge) tc.vMerge = cell.vMerge
        if (cell.fill) tc.fill = cell.fill
        if (cell.vAlign) tc.vAlign = cell.vAlign
        // split-run cells (P27): the vertical rule between them was never
        // drawn in the source — suppress exactly those fabricated edges so
        // the lattice render stays faithful
        if (block.confidence === undefined && cell.softEdges !== undefined) {
          tc.borders = {
            ...(cell.softEdges.left ? { left: NO_BORDER } : {}),
            ...(cell.softEdges.right ? { right: NO_BORDER } : {}),
            ...(cell.softEdges.top ? { top: NO_BORDER } : {}),
            ...(cell.softEdges.bottom ? { bottom: NO_BORDER } : {}),
          }
        }
        return tc
      })
    }),
    colWidthsTwips: block.colWidthsPt.map((w) => Math.max(1, Math.round(w * PT_TO_TWIPS))),
    // measured heights keep the docx table's extent close to the PDF's
    // (w:trHeight hRule=atLeast — content may still grow a row)
    rowHeightsTwips: block.rows.map((row) => {
      const h = rowHeightTwips(row)
      if (h === null) return null
      // bordered rows render trHeight + the horizontal border width in
      // LibreOffice — deduct it or a 35-row form grows ~18pt (P16 H)
      const borderEat = block.confidence === undefined ? BORDER_EAT_TWIPS : 0
      // near-full-page shave (P17): spread across rows; padded (height-bound)
      // rows give the hair back, content-bound rows are atLeast and keep it
      const rowShave = Math.round((shavePt * PT_TO_TWIPS) / Math.max(1, block.rows.length))
      return Math.max(1, Math.round(h * rowScale) - borderEat - rowShave)
    }),
  }
  return { kind: 'xml', xml: generateTableModelXml(model) }
}

function imageAlign(block: ImageBlock, geo: PageGeometry, pageWidthPt: number): NewImage['align'] {
  const bodyLeft = geo.marginLeftPt
  const bodyRight = pageWidthPt - geo.marginRightPt
  const tol = Math.max(6, geo.contentWidthPt * 0.02)
  if (approxEq(rectCenterX(block.box), (bodyLeft + bodyRight) / 2, 2 * tol)) return 'center'
  if (block.box.x1 >= bodyRight - tol && block.box.x0 > bodyLeft + tol) return 'right'
  return 'left'
}

/**
 * Explicit spacing for a picture's holder paragraph: 'atLeast' with a tiny
 * floor pins the line to the picture's own height without the docDefaults
 * 1.15× multiplier (exact would clip the picture), and w:after stays 0.
 */
function imageParaSpacing(spacingBeforePt: number): NewImage['paraSpacing'] {
  return {
    ...(spacingBeforePt > 0 ? { beforeTwips: Math.round(spacingBeforePt * PT_TO_TWIPS) } : {}),
    afterTwips: 0,
    lineTwips: TIGHT_LINE_TWIPS,
    lineRule: 'atLeast',
  }
}

function imageToSave(
  block: ImageBlock,
  geo: PageGeometry,
  pageWidthPt: number,
  spacingBeforePt = 0,
  heightScale = 1,
): SaveBlock {
  const boxW = Math.max(1, rectWidth(block.box))
  const boxH = Math.max(1, rectHeight(block.box))
  // whole-page compression (P9 C) shrinks pictures with the text around them
  const displayW = Math.min(boxW, geo.contentWidthPt) * heightScale
  const displayH = boxH * (displayW / boxW)
  const image: NewImage = {
    base64: bytesToBase64(block.data),
    mime: block.mime,
    widthPx: Math.max(1, Math.round(displayW * PT_TO_PX)),
    heightPx: Math.max(1, Math.round(displayH * PT_TO_PX)),
    align: imageAlign(block, geo, pageWidthPt),
    paraSpacing: imageParaSpacing(spacingBeforePt),
  }
  return { kind: 'image', image }
}

const EMU_PER_PT = 12700

/**
 * floating image → anchored picture (wp:anchor) pinned to its measured page
 * position. Page-relative (P9 C): a paragraph-relative offset put tall floats
 * wherever their holder paragraph happened to land — a page-height photo
 * anchored mid-flow then reached past the page bottom and pushed every
 * wrapped paragraph after it onto a fresh page.
 */
/** square-wrapped floats at/above this share of the page height wrap 'behind' */
const SQUARE_FLOAT_MAX_HEIGHT_RATIO = 0.6

/**
 * The wrap mode actually written for a float: text cannot meaningfully run
 * beside a near-page-tall image — square wrap there only forces the following
 * paragraphs off the page (LibreOffice spills everything under a float that
 * reaches the bottom text edge), so it demotes to 'behind'.
 */
function floatWrapOf(block: ImageBlock, page: IrPage): FloatPlacement['wrap'] {
  const float = block.float!
  // landscape pages are slides (P11 D): PowerPoint has no text wrap — every
  // "text beside image" there is absolute positioning, and a square band
  // reserves phantom column space that spills the tail of the page
  if (float.wrap !== 'behind' && page.widthPt > page.heightPt) return 'behind'
  return float.wrap !== 'behind' &&
    Math.max(1, rectHeight(block.box)) >= page.heightPt * SQUARE_FLOAT_MAX_HEIGHT_RATIO
    ? 'behind'
    : float.wrap
}

function floatImageToSave(block: ImageBlock, page: IrPage, zOrder?: number): SaveBlock {
  const boxW = Math.max(1, rectWidth(block.box))
  const boxH = Math.max(1, rectHeight(block.box))
  const wrap = floatWrapOf(block, page)
  const image: NewImage = {
    base64: bytesToBase64(block.data),
    mime: block.mime,
    widthPx: Math.max(1, Math.round(boxW * PT_TO_PX)),
    heightPx: Math.max(1, Math.round(boxH * PT_TO_PX)),
    wrap,
    posOffsetEmu: {
      x: Math.round(block.box.x0 * EMU_PER_PT),
      y: Math.round((page.heightPt - block.box.y1) * EMU_PER_PT),
      relativeTo: 'page',
    },
    // stacked behind anchors keep their source paint order (P16 A)
    ...(zOrder !== undefined ? { zOrder } : {}),
    // the anchor's empty holder paragraph must not take flow space
    paraSpacing: { afterTwips: 0, lineTwips: TIGHT_LINE_TWIPS, lineRule: 'exact' },
  }
  return { kind: 'image', image }
}

/**
 * Full-page background render (P9 B) → behindDoc float pinned to the page box
 * (positionH/V relativeFrom="page", offset 0). Word/LibreOffice render it
 * under the page's text, restoring gradient/wallpaper backgrounds the flat
 * w:background color cannot carry.
 */
function bgRenderToSave(render: PageRender, page: IrPage): SaveBlock {
  const image: NewImage = {
    base64: bytesToBase64(render.data),
    mime: render.mime,
    widthPx: Math.max(1, Math.round(page.widthPt * PT_TO_PX)),
    heightPx: Math.max(1, Math.round(page.heightPt * PT_TO_PX)),
    wrap: 'behind',
    posOffsetEmu: { x: 0, y: 0, relativeTo: 'page' },
    // the anchor's empty holder paragraph must not take flow space
    paraSpacing: { afterTwips: 0, lineTwips: TIGHT_LINE_TWIPS, lineRule: 'exact' },
  }
  return { kind: 'image', image }
}

/**
 * An empty utility paragraph: near-zero height (exact 1pt line, no after)
 * unless it doubles as a spacer carrying measured whitespace of `heightPt`.
 */
function spacerParagraph(heightPt: number, pageBreakBefore: boolean): SaveBlock {
  return {
    kind: 'generated',
    block: {
      type: 'paragraph',
      runs: [],
      format: {
        ...(pageBreakBefore ? { pageBreakBefore: true } : {}),
        spaceAfter: 0,
        lineRule: 'exact',
        lineRawTwips: Math.max(TIGHT_LINE_TWIPS, Math.round(heightPt * PT_TO_TWIPS)),
      },
    },
  }
}

/** an empty paragraph carrying the page break (used when a page starts with an image) */
const pageBreakParagraph = (leadPt = 0): SaveBlock => spacerParagraph(leadPt, true)

const emptyParagraph = (spacingPt = 0): SaveBlock => spacerParagraph(spacingPt, false)

/**
 * Explicit column break (P10 A). Word/LibreOffice BALANCE a continuous-ended
 * multi-column section: tail blocks of a tall first column get redistributed
 * into (often narrower) neighbor columns, re-wrap, and push the rest of the
 * page out. The break pins the measured column split — which is also the
 * model the page budget already assumes (tallest column pays the bill).
 */
const columnBreakParagraph = (): SaveBlock => ({
  kind: 'xml',
  xml:
    `<w:p><w:pPr><w:spacing w:after="0" w:line="${TIGHT_LINE_TWIPS}" w:lineRule="exact"/></w:pPr>` +
    '<w:r><w:br w:type="column"/></w:r></w:p>',
})

// ── multi-column sections (P3) ──

/** merge sections whose column geometry differs by less than this (twips, 5pt) */
const SIG_TWIPS_TOL = 100
/** column widths within this ratio of each other are "equal" (plain w:num) */
const EQUAL_WIDTH_RATIO = 1.05
/** fallback column gap when the IR carries no gutter measurement (pt) */
const DEFAULT_GUTTER_PT = 18
/** page-top leading under this many points is noise — omit */
const MIN_LEAD_EMIT_PT = 2
/**
 * a pinned column's first block keeps its offset from the section top when
 * it is at least this many points (P15 B: a photo caption starting a stats
 * row's height below the section top); near-balanced column heads stay flush
 */
const COLUMN_LEAD_MIN_PT = 24

/** what one docx section looks like: enough to compare and to emit w:cols + pgSz */
interface SectionSignature {
  columns: number
  /** twips, LAYOUT order (left → right); present only for unequal columns */
  colWidths?: number[]
  /** column gap, twips */
  spaceTwips: number
  /** draw a rule between columns (w:cols w:sep, P14 C) */
  sep?: boolean
  /** RTL section (w:bidi): Word fills columns right → left */
  bidi: boolean
  /**
   * page size (twips): AI-generated documents mix page sizes freely (each
   * page its own height); a size change must open a new section with its own
   * w:pgSz or every over-tall page overflows into extra pages
   */
  pageWidthTwips: number
  pageHeightTwips: number
}

const singleColumnSig = (page: IrPage): SectionSignature => ({
  columns: 1,
  spaceTwips: 0,
  bidi: false,
  pageWidthTwips: Math.round(page.widthPt * PT_TO_TWIPS),
  pageHeightTwips: Math.round(page.heightPt * PT_TO_TWIPS),
})

/**
 * Output width of each READING-ORDER column (pt): the same content-width
 * scaling signatureOf writes into w:cols — the width a flow paragraph in
 * that column really gets (the preflight measures lines against it).
 */
const columnWidthsMemo = new WeakMap<PageSection, number[]>()
function columnWidthsPt(section: PageSection, geo: PageGeometry): number[] {
  let widths = columnWidthsMemo.get(section)
  if (widths) return widths
  const n = section.columns.length
  if (n <= 1) {
    widths = [geo.contentWidthPt]
  } else {
    const raw = section.columns.map((c) => Math.max(rectWidth(c.box), 1))
    const gutters =
      section.gutterWidthsPt.length === n - 1
        ? section.gutterWidthsPt
        : Array.from({ length: n - 1 }, () => DEFAULT_GUTTER_PT)
    const total = raw.reduce((a, b) => a + b, 0) + gutters.reduce((a, b) => a + b, 0)
    const scale = geo.contentWidthPt / Math.max(total, 1)
    widths = raw.map((w) => w * scale)
  }
  columnWidthsMemo.set(section, widths)
  return widths
}

/** measured column geometry → docx signature, scaled into the content width */
function signatureOf(section: PageSection, geo: PageGeometry, page: IrPage): SectionSignature {
  const n = section.columns.length
  if (n <= 1) return singleColumnSig(page)
  // IR columns are in reading order; docx wants layout order
  const layoutCols = section.dir === 'rtl' ? [...section.columns].reverse() : section.columns
  const widths = layoutCols.map((c) => Math.max(rectWidth(c.box), 1))
  const gutters =
    section.gutterWidthsPt.length === n - 1
      ? section.gutterWidthsPt
      : Array.from({ length: n - 1 }, () => DEFAULT_GUTTER_PT)
  const total = widths.reduce((a, b) => a + b, 0) + gutters.reduce((a, b) => a + b, 0)
  const scale = geo.contentWidthPt / Math.max(total, 1)
  const meanGutter = gutters.reduce((a, b) => a + b, 0) / gutters.length
  const sig: SectionSignature = {
    columns: n,
    spaceTwips: Math.max(0, Math.round(meanGutter * scale * PT_TO_TWIPS)),
    bidi: section.dir === 'rtl',
    pageWidthTwips: Math.round(page.widthPt * PT_TO_TWIPS),
    pageHeightTwips: Math.round(page.heightPt * PT_TO_TWIPS),
  }
  if (section.colSep) sig.sep = true
  if (Math.max(...widths) > EQUAL_WIDTH_RATIO * Math.min(...widths)) {
    sig.colWidths = widths.map((w) => Math.max(1, Math.round(w * scale * PT_TO_TWIPS)))
  }
  return sig
}

function sameSignature(a: SectionSignature, b: SectionSignature): boolean {
  if (a.columns !== b.columns || a.bidi !== b.bidi) return false
  if ((a.sep ?? false) !== (b.sep ?? false)) return false
  if (
    Math.abs(a.pageWidthTwips - b.pageWidthTwips) > SIG_TWIPS_TOL ||
    Math.abs(a.pageHeightTwips - b.pageHeightTwips) > SIG_TWIPS_TOL
  ) {
    return false
  }
  if (a.columns === 1) return true
  if ((a.colWidths === undefined) !== (b.colWidths === undefined)) return false
  if (Math.abs(a.spaceTwips - b.spaceTwips) > SIG_TWIPS_TOL) return false
  if (a.colWidths && b.colWidths) {
    return a.colWidths.every((w, i) => Math.abs(w - b.colWidths![i]!) <= SIG_TWIPS_TOL)
  }
  return true
}

function colsXml(sig: SectionSignature): string {
  if (sig.columns <= 1) return ''
  const sep = sig.sep ? ' w:sep="1"' : ''
  if (sig.colWidths) {
    const children = sig.colWidths
      .map((w, i) =>
        i < sig.colWidths!.length - 1
          ? `<w:col w:w="${w}" w:space="${sig.spaceTwips}"/>`
          : `<w:col w:w="${w}"/>`,
      )
      .join('')
    return `<w:cols w:num="${sig.columns}" w:space="${sig.spaceTwips}"${sep} w:equalWidth="0">${children}</w:cols>`
  }
  return `<w:cols w:num="${sig.columns}" w:space="${sig.spaceTwips}"${sep}/>`
}

/**
 * A section-break paragraph. In OOXML the embedded sectPr CLOSES the section
 * whose content precedes it, and its w:type says how that section started
 * relative to the one before (continuous = same page, column change in place).
 */
function sectionBreakParagraph(
  sig: SectionSignature,
  geo: PageGeometry,
  startType: 'nextPage' | 'continuous',
  titlePg = false,
): SaveBlock {
  const s = geo.section
  const orient = sig.pageWidthTwips > sig.pageHeightTwips ? ' w:orient="landscape"' : ''
  const xml =
    // near-zero spacing: the break paragraph itself must not take flow space
    `<w:p><w:pPr><w:spacing w:after="0" w:line="${TIGHT_LINE_TWIPS}" w:lineRule="exact"/><w:sectPr>` +
    (startType === 'continuous' ? '<w:type w:val="continuous"/>' : '') +
    `<w:pgSz w:w="${sig.pageWidthTwips}" w:h="${sig.pageHeightTwips}"${orient}/>` +
    `<w:pgMar w:top="${s.marginTop}" w:right="${s.marginRight}"` +
    ` w:bottom="${s.marginBottom}" w:left="${s.marginLeft}"` +
    ` w:header="${s.headerDist ?? 708}" w:footer="${s.footerDist ?? 708}" w:gutter="0"/>` +
    colsXml(sig) +
    (titlePg ? '<w:titlePg/>' : '') +
    (sig.bidi ? '<w:bidi/>' : '') +
    '</w:sectPr></w:pPr></w:p>'
  return { kind: 'xml', xml }
}

// ── re-emitted page furniture (P17) ──
// Dropped repeated headers/footers/page numbers come back as REAL docx
// header/footer parts: one line per furniture slot, PAGE fields where the
// running number sat, positioned via pgMar w:header/w:footer so the body
// geometry (and pagination) is untouched.

const HF_CJK_RE = /[⺀-鿿豈-﫿぀-ヿ]/
/** keep the header/footer line inside the margin: dist + this × font ≤ margin */
const HF_LINE_FACTOR = 1.4

interface FurnitureHfSave {
  header?: HeaderFooter
  footer?: HeaderFooter
  headerFirst?: HeaderFooter
  footerFirst?: HeaderFooter
  titlePg?: boolean
}

function hfParagraph(s: FurnitureHf, geo: PageGeometry, pageWidthPt: number): HfParagraph {
  const run: Run = { text: s.text }
  const half = Math.round(s.fontSizePt * 2)
  if (half > 0) run.sizeHalfPoints = half
  if (s.bold) run.bold = true
  if (s.italic) run.italic = true
  if (s.color && s.color !== '000000') run.color = s.color
  if (s.fontFamily) {
    if (HF_CJK_RE.test(s.text)) run.font = s.fontFamily
    else run.fontAscii = s.fontFamily
  }
  // exact line height: template docDefaults (line=276 auto + after=120)
  // would grow a small-print header far past its source band and push the
  // body down — pagination must not move when furniture is re-emitted
  const para: HfParagraph = {
    runs: [run],
    spaceAfter: 0,
    lineRule: 'exact',
    lineRawTwips: Math.max(TIGHT_LINE_TWIPS, Math.round(s.fontSizePt * LINE_LEADING * PT_TO_TWIPS)),
  }
  const center = (s.x0 + s.x1) / 2
  if (Math.abs(center - pageWidthPt / 2) <= pageWidthPt * 0.06) para.align = 'center'
  else if (s.x1 >= pageWidthPt - geo.marginRightPt - 8) para.align = 'right'
  else {
    const indent = Math.round((s.x0 - geo.marginLeftPt) * PT_TO_TWIPS)
    if (indent > 20) para.indentLeft = indent
  }
  return para
}

/** furniture slots → header/footer SaveOptions + pgMar distances (twips, set on geo.section) */
function furnitureHfToSave(slots: readonly FurnitureHf[], geo: PageGeometry): FurnitureHfSave {
  const pageWidthPt = geo.section.pageWidth / PT_TO_TWIPS
  const bandOf = (band: 'top' | 'bottom') =>
    slots
      .filter((s) => s.band === band)
      .sort((a, b) => (band === 'top' ? a.edgeDistPt - b.edgeDistPt : b.edgeDistPt - a.edgeDistPt))
  const hfOf = (band: readonly FurnitureHf[]): HeaderFooter | undefined =>
    band.length > 0
      ? {
          text: band.map((s) => s.text).join('\n'),
          paras: band.map((s) => hfParagraph(s, geo, pageWidthPt)),
        }
      : undefined
  const tops = bandOf('top')
  const bottoms = bandOf('bottom')
  const out: FurnitureHfSave = {}
  const header = hfOf(tops)
  const footer = hfOf(bottoms)
  // the whole band's rendered height (its exact lines stacked) must stay
  // inside the margin, or the header/footer pushes the body and pagination
  const bandHeightTwips = (band: readonly FurnitureHf[]) =>
    band.reduce((t, s) => t + Math.ceil(s.fontSizePt * HF_LINE_FACTOR * PT_TO_TWIPS), 0)
  if (header) {
    out.header = header
    // the topmost line's TOP from the page top, clamped as above
    const distPt = Math.min(...tops.map((s) => Math.max(0, s.edgeDistPt - s.fontSizePt)))
    const capTwips = Math.max(0, geo.section.marginTop - bandHeightTwips(tops))
    geo.section.headerDist = Math.min(Math.round(distPt * PT_TO_TWIPS), capTwips)
  }
  if (footer) {
    out.footer = footer
    const distPt = Math.min(...bottoms.map((s) => Math.max(0, s.edgeDistPt - s.fontSizePt)))
    const capTwips = Math.max(0, geo.section.marginBottom - bandHeightTwips(bottoms))
    geo.section.footerDist = Math.min(Math.round(distPt * PT_TO_TWIPS), capTwips)
  }
  // slots that skip the document's first page (cover pages) blank it via
  // w:titlePg; slots that DO cover it come back as explicit first-page parts
  if (slots.some((s) => !s.coversFirstPage)) {
    out.titlePg = true
    const firstTops = tops.filter((s) => s.coversFirstPage)
    const firstBottoms = bottoms.filter((s) => s.coversFirstPage)
    const headerFirst = hfOf(firstTops)
    const footerFirst = hfOf(firstBottoms)
    if (headerFirst) out.headerFirst = headerFirst
    if (footerFirst) out.footerFirst = footerFirst
  }
  return out
}

const isFloatImage = (b: PageBlock): b is ImageBlock => b.kind === 'image' && b.float !== undefined

// ── canvas pages (P19) ──
// High-confidence slide pages bypass the flow entirely: every text block
// becomes a page-anchored w:framePr container at its measured coordinates,
// tables float via w:tblpPr, images/backgrounds ride the existing behindDoc
// pins. No spacing chain, no page budget — the geometry IS the layout.

/**
 * Wrap headroom added to a frame's width: substituted fonts render a hair
 * wider, and one wrapped line costs a whole pitch where a few points of extra
 * width are invisible. The pad extends AWAY from the block's anchored edge
 * (left-aligned text keeps its left edge, right-aligned its right, centered
 * splits the pad) so the visible ink never drifts.
 */
const FRAME_WRAP_PAD_RATIO = 0.04
const FRAME_WRAP_PAD_MIN_PT = 4

// Container choice for canvas regions (P20 C comparison): page-anchored
// text boxes scored 0.9138 vs framePr 0.9124 on lo-tableSectionColumns —
// inside LO's render noise (±0.005). framePr keeps the P19-proven pipeline
// (editability, literal list markers, tblpPr interplay), so newsletter
// regions stay framePr; text boxes are reserved for FILLED card regions
// (P20 B), where framePr cannot express the plate fill + insets.
function canvasTextToSave(block: TextBlock, page: IrPage): SaveBlock {
  const format = blockFormat(block)
  const widthPt = Math.max(1, rectWidth(block.box))
  const padPt = Math.max(FRAME_WRAP_PAD_MIN_PT, widthPt * FRAME_WRAP_PAD_RATIO)
  let xPt = block.box.x0
  if (block.align === 'center') xPt -= padPt / 2
  else if (block.align === 'right') xPt -= padPt
  format.frame = {
    wTwips: Math.round((widthPt + padPt) * PT_TO_TWIPS),
    xTwips: Math.round(xPt * PT_TO_TWIPS),
    yTwips: Math.round((page.heightPt - block.box.y1) * PT_TO_TWIPS),
  }
  // decorative rules stay paragraph borders; no spacing chain to fund on a
  // canvas page, so the drawn gap is written as-is (Word clamps to 0–31pt)
  if (block.border) {
    const b = block.border
    format.borders = b.side === 'top' ? 't' : b.side === 'left' ? 'l' : 'b'
    format.borderStyle = {
      color: b.color,
      szEighths: Math.max(2, Math.round(b.widthPt * 8)),
      spacePt: b.spacePt,
    }
    if (block.lines.length === 0) format.lineRawTwips = TIGHT_LINE_TWIPS
  }
  const runs = paragraphRuns(block)
  // list items keep their literal marker on canvas pages (P20): w:numPr's
  // hanging indent eats into the fixed frame width, wraps the measured line
  // and collides the exact-pitch frames below — geometry IS the layout here,
  // so the marker rides as plain text at its drawn position instead
  if (block.list?.marker !== undefined && runs.length > 0) {
    const first = runs[0]!
    runs.unshift({ ...first, text: `${block.list.marker.trimEnd()} ` })
  }
  const generated: GeneratedBlock = { type: 'paragraph', runs }
  generated.format = format
  return { kind: 'generated', block: generated }
}

// ── card regions (P20) ──
// A backdrop plate + its text blocks rebuild as ONE paragraph-anchored text
// box (wrap topAndBottom): the plate is the box fill, the measured plate/text
// gaps are the box insets, and the whole card rides the flow together — the
// absolute-panel + flowed-text pairing misaligned the moment anything above
// reflowed (byte-deck FINAL TAKEAWAY). The box's occupied height IS the plate's
// measured height, so the page's spacing chain stays conserved.

/** rounded plates draw with this corner radius (pt, capped to a quarter height) */
const CARD_CORNER_RADIUS_PT = 6
/** anchored text boxes stack above every behindDoc pin (they carry the text) */
const CARD_Z_BASE = 2000
/** docPr id base for card boxes (patch-time images mint from 9000 up) */
const CARD_DOCPR_BASE = 7000

/** the card a text block opens, when this page rebuilds it as a text box */
const cardOf = (block: PageBlock, page: IrPage): CardRegion | undefined =>
  block.kind === 'text' && block.cardId !== undefined && !page.canvas
    ? page.cards?.[block.cardId]
    : undefined

/** holder spacing for a card: the chain measured to the first TEXT ink, and
 * the box top sits one top-inset above it */
const cardSpacingBeforePt = (first: TextBlock, card: CardRegion): number =>
  Math.max(0, (first.spacingBeforePt ?? 0) - Math.max(0, card.box.y1 - first.box.y1))

function cardToSave(
  card: CardRegion,
  cardId: number,
  members: TextBlock[],
  pageBreakBefore: boolean,
  spacingBeforePt: number,
  heightScale: number,
  page: IrPage,
): SaveBlock {
  const textUnion = rectUnionAll(members.map((m) => m.box))
  const insLPt = Math.max(0, textUnion.x0 - card.box.x0)
  const insTPt = Math.max(0, card.box.y1 - textUnion.y1) * heightScale
  // wrap headroom (P19 frame lesson): substituted fonts render a hair wider,
  // and inside a fixed box one wrapped line paints OUTSIDE the plate. The
  // right inset gives up the pad — the left edge keeps the measured anchor.
  const wrapPadPt = Math.max(FRAME_WRAP_PAD_MIN_PT, rectWidth(card.box) * FRAME_WRAP_PAD_RATIO)
  const insRPt = Math.max(0, card.box.x1 - textUnion.x1 - wrapPadPt)
  const insBPt = Math.max(0, textUnion.y0 - card.box.y0) * heightScale
  const contentLeftPt = card.box.x0 + insLPt
  const paragraphs: TextboxContentParagraph[] = members.map((m, i) => {
    const format = blockFormat(m, heightScale)
    // inter-member gaps stay the measured chain values — inside the box they
    // are interior rhythm, not page whitespace, so no budget scaling
    if (i > 0 && (m.spacingBeforePt ?? 0) > 0) {
      format.spaceBefore = Math.round((m.spacingBeforePt ?? 0) * heightScale * PT_TO_TWIPS)
    }
    const indentPt = m.box.x0 - contentLeftPt
    if (
      (format.align === undefined || format.align === 'left') &&
      m.dir !== 'rtl' &&
      indentPt >= 1
    ) {
      format.indentLeft = Math.round(indentPt * PT_TO_TWIPS)
    }
    const runs = paragraphRuns(m)
    // literal list markers, like canvas pages (P20): w:numPr's hanging indent
    // eats into the fixed box width and wraps the measured line
    if (m.list?.marker !== undefined && runs.length > 0) {
      runs.unshift({ ...runs[0]!, text: `${m.list.marker.trimEnd()} ` })
    }
    return { runs, format }
  })
  // page anchor at the plate's MEASURED coordinates (like the behindDoc panel
  // it replaces — the plate's position is gold), wrap topAndBottom so the
  // flow can never run through the band. The flow around it still pays the
  // plate's height via the caller's budget bookkeeping, so a reflow above
  // shifts the neighbors, never the card.
  return {
    kind: 'xml',
    xml: buildAnchoredTextboxParagraphXml({
      anchor: 'page',
      wrap: 'topAndBottom',
      xEmu: Math.round(card.box.x0 * EMU_PER_PT),
      yEmu: Math.round((page.heightPt - card.box.y1) * EMU_PER_PT),
      widthEmu: Math.round(rectWidth(card.box) * EMU_PER_PT),
      heightEmu: Math.round(rectHeight(card.box) * heightScale * EMU_PER_PT),
      fillHex: card.color,
      insetsEmu: {
        l: Math.round(insLPt * EMU_PER_PT),
        t: Math.round(insTPt * EMU_PER_PT),
        r: Math.round(insRPt * EMU_PER_PT),
        b: Math.round(insBPt * EMU_PER_PT),
      },
      ...(card.rounded
        ? {
            cornerRadiusEmu: Math.round(
              Math.min(CARD_CORNER_RADIUS_PT, rectHeight(card.box) / 4) * EMU_PER_PT,
            ),
          }
        : {}),
      zOrder: CARD_Z_BASE + cardId,
      id: CARD_DOCPR_BASE + page.index * 8 + cardId,
      paragraphs,
      holderLineTwips: TIGHT_LINE_TWIPS,
      ...(spacingBeforePt > 0
        ? { holderSpacingBeforeTwips: Math.round(spacingBeforePt * PT_TO_TWIPS) }
        : {}),
      ...(pageBreakBefore ? { holderPageBreakBefore: true } : {}),
    }),
  }
}

// ── page vertical budget (P8) ──
// The before_space chain is a faithful record of the source whitespace, but
// Word never truncates w:before at a page boundary: a page-bottom block whose
// chain gap is near the page height (decorative footers under sparse pages)
// gets pushed WHOLE onto the next page, doubling the page count. Emitted
// spacing is therefore clamped so no block's spacing can push the block
// itself out of its source page's usable height.

/**
 * Flow height (pt) one block occupies in the rebuilt page, mirroring what the
 * save functions write: text = lines × exact line height (incl. the P8 B
 * font-leading floor), inline image = display height, table = sum of the
 * measured row heights. Floats are outside the flow and never call this.
 */
/** display-size text only risks an extra wrapped line when its widest line
 * nearly fills its column — a short title never wraps under substitution */
const WRAP_RISK_LINE_FILL = 0.85

function wrapsAtRisk(b: TextBlock, colWidthPt: number): boolean {
  if (maxFontSizePt(b) < WRAP_RISK_FONT_PT) return false
  const widest = Math.max(...b.lines.map((l) => l.box.x1 - l.box.x0), 0)
  return widest >= WRAP_RISK_LINE_FILL * Math.max(colWidthPt, 1)
}

/**
 * Body-size lines wrap under substituted metrics too (P18 A): a source line
 * that fills its column wall-to-wall (justified body text, a caption+sentence
 * sharing an exact-pitch paragraph, an unbreakable formula token) re-wraps in
 * LibreOffice the moment its glyph advances drift a hair wider, and with
 * exact line pitch every gained line costs a full pitch — near-full pages
 * spill their tail block (a lone caption / two formula lines) onto a ghost
 * page. Display sizes are wrapsAtRisk's business; this rule is deliberately
 * tighter (edge fill, not 0.85) because justified text hits it by
 * construction, and the charge is capped per page and only ever squeezes
 * spacing (never triggers whole-page height compression).
 */
/** a line whose right end sits within this share of the content width from
 * the document's right text edge is wrap-bait; measured against the PAGE
 * text column (geo), not the ink-hugging section box, which every widest
 * line fills by construction */
const BODY_WRAP_RISK_EDGE_TOL = 0.015
/** …and it must be a real line, not a short right-aligned scrap */
const BODY_WRAP_RISK_MIN_FILL = 0.5
/** at most ~2 gained lines of headroom per page — beyond that the page was
 * genuinely overfull and the height chain (squeeze/compression) owns it */
const BODY_WRAP_RISK_CAP_PT = 40
/** skipping the floor-surplus deduction is pagination-safe only when the
 * page's spacing slack covers every unpaid surplus with room to spare */
const SURPLUS_SKIP_SLACK_PT = 96
const CJK_RE = /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF\u3000-\u303F]/u
/**
 * Only wrap-BRITTLE content gets the body charge: CJK wraps per character,
 * so any advance drift moves the last glyph down a line. Latin prose —
 * justified or ragged — absorbs drift in its word spaces, and re-squeezing
 * those pages just lifts their ink (P17's dmp_he lesson; a long-token rule
 * for URLs/formulas was tried and re-hit exactly that sample).
 */
function lineWrapBrittle(l: Line): boolean {
  const text = l.spans.map((s) => s.text).join('')
  let cjk = 0
  for (const ch of text) if (CJK_RE.test(ch)) cjk++
  const glyphs = [...text].filter((c) => c.trim() !== '').length
  return glyphs > 0 && cjk / glyphs >= 0.5
}
function bodyWrapsAtRisk(b: TextBlock, geo: PageGeometry): boolean {
  if (maxFontSizePt(b) >= WRAP_RISK_FONT_PT) return false
  // a long paragraph re-flows as a whole and swallows the drift; only a
  // SHORT block (a caption+sentence pair, a two-line formula) turns one
  // wrapped glyph into a whole gained line (charging every justified line
  // of a dense report re-squeezed pages LibreOffice renders true)
  if (b.lines.length > 3) return false
  const rightEdge = geo.marginLeftPt + geo.contentWidthPt
  const tol = BODY_WRAP_RISK_EDGE_TOL * geo.contentWidthPt
  return b.lines.some(
    (l) =>
      rightEdge - l.box.x1 <= tol &&
      l.box.x1 - l.box.x0 >= BODY_WRAP_RISK_MIN_FILL * geo.contentWidthPt &&
      lineWrapBrittle(l),
  )
}

/**
 * Micro-section chrome (P18 A): a signature ruling / label+line row becomes a
 * SHORT multi-column continuous section (label column beside a bar column).
 * LibreOffice renders each such transition taller than the measured model —
 * its column balancing rounds the section and its break paragraphs up — and
 * a form page holding three signature rows creeps ~70pt and spills its last
 * line (form-gov's footer) onto a ghost page. Reserve a fixed charge per
 * short multi-column section in the spacing budget. Tall multi-column
 * sections (newsletter/deck columns) are excluded: their balancing cost is
 * proportionally invisible and deck pages ride a knife edge.
 */
const MICRO_SECTION_MAX_PT = 48
const MICRO_SECTION_CHROME_PT = 16
const MICRO_SECTION_CHROME_CAP_PT = 48
function isMicroSection(s: PageSection, geo: PageGeometry): boolean {
  if (s.columns.length !== 2) return false
  return s.columns.every((c) => {
    let h = 0
    for (const b of c.blocks) if (!isFloatImage(b)) h += flowBlockHeightPt(b, geo)
    return h <= MICRO_SECTION_MAX_PT
  })
}

function flowBlockHeightPt(block: PageBlock, geo: PageGeometry, heightScale = 1): number {
  if (block.kind === 'text') {
    // standalone decorative rules shrink their carrier line to TIGHT
    const lineTwips =
      block.border && block.lines.length === 0
        ? TIGHT_LINE_TWIPS
        : (blockFormat(block, heightScale).lineRawTwips ?? MIN_LINE_TWIPS)
    return (Math.max(1, block.lines.length) * lineTwips) / PT_TO_TWIPS
  }
  if (block.kind === 'table') {
    return (
      (block.rows.reduce((sum, row) => sum + (rowHeightTwips(row) ?? 0), 0) / PT_TO_TWIPS) *
      heightScale
    )
  }
  const boxW = Math.max(1, rectWidth(block.box))
  const displayW = Math.min(boxW, geo.contentWidthPt) * heightScale
  return Math.max(1, rectHeight(block.box)) * (displayW / boxW)
}

/**
 * How much taller a text block RENDERS than its measured ink box (P17).
 * Single-line blocks floor their exact line to fontSize × LINE_LEADING —
 * needed so ascenders/descenders stay unclipped — but the spacing chain
 * measured the gap ink-to-ink. Left unpaid, every floored line grows the
 * page by (floor − ink): forms whose pages are full of short one-line
 * paragraphs (answer rulings, radio options) creep 30–80pt per page and
 * spill extra pages. The surplus is deducted from the block's OWN
 * spacing-before, so emitted pitch (spacing + exact line) matches source.
 *
 * Body text only: display-size slide type has surpluses of 20pt+ per line,
 * and deducting them re-choreographs whole flush-full slides (deck.pdf's
 * knife-edge pages) for no pagination win — overflow accumulation is a
 * body-text-repetition disease (forms, reports), not a slide one.
 */
const SURPLUS_FONT_MAX_PT = 14
/** a table this share of the page's usable height is a knife-edge full-pager
 * (0.93: a spilling DEFINITIONS page measures 0.97, while lo-tSC's layout
 * table at 0.91 must keep its rows — the shave shifted its whole panel) */
const FULL_PAGE_TABLE_RATIO = 0.93
/** total row-height shave for a full-page table (spread across its rows) */
const FULL_PAGE_TABLE_SHAVE_PT = 12
function lineSurplusPt(block: PageBlock): number {
  if (block.kind !== 'text' || block.lines.length === 0) return 0
  if (maxFontSizePt(block) > SURPLUS_FONT_MAX_PT) return 0
  const lineTwips = blockFormat(block).lineRawTwips ?? MIN_LINE_TWIPS
  return Math.max(0, (lineTwips * block.lines.length) / PT_TO_TWIPS - rectHeight(block.box))
}

/** pages built by hand (tests) may lack sections — wrap their flat blocks */
const sectionsFromBlocks = (blocks: PageBlock[]): PageSection[] =>
  blocks.length === 0
    ? []
    : [
        {
          box: rectUnionAll(blocks.map((b) => b.box)),
          columns: [{ box: rectUnionAll(blocks.map((b) => b.box)), blocks }],
          gutterWidthsPt: [],
          dir: 'ltr',
        },
      ]

export function pagesToSaveBlocks(
  pages: IrPage[],
  furnitureHf: readonly FurnitureHf[] = [],
): {
  blocks: SaveBlock[]
  section: SectionSettings
  /** present only when section breaks were emitted (multi-section docs) */
  sectionStartType?: 'nextPage' | 'continuous'
  /** numbering definitions the detected lists need (docx-engine SaveOptions) */
  numbering?: SaveOptions['numbering']
  /** detected footnotes (word/footnotes.xml content), in reference order */
  footnotes?: NoteInfo[]
  /** re-emitted page furniture (P17): header/footer parts + first-page flags */
  hf?: FurnitureHfSave
} {
  const geo = computeGeometry(pages)
  // header/footer re-emission first: it sets pgMar header/footer distances on
  // geo.section, which every emitted sectPr copies
  const hf = furnitureHf.length > 0 ? furnitureHfToSave(furnitureHf, geo) : undefined
  // multi-section docs put w:titlePg on the FIRST section's sectPr (the
  // SaveOptions flag only reaches the trailing one); consumed on first use
  let titlePgPending = hf?.titlePg === true
  const numbering = new ListNumbering()
  const blocks: SaveBlock[] = []
  // OOXML forbids table-after-table and a table as the last body element
  // without a paragraph between/after — track and separate
  let lastWasTable = false
  let needBreak: boolean
  // the currently open docx section; a signature change closes it with a
  // section-break paragraph carrying ITS properties and start type
  let curSig: SectionSignature | null = null
  let curStart: 'nextPage' | 'continuous' = 'nextPage'
  let sectionBreaks = 0
  // canvas pages (P19) live in per-page sections: LO misplaces page-anchored
  // frames across in-section page breaks (probe: frames land on the wrong
  // page with y resolved against the unrotated page box), so a canvas page
  // must both OPEN and CLOSE its own section even when signatures match
  let forceClose = false

  const openSection = (sig: SectionSignature, atPageStart: boolean): void => {
    if (curSig === null) {
      curSig = sig
      forceClose = false
      return
    }
    // a multi-column section never continues across a page boundary (P14 A):
    // when consecutive pages open with same-signature column layouts, the
    // page break degrades to a COLUMN jump inside the still-open section
    // (Word/LibreOffice semantics) and the next page's content pours into
    // the previous page's columns. Close and reopen with a hard nextPage
    // start instead; single-column runs keep flowing (pageBreakBefore works).
    if (!forceClose && sameSignature(curSig, sig) && !(atPageStart && sig.columns > 1)) return
    forceClose = false
    blocks.push(sectionBreakParagraph(curSig, geo, curStart, titlePgPending))
    titlePgPending = false
    sectionBreaks++
    lastWasTable = false
    curStart = atPageStart ? 'nextPage' : 'continuous'
    if (atPageStart) needBreak = false // the nextPage section break IS the page break
    curSig = sig
  }

  for (const page of pages) {
    needBreak = page.index > 0
    const pageStartBlockCount = blocks.length

    if ((page.scanned || page.degraded) && page.render) {
      openSection(singleColumnSig(page), needBreak)
      if (needBreak || lastWasTable)
        blocks.push(needBreak ? pageBreakParagraph() : emptyParagraph())
      lastWasTable = false
      // full-page behindDoc pin (P11 D): as an inline picture the render is
      // taller than the usable flow band (page height minus margins), and
      // LibreOffice pushes it out — leaving a blank page — before drawing it
      blocks.push(bgRenderToSave(page.render, page))
      continue
    }

    // ── canvas page (P19): absolute containers, per-page section, no budget ──
    if (page.canvas) {
      forceClose = true
      openSection(singleColumnSig(page), needBreak)
      if (page.bgRender) {
        if (needBreak) blocks.push(pageBreakParagraph())
        blocks.push(bgRenderToSave(page.bgRender, page))
        lastWasTable = false
        needBreak = false
      }
      const pinned = [
        ...(page.bgPanels ?? []),
        ...(page.decorImages ?? []),
        ...page.blocks.filter(isFloatImage),
      ]
        .map((block, order) => ({ block, order }))
        .sort((a, b) => (a.block.z ?? 0) - (b.block.z ?? 0) || a.order - b.order)
      for (const [rank, { block: pin }] of pinned.entries()) {
        if (needBreak) blocks.push(pageBreakParagraph())
        blocks.push(floatImageToSave(pin, page, rank + 1))
        lastWasTable = false
        needBreak = false
      }
      for (const block of page.blocks) {
        // isFloatImage would narrow ImageBlock out of the union entirely
        if (block.kind === 'image' && block.float !== undefined) continue
        if (block.kind === 'text') {
          if (block.lines.length === 0 && !block.border) continue
          if (needBreak) blocks.push(pageBreakParagraph())
          blocks.push(canvasTextToSave(block, page))
          lastWasTable = false
        } else if (block.kind === 'table') {
          if (needBreak) blocks.push(pageBreakParagraph())
          else if (lastWasTable) blocks.push(emptyParagraph())
          blocks.push(
            tableToSave(block, 1, 0, 0, {
              xTwips: Math.round(block.box.x0 * PT_TO_TWIPS),
              yTwips: Math.round((page.heightPt - block.box.y1) * PT_TO_TWIPS),
            }),
          )
          lastWasTable = true
        } else {
          // stray inline image (analyze pins all images on landscape pages;
          // this is belt-and-braces) — pin it at its measured position
          if (needBreak) blocks.push(pageBreakParagraph())
          blocks.push(
            floatImageToSave(
              { ...block, float: { wrap: 'behind', xOffsetPt: block.box.x0 } },
              page,
            ),
          )
          lastWasTable = false
        }
        needBreak = false
      }
      if (needBreak) {
        blocks.push(pageBreakParagraph())
        lastWasTable = false
      }
      // the canvas section must not swallow the next page's content
      forceClose = true
      continue
    }

    const sections = page.sections?.length ? page.sections : sectionsFromBlocks(page.blocks)
    // page-top leading: whitespace between the content top (below the margin)
    // and the page's first flow block joins the before_space chain here, where
    // the final margins are known
    let leadPt = 0
    const firstFlow = sections[0]?.columns[0]?.blocks.find((b) => !isFloatImage(b))
    if (firstFlow) {
      leadPt = Math.max(0, page.heightPt - geo.marginTopPt - firstFlow.box.y1)
      if (leadPt < MIN_LEAD_EMIT_PT) leadPt = 0
    }

    // vertical budget (P8): emitted spacing may not push a block out of its
    // source page. Sections stack (continuous breaks), columns within one
    // section run side by side — the tallest column pays the section's bill.
    // The budget stops BUDGET_SLACK above the (already slack-reduced) bottom
    // margin: a clamped page fills flush, and that headroom must stay free
    // to absorb rendering growth (substituted font metrics wrapping extra
    // lines, atLeast table rows) — filling into it re-spills every clamped
    // page.
    // footnotes lifted into word/footnotes.xml re-materialize at THIS page's
    // bottom when Word lays it out (P12 D): their area comes out of the flow
    // budget, or bottom-anchored content (a footnote-region hairline rule
    // claiming spacing down to the margin) spills an entirely blank page
    const footnoteAreaPt = (page.footnotes ?? []).reduce(
      (sum, f) => sum + f.blocks.reduce((t, b) => t + rectHeight(b.box), 0),
      0,
    )
    const usablePt = Math.max(
      0,
      page.heightPt -
        geo.marginTopPt -
        geo.marginBottomPt -
        BUDGET_SLACK_PT -
        (footnoteAreaPt > 0 ? footnoteAreaPt + FOOTNOTE_SEPARATOR_PT : 0),
    )
    // pre-pass: when the page's whitespace does not fit next to its block
    // heights (document margins are page-1-wide, not per-page), every gap on
    // the page shrinks by the same factor — proportional scaling keeps the
    // page's rhythm, where clamping alone would keep early gaps at full size
    // and zero out the tail
    let heightsPt = 0
    let wantTotalPt = 0
    // anchor holder paragraphs (bgRender / panels / floats) each keep a 1pt
    // exact line in the flow — on a flush-full slide with several floats they
    // are the hair that spills a blank page (P11 D), so they are budgeted
    heightsPt +=
      ((page.bgRender ? 1 : 0) +
        // card plates leave the pin list (P20): their text box pays instead
        (page.bgPanels?.filter((p) => p.cardId === undefined).length ?? 0) +
        sections.reduce(
          (n, s) => n + s.columns.reduce((m, c) => m + c.blocks.filter(isFloatImage).length, 0),
          0,
        )) *
      (TIGHT_LINE_TWIPS / PT_TO_TWIPS)
    // wrap-growth reserve: display-size multi-line titles gain a line under a
    // substituted font; one pitch per such block stays budgeted (not emitted)
    let wrapRiskPt = 0
    // edge-filling body lines (P18 A): charged only against the spacing
    // budget below — single-column flow only (a multi-column page pays its
    // height per column and its pinning has its own headroom), and NEVER
    // into heightsPt, where it would trip whole-page height compression on
    // pages LibreOffice renders fine
    let bodyWrapRiskPt = 0
    const prePassCards = new Set<number>()
    // bottom inset of an open card (P20): the next block's measured gap
    // starts at the last member's INK, but flow resumes at the plate edge.
    // Page-level — card members may span a section boundary.
    let cardPadPt = 0
    for (const s of sections) {
      let tallest = 0
      for (const c of s.columns) {
        let col = 0
        for (const b of c.blocks) {
          if (isFloatImage(b)) continue
          // card group (P20): the whole group charges the plate's height once
          const card = cardOf(b, page)
          if (card !== undefined) {
            const tb = b as TextBlock
            if (!prePassCards.has(tb.cardId!)) {
              prePassCards.add(tb.cardId!)
              col += rectHeight(card.box)
              wantTotalPt += Math.max(0, cardSpacingBeforePt(tb, card) - cardPadPt)
              const members = page.blocks.filter(
                (x): x is TextBlock => x.kind === 'text' && x.cardId === tb.cardId,
              )
              cardPadPt = Math.max(0, Math.min(...members.map((m) => m.box.y0)) - card.box.y0)
            }
            continue
          }
          const h = flowBlockHeightPt(b, geo)
          col += h
          wantTotalPt += Math.max(0, (b.spacingBeforePt ?? 0) - cardPadPt)
          cardPadPt = 0
          if (b.kind === 'text' && b.lines.length >= 2 && wrapsAtRisk(b, rectWidth(c.box))) {
            wrapRiskPt += h / b.lines.length
          } else if (
            b.kind === 'text' &&
            b.lines.length >= 2 &&
            s.columns.length === 1 &&
            bodyWrapsAtRisk(b, geo)
          ) {
            bodyWrapRiskPt += h / Math.max(1, b.lines.length)
          }
        }
        tallest = Math.max(tallest, col)
      }
      heightsPt += tallest
    }
    heightsPt += wrapRiskPt
    bodyWrapRiskPt = Math.min(bodyWrapRiskPt, BODY_WRAP_RISK_CAP_PT)
    // one lone micro-row (a table's date header) measured fine in practice —
    // the creep that spills pages needs a run of them (signature rows), and
    // table pages are excluded outright: their micro rows are table header
    // captions whose pages render true to measure (charging them re-squeezed
    // two settled tax-form samples), and page-filling tables have their own
    // shave chain (P17 G). Unlike the wrap charge, this cost is
    // unconditional LibreOffice behaviour, so it is NOT gated on a
    // pre-existing squeeze.
    const pageHasTable = sections.some((s) =>
      s.columns.some((c) => c.blocks.some((b) => b.kind === 'table')),
    )
    const microSections = pageHasTable ? 0 : sections.filter((s) => isMicroSection(s, geo)).length
    const sectionChromePt =
      microSections >= 2
        ? Math.min(microSections * MICRO_SECTION_CHROME_PT, MICRO_SECTION_CHROME_CAP_PT)
        : 0
    // whole-page compression (P9 C): when the block heights ALONE exceed the
    // budget (full-bleed slides: title + a table as tall as the page), no
    // spacing squeeze can save the page — Word/LibreOffice will spill its
    // tail rows onto extra pages. Every pitch/row/picture on the page shrinks
    // by one proportional factor instead so the page stays whole. Below the
    // floor the compression is unreadable and the page spills anyway.
    const heightScale =
      heightsPt > usablePt
        ? Math.max((usablePt - COMPRESS_SAFETY_PT) / heightsPt, MIN_HEIGHT_SCALE)
        : 1
    const scaledHeightsPt = heightsPt * heightScale
    // measured whitespace may consume half the budget slack (P16 C): the
    // slack guards wrap growth in BLOCK heights, and crushing every gap to a
    // fraction of its measured size lifts the whole page's rhythm upward — a
    // bigger visual lie than the rare spilled tail line the slack prevents
    const spacingBudgetPt = usablePt + BUDGET_SLACK_PT / 2
    // the page-top lead anchors every block below it — fund it in full
    // before the inter-block gaps share what remains (P16 F): scaling the
    // lead like a gap shifted whole crushed pages upward
    let leadFundPt = Math.min(leadPt, Math.max(0, spacingBudgetPt - scaledHeightsPt))
    // the wrap charge only deepens an EXISTING squeeze (P18): a page whose
    // measured content already fits keeps its rhythm — charging it flips a
    // fitting page into a squeeze for a wrap that usually never happens
    // (the research-report sample -0.010); a page already over budget is the creep-overflow
    // candidate the charge exists for
    if (
      wantTotalPt <= 0 ||
      spacingBudgetPt - scaledHeightsPt - sectionChromePt - leadFundPt >= wantTotalPt
    ) {
      bodyWrapRiskPt = 0
    }
    const spacingScale =
      wantTotalPt > 0
        ? Math.max(
            0,
            Math.min(
              1,
              (spacingBudgetPt - scaledHeightsPt - bodyWrapRiskPt - sectionChromePt - leadFundPt) /
                wantTotalPt,
            ),
          )
        : 1
    // the floor-surplus deduction (P17) exists to stop creep-overflow; on a
    // page whose slack absorbs every surplus even unpaid, it has no
    // pagination benefit and only lifts the ink off its measured rhythm
    // (P17's poi-heading123/dmp_he drops). Pay it only when needed.
    const pageSurplusPt = sections.reduce(
      (sum, s) =>
        s.columns.length === 1
          ? sum +
            s.columns[0]!.blocks.reduce((t, b) => t + (b.kind === 'text' ? lineSurplusPt(b) : 0), 0)
          : sum,
      0,
    )
    const paySurplus =
      spacingBudgetPt -
        scaledHeightsPt -
        bodyWrapRiskPt -
        sectionChromePt -
        leadFundPt -
        wantTotalPt <
      pageSurplusPt + SURPLUS_SKIP_SLACK_PT
    let pageConsumedPt = 0

    // P9 B: the page's rendered background pins behind the text, anchored on
    // an empty holder paragraph at the page start (like floats, the break
    // paragraph moves in front of it so the anchor lands on THIS page)
    if (sections.length === 0 && page.bgRender) {
      openSection(singleColumnSig(page), needBreak)
      if (needBreak) blocks.push(pageBreakParagraph())
      blocks.push(bgRenderToSave(page.bgRender, page))
      lastWasTable = false
      needBreak = false
    }

    // floats are page-positioned (P9 C): their anchors all move to the page
    // start (P10 A). An anchor paragraph mid-flow inside a multi-column
    // section makes LibreOffice push the rest of the section onto a fresh
    // page; at the page start it is outside every column's content.
    const pageFloats = sections.flatMap((s) =>
      s.columns.flatMap((c) => c.blocks.filter(isFloatImage)),
    )
    // square-wrapped floats eat column band space: pinning the measured column
    // split with explicit breaks then overflows the page. Left alone,
    // Word/LibreOffice re-balance the columns around the floats — so only
    // pages whose floats all sit behind the text get their splits pinned.
    // A pinned column must also FIT with wrap headroom (P11 D): substituted
    // fonts wrap one extra line per block, and a pinned overflow cascades the
    // whole tail of the section onto a fresh page — near the budget the
    // renderer's own balancing recovers, explicit breaks cannot.
    const fitPassCards = new Set<number>()
    const columnsFit = sections.every((s) =>
      s.columns.every((c) => {
        let colPt = 0
        for (const b of c.blocks) {
          if (isFloatImage(b)) continue
          const card = cardOf(b, page)
          if (card !== undefined) {
            const tb = b as TextBlock
            if (!fitPassCards.has(tb.cardId!)) {
              fitPassCards.add(tb.cardId!)
              colPt +=
                rectHeight(card.box) * heightScale + cardSpacingBeforePt(tb, card) * spacingScale
            }
            continue
          }
          const h = flowBlockHeightPt(b, geo, heightScale)
          colPt += h + (b.spacingBeforePt ?? 0) * spacingScale
          if (b.kind === 'text' && b.lines.length >= 1 && wrapsAtRisk(b, rectWidth(c.box))) {
            colPt += h / b.lines.length
          }
        }
        // half the budget slack backs the pin, like the whitespace rebate
        // (P16 C/E): a table-anchored column a hair over the raw gate is
        // still far from LO's real usable height, and losing the pin lets
        // the balancer pour the neighbor column's text under the table
        return colPt <= usablePt + BUDGET_SLACK_PT / 2 - COMPRESS_SAFETY_PT
      }),
    )
    const pinColumnSplits = columnsFit && pageFloats.every((f) => floatWrapOf(f, page) === 'behind')

    // card groups already emitted on this page (P20); the pending pad is
    // page-level because members may span a section boundary
    const emittedCards = new Set<number>()
    let cardPadEmitPt = 0
    for (const [si, section] of sections.entries()) {
      openSection(signatureOf(section, geo, page), si === 0 && needBreak)
      if (si === 0) {
        if (page.bgRender) {
          if (needBreak) blocks.push(pageBreakParagraph())
          blocks.push(bgRenderToSave(page.bgRender, page))
          lastWasTable = false
          needBreak = false
        }
        // panels and floats stack by source paint order (P16 A): behindDoc
        // anchors tie on relativeHeight otherwise, and a full-page wallpaper
        // drawn first would paint OVER card panels drawn later, hiding the
        // light text that flows on them. Outside the flow: no spacing, no
        // budget; the page-top lead still belongs to the first flow block.
        // card plates skip the pin (P20): their anchored text box paints them
        const pinned = [
          ...(page.bgPanels ?? []).filter((p) => p.cardId === undefined),
          ...pageFloats,
        ]
          .map((block, order) => ({ block, order }))
          .sort((a, b) => (a.block.z ?? 0) - (b.block.z ?? 0) || a.order - b.order)
        for (const [rank, { block: pin }] of pinned.entries()) {
          if (needBreak) blocks.push(pageBreakParagraph())
          blocks.push(floatImageToSave(pin, page, rank + 1))
          lastWasTable = false
          needBreak = false
        }
      }
      let sectionTallestPt = 0
      for (const [ci, column] of section.columns.entries()) {
        if (ci > 0 && pinColumnSplits) {
          if (needBreak) {
            blocks.push(pageBreakParagraph())
            needBreak = false
          }
          blocks.push(columnBreakParagraph())
          lastWasTable = false
        }
        let colConsumedPt = 0
        let atColumnTop = ci > 0 && pinColumnSplits
        for (const block of column.blocks) {
          if (isFloatImage(block)) continue // anchored at the page start above
          // card group (P20): first member emits the whole group as one
          // paragraph-anchored text box; the rest are already inside it
          const card = cardOf(block, page)
          if (card !== undefined) {
            const tb = block as TextBlock
            if (emittedCards.has(tb.cardId!)) continue
            emittedCards.add(tb.cardId!)
            const members = page.blocks.filter(
              (x): x is TextBlock => x.kind === 'text' && x.cardId === tb.cardId,
            )
            const cardHeightPt = rectHeight(card.box) * heightScale
            const sectionLeadC = section.box.y1 - card.box.y1
            const columnLeadC =
              atColumnTop && sectionLeadC >= COLUMN_LEAD_MIN_PT ? sectionLeadC * heightScale : 0
            atColumnTop = false
            const measuredC = Math.max(0, cardSpacingBeforePt(tb, card) - cardPadEmitPt)
            const wantC = measuredC * spacingScale + leadFundPt + columnLeadC
            const spacingC = Math.max(
              0,
              Math.min(wantC, spacingBudgetPt - pageConsumedPt - colConsumedPt - cardHeightPt),
            )
            blocks.push(
              cardToSave(card, tb.cardId!, members, needBreak, spacingC, heightScale, page),
            )
            cardPadEmitPt = Math.max(0, Math.min(...members.map((m) => m.box.y0)) - card.box.y0)
            lastWasTable = false
            needBreak = false
            leadFundPt = 0
            colConsumedPt += spacingC + cardHeightPt
            continue
          }
          const heightPt = flowBlockHeightPt(block, geo, heightScale)
          // column-top offset (P15 B): a pinned column whose content starts
          // well below the section top (a photo caption beside a stats row)
          // keeps that offset — flush-started it rides up over the float
          // above. Parallel to the tallest column, so outside the page
          // spacing budget (unscaled); geometry bounds it to the section.
          const sectionLead = section.box.y1 - block.box.y1
          const columnLeadPt =
            atColumnTop && sectionLead >= COLUMN_LEAD_MIN_PT ? sectionLead * heightScale : 0
          atColumnTop = false
          // the font-floor surplus is settled only at emission (the budget
          // pre-pass keeps the full measured gaps, so squeeze decisions are
          // exactly what they were), and it does NOT stack with the page
          // squeeze: the binding cut wins. min(scaled, source-corrected)
          // never exceeds the baseline squeeze (knife-edge slides can only
          // shrink) and never sinks below the source-matching pitch — where
          // both apply the block is reduced by max(squeeze, surplus), not
          // their sum (stacking over-compressed nearly-full form pages)
          // single-column sections only: creep-overflow is a stacked-flow
          // disease (forms, reports); a multi-column newsletter page pays its
          // height per column and the deduction just shifted its panels up
          const measuredPt = Math.max(0, (block.spacingBeforePt ?? 0) - cardPadEmitPt)
          cardPadEmitPt = 0
          const surplusPt = paySurplus && section.columns.length === 1 ? lineSurplusPt(block) : 0
          const wantPt =
            Math.min(measuredPt * spacingScale, Math.max(0, measuredPt - surplusPt)) +
            leadFundPt +
            columnLeadPt
          const spacingPt = Math.max(
            0,
            Math.min(wantPt, spacingBudgetPt - pageConsumedPt - colConsumedPt - heightPt),
          )
          // indents measure from the column's text edge (single-column: the
          // page margin; extra columns: the column box the splitter built)
          const colBasePt = section.columns.length > 1 && ci > 0 ? column.box.x0 : geo.marginLeftPt
          if (block.kind === 'text') {
            const leader = leaderStyleOf(block)
            // width preflight (P21 A): flow paragraphs get the output column's
            // width — leaders end on a measured tab, TOC entries on a leader tab
            if (!block.tocEntry && leader === null) {
              // verse stacks (P22 E): every break is author-pinned, so a line
              // that re-wraps drops its tail (often a lone em-dash) onto its
              // own row — no wrap slack for them
              preflightFitBlock(
                block,
                columnWidthsPt(section, geo)[ci] ?? geo.contentWidthPt,
                undefined,
                {
                  strict: block.lines.some((l) => l.hardBreakBefore),
                },
              )
            }
            blocks.push(
              block.tocEntry
                ? tocToSave(
                    block,
                    colBasePt,
                    columnWidthsPt(section, geo)[ci] ?? geo.contentWidthPt,
                    needBreak,
                    spacingPt,
                    heightScale,
                  )
                : leader !== null
                  ? leaderToSave(block, leader, colBasePt, needBreak, spacingPt, heightScale)
                  : textBlockToSave(
                      block,
                      needBreak,
                      spacingPt,
                      page.index,
                      numbering,
                      heightScale,
                    ),
            )
            lastWasTable = false
          } else if (block.kind === 'table') {
            // w:tbl carries no spacing of its own — the utility paragraph
            // doubles as the spacer for the block's measured whitespace
            if (needBreak) blocks.push(pageBreakParagraph(spacingPt))
            else if (lastWasTable || spacingPt >= MIN_LEAD_EMIT_PT) {
              blocks.push(emptyParagraph(spacingPt))
            }
            const tableW = block.colWidthsPt.reduce((a, b) => a + b, 0)
            const tableIndentPt = Math.min(
              Math.max(0, block.box.x0 - colBasePt),
              Math.max(0, geo.contentWidthPt - tableW),
            )
            // a page-filling table rides the knife edge: one wrapped cell
            // line (substituted font) pushes its bottom border past the real
            // page bottom and a near-blank sliver page appears before the
            // pinned break (P17). Shave a hair off the row heights — padded
            // rows give it back, content-bound atLeast rows keep their size.
            const tableShavePt =
              heightPt >= FULL_PAGE_TABLE_RATIO * usablePt ? FULL_PAGE_TABLE_SHAVE_PT : 0
            blocks.push(tableToSave(block, heightScale, tableIndentPt, tableShavePt))
            lastWasTable = true
          } else {
            let imageSpacingPt = spacingPt
            if (needBreak) {
              blocks.push(pageBreakParagraph(imageSpacingPt))
              imageSpacingPt = 0
            }
            blocks.push(imageToSave(block, geo, page.widthPt, imageSpacingPt, heightScale))
            lastWasTable = false
          }
          needBreak = false
          leadFundPt = 0
          colConsumedPt += spacingPt + heightPt
        }
        sectionTallestPt = Math.max(sectionTallestPt, colConsumedPt)
      }
      pageConsumedPt += sectionTallestPt
    }
    // an entirely empty page still occupies one page in the docx
    if (needBreak) {
      blocks.push(pageBreakParagraph())
      lastWasTable = false
    } else if (blocks.length === pageStartBlockCount) {
      // blank document-FIRST page (only page 0 can get here: every other
      // emission path clears needBreak by pushing a block): Word and
      // LibreOffice ignore w:pageBreakBefore on the document's first
      // paragraph, so the NEXT page's break paragraph cannot stand in for
      // this page — without its own zero-height placeholder the blank page
      // collapses and the whole document loses one page (pdfbox2
      // AcrobatMerge-*/pagelabels: N blank pages round-tripped to N-1)
      openSection(singleColumnSig(page), false)
      blocks.push(emptyParagraph())
    }
  }
  if (lastWasTable) blocks.push(emptyParagraph())

  // the last open section's properties land in the trailing body sectPr
  const finalSig: SectionSignature =
    curSig ?? singleColumnSig(pages[0] ?? ({ widthPt: 612, heightPt: 792 } as IrPage))
  const section: SectionSettings = {
    ...geo.section,
    pageWidth: finalSig.pageWidthTwips,
    pageHeight: finalSig.pageHeightTwips,
    orientation: finalSig.pageWidthTwips > finalSig.pageHeightTwips ? 'landscape' : 'portrait',
  }
  if (finalSig.columns > 1) {
    section.columns = finalSig.columns
    section.colSpace = finalSig.spaceTwips
    if (finalSig.colWidths) section.colWidths = finalSig.colWidths
    if (finalSig.bidi) section.bidi = true
  }
  // footnotes lifted by the analysis layer → word/footnotes.xml entries.
  // Rich runs carry the measured size/font (P17): emitted as plain text the
  // note re-renders at the template's body size, ~40% taller than its source
  // area, and the host page's tail spills onto an extra page.
  const noteRunOf = (r: Run): NoteRun => ({
    text: r.text,
    ...(r.bold ? { bold: true } : {}),
    ...(r.italic ? { italic: true } : {}),
    ...(r.underline ? { underline: true } : {}),
    ...(r.strike ? { strike: true } : {}),
    ...(r.color ? { color: r.color } : {}),
    ...(r.sizeHalfPoints ? { sizeHalfPoints: r.sizeHalfPoints } : {}),
    ...(r.fontAscii ? { fontAscii: r.fontAscii } : {}),
    ...(r.font ? { font: r.font } : {}),
  })
  const footnotes: NoteInfo[] = pages.flatMap((page) =>
    (page.footnotes ?? []).map((f) => ({
      id: f.id,
      text: f.blocks
        .map((b) =>
          paragraphRuns(b)
            .map((r) => r.text)
            .join(''),
        )
        .join('\n'),
      richParas: f.blocks.map((b) =>
        paragraphRuns(b)
          .filter((r) => !r.noteRef && r.text !== '')
          .map(noteRunOf),
      ),
    })),
  )

  const numberingOptions = numbering.toSaveOptions()
  // titlePg not consumed by a mid-body sectPr (single-section doc) → the
  // trailing sectPr carries it via the SaveOptions flag
  const hfOut = hf ? { ...hf, titlePg: titlePgPending ? true : undefined } : undefined
  return {
    blocks,
    section,
    ...(hfOut !== undefined ? { hf: hfOut } : {}),
    ...(sectionBreaks > 0 ? { sectionStartType: curStart } : {}),
    ...(numberingOptions !== undefined ? { numbering: numberingOptions } : {}),
    ...(footnotes.length > 0 ? { footnotes } : {}),
  }
}

/** max per-channel difference for two page washes to count as the same tone */
const BG_CLUSTER_TOL = 18

function maxChannelDist(a: string, b: string): number {
  let d = 0
  for (let i = 0; i < 6; i += 2) {
    d = Math.max(d, Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)))
  }
  return d
}

/**
 * Dominant page background across the document. Word supports one background
 * per document (w:background), so a tone only wins when at least half of the
 * flow pages carry it — a lone colored cover page must not tint everything.
 * Near-identical washes (design systems shift the paper tone a few RGB steps
 * per page) vote as one cluster; the most frequent member represents it.
 */
function dominantBgColor(pages: IrPage[]): string | undefined {
  const flowPages = pages.filter((p) => !p.scanned && !p.degraded)
  if (flowPages.length === 0) return undefined
  const votes = new Map<string, number>()
  for (const p of flowPages) {
    if (p.bgColor) votes.set(p.bgColor, (votes.get(p.bgColor) ?? 0) + 1)
  }
  let best: string | undefined
  let bestCluster = 0
  for (const [color, n] of votes) {
    let cluster = 0
    for (const [other, on] of votes) {
      if (maxChannelDist(color, other) <= BG_CLUSTER_TOL) cluster += on
    }
    const moreVotes = best === undefined || (votes.get(best) ?? 0) < n
    if (cluster > bestCluster || (cluster === bestCluster && moreVotes)) {
      best = color
      bestCluster = cluster
    }
  }
  return bestCluster * 2 >= flowPages.length ? best : undefined
}

export interface RebuildOptions {
  /** page furniture to re-emit as real docx headers/footers (P17) */
  furnitureHf?: readonly FurnitureHf[]
}

/** IR → brand-new .docx bytes (fully local). */
export async function rebuildDocx(pages: IrPage[], opts: RebuildOptions = {}): Promise<Uint8Array> {
  // Word clamps w:pgSz to 22in per dimension and CROPS full-bleed content on
  // larger pages (wide-format scans) — shrink full-page-image pages to fit;
  // their only content is the render, which scales with the page box
  for (const page of pages) {
    if (!(page.scanned || page.degraded) || !page.render) continue
    const s = Math.min(1, WORD_MAX_PAGE_PT / page.widthPt, WORD_MAX_PAGE_PT / page.heightPt)
    if (s < 1) {
      page.widthPt *= s
      page.heightPt *= s
    }
  }
  // uninstalled families → metric-compatible stand-ins (P21 A); the assembly
  // loop then preflights every line against its output column's real width
  applyOutputFontSubstitutions(pages, opts.furnitureHf ?? [])
  const { blocks, section, sectionStartType, numbering, footnotes, hf } = pagesToSaveBlocks(
    pages,
    opts.furnitureHf ?? [],
  )
  const parsed = await parseDocx(await buildBlankDocx())
  const pageColor = dominantBgColor(pages)
  return saveDocx(parsed, blocks, {
    section,
    ...(sectionStartType !== undefined ? { sectionStartType } : {}),
    ...(numbering !== undefined ? { numbering } : {}),
    ...(footnotes !== undefined ? { footnotes } : {}),
    ...(pageColor !== undefined ? { pageColor } : {}),
    ...(hf?.header !== undefined ? { header: hf.header } : {}),
    ...(hf?.footer !== undefined ? { footer: hf.footer } : {}),
    ...(hf?.headerFirst !== undefined ? { headerFirst: hf.headerFirst } : {}),
    ...(hf?.footerFirst !== undefined ? { footerFirst: hf.footerFirst } : {}),
    ...(hf?.titlePg === true ? { titlePg: true } : {}),
    ...(hf !== undefined ? { hfAllSections: true } : {}),
  })
}
