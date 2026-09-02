/**
 * High-fidelity pagination F2 — assertion runner (line-level page splitting + Word
 * pagination constraints + table row-level page breaking)
 *
 * F2 improvements (over F1):
 *   - computeSectionedSlicesF2: main line-level splitting algorithm (widowControl/keepLines/keepNext)
 *   - BlockBox carries lineBoxes (per-line offset+height) for line-level break decisions
 *   - Table row-level page breaking: tableRows accumulated row by row, cantSplit/tblHeader aware
 *   - Footnote placeholders: paragraphs with footnote refs reserve footnote height on their page
 *   - Page-start line text location: mid-paragraph/mid-table cut points report the exact line text
 *
 * How to run:
 *   npx vitest run tests/pagination-parity.test.ts
 *
 * CI policy: skipIf when no baseline; runnable in any environment.
 */

import { describe, expect, it, beforeAll } from 'vitest'
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { parseDocx } from '@genoffice/docx-engine'
import { readSections } from '@genoffice/docx-engine'
import {
  computeSectionedSlicesF2,
  sectionColGeom,
  sectionColumns,
  sectionGeoms,
  type BlockBox,
  type TableRowBox,
} from '../src/renderer/pagination'
import {
  computeLineMetrics,
  estimateFootnoteHeight,
  resolveNoteStyle,
} from '../src/renderer/line-metrics'
import { loBaselineMetrics } from './helpers/lo-fonts'
import type { ParsedDoc, DocGrid, ParaFormat, StyleDisplay } from '@genoffice/docx-engine'

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)
const CORPUS_DIR = join(__dirname, 'pagination-corpus/docx')
const BASELINE_FILE = join(__dirname, 'pagination-corpus/baseline-lo.json')
const BASELINE_WORD_FILE = join(__dirname, 'pagination-corpus/baseline-word.json')
const REPO_ROOT = join(__dirname, '../../..')
const REPORT_FILE = join(REPO_ROOT, '.task/pagination-parity-report.md')

// ─── Constants ─────────────────────────────────────────────────────────────

const TWIPS_TO_PX = 96 / 1440

/** Metrics instance: real LO-bundled font metrics when installed (aligned with the baseline layout engine), else heuristics */
const metrics = loBaselineMetrics()

// ─── Style chain helpers ────────────────────────────────────────────────────

/**
 * Resolve effective style attributes from parsed.styles and parsed.docDefaults (style chain merge).
 * Returns font size (pt), line-height rule, spacing, pagination constraints, etc. for paragraph rendering.
 */
interface EffectiveStyle {
  fontSizePt: number
  fontFamily: string
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined
  lineRawTwips: number | undefined
  spaceBeforeTwips: number
  spaceAfterTwips: number
  keepNext: boolean
  keepLines: boolean
  widowControl: boolean | undefined
  contextualSpacing: boolean
}

function resolveStyle(
  styleId: string | undefined,
  format: ParaFormat | undefined,
  parsed: ParsedDoc,
): EffectiveStyle {
  const defaults = parsed.docDefaults

  // Read base values from docDefaults
  const defaultFontSizePt = defaults?.sizeHalfPoints ? defaults.sizeHalfPoints / 2 : 12
  const defaultFamily = 'Calibri' // Word default font

  // Read display attributes from the style chain (basedOn already merged in parseStyles)
  let styleFontSizePt: number | undefined
  let styleFamily: string | undefined
  let styleLineRule: 'auto' | 'atLeast' | 'exact' | undefined
  let styleLineRawTwips: number | undefined
  let styleBefore: number | undefined
  let styleAfter: number | undefined
  let styleKeepNext = false
  let styleContextual = false

  if (styleId) {
    const styleInfo = parsed.styles.get(styleId)
    const display: StyleDisplay | undefined = styleInfo?.display
    if (display) {
      if (display.sizeHalfPoints) styleFontSizePt = display.sizeHalfPoints / 2
      if (display.font) styleFamily = display.font
      if (display.lineRule) styleLineRule = display.lineRule
      if (display.lineRawTwips) styleLineRawTwips = display.lineRawTwips
      if (display.spaceBeforeTwips !== undefined) styleBefore = display.spaceBeforeTwips
      if (display.spaceAfterTwips !== undefined) styleAfter = display.spaceAfterTwips
      if (display.keepNext) styleKeepNext = true
      if (display.contextualSpacing) styleContextual = true
    }
  }

  // Direct paragraph attributes override the style chain
  const lineRule = format?.lineRule ?? styleLineRule ?? defaults?.lineRule
  const lineRawTwips = format?.lineRawTwips ?? styleLineRawTwips ?? defaults?.lineRawTwips

  // Space before/after: paragraph > style > docDefaults > Word Normal default 160
  // Note: format?.spaceAfter can be 0 (explicitly set to 0) or undefined (not set)
  // Use a ?? chain: continue only on undefined; stop at 0 (keep the 0)
  const spaceBeforeTwips = format?.spaceBefore ?? styleBefore ?? defaults?.spaceBeforeTwips ?? 0
  const spaceAfterTwips = format?.spaceAfter ?? styleAfter ?? defaults?.spaceAfterTwips ?? 160

  const keepNext = !!(format?.keepNext ?? styleKeepNext)
  const keepLines = !!format?.keepLines
  // widowControl: on by default in Word; explicitly off when format?.widowControl === false
  const widowControl = format?.widowControl === false ? false : undefined
  const contextualSpacing = !!(format?.contextualSpacing ?? styleContextual)

  return {
    fontSizePt: styleFontSizePt ?? defaultFontSizePt,
    fontFamily: styleFamily ?? defaultFamily,
    lineRule,
    lineRawTwips,
    spaceBeforeTwips,
    spaceAfterTwips,
    keepNext,
    keepLines,
    widowControl,
    contextualSpacing,
  }
}

// ─── F2 block box construction (line-box data + pagination constraints) ────

/** Line-box data for a paragraph (incl. heading/listItem/passthrough/image) */
interface ParaLineData {
  lineBoxes: Array<{ offsetInBlock: number; height: number }>
  /** Per-line text aligned with lineBoxes (for page-start line location; none for image) */
  lineTexts?: string[]
  spaceBeforePx: number
  spaceAfterPx: number
  totalHeight: number
}

function computeParaLineData(
  block: {
    type: string
    runs?: Array<{
      text: string
      font?: string
      sizeHalfPoints?: number
      bold?: boolean
      italic?: boolean
    }>
    originalXml?: string | null
    imageHeightPx?: number
    level?: number
    styleId?: string
    format?: ParaFormat
    previewText?: string
  },
  contentWidthPx: number,
  docGrid: DocGrid | undefined,
  parsed: ParsedDoc,
): ParaLineData {
  // ── Image ─────────────────────────────────────────────────────────────────
  if (block.type === 'image') {
    const style = resolveStyle(block.styleId, block.format, parsed)
    const spaceAfterPx = style.spaceAfterTwips * TWIPS_TO_PX
    const totalH = (block.imageHeightPx ?? 100) + spaceAfterPx
    return {
      lineBoxes: [{ offsetInBlock: 0, height: totalH }],
      spaceBeforePx: 0,
      spaceAfterPx,
      totalHeight: totalH,
    }
  }

  // ── passthrough (TOC entry / field etc., estimated from preview text) ─────
  if (block.type === 'passthrough') {
    const text = block.previewText ?? ''
    const result = computeLineMetrics({
      runs: text.trim() ? [{ text }] : [],
      availWidthPx: contentWidthPx,
      docGrid,
      defaultFontSizePt: 12,
      metrics,
      isEmpty: !text.trim(),
    })
    // Wrapped anchored objects (wrapTopAndBottom) take vertical space not covered by the text estimate
    const wrapCy = anchorWrapHeight(block.originalXml ?? '')
    const totalHeight = Math.max(result.totalHeight, wrapCy)
    return {
      lineBoxes:
        wrapCy > result.totalHeight
          ? [{ offsetInBlock: 0, height: totalHeight }]
          : result.lineBoxes,
      lineTexts: result.lineTexts,
      spaceBeforePx: result.spaceBeforePx,
      spaceAfterPx: result.spaceAfterPx,
      totalHeight,
    }
  }

  // ── Text paragraphs (paragraph / heading / listItem) ──────────────────────
  const style = resolveStyle(block.styleId, block.format, parsed)

  // Heading adjusts font size by level (Word heading 1=16pt, 2=13pt, 3=12pt etc.)
  let fontSizePt = style.fontSizePt
  if (block.type === 'heading') {
    const lvl = block.level ?? 1
    const headingSize = [0, 16, 13, 12, 12, 11, 11, 11, 11, 11]
    fontSizePt = Math.max(fontSizePt, headingSize[Math.min(lvl, 9)] ?? 12)
  }

  // Metrics policy (precise Word baseline revisit 2026-07-27): keep measuring uniformly with the
  // style-chain font. Honoring the run-declared font (r.font ?? style.fontFamily) gets the
  // openclaw-report fixture to 10 pages (=Word), but mid-page cut points drift 8/10→4/10 and the
  // overall match rate drops 91.2%→82.4% — substitution error for missing fonts exceeds the
  // "uniform proxy" error, so it is a net loss.
  const runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }> = (block.runs ?? []).map((r) => ({
    text: r.text ?? '',
    fontFamily: r.font ?? style.fontFamily,
    sizeHalfPoints: r.sizeHalfPoints ?? fontSizePt * 2,
    bold: r.bold,
    italic: r.italic,
  }))

  const result = computeLineMetrics({
    runs,
    availWidthPx: contentWidthPx,
    lineRule: style.lineRule,
    lineRawTwips: style.lineRawTwips,
    spaceBefore: style.spaceBeforeTwips,
    spaceAfter: style.spaceAfterTwips,
    docGrid,
    defaultFontSizePt: fontSizePt,
    defaultFontFamily: style.fontFamily,
    metrics,
    isEmpty: runs.length === 0 || runs.every((r) => !r.text),
  })

  return {
    lineBoxes: result.lineBoxes,
    lineTexts: result.lineTexts,
    spaceBeforePx: result.spaceBeforePx,
    spaceAfterPx: result.spaceAfterPx,
    totalHeight: result.totalHeight,
  }
}

// ─── Table row-box construction ─────────────────────────────────────────────

/**
 * Parse table XML and extract each row's cantSplit / tblHeader flags (indexed like table.rows).
 */
function parseTableRowFlags(
  tableXml: string,
): Array<{ cantSplit: boolean; isHeader: boolean; heightTwips?: number; heightRule?: string }> {
  const result: Array<{
    cantSplit: boolean
    isHeader: boolean
    heightTwips?: number
    heightRule?: string
  }> = []
  const trRe = /<w:tr[\s>][\s\S]*?(?=<w:tr[\s>]|<\/w:tbl>)/g
  for (const m of tableXml.matchAll(trRe)) {
    const trChunk = m[0]
    const trPr = trChunk.match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? ''
    const hasCantSplit = /<w:cantSplit(?!\s+w:val="(?:0|false)")/.test(trPr)
    const hasTblHeader = /<w:tblHeader(?!\s+w:val="(?:0|false)")/.test(trPr)
    const trHeight = /<w:trHeight w:val="(\d+)"(?:[^>]*w:hRule="(\w+)")?/.exec(trPr)
    result.push({
      cantSplit: hasCantSplit,
      isHeader: hasTblHeader,
      ...(trHeight
        ? { heightTwips: parseInt(trHeight[1], 10), heightRule: trHeight[2] ?? 'atLeast' }
        : {}),
    })
  }
  return result
}

/** Vertical space (px) taken by wrapped anchored objects (wrapTopAndBottom/wrapSquare/wrapTight/wrapThrough) */
function anchorWrapHeight(xml: string): number {
  let sum = 0
  for (const m of xml.matchAll(/<wp:anchor[^>]*>[\s\S]*?<\/wp:anchor>/g)) {
    if (!/<wp:wrap(?:TopAndBottom|Square|Tight|Through)/.test(m[0])) continue
    const ext = /<wp:extent cx="\d+" cy="(\d+)"/.exec(m[0])
    if (ext) sum += (parseInt(ext[1], 10) / 914400) * 96
  }
  return sum
}

/** Apply w:trHeight: exact pins the row height, atLeast takes the max */
function applyTrHeight(
  computedH: number,
  flag?: { heightTwips?: number; heightRule?: string },
): number {
  if (!flag?.heightTwips) return computedH
  const trPx = flag.heightTwips * TWIPS_TO_PX
  return flag.heightRule === 'exact' ? trPx : Math.max(computedH, trPx)
}

/**
 * Compute the table's tableRowBoxes (for F2 row-level splitting) + per-row text (for page-start
 * line location). Cell line heights share the body-text source (measured on the real-Word
 * baseline: 18pt line + 8pt space-after = 26.4pt/row, no cell-specific factor; the former
 * CELL_CJK_FACTOR=1.44 was an LO tuning artifact, now removed).
 */
function computeTableRows(
  block: {
    type: string
    originalXml?: string | null
    table?: {
      rows: Array<
        Array<{
          paras: string[]
          richParas?: Array<{
            runs: Array<{
              text: string
              sizeHalfPoints?: number
              font?: string
              bold?: boolean
              italic?: boolean
            }>
          }>
          vMerge?: string
        }>
      >
    }
  },
  contentWidthPx: number,
  docGrid: DocGrid | undefined,
  parsed: ParsedDoc,
): { rows: TableRowBox[]; rowTexts: string[] } {
  const tableXml = block.originalXml ?? ''
  const rows = block.table?.rows
  const rowFlags = parseTableRowFlags(tableXml)
  const defaults = parsed.docDefaults

  if (!rows || rows.length === 0) {
    // Fallback: count w:tr in the XML; estimate row height as an empty cell
    const rowCount = (tableXml.match(/<w:tr[\s>]/g) ?? []).length || 1
    const cellH = computeLineMetrics({
      runs: [],
      availWidthPx: contentWidthPx / 2,
      lineRule: defaults?.lineRule,
      lineRawTwips: defaults?.lineRawTwips,
      spaceAfter: defaults?.spaceAfterTwips ?? 0,
      docGrid,
      defaultFontSizePt: defaults?.sizeHalfPoints ? defaults.sizeHalfPoints / 2 : 12,
      metrics,
      isEmpty: true,
    }).totalHeight
    return {
      rows: Array.from({ length: rowCount }, (_, i) => ({
        height: cellH,
        cantSplit: rowFlags[i]?.cantSplit ?? false,
        isHeader: rowFlags[i]?.isHeader ?? false,
      })),
      rowTexts: Array.from({ length: rowCount }, () => ''),
    }
  }

  const cellLineRule = defaults?.lineRule
  const cellLineRawTwips = defaults?.lineRawTwips
  const cellSpaceAfter = defaults?.spaceAfterTwips ?? 0
  const cellFontSizePt = defaults?.sizeHalfPoints ? defaults.sizeHalfPoints / 2 : 12
  // Amortize 1 horizontal border (0.5pt) per row: measured Word row pitch 26.4pt = 26pt content + border
  const ROW_BORDER_PX = 0.5 * (96 / 72)

  const rowTexts: string[] = []
  const rowBoxes = rows.map((row, ri) => {
    // Whole row vMerge=continue (vertical-merge continuation row): height counted in the start row
    const allContinue = row.every((c) => (c as { vMerge?: string }).vMerge === 'continue')
    if (allContinue) {
      rowTexts.push('')
      return {
        height: 0,
        vMergeContinue: true,
        cantSplit: rowFlags[ri]?.cantSplit ?? false,
        isHeader: rowFlags[ri]?.isHeader ?? false,
      }
    }

    let maxCellH = 0
    const cellTexts: string[] = []
    for (const cell of row) {
      if ((cell as { vMerge?: string }).vMerge === 'continue') continue
      const richParas =
        (
          cell as {
            richParas?: Array<{
              runs: Array<{
                text: string
                sizeHalfPoints?: number
                font?: string
                bold?: boolean
                italic?: boolean
              }>
            }>
          }
        ).richParas ?? []
      cellTexts.push(
        richParas.length > 0
          ? richParas.map((p) => p.runs.map((r) => r.text).join('')).join(' ')
          : cell.paras.join(' '),
      )
      let cellH = 0
      if (richParas.length > 0) {
        for (const para of richParas) {
          cellH += computeLineMetrics({
            runs: para.runs,
            availWidthPx: contentWidthPx / Math.max(1, row.length),
            lineRule: cellLineRule,
            lineRawTwips: cellLineRawTwips,
            spaceAfter: cellSpaceAfter,
            docGrid,
            defaultFontSizePt: cellFontSizePt,
            metrics,
            isEmpty: para.runs.length === 0,
          }).totalHeight
        }
      } else {
        for (const paraText of cell.paras) {
          cellH += computeLineMetrics({
            runs: paraText ? [{ text: paraText }] : [],
            availWidthPx: contentWidthPx / Math.max(1, row.length),
            lineRule: cellLineRule,
            lineRawTwips: cellLineRawTwips,
            spaceAfter: cellSpaceAfter,
            docGrid,
            defaultFontSizePt: cellFontSizePt,
            metrics,
            isEmpty: !paraText,
          }).totalHeight
        }
      }
      maxCellH = Math.max(maxCellH, cellH)
    }

    rowTexts.push(cellTexts.join(' ').trim())
    return {
      height: applyTrHeight(Math.max(maxCellH, 1), rowFlags[ri]) + ROW_BORDER_PX,
      cantSplit: rowFlags[ri]?.cantSplit ?? false,
      isHeader: rowFlags[ri]?.isHeader ?? false,
    }
  })
  return { rows: rowBoxes, rowTexts }
}

// ─── Footnote helpers ───────────────────────────────────────────────────────
// footnote heights come from the shared estimator (FootnoteText/Normal style
// chain + rich run sizes), same as the app's reservation path

// ─── Main logic ────────────────────────────────────────────────────────────

interface BaselineEntry {
  source: string
  version: string
  pages: number
  pageStarts: string[]
  pageEnds: string[]
  error?: string
}

type Baseline = Record<string, BaselineEntry>

interface ParityResult {
  stem: string
  loPages: number
  ourPages: number
  pageDiff: number
  loPageStarts: string[]
  ourPageStarts: string[]
  matchedPageStarts: number
  matchRate: number
  error?: string
}

async function runParity(stem: string, baseline: BaselineEntry): Promise<ParityResult> {
  const docxPath = join(CORPUS_DIR, `${stem}.docx`)
  if (!existsSync(docxPath)) {
    return {
      stem,
      loPages: baseline.pages,
      ourPages: 0,
      pageDiff: baseline.pages,
      loPageStarts: baseline.pageStarts,
      ourPageStarts: [],
      matchedPageStarts: 0,
      matchRate: 0,
      error: 'docx file does not exist',
    }
  }

  const bytes = readFileSync(docxPath)
  const parsed = await parseDocx(new Uint8Array(bytes))
  const sections = readSections(parsed)
  // Multi-column flow: text wraps at column width; the engine splits as a true column flow with
  // "page = columns × column height" (overflow moves to the next column, last column turns the
  // page, widow/orphan constraints at column boundaries). sectionGeoms outputs cols;
  // equalWidth="0" (local layout columns from PDF-converted docs) is treated as 1 column.
  const geoms = sectionGeoms(sections)

  // Per-section content width + docGrid (precomputed; column sections split width evenly by w:space)
  const sectionWidths = sections.map((s) => sectionColGeom(s).colWidthPx)
  const sectionDocGrids = sections.map((s) => s.settings.docGrid)

  // Footnote id → text
  const footnoteById = new Map<string, (typeof parsed.footnotes)[number]>()
  for (const fn of parsed.footnotes) footnoteById.set(fn.id, fn)

  // Build the F2 BlockBox list block by block (line boxes + constraints + table rows + footnote placeholders)
  const blocks: BlockBox[] = []
  const lineTextsOf = new Map<BlockBox, string[]>()
  const rowTextsOf = new Map<BlockBox, string[]>()
  let cursor = 0
  let curSectionIdx = 0

  for (const block of parsed.blocks) {
    // Skip hidden blocks that are pure sectPr
    if (block.hidden) continue
    if (
      block.type === 'passthrough' &&
      block.originalXml?.includes('<w:sectPr') &&
      !block.originalXml?.includes('<w:t>')
    ) {
      continue
    }

    // Determine the owning section
    const blockDocxIdx = block.docxIndex ?? 0
    let si = curSectionIdx
    for (let i = 0; i < sections.length; i++) {
      if (blockDocxIdx <= sections[i].lastBlockIndex) {
        si = i
        break
      }
    }
    curSectionIdx = si

    const contentWidthPx = sectionWidths[si] ?? (11906 - 3600) * TWIPS_TO_PX
    const docGrid = sectionDocGrids[si]

    // Forced break before the block
    const breakBefore =
      block.format?.pageBreakBefore ||
      (block.rawPPr?.includes('<w:pageBreakBefore/>') ?? false) ||
      (block.rawPPr?.includes('<w:pageBreakBefore ') ?? false)

    // In-block page/column breaks (column breaks only apply in modeled equal-width column
    // sections: equalWidth="0" is unmodeled and treated as 1 column — its column breaks switch
    // columns on the same page and must not count as page turns)
    const breakAfter =
      (block.originalXml?.includes('w:type="page"') ?? false) &&
      (block.originalXml?.includes('<w:br') ?? false)
    const colBreakAfter =
      (block.originalXml?.includes('w:type="column"') ?? false) &&
      (block.originalXml?.includes('<w:br') ?? false) &&
      sections[si] !== undefined &&
      sectionColumns(sections[si]) > 1

    // ── Table blocks ───────────────────────────────────────────────────────
    if (block.type === 'table') {
      const { rows: tableRows, rowTexts } = computeTableRows(
        block as Parameters<typeof computeTableRows>[0],
        contentWidthPx,
        docGrid,
        parsed,
      )
      const totalH = tableRows.reduce((s, r) => s + r.height, 0) + 4
      const box: BlockBox = {
        top: cursor,
        height: Math.max(totalH, 1),
        tableRows,
        breakBefore: breakBefore || undefined,
        breakAfter: breakAfter || undefined,
        docxIndex: block.docxIndex ?? undefined,
        section: si,
      }
      blocks.push(box)
      rowTextsOf.set(box, rowTexts)
      cursor += Math.max(totalH, 1)
      continue
    }

    // ── Paragraph/image/passthrough ────────────────────────────────────────
    const paraData = computeParaLineData(block, contentWidthPx, docGrid, parsed)

    // Footnote refs: footnote height is added to the paragraph's height (reserving footnote space on that page)
    let footnoteExtra = 0
    if (block.runs) {
      for (const run of block.runs) {
        if (run.noteRef?.kind === 'footnote') {
          const fn = footnoteById.get(run.noteRef.id)
          footnoteExtra += estimateFootnoteHeight(
            fn?.text ?? '',
            contentWidthPx,
            docGrid,
            metrics,
            resolveNoteStyle(parsed, fn?.styleId, fn?.spacing),
            fn?.richParas,
          )
        }
      }
      // the once-per-page separator is charged by the pagination engine via footnoteExtraPx
    }

    const style = resolveStyle(block.styleId, block.format, parsed)
    const totalH = paraData.totalHeight + footnoteExtra

    const box: BlockBox = {
      top: cursor,
      height: Math.max(totalH, 1),
      lineBoxes: paraData.lineBoxes,
      spaceBeforePx: paraData.spaceBeforePx,
      spaceAfterPx: paraData.spaceAfterPx,
      ...(footnoteExtra > 0 ? { footnoteExtraPx: footnoteExtra } : {}),
      keepNext: style.keepNext || undefined,
      keepLines: style.keepLines || undefined,
      widowControl: style.widowControl,
      breakBefore: breakBefore || undefined,
      breakAfter: breakAfter || undefined,
      colBreakAfter: colBreakAfter || undefined,
      docxIndex: block.docxIndex ?? undefined,
      section: si,
    }
    blocks.push(box)
    if (paraData.lineTexts) lineTextsOf.set(box, paraData.lineTexts)
    cursor += Math.max(totalH, 1)
  }

  const slices = computeSectionedSlicesF2(blocks, geoms, cursor)
  const ourPages = slices.length

  // Diagnostics: DEBUG_STEM=14 npx vitest run tests/pagination-parity.test.ts
  const blockByDocxIdxDebug = process.env.DEBUG_STEM
    ? new Map(parsed.blocks.map((b) => [b.docxIndex, b]))
    : null
  if (process.env.DEBUG_STEM && stem.includes(process.env.DEBUG_STEM)) {
    console.log(
      `[${stem}] contentH=`,
      geoms.map((g) => g.contentHeight.toFixed(1)),
      'cursor=',
      cursor.toFixed(1),
    )
    for (const b of blocks) {
      const src = blockByDocxIdxDebug?.get(b.docxIndex ?? -1)
      const text = (src?.runs ?? [])
        .map((r: { text: string }) => r.text)
        .join('')
        .slice(0, 14)
      console.log(
        `  top=${b.top.toFixed(1)} h=${b.height.toFixed(1)} lines=${b.lineBoxes?.length ?? '-'}`,
        `spB=${(b.spaceBeforePx ?? 0).toFixed(1)} spA=${(b.spaceAfterPx ?? 0).toFixed(1)}`,
        b.keepNext ? 'keepNext' : '',
        b.tableRows ? `tableRows=${b.tableRows.length}` : '',
        '|',
        text,
      )
    }
    console.log(
      '  slices:',
      slices.map((s) => `${s.start.toFixed(0)}-${s.end.toFixed(0)}`).join(' '),
    )
  }

  // Resolve page-start text from each slice (block level → table row level → paragraph line level → next-block fallback)
  const blockByDocxIdx = new Map(parsed.blocks.map((b) => [b.docxIndex, b]))
  const ourPageStarts: string[] = slices.map((slice, i) => {
    if (i === 0) {
      const matchBlock = blockByDocxIdx.get(blocks[0]?.docxIndex ?? -1) ?? null
      return blockText(matchBlock)
    }
    // tblHeader repetition: pages that split a table render the header row at the top (reproduces the recorded baseline PDFs)
    if (slice.repeatHeader) {
      const src = blocks.find((b) => Math.abs(b.top - slice.repeatHeader!.top) < 2 && b.tableRows)
      const texts = src ? rowTextsOf.get(src) : undefined
      if (texts) {
        const headerRows = src!.tableRows!.filter((r) => r.isHeader).length
        const headerText = texts.slice(0, Math.max(headerRows, 1)).join(' ').trim()
        if (headerText) return headerText.slice(0, 40)
      }
    }
    // Block-level boundary: the page starts exactly at some block's top
    const exactBlock = blocks.find((b) => Math.abs(b.top - slice.start) < 2)
    if (exactBlock) {
      const matchBlock = blockByDocxIdx.get(exactBlock.docxIndex ?? -1) ?? null
      return blockText(matchBlock)
    }
    // The page start falls inside some block
    const container = blocks.find((b) => b.top < slice.start && slice.start < b.top + b.height)
    if (container) {
      // Table row-level page break: locate the page-start row
      if (container.tableRows && container.tableRows.length > 0) {
        const rowTexts = rowTextsOf.get(container) ?? []
        let rowOffset = container.top
        for (let ri = 0; ri < container.tableRows.length; ri++) {
          const rowH = container.tableRows[ri].height
          if (
            Math.abs(rowOffset - slice.start) < 2 ||
            (rowOffset <= slice.start && slice.start < rowOffset + rowH)
          ) {
            return (rowTexts[ri] ?? '').slice(0, 40)
          }
          rowOffset += rowH
        }
        return '[TABLE]'
      }
      // Paragraph line-level page break: locate the page-start line
      const texts = lineTextsOf.get(container)
      const k =
        container.lineBoxes?.findIndex(
          (lb) => Math.abs(container.top + lb.offsetInBlock - slice.start) < 2,
        ) ?? -1
      if (k >= 0 && texts) return (texts[k] ?? '').slice(0, 40)
    }
    // Fallback: the first block with top >= slice.start
    const firstBlock = blocks.find((b) => b.top >= slice.start)
    if (!firstBlock) return ''
    const matchBlock = blockByDocxIdx.get(firstBlock.docxIndex ?? -1) ?? null
    return blockText(matchBlock)
  })

  // Compare against the baseline
  const loPages = baseline.pages
  const loPageStarts = baseline.pageStarts

  let matched = 0
  const matchLen = Math.min(loPageStarts.length, ourPageStarts.length)
  for (let i = 0; i < matchLen; i++) {
    if (ourPageStarts[i] && loPageStarts[i]) {
      // The page-1 start is not a pagination decision; two-column/floating elements scramble
      // pdftotext's first-line order, so text comparison is only meaningful from page 2 on
      if (i === 0) {
        matched++
        continue
      }
      const loKey = loPageStarts[i].replace(/\s+/g, '').slice(0, 20)
      const ourKey = ourPageStarts[i].replace(/\s+/g, '').slice(0, 20)
      if (loKey && ourKey && loKey.includes(ourKey.slice(0, 10))) matched++
      else if (ourKey && loKey && ourKey.includes(loKey.slice(0, 10))) matched++
    }
  }

  if (process.env.DEBUG_STEM && stem.includes(process.env.DEBUG_STEM)) {
    for (let i = 0; i < Math.max(loPageStarts.length, ourPageStarts.length); i++) {
      console.log(
        `  p${i + 1} base:「${loPageStarts[i] ?? ''}」 ours:「${ourPageStarts[i] ?? ''}」`,
      )
    }
  }

  return {
    stem,
    loPages,
    ourPages,
    pageDiff: Math.abs(ourPages - loPages),
    loPageStarts,
    ourPageStarts,
    matchedPageStarts: matched,
    matchRate: loPages === 0 ? 0 : matched / loPages,
  }
}

function blockText(
  block: {
    type: string
    runs?: Array<{ text: string }>
    label?: string
    previewText?: string
  } | null,
): string {
  if (!block) return ''
  if (block.runs)
    return block.runs
      .map((r) => r.text)
      .join('')
      .slice(0, 40)
  return (block.label ?? block.previewText ?? '').slice(0, 40)
}

// ─── Tests ──────────────────────────────────────────────────────────────────

let baseline: Baseline = {}
/** Precise real-Word baseline (only successfully exported entries); empty object = not collected */
const wordBaseline: Baseline = {}

describe('pagination parity (F2: line-level pagination + page-break constraints)', () => {
  beforeAll(() => {
    if (existsSync(BASELINE_FILE)) {
      baseline = JSON.parse(readFileSync(BASELINE_FILE, 'utf-8'))
    }
    if (existsSync(BASELINE_WORD_FILE)) {
      const raw: Baseline = JSON.parse(readFileSync(BASELINE_WORD_FILE, 'utf-8'))
      for (const [stem, entry] of Object.entries(raw)) {
        if (!entry.error && entry.pages > 0) wordBaseline[stem] = entry
      }
    }
  })

  it.skipIf(!existsSync(BASELINE_FILE))('baseline-lo.json exists', () => {
    expect(Object.keys(baseline).length).toBeGreaterThan(0)
  })

  it.skipIf(!existsSync(BASELINE_FILE))(
    'generates the F2 pagination parity report',
    async () => {
      const stems = Object.keys(baseline).sort()
      const results: ParityResult[] = []

      for (const stem of stems) {
        const entry = baseline[stem]
        if (entry.error) {
          results.push({
            stem,
            loPages: entry.pages,
            ourPages: 0,
            pageDiff: 0,
            loPageStarts: [],
            ourPageStarts: [],
            matchedPageStarts: 0,
            matchRate: 0,
            error: entry.error,
          })
          continue
        }
        const r = await runParity(stem, entry)
        results.push(r)
      }

      // Run a separate pass on the precise Word baseline (valid entries) — the primary acceptance metric
      const wordResults: ParityResult[] = []
      for (const stem of Object.keys(wordBaseline).sort()) {
        wordResults.push(await runParity(stem, wordBaseline[stem]))
      }

      writeReport(results, wordResults)

      const summarize = (rs: ParityResult[], label: string) => {
        const valid = rs.filter((r) => !r.error)
        const total = valid.reduce((s, r) => s + r.loPages, 0)
        const matchedStarts = valid.reduce((s, r) => s + r.matchedPageStarts, 0)
        const rate = total === 0 ? 0 : matchedStarts / total
        console.log(
          `${label} parity: ${(rate * 100).toFixed(1)}% (${matchedStarts}/${total} page-start blocks matched), ` +
            `identical page counts ${valid.filter((r) => r.pageDiff === 0).length}/${valid.length} docs`,
        )
        return { rate, valid }
      }
      console.log('')
      const lo = summarize(results, 'F2 vs LO')
      const word = wordResults.length > 0 ? summarize(wordResults, 'F2 vs Word') : null

      const primary = word ?? lo
      if (primary.rate < 0.85) {
        console.warn(
          `⚠️  parity below 85% (currently ${(primary.rate * 100).toFixed(1)}%), see the report for details`,
        )
        primary.valid
          .filter((r) => r.matchRate < 0.5)
          .sort((a, b) => a.matchRate - b.matchRate)
          .slice(0, 5)
          .forEach((r) =>
            console.warn(
              `  ${r.stem}: ${(r.matchRate * 100).toFixed(0)}% (baseline=${r.loPages}, ours=${r.ourPages})`,
            ),
          )
      }

      expect(results.length).toBeGreaterThan(0)
    },
    120_000,
  )
})

// ─── Report generation ──────────────────────────────────────────────────────

function writeReport(results: ParityResult[], wordResults: ParityResult[] = []) {
  const valid = results.filter((r) => !r.error)
  const failed = results.filter((r) => r.error)
  const totalLo = valid.reduce((s, r) => s + r.loPages, 0)
  const matchedStarts = valid.reduce((s, r) => s + r.matchedPageStarts, 0)
  const overallRate = totalLo === 0 ? 0 : matchedStarts / totalLo
  const exactPageMatch = valid.filter((r) => r.pageDiff === 0).length

  const wordValid = wordResults.filter((r) => !r.error)
  const totalWord = wordValid.reduce((s, r) => s + r.loPages, 0)
  const matchedWord = wordValid.reduce((s, r) => s + r.matchedPageStarts, 0)
  const wordRate = totalWord === 0 ? 0 : matchedWord / totalWord

  const lines: string[] = [
    `# Pagination parity report (F2: line-level pagination + Word page-break constraints)`,
    ``,
    `Generated at: ${new Date().toISOString()}`,
    `Baseline source: LibreOffice headless (coarse baseline)${wordValid.length > 0 ? ` + real Word for Mac (precise baseline, ${wordValid.length} docs)` : ''}`,
    ``,
    `## Overview`,
    ``,
    `| Metric | F0 baseline | F2 vs LO | F2 vs Word |`,
    `|------|--------|--------|--------|`,
    `| Corpus | 28 docs | ${valid.length} docs | ${wordValid.length} docs |`,
    `| Baseline total pages | 108 pages | ${totalLo} pages | ${totalWord} pages |`,
    `| Page-start blocks matched | 52/108 = **48.1%** | ${matchedStarts}/${totalLo} = **${(overallRate * 100).toFixed(1)}%** | ${wordValid.length > 0 ? `${matchedWord}/${totalWord} = **${(wordRate * 100).toFixed(1)}%**` : 'not collected'} |`,
    `| Identical page counts | 12/28 docs | ${exactPageMatch}/${valid.length} docs | ${wordValid.length > 0 ? `${wordValid.filter((r) => r.pageDiff === 0).length}/${wordValid.length} docs` : '-'} |`,
    ``,
    `> Milestones: F1 (self-managed line heights + docGrid) 80.6% → F2 block+line fusion 85.2% → F2 unified engine 87.0%.`,
    `> Acceptance criterion (vs precise Word baseline): 90% page-start blocks matched.`,
    ``,
    `## Per-document details (vs LO)`,
    ``,
    `| File | LO pages | Our pages | Diff | Page starts matched | Match rate |`,
    `|------|--------|---------|------|---------|--------|`,
  ]

  for (const r of results) {
    if (r.error) {
      lines.push(`| ${r.stem} | ${r.loPages} | ✗ | - | - | ✗ error |`)
    } else {
      const diffStr =
        r.pageDiff === 0
          ? '✓ 0'
          : `${r.ourPages - r.loPages > 0 ? '+' : ''}${r.ourPages - r.loPages}`
      lines.push(
        `| ${r.stem} | ${r.loPages} | ${r.ourPages} | ${diffStr} | ${r.matchedPageStarts}/${r.loPages} | ${(r.matchRate * 100).toFixed(0)}% |`,
      )
    }
  }

  if (wordValid.length > 0) {
    lines.push(``, `## Per-document details (vs precise Word baseline)`, ``)
    lines.push(`| File | Word pages | Our pages | Diff | Page starts matched | Match rate |`)
    lines.push(`|------|--------|---------|------|---------|--------|`)
    for (const r of wordValid) {
      const diffStr =
        r.pageDiff === 0
          ? '✓ 0'
          : `${r.ourPages - r.loPages > 0 ? '+' : ''}${r.ourPages - r.loPages}`
      lines.push(
        `| ${r.stem} | ${r.loPages} | ${r.ourPages} | ${diffStr} | ${r.matchedPageStarts}/${r.loPages} | ${(r.matchRate * 100).toFixed(0)}% |`,
      )
    }
  }

  if (failed.length > 0) {
    lines.push(``, `## Files that failed processing`, ``)
    for (const r of failed) {
      lines.push(`- **${r.stem}**: ${r.error}`)
    }
  }

  lines.push(``, `## F2 completion status`, ``)
  lines.push(`- [x] Corpus built (28 docx) + LO baseline (108 pages)`)
  lines.push(`- [x] F1: self-managed line heights + docGrid (48.1% → 80.6%)`)
  lines.push(
    `- [x] F2 engine: line-level pagination + widowControl/keepLines/keepNext + table row-level page breaks (cantSplit/tblHeader)`,
  )
  lines.push(
    `- [x] Footnote placeholders: paragraphs with footnote references reserve page-bottom height`,
  )
  const achieved = overallRate >= 0.85
  lines.push(
    `- [${achieved ? 'x' : ' '}] F2 parity ${(overallRate * 100).toFixed(1)}% (stage target ≥85%, acceptance criterion 90%)`,
  )
  lines.push(``, `## Bottleneck analysis`, ``)

  const primaryValid = wordValid.length > 0 ? wordValid : valid
  const primaryLabel = wordValid.length > 0 ? 'Word' : 'LO'
  const underperforming = primaryValid
    .filter((r) => r.matchRate < 0.6)
    .sort((a, b) => a.matchRate - b.matchRate)
  if (underperforming.length > 0) {
    lines.push(`Documents with a match rate <60% (vs ${primaryLabel}, top priority):`, ``)
    for (const r of underperforming) {
      const pageDiff = r.ourPages - r.loPages
      lines.push(
        `- **${r.stem}**: ${(r.matchRate * 100).toFixed(0)}%, ${primaryLabel}=${r.loPages} pages, ours=${r.ourPages} pages (diff ${pageDiff > 0 ? '+' : ''}${pageDiff})`,
      )
      const n = Math.max(r.loPageStarts.length, r.ourPageStarts.length)
      for (let i = 0; i < n; i++) {
        lines.push(
          `  - p${i + 1} ${primaryLabel}:「${(r.loPageStarts[i] ?? '').slice(0, 20)}」 ours:「${(r.ourPageStarts[i] ?? '').slice(0, 20)}」`,
        )
      }
    }
  } else {
    lines.push(`No documents below a 60% match rate (vs ${primaryLabel}) 🎉`)
  }

  lines.push(``, `## Remaining long tail`, ``)
  lines.push(
    `1. openclaw (PDF-converted doc) off by 1 page: multi-section variable margins + text boxes/floating objects;`,
  )
  lines.push(
    `   fixed layout not fully modeled (trHeight/wrap anchoring are modeled but still insufficient)`,
  )
  lines.push(
    `2. 09-large-paragraph line-break points off by 2-3 chars (kerning precision explicitly out of scope per plan §2)`,
  )

  const reportText = lines.join('\n')

  try {
    mkdirSync(dirname(REPORT_FILE), { recursive: true })
    writeFileSync(REPORT_FILE, reportText)
    console.log(`\nReport written to: ${REPORT_FILE}`)
  } catch (e) {
    console.warn('Failed to write the report (path not writable):', e)
    console.log('\n--- report start ---')
    console.log(reportText)
    console.log('--- report end ---')
  }
}
