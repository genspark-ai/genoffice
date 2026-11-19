/**
 * F2 block boxes for a parsed corpus document, measured with the LO baseline
 * metrics. Shared by the parity runner and the invariant suite so both slice
 * exactly the same input.
 */
import { readSections } from '@genoffice/docx-engine'
import type { ParsedDoc, DocGrid, ParaFormat, StyleDisplay } from '@genoffice/docx-engine'
import {
  sectionColGeom,
  sectionColumns,
  sectionGeoms,
  type BlockBox,
  type SectionGeom,
  type TableRowBox,
} from '../../src/renderer/pagination'
import {
  computeLineMetrics,
  estimateFootnoteHeight,
  resolveNoteStyle,
} from '../../src/renderer/line-metrics'
import { loBaselineMetrics } from './lo-fonts'

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

export interface CorpusBlocks {
  blocks: BlockBox[]
  geoms: SectionGeom[]
  /** flow height (bottom of the last block) */
  totalHeight: number
  lineTextsOf: Map<BlockBox, string[]>
  rowTextsOf: Map<BlockBox, string[]>
}

export function buildCorpusBlocks(parsed: ParsedDoc): CorpusBlocks {
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
  return { blocks, geoms, totalHeight: cursor, lineTextsOf, rowTextsOf }
}
