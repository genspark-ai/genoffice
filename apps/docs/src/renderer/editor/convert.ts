import {
  mergePPrFormat,
  setPPrChange,
  stripPPrChange,
  ommlToLatex,
  patchFieldParagraphXml,
  applyImageWrap,
  patchImageParagraphXml,
  patchMathTokens,
  patchTableCellTexts,
  type CellTextsPatch,
  patchDrawingExtent,
  patchTextboxSizes,
  patchShapeStyles,
  type ShapeStylePatch,
  type TextboxSizePatch,
  patchTextboxParas,
  generateTableModelXml,
  type Block,
  type CellBorder,
  type ChartDisplay,
  type ChartPatch,
  type ChartSeriesPatch,
  type FieldDisplay,
  type FieldTextPatch,
  type FormulaDisplay,
  type GeneratedBlock,
  type ImageWrap,
  type NewChart,
  type NewImage,
  type ParaFormat,
  type Run,
  type SaveBlock,
  type SdtShell,
  type SectionInfo,
  type TableCell,
  type TableModel,
  type TextboxDisplay,
  type TextboxParaPatch,
} from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import { textHasComplexScript } from '../line-metrics'
import { inlineMathML } from './equation'
import { isStraightLineKind } from './shape-svg'

/** minimal ProseMirror JSON shapes */
export interface PmMark {
  type: string
  attrs?: Record<string, unknown>
}
export interface PmNode {
  type: string
  attrs?: Record<string, unknown>
  content?: PmNode[]
  text?: string
  marks?: PmMark[]
}

// ---- Block[] -> ProseMirror doc ----

export function blocksToPmDoc(blocks: Block[], sections?: SectionInfo[]): PmNode {
  const content: PmNode[] = []
  for (const block of blocks) {
    if (block.hidden) continue
    content.push(blockToPmNode(block, sectionRowCapTwips(sections, block.docxIndex)))
  }
  if (content.length === 0) {
    content.push({
      type: 'docParagraph',
      attrs: { docxIndex: null, styleId: null, aiChanged: false },
    })
  }
  return { type: 'doc', content }
}

function formatAttrs(format: ParaFormat | undefined): Record<string, unknown> {
  return {
    align: format?.align ?? null,
    lineSpacing: format?.lineSpacing ?? null,
    lineRule: format?.lineRule ?? null,
    lineRawTwips: format?.lineRawTwips ?? null,
    indentLeft: format?.indentLeft ?? null,
    indentRight: format?.indentRight ?? null,
    indentFirstLine: format?.indentFirstLine ?? null,
    spaceBefore: format?.spaceBefore ?? null,
    spaceAfter: format?.spaceAfter ?? null,
    pageBreakBefore: format?.pageBreakBefore ?? false,
    shadingFill: format?.shadingFill ?? null,
    borders: format?.borders ?? null,
    tabStops: format?.tabStops ? JSON.stringify(format.tabStops) : null,
    dropCap: format?.dropCap ? JSON.stringify(format.dropCap) : null,
    bidi: format?.bidi ?? false,
    autoSpace: format?.autoSpace ?? null,
  }
}

/** Effective row-height ceiling: one page of section content (Word truncates taller rows within the page) */
function sectionRowCapTwips(
  sections: SectionInfo[] | undefined,
  docxIndex: number | null,
): number | null {
  if (!sections?.length || docxIndex == null) return null
  const sec =
    sections.find((s) => docxIndex >= s.firstBlockIndex && docxIndex <= s.lastBlockIndex) ??
    sections[sections.length - 1]
  const cap = sec.settings.pageHeight - sec.settings.marginTop - sec.settings.marginBottom
  return cap > 0 ? cap : null
}

/** Cap declared row heights (incl. nested tables); returns the input model when nothing exceeds the cap */
export function capTableRowHeights(model: TableModel, capTwips: number): TableModel {
  let changed = false
  const rowHeightsTwips = model.rowHeightsTwips?.map((h) => {
    if (h != null && h > capTwips) {
      changed = true
      return capTwips
    }
    return h
  })
  const rows = model.rows.map((row) =>
    row.map((cell) => {
      if (!cell.nestedTables?.length) return cell
      const nested = cell.nestedTables.map((nt) => capTableRowHeights(nt, capTwips))
      if (nested.every((nt, i) => nt === cell.nestedTables![i])) return cell
      changed = true
      return { ...cell, nestedTables: nested }
    }),
  )
  return changed ? { ...model, rows, rowHeightsTwips } : model
}

function blockToPmNode(block: Block, rowCapTwips: number | null = null): PmNode {
  switch (block.type) {
    case 'heading':
      return {
        type: 'docHeading',
        attrs: {
          docxIndex: block.docxIndex,
          styleId: block.styleId ?? null,
          aiChanged: false,
          level: block.level ?? 1,
          bookmarks: block.bookmarks ?? null,
          hiddenBookmarks: block.hiddenBookmarks ?? null,
          commentStarts: block.commentStarts ?? null,
          commentEnds: block.commentEnds ?? null,
          pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
          blockRevision: block.blockRevision ?? null,
          ...formatAttrs(block.format),
        },
        content: runsToInline(block.runs ?? []),
      }
    case 'listItem':
      return {
        type: 'docListItem',
        attrs: {
          docxIndex: block.docxIndex,
          styleId: block.styleId ?? null,
          aiChanged: false,
          kind: block.list?.kind ?? 'bullet',
          numId: block.list?.numId ?? null,
          ilvl: block.list?.ilvl ?? 0,
          bookmarks: block.bookmarks ?? null,
          hiddenBookmarks: block.hiddenBookmarks ?? null,
          commentStarts: block.commentStarts ?? null,
          commentEnds: block.commentEnds ?? null,
          pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
          blockRevision: block.blockRevision ?? null,
          ...formatAttrs(block.format),
        },
        content: runsToInline(block.runs ?? []),
      }
    case 'paragraph':
      return {
        type: 'docParagraph',
        attrs: {
          docxIndex: block.docxIndex,
          styleId: block.styleId ?? null,
          aiChanged: false,
          bookmarks: block.bookmarks ?? null,
          hiddenBookmarks: block.hiddenBookmarks ?? null,
          commentStarts: block.commentStarts ?? null,
          commentEnds: block.commentEnds ?? null,
          sdtShell: block.sdtShell ? JSON.stringify(block.sdtShell) : null,
          moveRevision: block.moveRevision ?? null,
          pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
          blockRevision: block.blockRevision ?? null,
          ...formatAttrs(block.format),
        },
        content: runsToInline(block.runs ?? []),
      }
    case 'table':
      return tableModelToPmNode(
        block.table ?? { rows: [[{ paras: [''] }]] },
        block.docxIndex,
        block.blockRevision ?? null,
        rowCapTwips,
      )
    default:
      // image / passthrough: protected whole-unit blocks
      return {
        type: 'docProtected',
        attrs: {
          docxIndex: block.docxIndex,
          blockRevision: block.blockRevision ?? null,
          blockType: block.type,
          label: block.label ?? block.type,
          previewText: block.previewText ?? '',
          imageDataUrl: block.imageDataUrl ?? null,
          oleProgId: block.oleProgId ?? null,
          imageWidthPx: block.imageWidthPx ?? null,
          imageHeightPx: block.imageHeightPx ?? null,
          imageAlign: block.imageAlign ?? null,
          imageWrap: block.imageWrap ?? null,
          imageOffsetXEmu: block.imageOffsetXEmu ?? null,
          imageOffsetYEmu: block.imageOffsetYEmu ?? null,
          imagePosH: block.imagePosH ?? null,
          imagePosV: block.imagePosV ?? null,
          imageRotDeg: block.imageRotDeg ?? null,
          imageFlipH: block.imageFlipH ?? false,
          imageFlipV: block.imageFlipV ?? false,
          table:
            block.table && rowCapTwips != null
              ? capTableRowHeights(block.table, rowCapTwips)
              : (block.table ?? null),
          fieldDisplay: block.fieldDisplay ?? null,
          decorative: block.decorative ?? false,
          ruleColorHex: block.ruleColorHex ?? null,
          ruleThicknessPx: block.ruleThicknessPx ?? null,
          ruleWidthPx: block.ruleWidthPx ?? null,
          brokenImage: block.brokenImage ?? false,
          invisibleMarker: block.invisibleMarker ?? false,
          textboxes: block.textboxes ?? null,
          formulaDisplay: block.formulaDisplay ?? null,
          chartDisplay: block.chartDisplay ?? null,
        },
      }
  }
}

function cellRowSpan(model: TableModel, row: number, cell: number, positions: number[][]): number {
  const current = model.rows[row][cell]
  if (current.vMerge !== 'restart') return 1
  const gridColumn = positions[row][cell]
  let span = 1
  for (let r = row + 1; r < model.rows.length; r++) {
    const index = positions[r].indexOf(gridColumn)
    if (index === -1 || model.rows[r][index].vMerge !== 'continue') break
    span++
  }
  return span
}

function cellBorderPx(b: CellBorder | undefined): number {
  if (!b || b.style === 'none' || b.style === 'nil') return 0
  return Math.max(1, Math.round(((b.szEighths ?? 4) / 8 / 72) * 96))
}

/** Inner clip-box height (twips) for cells of hRule="exact" rows — Word clips overflow
 *  instead of growing the row. null = not exact, or the cell spans rows (its clip height
 *  would be the sum of the spanned rows; not handled yet). */
export function cellClipTwips(
  model: TableModel,
  rowIndex: number,
  cell: TableCell,
  rowSpan: number,
): number | null {
  if (rowSpan > 1 || model.rowHeightRules?.[rowIndex] !== 'exact') return null
  const h = model.rowHeightsTwips?.[rowIndex]
  if (!h) return null
  const padTop = cell.cellMarTwips?.top ?? model.cellMarTwips?.top ?? 0
  const padBottom = cell.cellMarTwips?.bottom ?? model.cellMarTwips?.bottom ?? 0
  // collapsed borders straddle the row edges: half of each eats into the row height
  const bTop = cell.borders?.top ?? (rowIndex === 0 ? model.borders?.top : model.borders?.insideH)
  const bBottom =
    cell.borders?.bottom ??
    (rowIndex === model.rows.length - 1 ? model.borders?.bottom : model.borders?.insideH)
  const borderTwips = Math.round(((cellBorderPx(bTop) + cellBorderPx(bBottom)) / 2) * 15)
  return Math.max(0, h - padTop - padBottom - borderTwips)
}

export function tableModelToPmNode(
  model: TableModel,
  docxIndex: number | null = null,
  blockRevision: Block['blockRevision'] | null = null,
  rowCapTwips: number | null = null,
): PmNode {
  if (rowCapTwips != null) model = capTableRowHeights(model, rowCapTwips)
  const positions = model.rows.map((row) => {
    let column = 0
    return row.map((cell) => {
      const at = column
      column += cell.colSpan ?? 1
      return at
    })
  })
  // 1 px = 15 twips (96dpi); use real values when absolute grid widths exist, fall back to percentage approximation otherwise
  const widthPx = model.colWidthsTwips
    ? model.colWidthsTwips.map((width) => Math.max(24, Math.round(width / 15)))
    : model.colWidthsPct?.map((width) => Math.max(24, Math.round(width * 6.24)))
  const hasHeaderRow =
    model.rows.length > 1 && model.rows[0].every((cell) => cell.bold || cell.fill !== undefined)
  const table: PmNode = {
    type: 'docTable',
    attrs: {
      docxIndex,
      blockRevision,
      colWidthsPct: model.colWidthsPct ?? null,
      widthPx:
        !model.widthPct && model.colWidthsTwips && widthPx
          ? widthPx.reduce((sum, width) => sum + width, 0)
          : null,
      widthPct: model.widthPct ?? null,
      cellMar: model.cellMarTwips ?? null,
      borders: model.borders ?? null,
      tblAlign: model.align ?? null,
      indentTwips: model.indentTwips ?? null,
      tblStyleId: model.tblStyleId ?? null,
      bidiVisual: model.bidiVisual ?? false,
      originalStructure: null,
      originalFormatting: null,
    },
    content: model.rows.map((row, rowIndex) => ({
      type: 'docTableRow',
      attrs: {
        heightTwips: model.rowHeightsTwips?.[rowIndex] ?? null,
        heightRule: model.rowHeightRules?.[rowIndex] ?? null,
        rawTrPr: model.rawTrPrs?.[rowIndex] ?? null,
        rowRevision: model.rowRevisions?.[rowIndex] ?? null,
      },
      content: row.flatMap((cell, cellIndex) => {
        if (cell.vMerge === 'continue') return []
        const start = positions[rowIndex][cellIndex]
        const colspan = cell.colSpan ?? 1
        const rowspan = cellRowSpan(model, rowIndex, cellIndex, positions)
        return [
          {
            type: hasHeaderRow && rowIndex === 0 ? 'docTableHeader' : 'docTableCell',
            attrs: {
              colspan,
              rowspan,
              clipHeightTwips: cellClipTwips(model, rowIndex, cell, rowspan),
              colwidth: widthPx ? widthPx.slice(start, start + colspan) : null,
              cellMar: cell.cellMarTwips ?? null,
              textDirection: cell.textDirection ?? null,
              fill: cell.fill ?? null,
              color: cell.color ?? null,
              bold: cell.bold ?? false,
              align: cell.align ?? null,
              vAlign: cell.vAlign ?? null,
              borders: cell.borders ?? null,
              rawTcPr: cell.rawTcPr ?? null,
              cellRevision: cell.cellRevision ?? null,
            },
            content: cellContentNodes(cell),
          },
        ]
      }),
    })),
  }
  table.attrs!.originalStructure = tableStructureSignature(table)
  table.attrs!.originalFormatting = tableFormattingSignature(table)
  return table
}

/** cell paragraphs with nested tables (read-only atoms) spliced in at their anchors */
function cellContentNodes(cell: TableCell): PmNode[] {
  const paraNodes: PmNode[] = (
    cell.richParas?.length
      ? cell.richParas
      : (cell.paras.length > 0 ? cell.paras : ['']).map((text) => ({
          align: cell.align,
          runs: text === '' ? [] : [{ text, bold: cell.bold, color: cell.color }],
        }))
  ).map((paragraph) => {
    const list = 'list' in paragraph ? paragraph.list : undefined
    return list
      ? {
          type: 'docListItem',
          attrs: {
            ...formatAttrs(paragraph),
            kind: list.kind,
            numId: list.numId,
            ilvl: list.ilvl,
          },
          content: runsToInline(paragraph.runs),
        }
      : {
          type: 'docParagraph',
          attrs: formatAttrs(paragraph),
          content: runsToInline(paragraph.runs),
        }
  })
  // fidelity on save comes from the outer table's bytes; reverse insertion keeps anchors valid
  const nested = cell.nestedTables ?? []
  const content = [...paraNodes]
  for (let i = nested.length - 1; i >= 0; i--) {
    const at = Math.min(cell.nestedTableAnchors?.[i] ?? paraNodes.length, paraNodes.length)
    content.splice(at, 0, { type: 'docNestedTable', attrs: { model: nested[i] } })
  }
  return content
}

export function tableStructureSignature(table: PmNode): string {
  return JSON.stringify({
    widths: table.attrs?.colWidthsPct ?? null,
    tblStyle: table.attrs?.tblStyleId ?? null,
    rows: (table.content ?? []).map((row) => [
      row.attrs?.heightTwips ?? null,
      // accepting/rejecting revisions strips records from trPr/tcPr: include them in the signature to trigger regeneration (works for empty rows too)
      row.attrs?.rawTrPr ?? null,
      row.attrs?.rowRevision ?? null,
      ...(row.content ?? []).map((cell) => [
        cell.type,
        cell.attrs?.colspan ?? 1,
        cell.attrs?.rowspan ?? 1,
        cell.attrs?.colwidth ?? null,
        cell.attrs?.fill ?? null,
        cell.attrs?.color ?? null,
        cell.attrs?.bold ?? false,
        cell.attrs?.align ?? null,
        cell.attrs?.vAlign ?? null,
        cell.attrs?.borders ?? null,
        cell.attrs?.rawTcPr ?? null,
        cell.attrs?.cellRevision ?? null,
      ]),
    ]),
  })
}

function inlineFormattingSignature(content: PmNode[] | undefined): string[] {
  const styles: string[] = []
  for (const node of content ?? []) {
    const style =
      node.type === 'text'
        ? JSON.stringify((node.marks ?? []).map((mark) => [mark.type, mark.attrs ?? null]))
        : node.type
    if (styles[styles.length - 1] !== style) styles.push(style)
  }
  return styles
}

export function tableFormattingSignature(table: PmNode): string {
  return JSON.stringify(
    (table.content ?? []).map((row) =>
      (row.content ?? []).map((cell) =>
        (cell.content ?? []).map((paragraph) => [
          normalizedFormat(nodeFormat(paragraph)),
          // list toggles must force regeneration (the text patch keeps old pPr)
          paragraph.type === 'docListItem'
            ? [paragraph.attrs?.kind, paragraph.attrs?.numId, paragraph.attrs?.ilvl]
            : null,
          inlineFormattingSignature(paragraph.content),
        ]),
      ),
    ),
  )
}

function paragraphText(node: PmNode): string {
  return inlineToRuns(node.content ?? [])
    .map((run) => run.text)
    .join('')
}

function unwrapBlockRevisionXml(xml: string): string {
  const match = /^<w:(ins|del)\b[^>]*>([\s\S]*)<\/w:\1>$/.exec(xml)
  return match ? match[2] : xml
}

/** Convert native table cells back to the physical OOXML-style TableModel grid. */
export function pmTableToModel(table: PmNode): TableModel {
  type ActiveSpan = {
    remaining: number
    span: number
    cell: Omit<TableModel['rows'][number][number], 'paras'>
  }
  let active = new Map<number, ActiveSpan>()
  const rows: TableModel['rows'] = []
  let maximumColumns = 0

  const rowHeightsTwips: Array<number | null> = []
  const rowHeightRules: NonNullable<TableModel['rowHeightRules']> = []
  const rawTrPrs: Array<string | null> = []
  const rowRevisions: TableModel['rowRevisions'] = []
  for (const rowNode of table.content ?? []) {
    rowHeightsTwips.push((rowNode.attrs?.heightTwips as number | null) ?? null)
    rowHeightRules.push(
      (rowNode.attrs?.heightRule as NonNullable<TableModel['rowHeightRules']>[number]) ?? null,
    )
    rawTrPrs.push((rowNode.attrs?.rawTrPr as string | null) ?? null)
    rowRevisions.push(
      (rowNode.attrs?.rowRevision as NonNullable<TableModel['rowRevisions']>[number]) ?? null,
    )
    const entries: Array<{ column: number; cell: TableModel['rows'][number][number] }> = []
    const occupied = new Set<number>()
    for (const [column, span] of active) {
      entries.push({
        column,
        cell: {
          ...span.cell,
          paras: [''],
          colSpan: span.span > 1 ? span.span : undefined,
          vMerge: 'continue',
        },
      })
      for (let i = 0; i < span.span; i++) occupied.add(column + i)
    }

    const added = new Map<number, ActiveSpan>()
    let cursor = 0
    for (const cellNode of rowNode.content ?? []) {
      while (occupied.has(cursor)) cursor++
      const colspan = Math.max(1, Number(cellNode.attrs?.colspan) || 1)
      const rowspan = Math.max(1, Number(cellNode.attrs?.rowspan) || 1)
      const style = {
        fill: (cellNode.attrs?.fill as string | null) ?? undefined,
        color: (cellNode.attrs?.color as string | null) ?? undefined,
        bold: cellNode.attrs?.bold ? true : undefined,
        align: (cellNode.attrs?.align as TableModel['rows'][number][number]['align']) ?? undefined,
        vAlign: (cellNode.attrs?.vAlign as TableCell['vAlign'] | null) ?? undefined,
        borders: (cellNode.attrs?.borders as TableCell['borders'] | null) ?? undefined,
        rawTcPr: (cellNode.attrs?.rawTcPr as string | null) ?? undefined,
      }
      const cellParas = (cellNode.content ?? []).filter(
        (n) => n.type === 'docParagraph' || n.type === 'docListItem',
      )
      // nested tables round-trip with the model: text edits use the surgical patch, structural regeneration emits the whole table (nested tables no longer dropped)
      const nestedModels: TableModel[] = []
      const nestedAnchors: number[] = []
      let paraCount = 0
      for (const n of cellNode.content ?? []) {
        if (n.type === 'docParagraph' || n.type === 'docListItem') paraCount++
        else if (n.type === 'docNestedTable' && n.attrs?.model) {
          nestedModels.push(n.attrs.model as TableModel)
          nestedAnchors.push(paraCount)
        }
      }
      entries.push({
        column: cursor,
        cell: {
          ...style,
          paras: cellParas.map(paragraphText),
          richParas: cellParas.map((paragraph) => ({
            ...nodeFormat(paragraph),
            ...(paragraph.type === 'docListItem' && paragraph.attrs?.numId
              ? {
                  list: {
                    kind: (paragraph.attrs.kind as 'bullet' | 'ordered') ?? 'bullet',
                    numId: String(paragraph.attrs.numId),
                    ilvl: Number(paragraph.attrs.ilvl) || 0,
                  },
                }
              : {}),
            runs: inlineToRuns(paragraph.content ?? []),
          })),
          ...(nestedModels.length > 0
            ? { nestedTables: nestedModels, nestedTableAnchors: nestedAnchors }
            : {}),
          colSpan: colspan > 1 ? colspan : undefined,
          vMerge: rowspan > 1 ? 'restart' : undefined,
        },
      })
      if (rowspan > 1) added.set(cursor, { remaining: rowspan - 1, span: colspan, cell: style })
      for (let i = 0; i < colspan; i++) occupied.add(cursor + i)
      cursor += colspan
    }
    maximumColumns = Math.max(maximumColumns, occupied.size)
    rows.push(entries.sort((a, b) => a.column - b.column).map((entry) => entry.cell))

    const next = new Map<number, ActiveSpan>()
    for (const [column, span] of active) {
      if (span.remaining > 1) next.set(column, { ...span, remaining: span.remaining - 1 })
    }
    for (const [column, span] of added) next.set(column, span)
    active = next
  }

  let colWidthsPct = (table.attrs?.colWidthsPct as number[] | null) ?? undefined
  let colWidthsTwips: number[] | undefined
  const firstRow = table.content?.[0]?.content ?? []
  const explicit = firstRow.flatMap((cell) => (cell.attrs?.colwidth as number[] | null) ?? [])
  if (explicit.length === maximumColumns && explicit.every((width) => width > 0)) {
    const total = explicit.reduce((sum, width) => sum + width, 0)
    colWidthsPct = explicit.map((width) => (width / total) * 100)
    colWidthsTwips = explicit.map((width) => Math.round(width * 15))
  } else if (colWidthsPct?.length !== maximumColumns) {
    colWidthsPct = Array.from({ length: maximumColumns }, () => 100 / maximumColumns)
  }
  const tblStyleAttr = table.attrs?.tblStyleId as string | null | undefined
  return {
    rows,
    ...(colWidthsPct ? { colWidthsPct } : {}),
    ...(colWidthsTwips ? { colWidthsTwips } : {}),
    ...(rowHeightsTwips.some((h) => h !== null) ? { rowHeightsTwips, rowHeightRules } : {}),
    ...(rawTrPrs.some((r) => r !== null) ? { rawTrPrs } : {}),
    ...(rowRevisions.some((revision) => revision !== null) ? { rowRevisions } : {}),
    // null (cleared) → '' removes explicitly; undefined leaves it alone
    ...(tblStyleAttr !== undefined ? { tblStyleId: tblStyleAttr ?? '' } : {}),
  }
}

/**
 * w:w horizontal character scaling → letter-spacing approximation (em): estimate,
 * from the run text's average character width, how much layout width to add or
 * remove per character. Error is tiny for monospaced CJK; glyphs themselves are not scaled.
 */
function charScaleEm(text: string, scalePct: number): number {
  let sum = 0
  let n = 0
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    const wide =
      (cp >= 0x2e80 && cp <= 0x9fff) ||
      (cp >= 0xac00 && cp <= 0xd7af) ||
      (cp >= 0xf900 && cp <= 0xfaff) ||
      (cp >= 0xff00 && cp <= 0xff60) ||
      cp >= 0x20000
    sum += wide ? 1 : 0.52
    n++
  }
  const avg = n > 0 ? sum / n : 0.52
  return Math.round((scalePct / 100 - 1) * avg * 1000) / 1000
}

export function runsToInline(runs: Run[]): PmNode[] {
  const nodes: PmNode[] = []
  for (const run of runs) {
    if (run.math) {
      nodes.push({
        type: 'docInlineMath',
        attrs: {
          omml: run.math.omml,
          mathml: inlineMathML(run.math.omml),
          // recovered LaTeX makes parsed formulas re-editable; null = atom only
          latex: ommlToLatex(run.math.omml),
          text: run.text,
        },
      })
      continue
    }
    if (run.ruby) {
      nodes.push({
        type: 'docRuby',
        attrs: { base: run.text, rt: run.ruby.rt, xml: run.ruby.xml },
      })
      continue
    }
    if (run.noteRef) {
      nodes.push({
        type: 'docNoteRef',
        attrs: {
          kind: run.noteRef.kind,
          id: run.noteRef.id,
          num: parseInt(run.text, 10) || 1,
        },
      })
      continue
    }
    const marks = runMarks(run)
    // \n = soft line break, \f = in-paragraph page break (w:br w:type="page")
    for (const segment of run.text.split(/([\n\f])/)) {
      if (segment === '\n') nodes.push({ type: 'hardBreak' })
      else if (segment === '\f') nodes.push({ type: 'hardBreak', attrs: { pageBreak: true } })
      else if (segment !== '') {
        nodes.push({ type: 'text', text: segment, ...(marks.length > 0 ? { marks } : {}) })
      }
    }
    // A run can carry both w:t text and a w:drawing; generate.ts writes the
    // text before the drawing, so emit the image after the text segments
    if (run.image) {
      nodes.push({
        type: 'docInlineImage',
        attrs: {
          dataUrl: run.image.dataUrl,
          widthPx: run.image.widthPx ?? null,
          heightPx: run.image.heightPx ?? null,
          xml: run.image.xml,
        },
      })
    }
    if (run.xeTerm !== undefined) {
      nodes.push({ type: 'docXeMark', attrs: { term: run.xeTerm } })
    }
  }
  return nodes
}

function runMarks(run: Run): PmMark[] {
  const marks: PmMark[] = []
  if (run.bold) marks.push({ type: 'bold' })
  if (run.italic) marks.push({ type: 'italic' })
  if (run.underline) marks.push({ type: 'underline' })
  if (run.strike) marks.push({ type: 'strike' })
  if (run.link)
    marks.push({
      type: 'link',
      attrs: { href: run.link.href, rId: run.link.rId ?? null, tooltip: run.link.tooltip ?? null },
    })
  if (run.refField !== undefined) marks.push({ type: 'refField', attrs: { name: run.refField } })
  if (run.instrField !== undefined)
    marks.push({ type: 'instrField', attrs: { instr: run.instrField } })
  if (run.commentIds?.length)
    marks.push({ type: 'comment', attrs: { ids: run.commentIds.join(' ') } })
  if (run.ins) {
    marks.push({
      type: 'ins',
      attrs: { author: run.ins.author, date: run.ins.date ?? null, id: run.ins.id ?? null },
    })
  }
  if (run.del) {
    marks.push({
      type: 'del',
      attrs: { author: run.del.author, date: run.del.date ?? null, id: run.del.id ?? null },
    })
  }
  if (run.rPrChange) {
    marks.push({
      type: 'rprChange',
      attrs: {
        author: run.rPrChange.author,
        date: run.rPrChange.date ?? null,
        id: run.rPrChange.id ?? null,
        old: run.rPrChange.old ?? null,
      },
    })
  }
  if (
    run.color ||
    run.sizeHalfPoints ||
    run.font ||
    run.fontAscii ||
    run.charSpacingTwips ||
    run.charScalePct ||
    run.highlight ||
    run.vertAlign ||
    run.em ||
    run.styleId ||
    run.rawRPr
  ) {
    marks.push({
      type: 'docTextStyle',
      attrs: {
        color: run.color ?? null,
        sizeHalfPoints: run.sizeHalfPoints ?? null,
        font: run.font ?? null,
        fontAscii: run.fontAscii ?? null,
        // cs chain only kicks in for complex-script text (Word's w:cs semantics)
        csFont: run.csFont && textHasComplexScript(run.text) ? run.csFont : null,
        charSpacingTwips: run.charSpacingTwips ?? null,
        charScaleEm: run.charScalePct ? charScaleEm(run.text, run.charScalePct) : null,
        highlight: run.highlight ?? null,
        vertAlign: run.vertAlign ?? null,
        em: run.em ?? null,
        styleId: run.styleId ?? null,
        rawRPr: run.rawRPr ?? null,
      },
    })
  }
  return marks
}

// ---- ProseMirror doc -> SaveBlock[] (dirty detection via content signatures) ----

export interface SavePlan {
  saveBlocks: SaveBlock[]
  /**
   * Chart data edits. These patch the chart's own zip part, not the body
   * paragraph (which stays byte-identical); App applies patchChartPartXml
   * and hands the result to saveDocx via options.partXml.
   */
  chartPatches: Array<{ partPath: string; patch: ChartPatch }>
  /**
   * docxIndex of an original block -> its index in saveBlocks. Ink
   * annotations anchor by docxIndex and the engine wants finalBlocks indexes.
   */
  saveBlockIndexByDocx: Map<number, number>
  /** count of blocks whose content was regenerated */
  changedCount: number
  /** count of original blocks removed */
  deletedCount: number
  totalOriginal: number
}

export function pmDocToSavePlan(doc: PmNode, originalBlocks: Block[]): SavePlan {
  const originalByIndex = new Map<number, Block>()
  for (const block of originalBlocks) {
    if (!block.hidden && block.docxIndex !== null) originalByIndex.set(block.docxIndex, block)
  }

  // Multi-block sdt groups (one w:sdt split into N blocks): the shell open/close
  // bytes live on the first/last member. Track which surviving member emits
  // them so deleting or regenerating a boundary member never unbalances the sdt.
  const sdtGroupOf = new Map<number, number>()
  const sdtGroupShells = new Map<number, SdtShell>()
  for (const block of originalBlocks) {
    const shell = block.sdtShell
    if (block.hidden || block.docxIndex === null || shell?.group === undefined) continue
    sdtGroupOf.set(block.docxIndex, shell.group)
    const agg = sdtGroupShells.get(shell.group)
    if (!agg) sdtGroupShells.set(shell.group, { ...shell })
    else {
      if (shell.openXml) agg.openXml = shell.openXml
      if (shell.closeXml) agg.closeXml = shell.closeXml
    }
  }
  const sdtGroupEmits = new Map<
    number,
    Array<{ at: number; idx: number; open: boolean; close: boolean }>
  >()

  const saveBlocks: SaveBlock[] = []
  const chartPatches: Array<{ partPath: string; patch: ChartPatch }> = []
  const saveBlockIndexByDocx = new Map<number, number>()
  let changedCount = 0
  const usedIndexes = new Set<number>()
  /** record where an original anchor ended up in saveBlocks (set before push) */
  const mapAnchor = (idx: number) => saveBlockIndexByDocx.set(idx, saveBlocks.length)

  const recordSdtEmit = (node: PmNode, sb: SaveBlock, at: number) => {
    const idx = node.attrs?.docxIndex as number | null | undefined
    if (idx === null || idx === undefined) return
    if (saveBlockIndexByDocx.get(idx) !== at) return
    const group = sdtGroupOf.get(idx)
    if (group === undefined) return
    const shell = sdtGroupShells.get(group)!
    let open = false
    let close = false
    if (sb.kind === 'original') {
      const original = originalByIndex.get(idx)
      open = !!original?.sdtShell?.openXml
      close = !!original?.sdtShell?.closeXml
    } else if (sb.kind === 'generated') {
      open = !!sb.block.sdtShell?.openXml
      close = !!sb.block.sdtShell?.closeXml
    } else if (sb.kind === 'xml') {
      open = !!shell.openXml && sb.xml.startsWith(shell.openXml)
      close = !!shell.closeXml && sb.xml.endsWith(shell.closeXml)
    }
    let emits = sdtGroupEmits.get(group)
    if (!emits) sdtGroupEmits.set(group, (emits = []))
    emits.push({ at, idx, open, close })
  }

  for (const node of doc.content ?? []) {
    const revision = node.attrs?.blockRevision as {
      kind: 'ins' | 'del'
      author: string
      date?: string
      id?: string
    } | null
    const pushBlock = (block: SaveBlock) => {
      recordSdtEmit(node, block, saveBlocks.length)
      saveBlocks.push(revision ? { ...block, revision } : block)
    }
    if (node.type === 'docTable') {
      const idx = node.attrs?.docxIndex as number | null
      const original =
        idx !== null && idx !== undefined && !usedIndexes.has(idx)
          ? originalByIndex.get(idx)
          : undefined
      const model = pmTableToModel(node)
      if (original?.type === 'table') {
        usedIndexes.add(idx!)
        mapAnchor(idx!)
        const structurallyUnchanged =
          node.attrs?.originalStructure === tableStructureSignature(node)
        const formattingUnchanged =
          node.attrs?.originalFormatting === tableFormattingSignature(node)
        if (
          original.blockRevision &&
          !revision &&
          structurallyUnchanged &&
          formattingUnchanged &&
          original.originalXml
        ) {
          changedCount++
          pushBlock({ kind: 'xml', xml: unwrapBlockRevisionXml(original.originalXml) })
          continue
        }
        const tableTexts =
          structurallyUnchanged && formattingUnchanged
            ? tableTextsPatchFromModel(model, original)
            : null
        if (structurallyUnchanged && formattingUnchanged && tableTexts && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchTableCellTexts(original.originalXml, tableTexts) })
        } else if (structurallyUnchanged && formattingUnchanged) {
          pushBlock({ kind: 'original', docxIndex: idx! })
        } else {
          changedCount++
          pushBlock({
            kind: 'xml',
            xml: generateTableModelXml(model, original.originalXml ?? undefined),
          })
        }
      } else {
        changedCount++
        pushBlock({ kind: 'xml', xml: generateTableModelXml(model) })
      }
      continue
    }
    if (node.type === 'docProtected') {
      const idx = node.attrs?.docxIndex as number | null
      if (idx !== null && idx !== undefined && originalByIndex.has(idx) && !usedIndexes.has(idx)) {
        usedIndexes.add(idx)
        mapAnchor(idx)
        const original = originalByIndex.get(idx)!
        if (original.blockRevision && !revision && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: unwrapBlockRevisionXml(original.originalXml) })
          continue
        }
        const imagePatch = imagePatchOf(node, original)
        const tableTexts = tableTextsPatch(node, original)
        const textboxTexts = textboxParasPatch(node, original)
        const textboxSizes = textboxSizesPatch(node, original)
        const textboxStyles = textboxStylesPatch(node, original)
        const textboxOffsetX =
          node.attrs?.imageOffsetXEmu != null ? Number(node.attrs.imageOffsetXEmu) : undefined
        const textboxOffsetY =
          node.attrs?.imageOffsetYEmu != null ? Number(node.attrs.imageOffsetYEmu) : undefined
        const textboxPositionChanged =
          !!original.textboxes?.length &&
          textboxOffsetX !== undefined &&
          textboxOffsetY !== undefined &&
          (textboxOffsetX !== original.imageOffsetXEmu ||
            textboxOffsetY !== original.imageOffsetYEmu)
        const fieldText = fieldTextPatch(node, original)
        const formulaTokens = formulaTokensPatch(node, original)
        // chart data edits live in the chart's own part; the body block stays original
        const chartPatch = chartPatchOf(node, original)
        if (chartPatch) {
          changedCount++
          chartPatches.push(chartPatch)
        }
        // chart resize rewrites the body drawing's extent
        const chartDisplay = node.attrs?.chartDisplay as ChartDisplay | null
        const chartSize =
          chartDisplay?.widthPx &&
          chartDisplay.heightPx &&
          (chartDisplay.widthPx !== original.chartDisplay?.widthPx ||
            chartDisplay.heightPx !== original.chartDisplay?.heightPx)
            ? { w: chartDisplay.widthPx, h: chartDisplay.heightPx }
            : null
        const imageReplace =
          original.type === 'image'
            ? (node.attrs?.imageReplace as { base64: string; mime: string } | null)
            : null
        if ((imagePatch || imageReplace) && original.originalXml) {
          changedCount++
          let xml = imagePatch
            ? patchImageParagraphXml(original.originalXml, imagePatch)
            : original.originalXml
          if (imagePatch?.wrap !== undefined) {
            const posOffset =
              imagePatch.posOffsetX !== undefined && imagePatch.posOffsetY !== undefined
                ? { x: imagePatch.posOffsetX, y: imagePatch.posOffsetY }
                : undefined
            const marginAlign =
              posOffset === undefined && imagePatch.posH && imagePatch.posV
                ? { h: imagePatch.posH, v: imagePatch.posV }
                : undefined
            xml = applyImageWrap(xml, imagePatch.wrap, posOffset, marginAlign)
          } else if (
            imagePatch &&
            (imagePatch.posOffsetX !== undefined || imagePatch.posOffsetY !== undefined)
          ) {
            // posOffset changed without wrap change: patchImageParagraphXml already rewrote it
          }
          pushBlock({
            kind: 'xml',
            xml,
            ...(imageReplace
              ? {
                  replaceImage: {
                    base64: imageReplace.base64,
                    mime: imageReplace.mime as NewImage['mime'],
                  },
                }
              : {}),
          })
        } else if (tableTexts && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchTableCellTexts(original.originalXml, tableTexts) })
        } else if (
          (textboxTexts || textboxSizes || textboxStyles || textboxPositionChanged) &&
          original.originalXml
        ) {
          changedCount++
          let xml = original.originalXml
          if (textboxTexts) xml = patchTextboxParas(xml, textboxTexts)
          if (textboxSizes) xml = patchTextboxSizes(xml, textboxSizes)
          if (textboxStyles) xml = patchShapeStyles(xml, textboxStyles)
          if (textboxPositionChanged) {
            const wrap =
              (node.attrs?.imageWrap as ImageWrap | null) ?? original.imageWrap ?? 'square-left'
            xml = applyImageWrap(xml, wrap, { x: textboxOffsetX!, y: textboxOffsetY! })
          }
          pushBlock({ kind: 'xml', xml })
        } else if (fieldText && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchFieldParagraphXml(original.originalXml, fieldText) })
        } else if (formulaTokens && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchMathTokens(original.originalXml, formulaTokens) })
        } else if (chartSize && original.originalXml) {
          changedCount++
          pushBlock({
            kind: 'xml',
            xml: patchDrawingExtent(original.originalXml, chartSize.w, chartSize.h),
          })
        } else {
          pushBlock({ kind: 'original', docxIndex: idx })
        }
      } else if (idx !== null && idx !== undefined && originalByIndex.has(idx)) {
        // pasted copy: the anchor is consumed by the first occurrence, so clone from the block's own data
        const original = originalByIndex.get(idx)!
        const image = imageFromProtectedAttrs(node)
        if (image) {
          changedCount++
          pushBlock({ kind: 'image', image })
        } else if (original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: stripAnchorMarkers(original.originalXml) })
        } else {
          console.warn('dropping unclonable copy of protected block', node.attrs?.label)
        }
      } else if (node.attrs?.genXml) {
        // editor-created table or other self-contained fragment
        changedCount++
        let xml = String(node.attrs.genXml)
        const table = node.attrs?.table as TableModel | null
        if (table && node.attrs?.blockType === 'table') {
          xml = patchTableCellTexts(
            xml,
            table.rows.map((row) => row.map((cell) => cell.paras)),
          )
        }
        // editor-generated TOC lines: write the auto-refreshed page number back
        // (right only — the title is already in the genXml, and a left+right
        // patch bails out entirely when the title text nodes don't line up)
        const genField = node.attrs?.fieldDisplay as FieldDisplay | null
        if (genField?.kind === 'tocLine' && genField.right) {
          xml = patchFieldParagraphXml(xml, { right: genField.right })
        }
        // patch textbox paragraphs for newly-inserted textboxes/shapes with text
        const genTextboxes = node.attrs?.textboxes as TextboxDisplay[] | null
        const hasNonEmptyTextbox =
          genTextboxes != null &&
          genTextboxes.length > 0 &&
          genTextboxes.some((box) => box.paras.some((p) => p.runs.some((r) => r.text !== '')))
        if (hasNonEmptyTextbox) {
          xml = patchTextboxParas(
            xml,
            genTextboxes!.map((box) =>
              box.paras.map((p) => ({ runs: mergeRuns(p.runs), align: p.align ?? null })),
            ),
          )
        }
        // persist resize/autogrow of newly-inserted single-box shapes/lines;
        // horizontal lines keep their zero-height extent (display box is a grab band)
        const genBox = genTextboxes?.length === 1 ? genTextboxes[0] : null
        if (genBox && (genBox.widthPx || genBox.heightPx)) {
          xml = patchTextboxSizes(xml, [
            {
              wPx: genBox.widthPx ?? null,
              hPx: isStraightLineKind(genBox.prst) ? null : (genBox.heightPx ?? null),
            },
          ])
        }
        if (genBox) {
          xml = patchShapeStyles(xml, [
            { fillHex: genBox.fill ?? null, borderHex: genBox.borderColor ?? null },
          ])
        }
        // apply wrap changes for floating textboxes/shapes
        const genWrap = node.attrs?.imageWrap as ImageWrap | null
        const genOffsetX =
          node.attrs?.imageOffsetXEmu != null ? Number(node.attrs.imageOffsetXEmu) : undefined
        const genOffsetY =
          node.attrs?.imageOffsetYEmu != null ? Number(node.attrs.imageOffsetYEmu) : undefined
        if (genWrap !== undefined && genWrap !== null) {
          const posOffset =
            genOffsetX !== undefined && genOffsetY !== undefined
              ? { x: genOffsetX, y: genOffsetY }
              : undefined
          xml = applyImageWrap(xml, genWrap, posOffset)
        }
        pushBlock({ kind: 'xml', xml })
      } else if (node.attrs?.genImage) {
        changedCount++
        const image = { ...(node.attrs.genImage as NewImage) }
        if (node.attrs.imageWidthPx) image.widthPx = Number(node.attrs.imageWidthPx)
        if (node.attrs.imageHeightPx) image.heightPx = Number(node.attrs.imageHeightPx)
        const align = node.attrs.imageAlign as NewImage['align'] | null
        if (align) image.align = align
        const wrap = node.attrs.imageWrap as ImageWrap | null
        if (wrap) image.wrap = wrap
        if (node.attrs.imageRotDeg) image.rotDeg = Number(node.attrs.imageRotDeg)
        if (node.attrs.imageFlipH) image.flipH = true
        if (node.attrs.imageFlipV) image.flipV = true
        pushBlock({ kind: 'image', image })
      } else if (node.attrs?.genChart) {
        // in-place edits live in chartDisplay; the saved part reflects them
        changedCount++
        const spec = { ...(node.attrs.genChart as NewChart) }
        const display = node.attrs.chartDisplay as ChartDisplay | null
        if (display) {
          if (display.title !== undefined) spec.title = display.title
          spec.categories = display.categories
          spec.series = display.series.map((s, i) => ({
            name: s.name ?? spec.series[i]?.name ?? t('editorChartSeries', { num: i + 1 }),
            values: s.values,
          }))
        }
        pushBlock({
          kind: 'chart',
          chart: spec,
          ...(display?.widthPx && display.heightPx
            ? { extentPx: { w: display.widthPx, h: display.heightPx } }
            : {}),
        })
      }
      // protected node without an anchor or generated payload cannot be regenerated; drop it
      continue
    }

    const generated = pmNodeToGeneratedBlock(node)
    const idx = node.attrs?.docxIndex as number | null
    const original = idx !== null && idx !== undefined ? originalByIndex.get(idx) : undefined
    if (
      original &&
      !usedIndexes.has(idx!) &&
      signatureOfBlock(original) === signatureOfGenerated(generated)
    ) {
      usedIndexes.add(idx!)
      mapAnchor(idx!)
      pushBlock({ kind: 'original', docxIndex: idx! })
    } else {
      changedCount++
      // an in-place edit consumes its original anchor: it is replaced, not
      // deleted — and may reuse the anchor's original pPr bytes. A split twin
      // (anchor already consumed) gets no rawPPr, so revision records and
      // section props are never duplicated.
      if (original && !usedIndexes.has(idx!)) {
        usedIndexes.add(idx!)
        mapAnchor(idx!)
        applyRawPPr(generated, original)
      } else if (original) {
        // A split twin inherits the source node's attrs, but pPrChange and
        // bookmark/comment anchors belong only to the anchored original.
        delete generated.pPrChange
        delete generated.bookmarks
        delete generated.hiddenBookmarks
        delete generated.commentStarts
        delete generated.commentEnds
        // A grouped shell (partial open/close of a multi-block sdt) must be
        // emitted exactly once, by the anchor.
        if (generated.sdtShell?.group !== undefined) delete generated.sdtShell
      }
      pushBlock({ kind: 'generated', block: generated })
    }
  }

  const attachShell = (emit: { at: number; idx: number }, side: 'open' | 'close') => {
    const group = sdtGroupOf.get(emit.idx)!
    const shell = sdtGroupShells.get(group)!
    const add = side === 'open' ? shell.openXml : shell.closeXml
    const sb = saveBlocks[emit.at]
    if (sb.kind === 'original') {
      const xml = originalByIndex.get(emit.idx)?.originalXml
      if (xml === undefined) return
      saveBlocks[emit.at] = {
        kind: 'xml',
        xml: side === 'open' ? add + xml : xml + add,
        docxIndex: emit.idx,
        ...(sb.revision ? { revision: sb.revision } : {}),
      }
    } else if (sb.kind === 'generated') {
      const gShell = sb.block.sdtShell ?? { ...shell, openXml: '', closeXml: '' }
      if (side === 'open') gShell.openXml = add
      else gShell.closeXml = add
      sb.block.sdtShell = gShell
    } else if (sb.kind === 'xml') {
      sb.xml = side === 'open' ? add + sb.xml : sb.xml + add
    }
  }
  for (const [group, emits] of sdtGroupEmits) {
    const shell = sdtGroupShells.get(group)!
    if (shell.openXml && !emits.some((e) => e.open)) attachShell(emits[0], 'open')
    if (shell.closeXml && !emits.some((e) => e.close)) attachShell(emits[emits.length - 1], 'close')
  }

  const deletedCount = [...originalByIndex.keys()].filter((i) => !usedIndexes.has(i)).length
  return {
    saveBlocks,
    chartPatches,
    saveBlockIndexByDocx,
    changedCount,
    deletedCount,
    totalOriginal: originalByIndex.size,
  }
}

/** rebuild a pasted image copy from its preview bytes; null when not materializable */
function imageFromProtectedAttrs(node: PmNode): NewImage | null {
  if (node.attrs?.blockType !== 'image') return null
  const m = /^data:(image\/(?:png|jpeg|gif));base64,(.+)$/.exec(
    String(node.attrs?.imageDataUrl ?? ''),
  )
  const widthPx = Number(node.attrs?.imageWidthPx)
  const heightPx = Number(node.attrs?.imageHeightPx)
  if (!m || !widthPx || !heightPx) return null
  const image: NewImage = { base64: m[2], mime: m[1] as NewImage['mime'], widthPx, heightPx }
  const align = node.attrs?.imageAlign as NewImage['align'] | null
  if (align) image.align = align
  const wrap = node.attrs?.imageWrap as ImageWrap | null
  if (wrap) image.wrap = wrap
  if (node.attrs?.imageRotDeg) image.rotDeg = Number(node.attrs.imageRotDeg)
  if (node.attrs?.imageFlipH) image.flipH = true
  if (node.attrs?.imageFlipV) image.flipV = true
  return image
}

/** cloned XML must not repeat the anchor's bookmark/comment ids */
function stripAnchorMarkers(xml: string): string {
  return xml.replace(
    /<w:(?:bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\/>/g,
    '',
  )
}

/** chart data changes vs the parsed model; null when untouched or unpatchable */
function chartPatchOf(
  node: PmNode,
  original: Block,
): { partPath: string; patch: ChartPatch } | null {
  const current = node.attrs?.chartDisplay as ChartDisplay | null
  const initial = original.chartDisplay
  if (!current || !initial || current.partPath !== initial.partPath) return null
  if (
    current.series.length !== initial.series.length ||
    current.categories.length !== initial.categories.length
  ) {
    return null
  }
  const patch: ChartPatch = {}
  if (initial.title !== undefined && (current.title ?? '') !== initial.title) {
    patch.title = current.title ?? ''
  }
  let catChanged = false
  const categories = current.categories.map((cat, i) => {
    if (cat === initial.categories[i]) return null
    catChanged = true
    return cat
  })
  if (catChanged) patch.categories = categories
  let serChanged = false
  const series = current.series.map((ser, i): ChartSeriesPatch | null => {
    const orig = initial.series[i]
    if (ser.values.length !== orig.values.length) return null
    const serPatch: ChartSeriesPatch = {}
    if (orig.name !== undefined && (ser.name ?? '') !== orig.name) serPatch.name = ser.name ?? ''
    let valChanged = false
    const values = ser.values.map((value, j) => {
      // gaps in the cache have no pt to patch; they stay read-only
      if (value === orig.values[j] || orig.values[j] === null) return null
      valChanged = true
      return value
    })
    if (valChanged) serPatch.values = values
    if (Object.keys(serPatch).length === 0) return null
    serChanged = true
    return serPatch
  })
  if (serChanged) patch.series = series
  if (Object.keys(patch).length === 0) return null
  return { partPath: initial.partPath, patch }
}

interface ImageBlockPatch {
  widthPx?: number
  heightPx?: number
  align?: 'left' | 'center' | 'right' | null
  /** null = back to inline; undefined = wrap unchanged */
  wrap?: ImageWrap | null
  /** posOffset in EMU for free-position floating images; undefined = keep */
  posOffsetX?: number
  posOffsetY?: number
  /** margin-relative align pair (Word position-gallery presets); undefined = keep */
  posH?: 'left' | 'center' | 'right'
  posV?: 'top' | 'center' | 'bottom'
  /** rotation (deg clockwise, 0 removes) / mirror flips; undefined = keep */
  rotDeg?: number
  flipH?: boolean
  flipV?: boolean
}

/** size/align/wrap changes on an original image block; null when untouched */
function imagePatchOf(node: PmNode, original: Block): ImageBlockPatch | null {
  if (original.type !== 'image' || node.attrs?.blockType !== 'image') return null
  const patch: ImageBlockPatch = {}
  const w = node.attrs?.imageWidthPx ? Number(node.attrs.imageWidthPx) : null
  const h = node.attrs?.imageHeightPx ? Number(node.attrs.imageHeightPx) : null
  if (w && h && (w !== (original.imageWidthPx ?? null) || h !== (original.imageHeightPx ?? null))) {
    patch.widthPx = w
    patch.heightPx = h
  }
  const align = (node.attrs?.imageAlign as 'left' | 'center' | 'right' | null) ?? null
  if (align !== (original.imageAlign ?? null)) {
    patch.align = align
  }
  const wrap = (node.attrs?.imageWrap as ImageWrap | null) ?? null
  if (wrap !== (original.imageWrap ?? null)) {
    patch.wrap = wrap
  }
  // posOffset for free-position floating images
  const posOffsetX =
    node.attrs?.imageOffsetXEmu != null ? Number(node.attrs.imageOffsetXEmu) : undefined
  const posOffsetY =
    node.attrs?.imageOffsetYEmu != null ? Number(node.attrs.imageOffsetYEmu) : undefined
  if (posOffsetX !== undefined && posOffsetX !== (original.imageOffsetXEmu ?? undefined)) {
    patch.posOffsetX = posOffsetX
  }
  if (posOffsetY !== undefined && posOffsetY !== (original.imageOffsetYEmu ?? undefined)) {
    patch.posOffsetY = posOffsetY
  }
  const rotDeg = node.attrs?.imageRotDeg != null ? Number(node.attrs.imageRotDeg) : 0
  if (rotDeg !== (original.imageRotDeg ?? 0)) patch.rotDeg = rotDeg
  const flipH = !!node.attrs?.imageFlipH
  const flipV = !!node.attrs?.imageFlipV
  if (flipH !== (original.imageFlipH ?? false)) patch.flipH = flipH
  if (flipV !== (original.imageFlipV ?? false)) patch.flipV = flipV
  const posH = (node.attrs?.imagePosH as ImageBlockPatch['posH'] | null) ?? null
  const posV = (node.attrs?.imagePosV as ImageBlockPatch['posV'] | null) ?? null
  if (
    posH &&
    posV &&
    (posH !== (original.imagePosH ?? null) || posV !== (original.imagePosV ?? null))
  ) {
    patch.posH = posH
    patch.posV = posV
    // the position pair is written by applyImageWrap, which only runs when wrap is in the patch
    if (patch.wrap === undefined) patch.wrap = wrap
  }
  return Object.keys(patch).length > 0 ? patch : null
}

/** per-cell text changes on an original table block; null when untouched */
function tableTextsPatch(node: PmNode, original: Block): (string[] | null)[][] | null {
  if (original.type !== 'table' || node.attrs?.blockType !== 'table') return null
  const current = node.attrs?.table as TableModel | null
  const orig = original.table
  if (!current || !orig) return null
  let changed = false
  const texts = current.rows.map((row, r) =>
    row.map((cell, c) => {
      const originalCell = orig.rows[r]?.[c]
      if (!originalCell) return null
      if (cell.paras.join('\n') === originalCell.paras.join('\n')) return null
      changed = true
      return cell.paras
    }),
  )
  return changed ? texts : null
}

function tableTextsPatchFromModel(
  current: TableModel,
  original: Block,
): (CellTextsPatch | null)[][] | null {
  if (original.type !== 'table' || !original.table) return null
  let changed = false
  const texts = current.rows.map((row, r) =>
    row.map((cell, c): CellTextsPatch | null => {
      const originalCell = original.table!.rows[r]?.[c]
      if (!originalCell) return null
      // nested tables: diff cell texts per table; changes go through the recursive surgical patch
      if (cell.nestedTables?.length && originalCell.nestedTables?.length) {
        const nested = cell.nestedTables.map((nt, i) => {
          const origNt = originalCell.nestedTables![i]
          if (!origNt) return null
          const grid = nestedTextsDiff(nt, origNt)
          return grid
        })
        if (nested.some((g) => g !== null)) {
          changed = true
          return { nested }
        }
        return null
      }
      if (cell.paras.join('\n') === originalCell.paras.join('\n')) return null
      changed = true
      return cell.paras
    }),
  )
  return changed ? texts : null
}

/** per-cell text diff of one nested table; null = untouched */
function nestedTextsDiff(current: TableModel, original: TableModel): (string[] | null)[][] | null {
  if (current.rows.length !== original.rows.length) return null
  let changed = false
  const grid = current.rows.map((row, r) =>
    row.map((cell, c) => {
      const originalCell = original.rows[r]?.[c]
      if (!originalCell || cell.paras.join('\n') === originalCell.paras.join('\n')) return null
      changed = true
      return cell.paras
    }),
  )
  return changed ? grid : null
}

/** rich-run signature of one textbox paragraph, for per-paragraph change detection */
export function textboxParaSignature(para: TextboxDisplay['paras'][number]): string {
  return JSON.stringify([para.align ?? null, normalizedRuns(para.runs)])
}

/**
 * Per-box, per-paragraph rich-run patches on an original textbox block; null
 * when untouched. Unchanged paragraphs stay null so the engine keeps their
 * original bytes.
 */
function textboxParasPatch(
  node: PmNode,
  original: Block,
): ((TextboxParaPatch | null)[] | null)[] | null {
  const current = node.attrs?.textboxes as TextboxDisplay[] | null
  const orig = original.textboxes
  if (!current || !orig || current.length !== orig.length) return null
  let changed = false
  const boxes = current.map((box, i) => {
    const origParas = orig[i].paras
    const same =
      box.paras.length === origParas.length &&
      box.paras.every((p, j) => textboxParaSignature(p) === textboxParaSignature(origParas[j]))
    if (same) return null
    changed = true
    return box.paras.map((p, j) => {
      if (j < origParas.length && textboxParaSignature(p) === textboxParaSignature(origParas[j])) {
        return null
      }
      // align is always explicit so a removed alignment also clears w:jc
      return { runs: mergeRuns(p.runs), align: p.align ?? null }
    })
  })
  return changed ? boxes : null
}

function textboxSizesPatch(node: PmNode, original: Block): (TextboxSizePatch | null)[] | null {
  const current = node.attrs?.textboxes as TextboxDisplay[] | null
  const initial = original.textboxes
  if (!current || !initial || current.length !== initial.length) return null
  let changed = false
  const sizes = current.map((box, index) => {
    // a resize on an autofit box (no initial heightPx) pins its height:
    // patchTextboxSizes drops spAutoFit so Word honors the fixed extent
    const wPx = box.widthPx && box.widthPx !== initial[index].widthPx ? box.widthPx : null
    const hPx = box.heightPx && box.heightPx !== initial[index].heightPx ? box.heightPx : null
    if (wPx == null && hPx == null) return null
    changed = true
    return { wPx, hPx }
  })
  return changed ? sizes : null
}

function textboxStylesPatch(node: PmNode, original: Block): (ShapeStylePatch | null)[] | null {
  const current = node.attrs?.textboxes as TextboxDisplay[] | null
  const initial = original.textboxes
  if (!current || !initial || current.length !== initial.length) return null
  let changed = false
  const styles = current.map((box, index) => {
    const fillHex =
      (box.fill ?? null) !== (initial[index].fill ?? null) ? (box.fill ?? null) : undefined
    const borderHex =
      (box.borderColor ?? null) !== (initial[index].borderColor ?? null)
        ? (box.borderColor ?? null)
        : undefined
    if (fillHex === undefined && borderHex === undefined) return null
    changed = true
    return { fillHex, borderHex }
  })
  return changed ? styles : null
}

function fieldTextPatch(node: PmNode, original: Block): FieldTextPatch | null {
  const current = node.attrs?.fieldDisplay as FieldDisplay | null
  const initial = original.fieldDisplay
  if (!current || !initial || current.kind !== initial.kind) return null
  const patch: FieldTextPatch = {}
  if ((current.left ?? '') !== (initial.left ?? '')) patch.left = current.left ?? ''
  if (current.kind === 'tocLine' && (current.right ?? '') !== (initial.right ?? '')) {
    patch.right = current.right ?? ''
  }
  return Object.keys(patch).length > 0 ? patch : null
}

function formulaTokensPatch(node: PmNode, original: Block): string[] | null {
  const current = node.attrs?.formulaDisplay as FormulaDisplay | null
  const initial = original.formulaDisplay
  if (!current || !initial || current.tokens.length !== initial.tokens.length) return null
  return current.tokens.some((token, i) => token !== initial.tokens[i]) ? current.tokens : null
}

/**
 * High-fidelity pPr passthrough for in-place paragraph edits: text-only edits
 * reuse the original <w:pPr> bytes verbatim; format edits merge the six
 * format-model children into it (everything else keeps its bytes). Structure
 * changes (type / style / list) fall back to a full rebuild.
 */
function applyRawPPr(generated: GeneratedBlock, original: Block): void {
  if (original.rawPPr === undefined) return
  const structureSame =
    original.type === generated.type &&
    (original.styleId ?? null) === (generated.styleId ?? null) &&
    JSON.stringify(original.list ?? null) === JSON.stringify(generated.list ?? null)
  if (!structureSame) return
  const formatSame =
    JSON.stringify(normalizedFormat(original.format)) ===
    JSON.stringify(normalizedFormat(generated.format))
  let rawPPr = formatSame ? original.rawPPr : mergePPrFormat(original.rawPPr, generated.format)
  // Strip pPrChange when the user accepted/rejected it (pPrChange: null in generated block)
  if (generated.pPrChange === null && original.pPrChangeInfo !== undefined) {
    rawPPr = stripPPrChange(rawPPr)
  } else if (generated.pPrChange) {
    rawPPr = setPPrChange(rawPPr, generated.pPrChange)
  }
  generated.rawPPr = rawPPr
}

function nodeFormat(node: PmNode): ParaFormat | undefined {
  const format: ParaFormat = {}
  if (node.attrs?.align) format.align = node.attrs.align as ParaFormat['align']
  if (node.attrs?.lineSpacing) format.lineSpacing = Number(node.attrs.lineSpacing)
  if (node.attrs?.lineRule) format.lineRule = node.attrs.lineRule as ParaFormat['lineRule']
  if (node.attrs?.lineRawTwips) format.lineRawTwips = Number(node.attrs.lineRawTwips)
  if (node.attrs?.indentLeft) format.indentLeft = Number(node.attrs.indentLeft)
  if (node.attrs?.indentRight) format.indentRight = Number(node.attrs.indentRight)
  if (node.attrs?.indentFirstLine) format.indentFirstLine = Number(node.attrs.indentFirstLine)
  if (node.attrs?.spaceBefore != null) format.spaceBefore = Number(node.attrs.spaceBefore)
  if (node.attrs?.spaceAfter != null) format.spaceAfter = Number(node.attrs.spaceAfter)
  if (node.attrs?.pageBreakBefore) format.pageBreakBefore = true
  if (node.attrs?.bidi) format.bidi = true
  if (node.attrs?.autoSpace != null) format.autoSpace = node.attrs.autoSpace as boolean
  if (node.attrs?.shadingFill) format.shadingFill = String(node.attrs.shadingFill)
  if (node.attrs?.borders) format.borders = String(node.attrs.borders)
  if (node.attrs?.tabStops) {
    try {
      const parsed = JSON.parse(String(node.attrs.tabStops))
      if (Array.isArray(parsed) && parsed.length > 0) format.tabStops = parsed
    } catch {
      /* ignore malformed */
    }
  }
  if (node.attrs?.dropCap) {
    try {
      const parsed = JSON.parse(String(node.attrs.dropCap))
      if (parsed && typeof parsed === 'object') format.dropCap = parsed
    } catch {
      /* ignore malformed */
    }
  }
  return Object.keys(format).length > 0 ? format : undefined
}

export function pmNodeToGeneratedBlock(node: PmNode): GeneratedBlock {
  const runs = inlineToRuns(node.content ?? [])
  const format = nodeFormat(node)
  const bookmarks = nodeBookmarks(node, 'bookmarks')
  const hiddenBookmarks = nodeBookmarks(node, 'hiddenBookmarks')
  const commentStarts = nodeBookmarks(node, 'commentStarts')
  const commentEnds = nodeBookmarks(node, 'commentEnds')
  if (node.type === 'docHeading') {
    const pPrChange = node.attrs?.pPrChange as string | null | undefined
    const blockRevision = node.attrs?.blockRevision as GeneratedBlock['blockRevision']
    return {
      type: 'heading',
      level: Number(node.attrs?.level) || 1,
      styleId: (node.attrs?.styleId as string) ?? undefined,
      format,
      bookmarks,
      hiddenBookmarks,
      commentStarts,
      commentEnds,
      runs,
      ...(pPrChange !== undefined ? { pPrChange: pPrChange ?? null } : {}),
      ...(blockRevision !== undefined ? { blockRevision: blockRevision ?? null } : {}),
    }
  }
  if (node.type === 'docListItem') {
    const numId = node.attrs?.numId as string | null
    if (numId) {
      const pPrChange = node.attrs?.pPrChange as string | null | undefined
      const blockRevision = node.attrs?.blockRevision as GeneratedBlock['blockRevision']
      return {
        type: 'listItem',
        styleId: (node.attrs?.styleId as string) ?? undefined,
        list: {
          kind: (node.attrs?.kind as 'bullet' | 'ordered') ?? 'bullet',
          numId,
          ilvl: Number(node.attrs?.ilvl) || 0,
        },
        format,
        bookmarks,
        hiddenBookmarks,
        commentStarts,
        commentEnds,
        runs,
        ...(pPrChange !== undefined ? { pPrChange: pPrChange ?? null } : {}),
        ...(blockRevision !== undefined ? { blockRevision: blockRevision ?? null } : {}),
      }
    }
    // list item without a docx numbering id degrades to a plain paragraph with marker
    return {
      type: 'paragraph',
      format,
      bookmarks,
      hiddenBookmarks,
      commentStarts,
      commentEnds,
      runs: [{ text: '• ' }, ...runs],
    }
  }
  // Extract sdtShell from node attrs (serialized as JSON)
  const sdtShell = node.attrs?.sdtShell
    ? (() => {
        try {
          const parsed = JSON.parse(String(node.attrs.sdtShell))
          return parsed && typeof parsed === 'object' ? parsed : undefined
        } catch {
          return undefined
        }
      })()
    : undefined
  // pPrChange: propagate to GeneratedBlock so applyRawPPr can strip it when accepted
  const pPrChange = node.attrs?.pPrChange as string | null | undefined
  const blockRevision = node.attrs?.blockRevision as GeneratedBlock['blockRevision']
  return {
    type: 'paragraph',
    styleId: (node.attrs?.styleId as string) ?? undefined,
    format,
    bookmarks,
    hiddenBookmarks,
    commentStarts,
    commentEnds,
    runs,
    ...(sdtShell ? { sdtShell } : {}),
    // carry pPrChange through: null means "user accepted/rejected, strip from rawPPr"
    ...(pPrChange !== undefined ? { pPrChange: pPrChange ?? null } : {}),
    ...(blockRevision !== undefined ? { blockRevision: blockRevision ?? null } : {}),
  }
}

function nodeBookmarks(
  node: PmNode,
  attr: 'bookmarks' | 'hiddenBookmarks' | 'commentStarts' | 'commentEnds',
): string[] | undefined {
  const value = node.attrs?.[attr] as string[] | null | undefined
  return Array.isArray(value) && value.length > 0 ? value : undefined
}

export function inlineToRuns(content: PmNode[]): Run[] {
  const runs: Run[] = []
  for (const node of content) {
    if (node.type === 'hardBreak') {
      const ch = node.attrs?.pageBreak ? '\f' : '\n'
      const prev = runs[runs.length - 1]
      const prevAtomic =
        prev && (prev.noteRef || prev.xeTerm !== undefined || prev.math || prev.ruby || prev.image)
      if (prev && !prevAtomic) prev.text += ch
      else runs.push({ text: ch })
      continue
    }
    if (node.type === 'docNoteRef') {
      runs.push({
        text: String(node.attrs?.num ?? 1),
        noteRef: {
          kind: (node.attrs?.kind as 'footnote' | 'endnote') ?? 'footnote',
          id: String(node.attrs?.id ?? ''),
        },
      })
      continue
    }
    if (node.type === 'docXeMark') {
      runs.push({ text: '', xeTerm: String(node.attrs?.term ?? '') })
      continue
    }
    if (node.type === 'docInlineMath') {
      const omml = String(node.attrs?.omml ?? '')
      if (omml) runs.push({ text: String(node.attrs?.text ?? ''), math: { omml } })
      continue
    }
    if (node.type === 'docRuby') {
      const base = String(node.attrs?.base ?? '')
      const xml = String(node.attrs?.xml ?? '')
      if (base) {
        runs.push(
          xml ? { text: base, ruby: { rt: String(node.attrs?.rt ?? ''), xml } } : { text: base },
        )
      }
      continue
    }
    if (node.type === 'docInlineImage') {
      const dataUrl = String(node.attrs?.dataUrl ?? '')
      const xml = String(node.attrs?.xml ?? '')
      if (dataUrl && xml) {
        runs.push({
          text: '',
          image: {
            dataUrl,
            xml,
            ...(node.attrs?.widthPx ? { widthPx: Number(node.attrs.widthPx) } : {}),
            ...(node.attrs?.heightPx ? { heightPx: Number(node.attrs.heightPx) } : {}),
          },
        })
      }
      continue
    }
    if (node.type !== 'text' || !node.text) continue
    const run: Run = { text: node.text }
    for (const mark of node.marks ?? []) {
      if (mark.type === 'bold') run.bold = true
      else if (mark.type === 'italic') run.italic = true
      else if (mark.type === 'underline') run.underline = true
      else if (mark.type === 'strike') run.strike = true
      else if (mark.type === 'link') {
        run.link = {
          href: String(mark.attrs?.href ?? ''),
          rId: (mark.attrs?.rId as string) ?? undefined,
          tooltip: (mark.attrs?.tooltip as string) ?? undefined,
        }
      } else if (mark.type === 'refField') {
        run.refField = String(mark.attrs?.name ?? '')
      } else if (mark.type === 'instrField') {
        run.instrField = String(mark.attrs?.instr ?? '')
      } else if (mark.type === 'comment') {
        const ids = String(mark.attrs?.ids ?? '')
          .split(' ')
          .filter(Boolean)
        if (ids.length > 0) run.commentIds = ids
      } else if (mark.type === 'ins' || mark.type === 'del') {
        const info: NonNullable<Run['ins']> = { author: String(mark.attrs?.author ?? '') }
        if (mark.attrs?.date) info.date = String(mark.attrs.date)
        if (mark.attrs?.id) info.id = String(mark.attrs.id)
        if (mark.type === 'ins') run.ins = info
        else run.del = info
      } else if (mark.type === 'docTextStyle') {
        if (mark.attrs?.color) run.color = String(mark.attrs.color)
        if (mark.attrs?.sizeHalfPoints) run.sizeHalfPoints = Number(mark.attrs.sizeHalfPoints)
        if (mark.attrs?.font) run.font = String(mark.attrs.font)
        if (mark.attrs?.fontAscii) run.fontAscii = String(mark.attrs.fontAscii)
        if (mark.attrs?.csFont) run.csFont = String(mark.attrs.csFont)
        if (mark.attrs?.charSpacingTwips) run.charSpacingTwips = Number(mark.attrs.charSpacingTwips)
        if (mark.attrs?.highlight) run.highlight = String(mark.attrs.highlight)
        if (mark.attrs?.vertAlign === 'superscript' || mark.attrs?.vertAlign === 'subscript') {
          run.vertAlign = mark.attrs.vertAlign
        }
        if (mark.attrs?.em) run.em = mark.attrs.em as NonNullable<Run['em']>
        if (mark.attrs?.styleId) run.styleId = String(mark.attrs.styleId)
        if (mark.attrs?.rawRPr) run.rawRPr = String(mark.attrs.rawRPr)
      } else if (mark.type === 'rprChange') {
        run.rPrChange = {
          author: String(mark.attrs?.author ?? ''),
          ...(mark.attrs?.date ? { date: String(mark.attrs.date) } : {}),
          ...(mark.attrs?.id ? { id: String(mark.attrs.id) } : {}),
          ...(mark.attrs?.old
            ? { old: mark.attrs.old as NonNullable<Run['rPrChange']>['old'] }
            : {}),
        }
      }
    }
    runs.push(run)
  }
  return mergeRuns(runs)
}

function mergeRuns(runs: Run[]): Run[] {
  const merged: Run[] = []
  for (const run of runs) {
    const prev = merged[merged.length - 1]
    // reference markers / index entries / inline math are atomic and never merge
    const atomic =
      run.noteRef ||
      run.xeTerm !== undefined ||
      run.math ||
      run.ruby ||
      run.image ||
      prev?.noteRef ||
      prev?.xeTerm !== undefined ||
      prev?.math ||
      prev?.ruby ||
      prev?.image
    if (prev && !atomic && runStyleKey(prev) === runStyleKey(run)) prev.text += run.text
    else merged.push({ ...run })
  }
  return merged
}

function runStyleKey(run: Run): string {
  return JSON.stringify([
    run.rawRPr ?? null,
    run.styleId ?? null,
    !!run.bold,
    !!run.italic,
    !!run.underline,
    !!run.strike,
    run.color ?? null,
    run.sizeHalfPoints ?? null,
    run.font ?? null,
    run.fontAscii ?? null,
    run.highlight ?? null,
    run.vertAlign ?? null,
    run.link?.href ?? null,
    run.link?.rId ?? null,
    run.commentIds?.join(' ') ?? null,
    revisionKey(run),
    run.noteRef ? [run.noteRef.kind, run.noteRef.id] : null,
    run.xeTerm ?? null,
    run.refField ?? null,
    run.instrField ?? null,
    run.math?.omml ?? null,
    run.ruby?.xml ?? null,
  ])
}

function revisionKey(run: Run): string | null {
  if (!run.ins && !run.del) return null
  return JSON.stringify([
    run.ins?.author ?? null,
    run.ins?.date ?? null,
    run.ins?.id ?? null,
    run.del?.author ?? null,
    run.del?.date ?? null,
    run.del?.id ?? null,
  ])
}

// ---- signatures: "did the user actually change this block?" ----

function normalizedRuns(runs: Run[]): unknown[] {
  return mergeRuns(runs).map((r) => [
    r.text,
    r.rawRPr ?? null,
    r.styleId ?? null,
    !!r.bold,
    !!r.italic,
    !!r.underline,
    !!r.strike,
    r.color ?? null,
    r.sizeHalfPoints ?? null,
    r.font ?? null,
    r.fontAscii ?? null,
    r.highlight ?? null,
    r.vertAlign ?? null,
    r.link?.href ?? null,
    r.commentIds?.join(' ') ?? null,
    revisionKey(r),
    r.noteRef ? [r.noteRef.kind, r.noteRef.id] : null,
    r.xeTerm ?? null,
    r.refField ?? null,
    r.instrField ?? null,
    r.math?.omml ?? null,
    r.ruby?.xml ?? null,
  ])
}

function normalizedFormat(format: ParaFormat | undefined): unknown {
  if (!format) return null
  return [
    format.align ?? null,
    format.lineSpacing ?? null,
    format.lineRule ?? null,
    format.lineRawTwips ?? null,
    format.bidi ?? false,
    format.indentLeft ?? null,
    format.indentRight ?? null,
    format.indentFirstLine ?? null,
    format.spaceBefore ?? null,
    format.spaceAfter ?? null,
    format.pageBreakBefore ?? false,
    format.shadingFill ?? null,
    format.borders ?? null,
    format.tabStops ? JSON.stringify(format.tabStops) : null,
    format.dropCap ? JSON.stringify(format.dropCap) : null,
    format.autoSpace ?? null,
  ]
}

export function signatureOfBlock(block: Block): string {
  return JSON.stringify({
    type: block.type,
    level: block.level ?? null,
    styleId: block.styleId ?? null,
    list: block.list ? [block.list.kind, block.list.numId, block.list.ilvl] : null,
    format: normalizedFormat(block.format),
    bookmarks: block.bookmarks ?? null,
    hiddenBookmarks: block.hiddenBookmarks ?? null,
    commentStarts: block.commentStarts ?? null,
    commentEnds: block.commentEnds ?? null,
    runs: normalizedRuns(block.runs ?? []),
    // include pPrChangeInfo in signature: present = has pPrChange, absent = clean
    pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
    blockRevision: block.blockRevision ?? null,
  })
}

export function signatureOfGenerated(block: GeneratedBlock): string {
  return JSON.stringify({
    type: block.type,
    level: block.level ?? null,
    styleId: block.styleId ?? null,
    list: block.list ? [block.list.kind, block.list.numId, block.list.ilvl] : null,
    format: normalizedFormat(block.format),
    bookmarks: block.bookmarks ?? null,
    hiddenBookmarks: block.hiddenBookmarks ?? null,
    commentStarts: block.commentStarts ?? null,
    commentEnds: block.commentEnds ?? null,
    runs: normalizedRuns(block.runs),
    // include pPrChange in signature: null (accepted/rejected) != non-null (has pPrChange)
    pPrChange: block.pPrChange ?? null,
    blockRevision: block.blockRevision ?? null,
  })
}
