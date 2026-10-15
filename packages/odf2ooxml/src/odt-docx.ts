/**
 * .odt → .docx import filter (Tier 1). The ODF body is walked into
 * docx-engine SaveBlocks and rebuilt as a fresh .docx — the same
 * buildBlankDocx → parseDocx → saveDocx assembly pdf2docx uses, so the output
 * is an ordinary editable Word file, not a byte-preserving edit (ODF has no
 * editable representation here at all).
 *
 * Scope of the first cut, chosen by the same "prove it or drop it" rule the
 * .ods formula translator follows: headings, paragraphs, character formatting,
 * hyperlinks, lists, simple tables (with spans and fills), inline and framed
 * images, and the page geometry. Footnotes, comments, shapes, and unrecognized
 * blocks are counted into `dropped` instead of guessed at.
 */

import {
  buildBlankDocx,
  generateParagraphXml,
  parseDocx,
  saveDocx,
  type CustomNumberingLevel,
  type GeneratedBlock,
  type NewImage,
  type ParaFormat,
  type Run,
  type SaveBlock,
  type SectionSettings,
} from '../../docx-engine/src/index'
import { intrinsicImageSize } from './images'
import { OdfConvertError, childElements, loadOdfPackage, odfBody, type OdfPackage } from './odf'
import { OdfStyles } from './styles'
import {
  Drops,
  firstDescendant,
  imageRefOfFrame,
  runsFromParagraph,
  type InlineImageRef,
  type NeutralRun,
  type WalkContext,
} from './text'
import { emuToTwips, sniffImageMime } from './units'
import { XNode, attr, firstChild, nodeChildren, nodeName } from './xml'

// The error class lives with the package loader, which both converters reach
// for; re-exported here so the public entry point is unchanged.
export { OdfConvertError }

export interface OdtToDocxResult {
  docx: Uint8Array
  blocks: number
  /** constructs the filter could not represent, with how often each occurred */
  dropped: Array<{ what: string; count: number }>
}

type Draft =
  | { kind: 'generated'; block: GeneratedBlock }
  | { kind: 'xml'; xml: string }
  | { kind: 'image'; ref: InlineImageRef }

interface ListFrame {
  /** w:numId values are integers in OOXML; keep them clear of template ids */
  numId: string
  styleName?: string
  /** deepest 1-based level any item used */
  maxDepth: number
}

interface BlockContext {
  list: ListFrame | null
  /** 1-based depth inside the current list (0 outside one) */
  depth: number
}

const HEADING_MAX = 6

export async function convertOdtToDocx(odt: Uint8Array): Promise<OdtToDocxResult> {
  const pack = await loadOdfPackage(odt)
  if (!pack) throw new OdfConvertError('unreadable')
  const body = odfBody(pack, 'office:text')
  if (!body) throw new OdfConvertError('not-a-text-document')

  const styles = new OdfStyles(pack.contentRoots, pack.stylesRoots)
  const drops = new Drops()
  const ctx: WalkContext = { styles, families: ['text'], drops, images: [] }
  const drafts: Draft[] = []
  const listFrames: ListFrame[] = []
  let listCounter = 0

  const walkElements = (nodes: XNode[], context: BlockContext): void => {
    for (const element of elements(nodes)) {
      switch (element.name) {
        case 'text:h':
          emitParagraph(element.node, ctx, drafts, context, {
            type: 'heading',
            level: clampHeadingLevel(attr(element.node, 'text:outline-level')),
          })
          break
        case 'text:p':
          emitParagraph(element.node, ctx, drafts, context, { type: 'paragraph' })
          break
        case 'text:list':
          walkList(element.node, null, 0, context)
          break
        case 'table:table':
          drafts.push(tableDraft(element.node, ctx))
          break
        case 'text:section':
          if (attr(element.node, 'text:display') !== 'none') {
            walkElements(nodeChildren(element.node), context)
          }
          break
        case 'text:table-of-content':
        case 'text:illustration-index':
        case 'text:table-index':
        case 'text:object-index':
        case 'text:user-index':
        case 'text:alphabetical-index': {
          const indexBody = firstChild(element.node, 'text:index-body')
          if (indexBody) walkElements(nodeChildren(indexBody), context)
          else drops.add('stale-indexes')
          break
        }
        case 'text:sequence-decls':
        case 'text:soft-page-break':
          break
        case 'draw:frame':
          emitFrame(element.node, context)
          break
        case 'draw:rect':
        case 'draw:line':
        case 'draw:circle':
        case 'draw:ellipse':
        case 'draw:custom-shape':
        case 'draw:connector':
        case 'draw:control':
        case 'draw:object':
        case 'math:math':
          drops.add('shapes')
          break
        default:
          drops.add('unhandled-blocks')
          break
      }
    }
  }

  // -- lists ----------------------------------------------------------------

  function walkList(
    list: XNode,
    parent: ListFrame | null,
    parentDepth: number,
    context: BlockContext,
  ): void {
    const frame: ListFrame = parent ?? {
      numId: String(5000 + ++listCounter),
      styleName: attr(list, 'text:style-name'),
      maxDepth: 0,
    }
    if (!parent) listFrames.push(frame)
    const depth = parent ? parentDepth + 1 : 1
    frame.maxDepth = Math.max(frame.maxDepth, depth)
    for (const { name, node } of childElements(list)) {
      if (name === 'text:list-item') {
        for (const inner of childElements(node)) {
          if (inner.name === 'text:p') {
            emitParagraph(inner.node, ctx, drafts, { list: frame, depth }, { type: 'paragraph' })
          } else if (inner.name === 'text:h') {
            emitParagraph(
              inner.node,
              ctx,
              drafts,
              { list: frame, depth },
              {
                type: 'heading',
                level: clampHeadingLevel(attr(inner.node, 'text:outline-level')),
              },
            )
          } else if (inner.name === 'text:list') {
            walkList(inner.node, frame, depth, context)
          } else if (inner.name === 'table:table') {
            drafts.push(tableDraft(inner.node, ctx))
          } else if (inner.name === 'draw:frame') {
            emitFrame(inner.node, context)
          }
        }
      } else if (name === 'text:list-header') {
        // an untitled lead-in paragraph: keep the text without a marker
        for (const inner of childElements(node)) {
          if (inner.name === 'text:p') {
            emitParagraph(inner.node, ctx, drafts, { list: null, depth }, { type: 'paragraph' })
          }
        }
      }
    }
  }

  // -- frames at block level --------------------------------------------------

  function emitFrame(frame: XNode, context: BlockContext): void {
    const textBox = firstDescendant(frame, 'draw:text-box')
    if (textBox) {
      // a floating text frame flattens into the flow: the text survives, the
      // box geometry does not
      drops.add('textboxes-flattened')
      walkElements(nodeChildren(textBox), context)
      return
    }
    const ref = imageRefOfFrame(frame)
    if (ref) {
      drafts.push({ kind: 'image', ref })
      return
    }
    drops.add('shapes')
  }

  // -- paragraphs -------------------------------------------------------------

  function emitParagraph(
    node: XNode,
    walkCtx: WalkContext,
    out: Draft[],
    blockContext: BlockContext,
    base: { type: 'paragraph' | 'heading'; level?: number },
  ): void {
    const imagesBefore = walkCtx.images.length
    const runs = runsFromParagraph(node, walkCtx)
    const inlineImages = walkCtx.images.slice(imagesBefore)
    // these belong to this paragraph alone; leaving them behind would have the
    // next table cell report them as its own dropped images
    walkCtx.images.length = imagesBefore
    const block: GeneratedBlock = { type: base.type, runs: runs.map((run) => toDocxRun(run, true)) }
    if (base.type === 'heading') block.level = base.level
    const format = paraFormat(walkCtx.styles, attr(node, 'text:style-name'))
    if (format) block.format = format
    if (blockContext.list) {
      const frame = blockContext.list
      if (base.type === 'heading') block.type = 'heading'
      else block.type = 'listItem'
      block.list = {
        kind: listLevelKind(walkCtx.styles, frame, blockContext.depth),
        numId: frame.numId,
        ilvl: blockContext.depth - 1,
      }
    }
    out.push({ kind: 'generated', block })
    for (const ref of inlineImages) out.push({ kind: 'image', ref })
  }

  walkElements(nodeChildren(body), { list: null, depth: 0 })

  // -- assembly ---------------------------------------------------------------

  const numbering = numberingDefs(listFrames, ctx)
  const page = styles.pageSetup(pack.stylesRoots)
  const section: SectionSettings | undefined = page
    ? {
        pageWidth: page.widthTw,
        pageHeight: page.heightTw,
        orientation: page.widthTw >= page.heightTw ? 'landscape' : 'portrait',
        marginTop: page.marginTopTw,
        marginRight: page.marginRightTw,
        marginBottom: page.marginBottomTw,
        marginLeft: page.marginLeftTw,
        pageBorder: false,
        columns: 1,
      }
    : undefined

  const blocks: SaveBlock[] = []
  for (const draft of drafts) {
    if (draft.kind === 'generated') blocks.push({ kind: 'generated', block: draft.block })
    else if (draft.kind === 'xml') blocks.push({ kind: 'xml', xml: draft.xml })
    else {
      const image = await resolveImage(draft.ref, pack)
      if (image) blocks.push({ kind: 'image', image })
      else drops.add('images')
    }
  }
  // a document that ends on a table needs a paragraph after it (Word merges
  // two adjacent tables and swallows a trailing one into the sectPr otherwise)
  if (blocks.length > 0 && blocks[blocks.length - 1].kind === 'xml') {
    blocks.push({ kind: 'generated', block: { type: 'paragraph', runs: [] } })
  }

  const parsed = await parseDocx(await buildBlankDocx())
  const docx = await saveDocx(parsed, blocks, {
    ...(section ? { section } : {}),
    ...(numbering.length > 0 ? { numbering: { newDefs: numbering } } : {}),
  })
  return { docx, blocks: blocks.length, dropped: drops.list() }
}

// -- runs and formatting ----------------------------------------------------

/** element children of a node list, skipping text and comments — the array form of odf.ts's childElements */
function elements(nodes: XNode[]): Array<{ name: string; node: XNode }> {
  const out: Array<{ name: string; node: XNode }> = []
  for (const child of nodes) {
    const name = nodeName(child)
    if (name && name !== '#comment') out.push({ name, node: child })
  }
  return out
}

function clampHeadingLevel(text: string | undefined): number {
  const level = Number(text ?? '1')
  if (!Number.isInteger(level) || level < 1) return 1
  return Math.min(level, HEADING_MAX)
}

function paraFormat(styles: OdfStyles, styleName: string | undefined): ParaFormat | undefined {
  const props = styles.paraProps(styleName, ['paragraph'])
  if (!props) return undefined
  const out: ParaFormat = {}
  if (props.align) out.align = props.align
  if (props.spaceBeforeTw !== undefined) out.spaceBefore = props.spaceBeforeTw
  if (props.spaceAfterTw !== undefined) out.spaceAfter = props.spaceAfterTw
  if (props.lineSpacing !== undefined && props.lineSpacing > 0) out.lineSpacing = props.lineSpacing
  if (props.lineAtLeastTw !== undefined) {
    out.lineRule = 'atLeast'
    out.lineRawTwips = props.lineAtLeastTw
  } else if (props.lineExactTw !== undefined) {
    out.lineRule = 'exact'
    out.lineRawTwips = props.lineExactTw
  }
  if (props.indentLeftTw !== undefined) out.indentLeft = props.indentLeftTw
  if (props.indentRightTw !== undefined) out.indentRight = props.indentRightTw
  if (props.indentFirstLineTw !== undefined) out.indentFirstLine = props.indentFirstLineTw
  if (props.pageBreakBefore) out.pageBreakBefore = true
  return Object.keys(out).length > 0 ? out : undefined
}

function toDocxRun(run: NeutralRun, allowLinks: boolean): Run {
  const out: Run = { text: run.text }
  if (run.bold) out.bold = true
  if (run.italic) out.italic = true
  if (run.underline) out.underline = true
  if (run.strike) out.strike = true
  if (run.sizePt !== undefined) out.sizeHalfPoints = Math.round(run.sizePt * 2)
  if (run.color) out.color = run.color
  if (run.font) {
    out.font = run.font
    out.fontAscii = run.font
  }
  if (run.vertAlign) out.vertAlign = run.vertAlign
  if (run.caps) out.caps = run.caps
  if (run.shading) out.shading = run.shading
  if (run.href && allowLinks) out.link = { href: run.href }
  return out
}

const NUM_FORMATS: Record<string, string> = {
  '1': 'decimal',
  a: 'lowerLetter',
  A: 'upperLetter',
  i: 'lowerRoman',
  I: 'upperRoman',
}

function listLevelKind(styles: OdfStyles, frame: ListFrame, depth: number): 'bullet' | 'ordered' {
  const level = styles.listLevels(frame.styleName)?.get(depth)
  return level?.kind === 'ordered' ? 'ordered' : 'bullet'
}

/** one abstractNum per list actually used, covering every level it reached */
function numberingDefs(
  frames: ListFrame[],
  ctx: WalkContext,
): Array<{ numId: string; kind: 'bullet' | 'ordered'; levels?: CustomNumberingLevel[] }> {
  const out: Array<{ numId: string; kind: 'bullet' | 'ordered'; levels?: CustomNumberingLevel[] }> =
    []
  for (const frame of frames) {
    if (frame.maxDepth < 1) continue
    const levels = ctx.styles.listLevels(frame.styleName)
    const defs: CustomNumberingLevel[] = []
    let kind: 'bullet' | 'ordered' = 'bullet'
    for (let depth = 1; depth <= frame.maxDepth; depth++) {
      const level = levels?.get(depth)
      const indentLeft = level?.indentLeftTw ?? 720 * depth
      const hanging = level?.hangingTw ?? 360
      if (level?.kind === 'ordered') {
        if (depth === 1) kind = 'ordered'
        const format = NUM_FORMATS[level.numFormat ?? '1'] ?? 'decimal'
        defs.push({
          numFmt: format,
          lvlText: `${level.numPrefix ?? ''}%${depth}${level.numSuffix ?? '.'}`,
          indentLeft,
          hanging,
        })
      } else {
        defs.push({ numFmt: 'bullet', lvlText: level?.bulletChar ?? '•', indentLeft, hanging })
      }
    }
    out.push({ numId: frame.numId, kind, levels: defs.length > 0 ? defs : undefined })
  }
  return out
}

// -- tables -----------------------------------------------------------------

interface ColumnSpec {
  widthEmu: number | null
}

interface GridCell {
  row: number
  column: number
  colSpan: number
  rowSpan: number
  /** built paragraph XML (links stripped: relationships live with generated blocks only) */
  contentXml: string
  fill: string | null
}

const CELL_GEN_CTX = {
  headingStyleIds: new Map(
    Array.from({ length: HEADING_MAX }, (_, index) => [index + 1, `Heading${index + 1}`] as const),
  ),
  allocateHyperlinkRel: (href: string): string => {
    throw new Error(`hyperlink reached table cell XML: ${href}`)
  },
}

function tableDraft(node: XNode, ctx: WalkContext): Draft {
  const columns = readColumns(node, ctx)
  const rows = readRows(node)
  const grid = placeCells(rows, columns.length, ctx)
  return { kind: 'xml', xml: tableXml(columns, grid) }
}

function readColumns(table: XNode, ctx: WalkContext): ColumnSpec[] {
  const out: ColumnSpec[] = []
  const push = (node: XNode): void => {
    const width = ctx.styles.columnWidthEmu(attr(node, 'table:style-name'))
    const repeat = clampRepeat(attr(node, 'table:number-columns-repeated'))
    for (let index = 0; index < repeat; index++) out.push({ widthEmu: width })
  }
  const walk = (node: XNode): void => {
    for (const { name, node: child } of childElements(node)) {
      if (name === 'table:table-column') push(child)
      else if (name === 'table:table-header-columns' || name === 'table:table-columns') walk(child)
    }
  }
  walk(table)
  // a grid needs at least one column or the row parser cannot place cells
  if (out.length === 0) out.push({ widthEmu: null })
  return out
}

function readRows(table: XNode): XNode[] {
  const out: XNode[] = []
  const walk = (node: XNode): void => {
    for (const { name, node: child } of childElements(node)) {
      if (name === 'table:table-row') out.push(child)
      else if (name === 'table:table-header-rows' || name === 'table:table-rows') walk(child)
    }
  }
  walk(table)
  return out
}

function clampRepeat(text: string | undefined, cap = 1024): number {
  const value = Number(text ?? '1')
  if (!Number.isInteger(value) || value < 1) return 1
  return Math.min(value, cap)
}

/** resolve the occupancy grid: every (row, column) slot names the cell that paints it */
function placeCells(
  rows: XNode[],
  columnCount: number,
  ctx: WalkContext,
): Array<Array<GridCell | null>> {
  const occupancy: Array<Array<GridCell | null>> = []
  const cellsOfRow = (row: XNode): GridCell[] => {
    const out: GridCell[] = []
    for (const { name, node } of childElements(row)) {
      if (name === 'table:table-cell') {
        const repeat = clampRepeat(attr(node, 'table:number-columns-repeated'), 64)
        for (let index = 0; index < repeat; index++) out.push(makeCell(node, ctx))
      } else if (name === 'table:covered-table-cell') {
        const repeat = clampRepeat(attr(node, 'table:number-columns-repeated'), 64)
        for (let index = 0; index < repeat; index++) {
          out.push({
            row: -1,
            column: -1,
            colSpan: 1,
            rowSpan: 1,
            contentXml: '<w:p/>',
            fill: null,
          })
        }
      }
    }
    return out
  }
  let rowIndex = 0
  for (const row of rows) {
    while (occupancy.length <= rowIndex) occupancy.push(new Array(columnCount).fill(null))
    const slotRow = occupancy[rowIndex]
    let column = 0
    for (const cell of cellsOfRow(row)) {
      while (column < columnCount && slotRow[column]) column++
      if (column >= columnCount) break
      const colSpan = Math.max(1, Math.min(cell.colSpan, columnCount - column))
      for (let dr = 0; dr < cell.rowSpan; dr++) {
        while (occupancy.length <= rowIndex + dr) occupancy.push(new Array(columnCount).fill(null))
        const slot = occupancy[rowIndex + dr]
        for (let dc = 0; dc < colSpan; dc++) slot[column + dc] = cell
      }
      cell.row = rowIndex
      cell.column = column
      cell.colSpan = colSpan
      column += colSpan
    }
    rowIndex++
  }
  return occupancy
}

function makeCell(node: XNode, ctx: WalkContext): GridCell {
  const blocks: GeneratedBlock[] = []
  const emitInto = (
    child: XNode,
    base: { type: 'paragraph' | 'heading'; level?: number },
  ): void => {
    const block: GeneratedBlock = {
      type: base.type,
      runs: runsFromParagraph(child, ctx).map((run) => toDocxRun(run, false)),
    }
    if (base.type === 'heading') block.level = base.level
    const format = paraFormat(ctx.styles, attr(child, 'text:style-name'))
    if (format) block.format = format
    blocks.push(block)
  }
  const emitListFlat = (list: XNode): void => {
    // lists inside cells keep their text as indented paragraphs; a private
    // numbering definition per cell would multiply abstractNums for little gain
    for (const { name, node: item } of childElements(list)) {
      if (name !== 'text:list-item') continue
      for (const inner of childElements(item)) {
        if (inner.name === 'text:p') {
          emitInto(inner.node, { type: 'paragraph' })
          const block = blocks[blocks.length - 1]
          block.format = { ...block.format, indentLeft: (block.format?.indentLeft ?? 0) + 720 }
        } else if (inner.name === 'text:list') {
          emitListFlat(inner.node)
        }
      }
    }
  }
  for (const { name, node: child } of childElements(node)) {
    if (name === 'text:p') emitInto(child, { type: 'paragraph' })
    else if (name === 'text:h') {
      emitInto(child, {
        type: 'heading',
        level: clampHeadingLevel(attr(child, 'text:outline-level')),
      })
    } else if (name === 'text:list') emitListFlat(child)
    else if (name === 'table:table') ctx.drops.add('nested-tables')
  }
  // whatever inline images the cell paragraphs referenced cannot live inside
  // the flat cell XML — dropped and counted
  if (ctx.images.length > 0) {
    ctx.drops.add('images')
    ctx.images.length = 0
  }
  const contentXml =
    blocks.map((block) => generateParagraphXml(block, CELL_GEN_CTX)).join('') || '<w:p/>'
  return {
    row: -1,
    column: -1,
    colSpan: Math.max(1, Number(attr(node, 'table:number-columns-spanned') ?? '1') || 1),
    rowSpan: Math.max(1, Number(attr(node, 'table:number-rows-spanned') ?? '1') || 1),
    contentXml,
    fill: ctx.styles.cellFill(attr(node, 'table:style-name')),
  }
}

function tableXml(columns: ColumnSpec[], grid: Array<Array<GridCell | null>>): string {
  // unknown column widths share the usable width of a Letter page with 1in
  // margins; fixed layout plus real gridCols keeps the spans honest
  const known = columns.filter((column) => column.widthEmu !== null)
  const DEFAULT_USABLE_EMU = 914_400 * 6.5
  const knownSum = known.reduce((sum, column) => sum + (column.widthEmu ?? 0), 0)
  const unknownCount = columns.length - known.length
  const unknownWidth =
    unknownCount > 0
      ? Math.max(Math.round((DEFAULT_USABLE_EMU - knownSum) / unknownCount), 100_000)
      : 0
  const colTw = columns.map((column) => Math.max(1, emuToTwips(column.widthEmu ?? unknownWidth)))

  const border =
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
  const parts: string[] = []
  parts.push(
    `<w:tbl><w:tblPr><w:tblW w:w="${colTw.reduce((a, b) => a + b, 0)}" w:type="dxa"/>` +
      `<w:tblBorders>${border}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>`,
  )
  parts.push(
    `<w:tblGrid>${colTw.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>`,
  )
  grid.forEach((row, rowIndex) => {
    parts.push('<w:tr>')
    row.forEach((cell, columnIndex) => {
      if (!cell) {
        // an empty slot in an irregular table keeps the row's column count
        parts.push(
          `<w:tc><w:tcPr><w:tcW w:w="${colTw[columnIndex]}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`,
        )
        return
      }
      const isOrigin = cell.row === rowIndex && cell.column === columnIndex
      if (!isOrigin) {
        if (columnIndex !== cell.column) return // a horizontal span's shadow slot: gridSpan already covers it
        // vertical continuation: carry the merge (and its span) forward
        const span = cell.colSpan > 1 ? `<w:gridSpan w:val="${cell.colSpan}"/>` : ''
        parts.push(
          `<w:tc><w:tcPr><w:tcW w:w="${colTw[columnIndex]}" w:type="dxa"/>${span}<w:vMerge/></w:tcPr><w:p/></w:tc>`,
        )
        return
      }
      const widthTw = colTw
        .slice(cell.column, cell.column + cell.colSpan)
        .reduce((sum, width) => sum + width, 0)
      const tcPr: string[] = [`<w:tcW w:w="${widthTw}" w:type="dxa"/>`]
      if (cell.colSpan > 1) tcPr.push(`<w:gridSpan w:val="${cell.colSpan}"/>`)
      if (cell.rowSpan > 1) tcPr.push('<w:vMerge w:val="restart"/>')
      if (cell.fill) tcPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${cell.fill}"/>`)
      parts.push(`<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${cell.contentXml}</w:tc>`)
    })
    parts.push('</w:tr>')
  })
  parts.push('</w:tbl>')
  return parts.join('')
}

// -- images -----------------------------------------------------------------

/** fall back to 400×300 CSS px when neither the frame nor the bytes declare a size */
const IMAGE_FALLBACK_PX = { width: 400, height: 300 }

async function resolveImage(ref: InlineImageRef, pack: OdfPackage): Promise<NewImage | null> {
  let bytes: Uint8Array | null = null
  let path = ''
  if (ref.base64) {
    bytes = Uint8Array.from(atob(ref.base64), (character) => character.charCodeAt(0))
  } else if (ref.href) {
    path = ref.href
    bytes = await pack.read(ref.href)
  }
  if (!bytes || bytes.length === 0) return null
  const mime = sniffImageMime(bytes, path)
  if (!mime) return null
  let widthPx = ref.widthEmu !== undefined ? Math.max(1, Math.round(ref.widthEmu / 9525)) : 0
  let heightPx = ref.heightEmu !== undefined ? Math.max(1, Math.round(ref.heightEmu / 9525)) : 0
  if (!widthPx || !heightPx) {
    const intrinsic = intrinsicImageSize(bytes, mime)
    widthPx = widthPx || intrinsic?.width || IMAGE_FALLBACK_PX.width
    heightPx = heightPx || intrinsic?.height || IMAGE_FALLBACK_PX.height
  }
  return {
    base64: Buffer.from(bytes).toString('base64'),
    mime,
    widthPx,
    heightPx,
  }
}
