import type { JSONContent } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  TABLE_HEADER_FILL,
  buildBlankDocx,
  generateTableModelXml,
  parseDocx,
  saveDocx,
} from '@genoffice/docx-engine'
import type {
  GeneratedBlock,
  NewImage,
  ParaFormat,
  Run,
  SaveBlock,
  SaveOptions,
  TableCell,
  TableModel,
  TableParagraph,
} from '@genoffice/docx-engine'

/** Resolve an authored image src to embeddable bytes; null → fall back to alt text */
export type ImageLoader = (src: string) => Promise<NewImage | null>

export interface DocxMapping {
  blocks: SaveBlock[]
  options: SaveOptions
}

const MAX_LIST_LEVEL = 4
const INDENT_STEP = 360
/** decimal abstractNum of the blank template (numbering.xml: 0 = bullet, 1 = decimal) */
const DECIMAL_ABSTRACT_NUM_ID = '1'
const CODE_FONT = 'Consolas'
const CODE_FILL = 'F2F3F5'
const CALLOUT_FILL: Record<string, string> = {
  info: 'E8F0FE',
  tip: 'E6F4EA',
  warning: 'FEF3E2',
  danger: 'FCE8E6',
}

// ── inline content → Run[] ──

function runsFromInline(content: JSONContent[] | undefined): Run[] {
  const runs: Run[] = []
  for (const child of content ?? []) {
    if (child.type === 'hardBreak') {
      runs.push({ text: '\n' })
      continue
    }
    if (child.type !== 'text' || !child.text) continue
    const run: Run = { text: child.text }
    for (const mark of child.marks ?? []) {
      if (mark.type === 'bold') run.bold = true
      else if (mark.type === 'italic') run.italic = true
      else if (mark.type === 'strike') run.strike = true
      else if (mark.type === 'underline') run.underline = true
      else if (mark.type === 'highlight') run.highlight = 'yellow'
      else if (mark.type === 'code') run.font = CODE_FONT
      else if (mark.type === 'link' && typeof mark.attrs?.href === 'string') {
        run.link = { href: mark.attrs.href }
      }
    }
    runs.push(run)
  }
  return runs
}

function plainText(node: JSONContent): string {
  if (node.type === 'text') return node.text ?? ''
  return (node.content ?? []).map(plainText).join('')
}

// ── mapping walker ──

interface WalkContext {
  blocks: SaveBlock[]
  restartNums: NonNullable<NonNullable<SaveOptions['numbering']>['restartNums']>
  nextOrderedNumId: number
  loadImage: ImageLoader
  pendingImages: Array<{ index: number; src: string; alt: string }>
}

function mergeFormat(base: ParaFormat | undefined, extra: ParaFormat): ParaFormat {
  return { ...base, ...extra, indentLeft: (base?.indentLeft ?? 0) + (extra.indentLeft ?? 0) }
}

function pushParagraph(ctx: WalkContext, block: GeneratedBlock): void {
  ctx.blocks.push({ kind: 'generated', block })
}

function walkList(
  ctx: WalkContext,
  node: JSONContent,
  kind: 'bullet' | 'ordered',
  ilvl: number,
  numId: string,
  base?: ParaFormat,
): void {
  for (const item of node.content ?? []) {
    if (item.type !== 'listItem' && item.type !== 'taskItem') continue
    let firstPara = true
    for (const child of item.content ?? []) {
      if (child.type === 'paragraph') {
        const runs = runsFromInline(child.content)
        if (item.type === 'taskItem') {
          runs.unshift({ text: item.attrs?.checked ? '☑ ' : '☐ ' })
        }
        if (firstPara && item.type !== 'taskItem') {
          pushParagraph(ctx, {
            type: 'listItem',
            list: { kind, numId, ilvl: Math.min(ilvl, MAX_LIST_LEVEL) },
            runs,
            format: base,
          })
        } else {
          // task items and continuation paragraphs indent under the marker
          pushParagraph(ctx, {
            type: 'paragraph',
            runs,
            format: mergeFormat(base, { indentLeft: INDENT_STEP * (ilvl + 1) }),
          })
        }
        firstPara = false
      } else if (child.type === 'bulletList' || child.type === 'taskList') {
        walkList(ctx, child, 'bullet', ilvl + 1, BLANK_BULLET_NUM_ID, base)
      } else if (child.type === 'orderedList') {
        walkList(ctx, child, 'ordered', ilvl + 1, allocOrderedNumId(ctx), base)
      } else {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: INDENT_STEP * (ilvl + 1) }))
      }
    }
  }
}

function allocOrderedNumId(ctx: WalkContext): string {
  const numId = String(ctx.nextOrderedNumId++)
  ctx.restartNums.push({
    numId,
    abstractNumId: DECIMAL_ABSTRACT_NUM_ID,
    startOverrides: { 0: 1, 1: 1, 2: 1, 3: 1, 4: 1 },
  })
  return numId
}

function tableParagraphs(cell: JSONContent): TableParagraph[] {
  const paras: TableParagraph[] = []
  for (const child of cell.content ?? []) {
    if (child.type === 'paragraph') paras.push({ runs: runsFromInline(child.content) })
  }
  return paras.length > 0 ? paras : [{ runs: [] }]
}

function mapTable(node: JSONContent): TableModel {
  const rows: TableCell[][] = []
  for (const row of node.content ?? []) {
    if (row.type !== 'tableRow') continue
    const cells: TableCell[] = []
    for (const cell of row.content ?? []) {
      if (cell.type !== 'tableCell' && cell.type !== 'tableHeader') continue
      const isHeader = cell.type === 'tableHeader'
      const rich = tableParagraphs(cell)
      cells.push({
        paras: rich.map((p) => p.runs.map((r) => r.text).join('')),
        richParas: rich,
        ...(Number(cell.attrs?.colspan) > 1 ? { colSpan: Number(cell.attrs?.colspan) } : {}),
        ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}),
      })
    }
    if (cells.length > 0) rows.push(cells)
  }
  return { rows }
}

function walkBlock(ctx: WalkContext, node: JSONContent, base?: ParaFormat): void {
  switch (node.type) {
    case 'paragraph':
      pushParagraph(ctx, { type: 'paragraph', runs: runsFromInline(node.content), format: base })
      break
    case 'heading': {
      const level = Math.min(Math.max(Number(node.attrs?.level) || 1, 1), 6) as number
      pushParagraph(ctx, {
        type: 'heading',
        level,
        runs: runsFromInline(node.content),
        format: base,
      })
      break
    }
    case 'bulletList':
      walkList(ctx, node, 'bullet', 0, BLANK_BULLET_NUM_ID, base)
      break
    case 'orderedList':
      walkList(ctx, node, 'ordered', 0, allocOrderedNumId(ctx), base)
      break
    case 'taskList':
      walkList(ctx, node, 'bullet', 0, BLANK_BULLET_NUM_ID, base)
      break
    case 'blockquote':
      for (const child of node.content ?? []) {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: INDENT_STEP, borders: 'l' }))
      }
      break
    case 'codeBlock':
      pushParagraph(ctx, {
        type: 'paragraph',
        runs: [{ text: plainText(node), font: CODE_FONT, sizeHalfPoints: 19 }],
        format: mergeFormat(base, { shadingFill: CODE_FILL }),
      })
      break
    case 'horizontalRule':
      pushParagraph(ctx, {
        type: 'paragraph',
        runs: [],
        format: mergeFormat(base, { borders: 'b' }),
      })
      break
    case 'image': {
      const src = String(node.attrs?.src ?? '')
      const alt = String(node.attrs?.alt ?? '')
      // placeholder now, replaced by the loaded image (or alt text) after the async pass
      ctx.pendingImages.push({ index: ctx.blocks.length, src, alt })
      ctx.blocks.push({ kind: 'generated', block: { type: 'paragraph', runs: [] } })
      break
    }
    case 'table':
      ctx.blocks.push({ kind: 'xml', xml: generateTableModelXml(mapTable(node)) })
      break
    case 'callout': {
      const fill = CALLOUT_FILL[String(node.attrs?.type ?? 'info')] ?? CALLOUT_FILL.info
      for (const child of node.content ?? []) {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: 120, shadingFill: fill }))
      }
      break
    }
    case 'toggle': {
      const summary = String(node.attrs?.summary ?? '')
      if (summary) {
        pushParagraph(ctx, {
          type: 'paragraph',
          runs: [{ text: summary, bold: true }],
          format: base,
        })
      }
      for (const child of node.content ?? []) {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: INDENT_STEP }))
      }
      break
    }
    default: {
      // unknown block: keep its text so nothing silently disappears
      const text = plainText(node)
      if (text.trim()) pushParagraph(ctx, { type: 'paragraph', runs: [{ text }], format: base })
    }
  }
}

/** Map a ProseMirror document to docx-engine SaveBlocks (pure except image loading) */
export async function mapDocToSaveBlocks(
  doc: JSONContent,
  loadImage: ImageLoader,
): Promise<DocxMapping> {
  const ctx: WalkContext = {
    blocks: [],
    restartNums: [],
    nextOrderedNumId: 100,
    loadImage,
    pendingImages: [],
  }
  for (const node of doc.content ?? []) walkBlock(ctx, node)

  for (const pending of ctx.pendingImages) {
    const image = await loadImage(pending.src).catch(() => null)
    if (image) {
      ctx.blocks[pending.index] = { kind: 'image', image }
    } else {
      ctx.blocks[pending.index] = {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: `[${pending.alt || pending.src}]`, italic: true, color: '888888' }],
        },
      }
    }
  }

  const options: SaveOptions =
    ctx.restartNums.length > 0 ? { numbering: { restartNums: ctx.restartNums } } : {}
  return { blocks: ctx.blocks, options }
}

/** md document (PM JSON) → brand-new .docx bytes, fully local */
export async function exportDocxBytes(
  doc: JSONContent,
  loadImage: ImageLoader,
): Promise<Uint8Array> {
  const mapping = await mapDocToSaveBlocks(doc, loadImage)
  const parsed = await parseDocx(await buildBlankDocx())
  return saveDocx(parsed, mapping.blocks, mapping.options)
}
