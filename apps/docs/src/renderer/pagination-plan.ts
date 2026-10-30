// Model-side pre-pass of the pagination facts that do not need layout: break
// classification (P5), table row keep policy (P3) and the space-before model
// (P2), computed from the ProseMirror document plus the resolved paragraph
// properties. Stage 1 runs it in shadow next to the DOM scrape in
// measureBlocks (invariant I9: plan == scrape for every block).
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EffectivePPr, ParsedDoc, SectionInfo, TableModel } from '@genoffice/docx-engine'
import type { BlockBox, BlockMetaOf, RowPlan } from './pagination-types'
import { outerTableRows } from './pagination-lines'
import { sectionGeoms } from './pagination-sections'
import { additiveParagraphSpacing, autoSpacingPt, spaceBeforeProp } from './line-metrics'
import {
  pictureBlockFacts,
  tableFloatFacts,
  type AnchorCtx,
  type PageRelVTarget,
} from './editor/anchor-geometry'
import type { FloatShiftState } from './editor/pagination-gaps'

export type { RowPlan } from './pagination-types'

export interface BlockPlan {
  pos: number
  node: PmNode
  docxIndex?: number
  kind: 'para' | 'table' | 'protected'
  /** the block's own w:pageBreakBefore (the page-break-before class) */
  pageBreakBeforeDirect: boolean
  /** direct or inherited from the style chain (applyBlockMeta's breakBefore) */
  pageBreakBefore: boolean
  /** page-type breaks that break the body flow (cell / textbox breaks excluded) */
  breakCount: number
  /** every hard break of the paragraph (a break-only paragraph lays out one line per break) */
  hardBreaks: number
  hasText: boolean
  leadingBreaks: number
  trailingBreaks: number
  /** ordinals (among the body breaks) whose post-break line turns the page mid-block */
  innerBreakIdx: number[]
  leadingBreakCut: boolean
  breakOnly: boolean
  colBreakBefore: boolean
  colBreakAfter: boolean
  hasCellBreak: boolean
  rowPlan?: RowPlan[]
  /** rendered space before (pt): direct, table-style, style chain; contextual spacing folded in */
  spaceBeforePt?: number
  /** the space before is laid out on the anchor line inside the host box (text-box
   *  and anchored-picture hosts), not as the block's outer margin */
  spaceBeforeInner?: true
  /** the space before vanishes when a natural break opens the page with this block */
  dropAtPageTop: boolean
  /** w:tblpPr geometry of a table block (the data-tblp-* twins) */
  tableGeom?: PlanFloatGeom
  /** anchored-drawing geometry of a picture / text-box block */
  anchorGeom?: PlanFloatGeom
  /** the engine's applied float shift from the previous pass (plugin state, not the DOM) */
  floatShift?: FloatShiftState
}

/** paginator-facing float geometry of one block, from the model */
export interface PlanFloatGeom {
  /** CSS-floated box (w:tblpPr table kept floated; side-wrapped picture) */
  floated: boolean
  /** page / margin Y the engine shifts the box to (BlockBox.pageRelVy*) */
  pageRelV?: { yPx: number; band: 'page' | 'margin'; spec?: 'top' | 'center' | 'bottom' }
  /** suppressed text-anchored float hanging above its anchor (BlockBox.liftPx) */
  liftPx?: number
  /** text-anchored float that turns the page whole (BlockBox.floatPushWhole) */
  pushWhole?: boolean
  /** column-spanning wrap band kept on the anchor's page (BlockBox.bandKeep) */
  bandKeep?: boolean
}

const PLAN_ANCHOR_CTX: AnchorCtx = { kind: 'body', surface: 'editor' }

function pageRelVOf(t: PageRelVTarget | null, zeroUnderSpec: boolean): PlanFloatGeom['pageRelV'] {
  if (!t || (t.band !== 'page' && t.band !== 'margin')) return undefined
  return {
    yPx: Math.round((zeroUnderSpec && t.spec ? 0 : t.yPx) * 10) / 10,
    band: t.band,
    ...(t.spec ? { spec: t.spec } : {}),
  }
}

/** the measure reads the float target only while the table is a CSS float
 *  (a flowed table loses it with the class), so the plan mirrors the gate */
function tableGeomOf(node: PmNode, shift: FloatShiftState | undefined): PlanFloatGeom | undefined {
  const f = tableFloatFacts(node.attrs, PLAN_ANCHOR_CTX)
  const floated = f.floated && !shift?.flow
  const pageRelV = floated ? pageRelVOf(f.pageRelV, true) : undefined
  if (!floated && !f.liftPx && !f.pushWhole) return undefined
  return {
    floated,
    ...(pageRelV ? { pageRelV } : {}),
    ...(f.liftPx > 0 ? { liftPx: f.liftPx } : {}),
    ...(f.pushWhole ? { pushWhole: true } : {}),
  }
}

const SIDE_WRAP_CLASS = /^(?:square|tight|through)-(?:left|right)$/

function anchorGeomOf(node: PmNode): PlanFloatGeom | undefined {
  const a = node.attrs
  // the renderer's precedence: a text-box / shape block before a picture block
  const boxes = a.textboxes as Array<{ floating?: boolean; bandOverflow?: boolean }> | null
  if (Array.isArray(boxes) && boxes.length > 0) {
    const diagram = a.diagramDisplay as { floating?: boolean } | null
    const anyFloating = boxes.some((b) => b.floating) && (!diagram || !!diagram.floating)
    if (!anyFloating) return undefined
    return { floated: false, ...(boxes.some((b) => b.bandOverflow) ? { bandKeep: true } : {}) }
  }
  if (!a.imageSrc && !a.imageWrap) return undefined
  const f = pictureBlockFacts(a, PLAN_ANCHOR_CTX)
  const wrap = String(a.imageWrap ?? '')
  const floated = !f.banded && SIDE_WRAP_CLASS.test(wrap)
  // the in-flow wrapTopAndBottom band carries the target like a float
  const pageRelV = floated || wrap === 'topBottom' ? pageRelVOf(f.pageRelV, false) : undefined
  if (!floated && !pageRelV) return undefined
  return { floated, ...(pageRelV ? { pageRelV } : {}) }
}

export interface PlanProps {
  /** resolved properties of the body paragraph at this docxIndex */
  paraEffective(docxIndex: number): EffectivePPr | undefined
  /** resolved properties of a cell paragraph (model row / cell / paragraph indices) */
  cellParaEffective?(
    docxIndex: number,
    row: number,
    cell: number,
    para: number,
  ): EffectivePPr | undefined
  tableModel?(docxIndex: number): TableModel | undefined
  /** w:doNotUseHTMLParagraphAutoSpacing: before/after add instead of collapsing */
  additiveSpacing: boolean
  /** Word's HTML auto spacing value (pt) */
  autoSpacingPt: number
  /** Original revision view: deleted breaks still break */
  showDeleted: boolean
  /** docxIndex of every section's first block whose start forces a page */
  forcedSectionStarts?: ReadonlySet<number>
  /** list group (abstractNum) of a numId; auto spacing collapses between items of one group */
  listGroupOf?(numId: string): string | undefined
}

/** plan inputs of a parsed document (the renderer's live spacing mode applies) */
export function planPropsOf(
  parsed: ParsedDoc,
  sections: readonly SectionInfo[],
  showDeleted: boolean,
): PlanProps {
  const byIndex = new Map<number, ParsedDoc['blocks'][number]>()
  for (const b of parsed.blocks)
    if (b.docxIndex !== null && !byIndex.has(b.docxIndex)) byIndex.set(b.docxIndex, b)
  const forced = new Set<number>()
  sectionGeoms([...sections]).forEach((g, i) => {
    if (g.forceBreak) forced.add(sections[i].firstBlockIndex)
  })
  return {
    paraEffective: (i) => byIndex.get(i)?.format?.effective,
    cellParaEffective: (i, r, c, p) =>
      byIndex.get(i)?.table?.rows[r]?.[c]?.richParas?.[p]?.effective,
    tableModel: (i) => byIndex.get(i)?.table ?? undefined,
    additiveSpacing: additiveParagraphSpacing(),
    autoSpacingPt: autoSpacingPt(),
    showDeleted,
    forcedSectionStarts: forced,
    listGroupOf: (numId) => parsed.numbering.get(numId)?.abstractNumId,
  }
}

const PARA_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])
/** inline atoms that render visible text (Range.toString sees them) */
const TEXT_ATOMS = new Set(['docNoteRef', 'docInlineMath'])

export function isParaNode(node: PmNode): boolean {
  return PARA_TYPES.has(node.type.name)
}

/** the near-invisible protected strips measureBlocks skips (section-break / field boundary paragraphs) */
export function isSkippedProtected(node: PmNode): boolean {
  if (node.type.name !== 'docProtected') return false
  const { label, previewText, fieldDisplay } = node.attrs
  if (previewText) return false
  if (label === 'Section break paragraph') return true
  return (
    !fieldDisplay &&
    (label === 'Auto TOC (updates when opened in Word)' || label === 'Field end marker')
  )
}

interface InlineBreak {
  kind: 'page' | 'col'
  before: boolean
  after: boolean
}

function isDeleted(node: PmNode): boolean {
  return node.marks.some((m) => m.type.name === 'del')
}

/** inline text of one child as the DOM Range sees it (atoms with a rendered label count) */
function inlineText(child: PmNode): string {
  if (child.isText) return child.text ?? ''
  if (child.type.name === 'hardBreak') return ''
  if (child.type.name === 'docInlineMath') return String(child.attrs.text ?? '')
  if (TEXT_ATOMS.has(child.type.name)) return 'x'
  return child.textContent
}

function hasPicture(child: PmNode): boolean {
  if (child.type.name === 'docInlineImage') return true
  let found = false
  child.descendants((n) => {
    if (n.type.name === 'docInlineImage') found = true
    return !found
  })
  return found
}

/** page / column breaks of a paragraph with the text on either side */
function paraBreaks(
  node: PmNode,
  showDeleted: boolean,
): { breaks: InlineBreak[]; text: boolean; picture: boolean; hardBreaks: number } {
  const parts: Array<{ br?: InlineBreak['kind']; text: string }> = []
  let picture = false
  let hardBreaks = 0
  node.forEach((child) => {
    if (child.type.name === 'hardBreak') {
      hardBreaks++
      const kind = child.attrs.pageBreak ? 'page' : child.attrs.colBreak ? 'col' : undefined
      if (kind && (showDeleted || !isDeleted(child))) parts.push({ br: kind, text: '' })
      return
    }
    if (hasPicture(child)) picture = true
    parts.push({ text: inlineText(child) })
  })
  const breaks: InlineBreak[] = []
  parts.forEach((p, i) => {
    if (!p.br) return
    const before = parts.slice(0, i).some((q) => !q.br && q.text.trim() !== '')
    const after = parts.slice(i + 1).some((q) => !q.br && q.text.trim() !== '')
    breaks.push({ kind: p.br, before, after })
  })
  const text = parts.some((p) => !p.br && p.text.trim() !== '')
  return { breaks, text, picture, hardBreaks }
}

/** measureBlocks' leading / trailing / inner classification over the body breaks */
function classifyBreaks(brs: InlineBreak[], fieldBreak: boolean) {
  const all = fieldBreak ? [{ kind: 'field' as const, before: false, after: false }] : brs
  const hasText = all.some((x) => x.before || x.after)
  const leading = all.filter((x) => !x.before).length
  const trailing = all.filter((x) => !x.after).length
  const inner: number[] = []
  let leadingCut = false
  all.forEach((x, i) => {
    if (!x.after || x.kind !== 'page') return
    if (!x.before && i !== leading - 1) return
    inner.push(i)
    if (!x.before) leadingCut = true
  })
  return { count: all.length, hasText, leading, trailing, inner, leadingCut }
}

function hasForcedBreakIn(node: PmNode): boolean {
  let found = false
  node.descendants((n) => {
    if (found) return false
    if (n.type.name === 'docParagraph' && n.attrs.pageBreakBefore) found = true
    return !found
  })
  return found
}

const INHERITED: ReadonlySet<string> = new Set([
  'paraStyle',
  'tableStyle',
  'numbering',
  'docDefaults',
])

function inheritedFlag(
  eff: EffectivePPr | undefined,
  key: 'keepNext' | 'pageBreakBefore',
): boolean {
  return !!eff && eff[key] && INHERITED.has(eff.source[key])
}

function rowPlanOf(table: PmNode, docxIndex: number | undefined, props: PlanProps): RowPlan[] {
  const model = docxIndex === undefined ? undefined : props.tableModel?.(docxIndex)
  const rows: RowPlan[] = []
  table.forEach((row, _off, ri) => {
    if (row.type.name !== 'docTableRow') return
    const modelRow = model?.rows[ri]
    const plan: RowPlan = {
      isHeader: !!row.attrs.repeatHeader,
      cantSplit: !!row.attrs.cantSplit,
    }
    const h = Number(row.attrs.heightTwips)
    if (h > 0 && row.attrs.heightRule !== 'exact') plan.minHPt = Math.min(h, 31680) / 20
    if (modelRow && modelRow.length > 0 && modelRow.every((c) => c.vMerge === 'continue'))
      plan.vMergeContinue = true
    // the cell's first paragraph; an anchored-shape strut or a nested table
    // ahead of it is not a paragraph (the XML scrape skips nested tables too)
    let firstPara: PmNode | undefined
    row.firstChild?.forEach((child) => {
      if (!firstPara && isParaNode(child)) firstPara = child
    })
    if (firstPara) {
      const direct = firstPara.attrs.keepNext
      if (direct === true) plan.keepNext = true
      else if (direct == null && docxIndex !== undefined) {
        // the PM row drops vMerge-continue cells: its first cell is the model's first open cell
        const modelCell = modelRow?.findIndex((c) => c.vMerge !== 'continue') ?? -1
        const eff =
          modelCell >= 0 ? props.cellParaEffective?.(docxIndex, ri, modelCell, 0) : undefined
        if (inheritedFlag(eff, 'keepNext')) plan.keepNext = true
      }
    }
    if (hasForcedBreakIn(row)) plan.hasForcedBreak = true
    rows.push(plan)
  })
  return rows
}

/** list group key of an item: its abstractNum, else the numId; non-items have none */
function listGroup(node: PmNode, props: PlanProps): string | undefined {
  if (node.type.name !== 'docListItem') return undefined
  const numId = node.attrs.numId == null ? null : String(node.attrs.numId)
  if (numId === null) return 'none'
  return props.listGroupOf?.(numId) ?? numId
}

/**
 * The rendered space before (pt) of a body paragraph, mirroring the stylesheet:
 * direct spacing, then the resolved chain; HTML auto spacing collapses on the
 * document's first block and between two items of one list; contextual
 * spacing between same-style neighbours (direct for unstyled, any source for
 * styled). The page-break indicator's chrome margin is not spacing.
 */
function paraSpaceBeforePt(
  node: PmNode,
  eff: EffectivePPr | undefined,
  prev: PmNode | undefined,
  index: number,
  props: PlanProps,
): number | undefined {
  const a = node.attrs
  const auto = a.spaceBeforeAuto ?? (a.spaceBefore == null && eff?.spacing.beforeAuto) ?? false
  let pt: number | undefined
  if (auto) pt = props.autoSpacingPt
  else if (a.spaceBefore != null) pt = Number(a.spaceBefore) / 20
  else if (eff) pt = eff.spacing.before / 20
  if (pt === undefined) return undefined
  if (auto) {
    if (index === 0) pt = 0
    else if (prev) {
      const g = listGroup(node, props)
      if (g !== undefined && listGroup(prev, props) === g) pt = 0
    }
  }
  const styleId = (a.styleId as string | null) ?? null
  // Word compares styles only (word-truths 2026-10-30): an outline-only heading
  // is a Normal paragraph; the unstyled rule (styles.css) skips the built-in
  // heading look, the style rule pairs any styled block
  const prevPara =
    !!prev && isParaNode(prev) && (!!styleId || (isBodyLook(prev) && isBodyLook(node)))
  const prevStyle = prevPara ? ((prev.attrs.styleId as string | null) ?? null) : undefined
  const ctx = styleId
    ? !!eff?.contextualSpacing && a.contextualSpacing !== false
    : a.contextualSpacing === true
  if (ctx && prevPara && prevStyle === styleId) pt = 0
  return Math.round(pt * 10) / 10
}

/** rendered as p / .doc-li / h*.doc-outline-only, not the built-in heading look */
function isBodyLook(node: PmNode): boolean {
  return node.type.name !== 'docHeading' || !!node.attrs.outlineOnly
}

/** Plans for the document's top-level blocks, keyed by node identity (and position). */
export function buildBlockPlan(
  doc: PmNode,
  props: PlanProps,
  /** the engine's applied float shifts of the current editor state (floatShiftsOf) */
  shiftOf?: (pos: number) => FloatShiftState | undefined,
): Map<PmNode, BlockPlan> {
  const plans = new Map<PmNode, BlockPlan>()
  let prev: PmNode | undefined
  let index = 0
  doc.forEach((node, pos) => {
    if (isSkippedProtected(node)) return
    const docxIndex =
      typeof node.attrs.docxIndex === 'number' ? (node.attrs.docxIndex as number) : undefined
    const base = {
      pos,
      node,
      ...(docxIndex !== undefined ? { docxIndex } : {}),
      colBreakBefore: false,
      colBreakAfter: false,
      hasCellBreak: false,
      innerBreakIdx: [] as number[],
      leadingBreakCut: false,
      breakOnly: false,
      hasText: false,
      leadingBreaks: 0,
      trailingBreaks: 0,
      breakCount: 0,
      hardBreaks: 0,
    }
    let plan: BlockPlan
    const shift = shiftOf?.(pos)
    if (node.type.name === 'docTable') {
      const tableGeom = tableGeomOf(node, shift)
      plan = {
        ...base,
        kind: 'table',
        pageBreakBeforeDirect: false,
        pageBreakBefore: false,
        hasCellBreak: hasForcedBreakIn(node),
        rowPlan: rowPlanOf(node, docxIndex, props),
        dropAtPageTop: !(
          index === 0 ||
          (docxIndex !== undefined && props.forcedSectionStarts?.has(docxIndex))
        ),
        ...(tableGeom ? { tableGeom } : {}),
        ...(shift ? { floatShift: shift } : {}),
      }
    } else if (isParaNode(node)) {
      const eff = docxIndex === undefined ? undefined : props.paraEffective(docxIndex)
      const { breaks, text, picture, hardBreaks } = paraBreaks(node, props.showDeleted)
      const pageBrs = breaks.filter((b) => b.kind === 'page')
      const colBrs = breaks.filter((b) => b.kind === 'col')
      const c = classifyBreaks(pageBrs, false)
      const direct = !!node.attrs.pageBreakBefore
      // a column break with nothing before it moves the whole paragraph
      const leadingCol = colBrs.length === 1 && !colBrs[0].before
      const pageBreakBefore = direct || inheritedFlag(eff, 'pageBreakBefore')
      const leadingBr = c.hasText && c.leading > 0 && !c.leadingCut
      plan = {
        ...base,
        kind: 'para',
        pageBreakBeforeDirect: direct,
        pageBreakBefore,
        breakCount: c.count,
        hardBreaks,
        hasText: c.hasText,
        leadingBreaks: c.leading,
        trailingBreaks: c.trailing,
        innerBreakIdx: c.inner,
        leadingBreakCut: c.leadingCut,
        breakOnly: c.count > 0 && !text && !picture,
        colBreakBefore: leadingCol,
        colBreakAfter: colBrs.length > 0 && !leadingCol,
        spaceBeforePt: paraSpaceBeforePt(node, eff, prev, index, props),
        dropAtPageTop: !(
          index === 0 ||
          pageBreakBefore ||
          leadingBr ||
          (docxIndex !== undefined && props.forcedSectionStarts?.has(docxIndex))
        ),
      }
    } else {
      const field = node.attrs.fieldDisplay as { kind?: string } | null
      const c = classifyBreaks([], field?.kind === 'pageBreak')
      const direct = !!node.attrs.pageBreakBefore
      // auto space before collapses on the document's first block like a paragraph's
      const sb = node.attrs.imageParagraphSpaceBeforeAuto
        ? index === 0
          ? 0
          : props.autoSpacingPt
        : node.attrs.imageParagraphSpaceBefore != null
          ? Number(node.attrs.imageParagraphSpaceBefore) / 20
          : undefined
      const anchorGeom = anchorGeomOf(node)
      // only an in-line picture's host carries the paragraph margin itself (doc-img-para)
      const boxes = node.attrs.textboxes as unknown[] | null
      const inlinePicture =
        !(Array.isArray(boxes) && boxes.length > 0) &&
        node.attrs.blockType === 'image' &&
        !!node.attrs.imageDataUrl &&
        !node.attrs.imageWrap
      plan = {
        ...base,
        kind: 'protected',
        pageBreakBeforeDirect: direct,
        pageBreakBefore: direct,
        breakCount: c.count,
        leadingBreaks: c.leading,
        trailingBreaks: c.trailing,
        ...(sb !== undefined ? { spaceBeforePt: Math.round(sb * 10) / 10 } : {}),
        ...(sb !== undefined && !inlinePicture ? { spaceBeforeInner: true as const } : {}),
        ...(anchorGeom ? { anchorGeom } : {}),
        ...(shift ? { floatShift: shift } : {}),
        dropAtPageTop: !(
          index === 0 ||
          direct ||
          (docxIndex !== undefined && props.forcedSectionStarts?.has(docxIndex))
        ),
      }
    }
    plans.set(node, plan)
    prev = node
    index++
  })
  return plans
}

export type PlanOf = (el: HTMLElement) => BlockPlan | undefined

/** the plan keyed by each block's rendered element (view.nodeDOM), for measureBlocks */
export function planByElement(
  doc: PmNode,
  view: { nodeDOM(pos: number): Node | null },
  props: PlanProps,
  shiftOf?: (pos: number) => FloatShiftState | undefined,
): PlanOf {
  const byEl = new Map<HTMLElement, BlockPlan>()
  for (const p of buildBlockPlan(doc, props, shiftOf).values()) {
    const dom = view.nodeDOM(p.pos)
    if (dom instanceof HTMLElement) byEl.set(dom, p)
  }
  return (el) => byEl.get(el)
}

/** derived BlockBox flags of a plan (the engine's vocabulary) */
export function planBreakFlags(p: BlockPlan): {
  breakBefore?: true
  breakBeforeBr?: true
  extraBreaksBefore?: number
  breakAfter?: true
  extraBreaksAfter?: number
  innerBreakCount: number
  leadingBreakCut?: true
  colBreakBefore?: true
  colBreakAfter?: true
  hasCellBreak?: true
} {
  const leadingBr = p.hasText && p.leadingBreaks > 0 && !p.leadingBreakCut
  return {
    ...(p.pageBreakBeforeDirect || leadingBr ? { breakBefore: true } : {}),
    ...(leadingBr ? { breakBeforeBr: true } : {}),
    ...(p.hasText && p.leadingBreaks > 1 ? { extraBreaksBefore: p.leadingBreaks - 1 } : {}),
    ...(p.trailingBreaks > 0 ? { breakAfter: true } : {}),
    ...(p.trailingBreaks > 1 ? { extraBreaksAfter: p.trailingBreaks - 1 } : {}),
    innerBreakCount: p.innerBreakIdx.length,
    ...(p.leadingBreakCut ? { leadingBreakCut: true } : {}),
    ...(p.colBreakBefore ? { colBreakBefore: true } : {}),
    ...(p.colBreakAfter ? { colBreakAfter: true } : {}),
    ...(p.hasCellBreak ? { hasCellBreak: true } : {}),
  }
}

export interface PlanDiff {
  /** disagreement class, e.g. `breakAfter`, `row.keepNext`, `spaceBefore`, `unplanned` */
  cls: string
  block: number
  docxIndex?: number
  plan: unknown
  scrape: unknown
}

/**
 * I9: compare the plan with what measureBlocks scraped. `planOf` maps a block
 * element to its plan (view.nodeDOM); `metaOf` supplies the parse-layer flags
 * applyBlockMeta / fillLineBoxes would add. Pure reads; never changes a block.
 */
export function comparePlanToBlocks(
  blocks: BlockBox[],
  planOf: (el: HTMLElement) => BlockPlan | undefined,
  metaOf?: BlockMetaOf,
): PlanDiff[] {
  const diffs: PlanDiff[] = []
  blocks.forEach((b, i) => {
    // decoration widgets (anchor-line clones) are layout, not model blocks
    if (!b.el || b.el.classList.contains('ProseMirror-widget')) return
    const p = planOf(b.el)
    const push = (cls: string, plan: unknown, scrape: unknown) =>
      diffs.push({
        cls,
        block: i,
        ...(b.docxIndex !== undefined ? { docxIndex: b.docxIndex } : {}),
        plan,
        scrape,
      })
    if (!p) {
      push('unplanned', undefined, b.el.className.slice(0, 40))
      return
    }
    const meta = b.docxIndex === undefined ? undefined : metaOf?.(b.docxIndex)
    const f = planBreakFlags(p)
    const flag = (
      cls: string,
      plan: boolean | number | undefined,
      scrape: boolean | number | undefined,
    ) => {
      if ((plan ?? false) !== (scrape ?? false)) push(cls, plan ?? false, scrape ?? false)
    }
    flag('pageBreakBefore', p.pageBreakBeforeDirect, b.el.classList.contains('page-break-before'))
    if (metaOf)
      flag(
        'pageBreakBeforeStyle',
        p.pageBreakBefore,
        p.pageBreakBeforeDirect || !!meta?.breakBefore,
      )
    // a plan-driven block carries the scrape's flags on shadowScrape; a
    // plan-less block's own flags are the scrape
    const sc: Record<string, unknown> = b.shadowScrape ?? {
      ...b,
      innerBreakCount: b.innerBreaks?.length ?? 0,
      breakOnly: b.breakOnlyLineH !== undefined,
    }
    flag('breakBefore', f.breakBefore, sc.breakBefore as boolean | undefined)
    flag('breakBeforeBr', f.breakBeforeBr, sc.breakBeforeBr as boolean | undefined)
    flag('extraBreaksBefore', f.extraBreaksBefore, sc.extraBreaksBefore as number | undefined)
    flag('breakAfter', f.breakAfter, sc.breakAfter as boolean | undefined)
    flag('extraBreaksAfter', f.extraBreaksAfter, sc.extraBreaksAfter as number | undefined)
    flag('innerBreakCount', f.innerBreakCount, sc.innerBreakCount as number)
    flag('leadingBreakCut', f.leadingBreakCut, sc.leadingBreakCut as boolean | undefined)
    flag('breakOnly', p.breakOnly, sc.breakOnly as boolean)
    flag('colBreakBefore', f.colBreakBefore, sc.colBreakBefore as boolean | undefined)
    flag('colBreakAfter', f.colBreakAfter, sc.colBreakAfter as boolean | undefined)
    flag('hasCellBreak', f.hasCellBreak, sc.hasCellBreak as boolean | undefined)
    if (p.rowPlan) {
      const trs = outerTableRows(b.el)
      if (trs.length !== p.rowPlan.length) push('row.count', p.rowPlan.length, trs.length)
      const rowFlags = meta?.tableRowFlags
      p.rowPlan.forEach((r, ri) => {
        const tr = trs[ri]
        if (!tr) return
        const live = (attr: string) =>
          tr.hasAttribute(attr) ? tr.getAttribute(attr) === '1' : undefined
        flag(`row.isHeader`, r.isHeader, live('data-repeat-header') ?? !!rowFlags?.[ri]?.isHeader)
        flag(`row.cantSplit`, r.cantSplit, live('data-cant-split') ?? !!rowFlags?.[ri]?.cantSplit)
        if (metaOf) {
          flag(`row.keepNext`, r.keepNext, rowFlags?.[ri]?.keepNext)
          const minPt =
            rowFlags?.[ri]?.minHPx === undefined ? undefined : (rowFlags[ri].minHPx! * 72) / 96
          if (Math.abs((r.minHPt ?? 0) - (minPt ?? 0)) > 0.05) push('row.minH', r.minHPt, minPt)
        }
        flag(`row.forcedBreak`, r.hasForcedBreak, !!tr.querySelector('p.page-break-before'))
      })
    }
    compareFloatGeom(p, sc, push)
    // the page-break indicator's 14pt chrome margin rides margin-top (styles.css);
    // Word reserves nothing there (word-truths 2026-09-23), so it is no plan fact
    // A page-gap widget between two blocks breaks the stylesheet's sibling
    // rules (list auto-spacing collapse, contextual spacing), so the margin
    // read there is decoration-dependent, not a document fact.
    const afterGap = b.el.previousElementSibling?.classList.contains('page-gap') ?? false
    if (
      p.spaceBeforePt !== undefined &&
      p.kind !== 'table' &&
      !p.pageBreakBeforeDirect &&
      !afterGap &&
      typeof getComputedStyle === 'function'
    ) {
      const prop = spaceBeforeProp()
      const host = p.spaceBeforeInner
        ? b.el.querySelector<HTMLElement>(':scope > .doc-anchor-line, :scope > .doc-textbox-stray')
        : b.el
      const css =
        host && getComputedStyle(host)[prop === 'padding-top' ? 'paddingTop' : 'marginTop']
      const pt = css ? cssLengthPt(css) : undefined
      if (pt !== undefined && Math.abs(pt - p.spaceBeforePt) > 0.15)
        push('spaceBefore', p.spaceBeforePt, Math.round(pt * 100) / 100)
    }
    // the page lead folded from the plan against the first block's DOM offset
    // (only where the DOM has computed geometry: jsdom reports no margin)
    const leadPx = b.shadowScrape?.leadPx
    if (
      typeof leadPx === 'number' &&
      !p.pageBreakBeforeDirect &&
      typeof getComputedStyle === 'function' &&
      cssLengthPt(getComputedStyle(b.el).marginTop) !== undefined
    ) {
      const planPt = ((b.leadFoldPx ?? 0) * 72) / 96
      const domPt = (leadPx * 72) / 96
      // a margin under the 0.5 px fold threshold is folded by neither side
      if (Math.abs(planPt - domPt) > 0.375)
        push('lead', Math.round(planPt * 100) / 100, Math.round(domPt * 100) / 100)
    }
  })
  return diffs
}

/** the plan's float geometry against the scraped data-tblp-* / band / shift facts */
function compareFloatGeom(
  p: BlockPlan,
  sc: Record<string, unknown>,
  push: (cls: string, plan: unknown, scrape: unknown) => void,
): void {
  const g = p.tableGeom ?? p.anchorGeom
  const num = (cls: string, plan: number | undefined, scrape: unknown) => {
    const s = typeof scrape === 'number' ? scrape : undefined
    if (Math.abs((plan ?? 0) - (s ?? 0)) > 0.05) push(cls, plan, s)
  }
  const same = (cls: string, plan: unknown, scrape: unknown) => {
    if ((plan ?? false) !== (scrape ?? false)) push(cls, plan ?? false, scrape ?? false)
  }
  num('float.pageRelVy', g?.pageRelV?.yPx, sc.pageRelVyPx)
  same('float.pageRelVAnchor', g?.pageRelV?.band, sc.pageRelVAnchor)
  same('float.pageRelVSpec', g?.pageRelV?.spec, sc.pageRelVSpec)
  num('float.lift', g?.liftPx, sc.liftPx)
  same('float.pushWhole', g?.pushWhole, sc.floatPushWhole)
  same('float.bandKeep', g?.bandKeep, sc.bandKeep)
  if (g?.pageRelV || sc.pageRelVyPx !== undefined)
    num('float.shift', p.floatShift?.flow ? 0 : p.floatShift?.dyPx, sc.floatShiftPx)
  if (p.tableGeom) same('float.flowed', p.floatShift?.flow, sc.floatFlowed)
}

/** computed lengths are px in a browser; jsdom hands back the declared value */
function cssLengthPt(css: string): number | undefined {
  const v = parseFloat(css)
  if (!Number.isFinite(v)) return undefined
  return css.trim().endsWith('pt') ? v : (v * 72) / 96
}

export function planDiffHistogram(diffs: readonly PlanDiff[]): Record<string, number> {
  const out: Record<string, number> = {}
  for (const d of diffs) out[d.cls] = (out[d.cls] ?? 0) + 1
  return out
}
