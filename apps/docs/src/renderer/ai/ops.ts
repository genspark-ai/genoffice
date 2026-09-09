import type { Editor } from '@tiptap/core'
import type { Node as PmDocNode, Schema } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { generateTocFieldXml, type TocEntry } from '@genoffice/docx-engine'
import { isEastAsianFontName } from '../font-list'
import { t } from '../i18n/locale'
import { blockRangePositions, isTrackedDeleted, liveText } from './doc-utils'

/**
 * Canonical edit ops for the document. The model (apply_ops) and, later, the
 * ribbon issue the same flat `{ op, target, ...fields }` records: a registry
 * validates the whole batch up front and applies it as one ProseMirror
 * transaction, so one undo reverts the batch. Fields are patches: a present
 * key sets the property, null clears it, an absent key leaves it untouched.
 * The docx round trip is untouched — changed nodes become dirty via the normal
 * signature comparison and regenerate on save; docxIndex anchors are never
 * modified here.
 */

// ---- target ----

/** All given conditions filter with AND; at least one condition is required */
export interface Target {
  /** 'image' matches protected image blocks (docProtected + blockType image) */
  nodeType?: 'docHeading' | 'docParagraph' | 'docListItem' | 'image'
  /** Restrict heading level (only with nodeType: 'docHeading') */
  headingLevel?: number
  /** The block's plain text contains this substring (case-sensitive unless matchCase: false) */
  containsText?: string
  matchCase?: boolean
  /** By block index (the index output by buildDocumentContext, 0-based, current PM doc top-level order) */
  blockIndexes?: number[]
  /** 'selection' = only blocks covered by the current selection (run-level ops narrow to the selected characters); default = whole document */
  scope?: 'selection' | 'document'
  /** explicit ProseMirror range standing in for the selection (UI callers that captured a range earlier) */
  range?: { from: number; to: number }
}

export interface FontFields {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  color?: string | null
  highlight?: string | null
  /** points */
  fontSize?: number | null
  fontFamily?: string | null
  baseline?: 'superscript' | 'subscript' | 'none' | null
  link?: { url: string } | null
}

export interface ParagraphFields {
  align?: 'left' | 'center' | 'right' | 'justify' | null
  lineSpacing?: number | null
  indentLeft?: number | null
  indentRight?: number | null
  /** twips; positive = first-line indent, negative = hanging indent */
  indentFirstLine?: number | null
  spaceBefore?: number | null
  spaceAfter?: number | null
  pageBreakBefore?: boolean
  shadingFill?: string | null
  /** Paragraph borders, a subset of "tblr" */
  borders?: string | null
}

export interface Op {
  op: string
  target?: Target
  [key: string]: unknown
}

export interface OpContext {
  /** numbering context so converted list items join real docx numbering on save */
  numIds?: { bullet: string | null; ordered: string | null } | null
  /** record the batch's changes as tracked revisions under this author */
  track?: { author: string } | null
  /**
   * Selection frozen when the model last saw the doc. scope:'selection'
   * resolves against this instead of the live selection, so the user clicking
   * elsewhere mid-run cannot retarget the op. from/to (PM positions) keep a
   * partial selection character-precise; without them the block range is used.
   */
  selection?: { startIndex: number; endIndex: number; from?: number; to?: number } | null
  /** who issues the batch: the model (default) or the app's own UI (ribbon, shortcuts, dialogs) */
  source?: 'ai' | 'ui'
  /** flag changed blocks with the yellow AI highlight (default: only for source 'ai') */
  markAi?: boolean
  /** validate and plan only; nothing is dispatched */
  dryRun?: boolean
}

/** the changed-block highlight belongs to AI edits only */
const wantsAiMark = (ctx: OpContext): boolean => ctx.markAi ?? ctx.source !== 'ui'

export interface OpResult {
  op: string
  matched: number
  changed: number
  skippedProtected: number
  /** targets skipped because the text is a pending tracked deletion */
  skippedDeleted?: number
  detail?: string
}

export interface ExecuteOutcome {
  ok: boolean
  results: OpResult[]
  summary: string
  error?: string
  /** one line per op, in order (always on success, alone on dryRun) */
  plan?: string[]
}

// ---- registry ----

interface SelRange {
  from: number
  to: number
}

interface RunEnv {
  tr: Transaction
  schema: Schema
  sel: SelRange
  ctx: OpContext
}

interface OpDef {
  name: string
  /** one-line usage: listed in the tool description, appended to validation errors */
  signature: string
  /** keys allowed besides op / target */
  keys: readonly string[]
  /** target is required (default), optional, or not accepted */
  target: 'required' | 'optional' | 'none'
  /** UI-only: never listed for the model, rejected as unknown when the source is the model */
  hidden?: boolean
  validate(op: Op, where: string): string | null
  apply(op: Op, env: RunEnv): OpResult
}

const REGISTRY = new Map<string, OpDef>()

function register(def: OpDef): void {
  REGISTRY.set(def.name, def)
}

const callable = () => [...REGISTRY.values()].filter((d) => !d.hidden)

/** ops the model may call */
export function opNames(): string[] {
  return callable().map((d) => d.name)
}

export function opSignatures(): string[] {
  return callable().map((d) => d.signature)
}

// ---- validation ----

const NODE_TYPES = ['docHeading', 'docParagraph', 'docListItem', 'image'] as const
const BASELINES = ['superscript', 'subscript', 'none'] as const
const ALIGNS = ['left', 'center', 'right', 'justify'] as const
const FONT_KEYS = [
  'bold',
  'italic',
  'underline',
  'strike',
  'color',
  'highlight',
  'fontSize',
  'fontFamily',
  'baseline',
  'link',
] as const
const PARA_KEYS = [
  'align',
  'lineSpacing',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'pageBreakBefore',
  'shadingFill',
  'borders',
] as const
const IMAGE_KEYS = ['widthPx', 'heightPx', 'align'] as const

function validateTarget(target: unknown, where: string): string | null {
  if (!target || typeof target !== 'object') return `${where}: missing target`
  const tg = target as Target
  if (tg.nodeType !== undefined && !NODE_TYPES.includes(tg.nodeType)) {
    return `${where}: unknown nodeType "${String(tg.nodeType)}"`
  }
  if (
    tg.headingLevel !== undefined &&
    (!Number.isInteger(tg.headingLevel) || tg.headingLevel < 1 || tg.headingLevel > 6)
  ) {
    return `${where}: headingLevel must be an integer between 1 and 6`
  }
  if (
    tg.blockIndexes !== undefined &&
    (!Array.isArray(tg.blockIndexes) || tg.blockIndexes.some((i) => !Number.isInteger(i) || i < 0))
  ) {
    return `${where}: blockIndexes must be an array of non-negative integers`
  }
  if (tg.range !== undefined) {
    const r = tg.range as { from?: unknown; to?: unknown } | null
    if (!r || typeof r.from !== 'number' || typeof r.to !== 'number' || r.from > r.to) {
      return `${where}: range must be { from, to } positions with from ≤ to`
    }
  }
  const hasCondition =
    tg.nodeType !== undefined ||
    tg.headingLevel !== undefined ||
    typeof tg.containsText === 'string' ||
    (Array.isArray(tg.blockIndexes) && tg.blockIndexes.length > 0) ||
    tg.scope === 'selection' ||
    tg.range !== undefined
  if (!hasCondition) return `${where}: target requires at least one condition`
  return null
}

function validateShape(op: Op, def: OpDef, where: string): string | null {
  if (def.target === 'none' && op.target !== undefined) {
    return `${where}: does not take a target`
  }
  if (def.target === 'required' || op.target !== undefined) {
    const error = validateTarget(op.target, where)
    if (error) return error
  }
  const unknown = Object.keys(op).filter(
    (k) => k !== 'op' && k !== 'target' && !def.keys.includes(k),
  )
  if (unknown.length > 0) {
    return `${where}: unknown field(s) ${unknown.join(', ')}; allowed: ${def.keys.join(', ')}`
  }
  return null
}

const present = (op: Op, keys: readonly string[]): string[] =>
  keys.filter((k) => op[k] !== undefined)

const HEX = /^#?([0-9a-f]{6})$/i

/** "#RRGGBB" or "RRGGBB" -> "RRGGBB"; null passes through */
function normalizeHex(value: unknown, where: string, field: string): string | null {
  if (value === null) return null
  const m = typeof value === 'string' ? HEX.exec(value.trim()) : null
  if (!m)
    throw new Error(
      `${where}: ${field} must be a 6-digit hex color like "#1A73E8" (or null to clear)`,
    )
  return m[1].toUpperCase()
}

function checkHex(value: unknown, where: string, field: string): string | null {
  try {
    normalizeHex(value, where, field)
    return null
  } catch (e) {
    return e instanceof Error ? e.message : String(e)
  }
}

const isNumberOrNull = (v: unknown) => v === null || (typeof v === 'number' && Number.isFinite(v))
/** a present number must be positive; null (= clear) passes */
const isPositiveOrNull = (v: unknown) =>
  v === null || (typeof v === 'number' && Number.isFinite(v) && v > 0)
const isStringOrNull = (v: unknown) => v === null || typeof v === 'string'

function validateFontFields(op: Op, where: string): string | null {
  for (const k of ['bold', 'italic', 'underline', 'strike'] as const) {
    if (op[k] !== undefined && typeof op[k] !== 'boolean') return `${where}: ${k} must be a boolean`
  }
  if (op.color !== undefined) {
    const error = checkHex(op.color, where, 'color')
    if (error) return error
  }
  if (op.highlight !== undefined && !isStringOrNull(op.highlight)) {
    return `${where}: highlight must be a color name / hex string or null`
  }
  if (op.fontSize !== undefined && !isPositiveOrNull(op.fontSize)) {
    return `${where}: fontSize must be a positive number of points (or null to clear)`
  }
  if (op.fontFamily !== undefined && !isStringOrNull(op.fontFamily)) {
    return `${where}: fontFamily must be a string or null`
  }
  if (
    op.baseline !== undefined &&
    op.baseline !== null &&
    !(BASELINES as readonly string[]).includes(String(op.baseline))
  ) {
    return `${where}: baseline must be superscript / subscript / none (or null)`
  }
  if (op.link !== undefined && op.link !== null) {
    const link = op.link as { url?: unknown }
    if (!link || typeof link !== 'object' || typeof link.url !== 'string') {
      return `${where}: link must be { url } or null`
    }
  }
  if (present(op, FONT_KEYS).length === 0) {
    return `${where}: give at least one font field (${FONT_KEYS.join(', ')})`
  }
  return null
}

function validateParagraphFields(op: Op, where: string): string | null {
  if (
    op.align !== undefined &&
    op.align !== null &&
    !(ALIGNS as readonly string[]).includes(String(op.align))
  ) {
    return `${where}: align must be left / center / right / justify (or null)`
  }
  for (const k of [
    'lineSpacing',
    'indentLeft',
    'indentRight',
    'indentFirstLine',
    'spaceBefore',
    'spaceAfter',
  ] as const) {
    if (op[k] !== undefined && !isNumberOrNull(op[k]))
      return `${where}: ${k} must be a number or null`
  }
  if (op.pageBreakBefore !== undefined && typeof op.pageBreakBefore !== 'boolean') {
    return `${where}: pageBreakBefore must be a boolean`
  }
  if (op.shadingFill !== undefined) {
    const error = checkHex(op.shadingFill, where, 'shadingFill')
    if (error) return error
  }
  if (op.borders !== undefined && op.borders !== null) {
    if (typeof op.borders !== 'string' || !/^[tblr]*$/.test(op.borders)) {
      return `${where}: borders must be a subset of "tblr" (or null)`
    }
  }
  if (present(op, PARA_KEYS).length === 0) {
    return `${where}: give at least one paragraph field (${PARA_KEYS.join(', ')})`
  }
  return null
}

function validateAfterBlockIndex(op: Op, where: string): string | null {
  if (!Number.isInteger(op.afterBlockIndex) || Number(op.afterBlockIndex) < -1) {
    return `${where}: afterBlockIndex must be an integer ≥ -1 (-1 = start of document)`
  }
  return null
}

// ---- target matching ----

interface TopBlock {
  index: number
  pos: number
  node: PmDocNode
  /** absolute range of the selected characters inside this block (partial selection); run-level ops stay inside it */
  clip?: SelRange
}

function topLevelBlocks(doc: PmDocNode): TopBlock[] {
  const out: TopBlock[] = []
  let index = 0
  doc.forEach((node, offset) => {
    out.push({ index, pos: offset, node })
    index++
  })
  return out
}

function blockText(node: PmDocNode): string {
  if (node.type.name === 'docProtected') {
    return [node.attrs.label, node.attrs.previewText].filter(Boolean).join(' ')
  }
  // match against current content only, or targets would hit already-deleted revision text
  return liveText(node)
}

function matchTarget(doc: PmDocNode, target: Target, sel: SelRange): TopBlock[] {
  return topLevelBlocks(doc).filter((b) => {
    if (target.blockIndexes && !target.blockIndexes.includes(b.index)) return false
    if (target.nodeType === 'image') {
      if (b.node.type.name !== 'docProtected' || b.node.attrs.blockType !== 'image') return false
    } else if (target.nodeType && b.node.type.name !== target.nodeType) {
      return false
    }
    if (target.headingLevel !== undefined) {
      if (b.node.type.name !== 'docHeading') return false
      if (Number(b.node.attrs.level) !== target.headingLevel) return false
    }
    if (target.containsText) {
      const text = blockText(b.node)
      const hit =
        target.matchCase === false
          ? text.toLowerCase().includes(target.containsText.toLowerCase())
          : text.includes(target.containsText)
      if (!hit) return false
    }
    const scoped = scopedRange(target, sel)
    if (scoped) {
      const from = b.pos
      const to = b.pos + b.node.nodeSize
      const overlaps = to > scoped.from && from < scoped.to
      const caretInside = scoped.from === scoped.to && scoped.from >= from && scoped.from <= to
      if (!overlaps && !caretInside) return false
      if (overlaps && scoped.from < scoped.to && b.node.isTextblock) {
        const contentFrom = from + 1
        const contentTo = contentFrom + b.node.content.size
        const clipFrom = Math.max(scoped.from, contentFrom)
        const clipTo = Math.min(scoped.to, contentTo)
        if (clipFrom > contentFrom || clipTo < contentTo) b.clip = { from: clipFrom, to: clipTo }
      }
    }
    return true
  })
}

/** the range a selection-style target resolves against: an explicit range, else the (frozen or live) selection */
function scopedRange(target: Target, sel: SelRange): SelRange | null {
  if (target.range) return target.range
  return target.scope === 'selection' ? sel : null
}

const PARAGRAPH_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

interface ParagraphHit {
  node: PmDocNode
  pos: number
}

/**
 * Paragraph-like nodes an op formats inside a matched block: the block itself
 * when it is a text block; for containers (tables) the nested paragraphs, kept
 * to the scoped range when the target is selection-based — the ribbon's
 * "align" inside a table cell must reach that cell's paragraph only.
 */
function paragraphsIn(b: TopBlock, scoped: SelRange | null): ParagraphHit[] {
  if (b.node.type.name === 'docProtected') return []
  if (b.node.isTextblock) return [{ node: b.node, pos: b.pos }]
  const hits: ParagraphHit[] = []
  b.node.descendants((node, offset) => {
    if (!PARAGRAPH_TYPES.has(node.type.name)) return true
    const pos = b.pos + 1 + offset
    const end = pos + node.nodeSize
    if (scoped) {
      const overlaps = end > scoped.from && pos < scoped.to
      const caretInside = scoped.from === scoped.to && scoped.from >= pos && scoped.from <= end
      if (!overlaps && !caretInside) return false
    }
    hits.push({ node, pos })
    return false
  })
  return hits
}

/** true when [from, to) lies inside the block's selected span (no clip = whole block is fair game) */
function insideClip(b: TopBlock, from: number, to: number): boolean {
  return !b.clip || (from >= b.clip.from && to <= b.clip.to)
}

const targetOf = (op: Op): Target => op.target as Target

// ---- executors (each reads tr.doc fresh, so sequential ops compose) ----

const BOOL_MARK_TYPES: Record<string, string> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strike: 'strike',
}

/** attrs for a block the op changed: the AI highlight rides along unless the caller opted out */
function changedAttrs(ctx: OpContext, attrs: Record<string, unknown>): Record<string, unknown> {
  return wantsAiMark(ctx) ? { ...attrs, aiChanged: true } : attrs
}

function markChanged(tr: Transaction, pos: number, ctx: OpContext): void {
  if (!wantsAiMark(ctx)) return
  const node = tr.doc.nodeAt(pos)
  if (node) tr.setNodeMarkup(pos, undefined, { ...node.attrs, aiChanged: true })
}

/**
 * Patch paragraph attrs on every paragraph the matched blocks contain. The one
 * implementation behind the model's setParagraphFormat and the ribbon's
 * setParagraphAttrs; `imageAlign` (UI alignment) also lands on matched image
 * blocks as their w:jc. An explicit spacing value turns Word's "Auto" spacing off.
 */
function applyParagraphPatch(
  env: RunEnv,
  op: Op,
  patch: Record<string, unknown>,
  imageAlign?: unknown,
): OpResult {
  const { tr, ctx, sel } = env
  const target = targetOf(op)
  const matched = matchTarget(tr.doc, target, sel)
  const scoped = scopedRange(target, sel)
  if ('spaceBefore' in patch && !('spaceBeforeAuto' in patch)) patch.spaceBeforeAuto = false
  if ('spaceAfter' in patch && !('spaceAfterAuto' in patch)) patch.spaceAfterAuto = false
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      if (imageAlign !== undefined && b.node.attrs.blockType === 'image') {
        const align = imageAlign ?? null
        if (b.node.attrs.imageAlign !== align) {
          tr.setNodeMarkup(b.pos, undefined, { ...b.node.attrs, imageAlign: align })
          changed++
        }
      } else {
        skippedProtected++
      }
      continue
    }
    for (const p of paragraphsIn(b, scoped)) {
      const attrs = { ...p.node.attrs, ...patch }
      const dirty = Object.keys(patch).some((k) => attrs[k] !== p.node.attrs[k])
      if (!dirty) continue
      tr.setNodeMarkup(p.pos, undefined, changedAttrs(ctx, attrs))
      changed++
    }
  }
  return { op: op.op, matched: matched.length, changed, skippedProtected }
}

/** patch applied onto each text node's existing docTextStyle attrs (fonts route to their rFonts slot) */
function buildTextMarkPatch(op: Op): Record<string, unknown> {
  const patch: Record<string, unknown> = {}
  if (op.fontFamily !== undefined) {
    // route to the matching rFonts slot; clearing clears both
    const v = op.fontFamily as string | null
    if (!v) Object.assign(patch, { font: null, fontAscii: null })
    else if (isEastAsianFontName(v)) patch.font = v
    else patch.fontAscii = v
  }
  if (op.color !== undefined) patch.color = normalizeHex(op.color, op.op, 'color')
  if (op.highlight !== undefined) patch.highlight = op.highlight ?? null
  if (op.fontSize !== undefined) {
    patch.sizeHalfPoints = op.fontSize === null ? null : Math.round((op.fontSize as number) * 2)
  }
  if (op.baseline !== undefined) {
    const v = op.baseline
    patch.vertAlign = v === 'superscript' || v === 'subscript' ? v : null
  }
  return patch
}

/** apply the op's mark-level font fields to [from, to) of a block's text nodes */
function applyFontRange(
  env: RunEnv,
  op: Op,
  boolKeys: string[],
  markPatch: Record<string, unknown>,
  from: number,
  to: number,
  node: PmDocNode,
  contentFrom: number,
): void {
  const { tr, schema } = env
  for (const k of boolKeys) {
    const markType = schema.marks[BOOL_MARK_TYPES[k]]
    if (op[k]) tr.addMark(from, to, markType.create())
    else tr.removeMark(from, to, markType)
  }
  if (op.link !== undefined) {
    const url = (op.link as { url?: string } | null)?.url
    if (url) tr.addMark(from, to, schema.marks.link.create({ href: url, rId: null }))
    else tr.removeMark(from, to, schema.marks.link)
  }
  if (Object.keys(markPatch).length === 0) return
  // merge per text node: only the given attrs change, the rest survive
  node.forEach((child, offset) => {
    if (!child.isText) return
    const childFrom = Math.max(contentFrom + offset, from)
    const childTo = Math.min(contentFrom + offset + child.nodeSize, to)
    if (childFrom >= childTo) return
    const existing = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs ?? {}
    // spread keeps every non-listed attr (fontAscii/styleId/rawRPr…) alive
    const merged: Record<string, unknown> = { ...existing, ...markPatch }
    const empty = Object.values(merged).every((v) => v === null)
    if (empty) tr.removeMark(childFrom, childTo, schema.marks.docTextStyle)
    else tr.addMark(childFrom, childTo, schema.marks.docTextStyle.create(merged))
  })
}

function runSetFont(op: Op, env: RunEnv): OpResult {
  const { tr, ctx, sel } = env
  const matched = matchTarget(tr.doc, targetOf(op), sel)
  const boolKeys = present(op, Object.keys(BOOL_MARK_TYPES))
  const markPatch = buildTextMarkPatch(op)
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    const before = b.node
    const contentFrom = b.pos + 1
    const from = b.clip?.from ?? contentFrom
    const to = b.clip?.to ?? contentFrom + b.node.content.size
    if (from < to) applyFontRange(env, op, boolKeys, markPatch, from, to, b.node, contentFrom)
    const after = tr.doc.nodeAt(b.pos)
    if (after && !after.eq(before)) {
      changed++
      markChanged(tr, b.pos, ctx)
    }
  }
  return { op: 'setFont', matched: matched.length, changed, skippedProtected }
}

function runSetParagraphFormat(op: Op, env: RunEnv): OpResult {
  const patch: Record<string, unknown> = {}
  for (const k of present(op, PARA_KEYS)) {
    if (k === 'shadingFill') patch.shadingFill = normalizeHex(op.shadingFill, op.op, 'shadingFill')
    else if (k === 'pageBreakBefore') patch.pageBreakBefore = op.pageBreakBefore === true
    else patch[k] = op[k] ?? null
  }
  return applyParagraphPatch(env, op, patch)
}

/** UI: raw paragraph attrs (any schema attr, e.g. lineRule / tabStops / dropCap); align also lands on images */
function runSetParagraphAttrs(op: Op, env: RunEnv): OpResult {
  const attrs = { ...(op.attrs as Record<string, unknown>) }
  return applyParagraphPatch(env, op, attrs, 'align' in attrs ? attrs.align : undefined)
}

/** Word's ribbon indent step: half an inch, in twips */
const INDENT_STEP = 720

/** UI: ribbon increase/decrease indent — list items change level, paragraphs and headings snap to the next half-inch stop */
function runStepIndent(op: Op, env: RunEnv): OpResult {
  const { tr, ctx, sel } = env
  const delta = Number(op.delta)
  const target = targetOf(op)
  const matched = matchTarget(tr.doc, target, sel)
  const scoped = scopedRange(target, sel)
  let changed = 0
  for (const b of matched) {
    for (const p of paragraphsIn(b, scoped)) {
      if (p.node.type.name === 'docListItem') {
        const ilvl = Number(p.node.attrs.ilvl) || 0
        const next = Math.min(Math.max(ilvl + delta, 0), 8)
        if (next === ilvl) continue
        tr.setNodeMarkup(p.pos, undefined, changedAttrs(ctx, { ...p.node.attrs, ilvl: next }))
        changed++
        continue
      }
      const cur = Number(p.node.attrs.indentLeft) || 0
      const next =
        delta > 0
          ? Math.floor(cur / INDENT_STEP) * INDENT_STEP + INDENT_STEP
          : Math.max(Math.ceil(cur / INDENT_STEP) * INDENT_STEP - INDENT_STEP, 0)
      if (next === cur) continue
      tr.setNodeMarkup(
        p.pos,
        undefined,
        changedAttrs(ctx, { ...p.node.attrs, indentLeft: next || null }),
      )
      changed++
    }
  }
  return { op: 'stepIndent', matched: matched.length, changed, skippedProtected: 0 }
}

/**
 * UI: Word's Ctrl+T / Ctrl+Shift+T — a hanging indent moves the left indent one
 * stop right while pulling the first line back by the same amount, so
 * continuation lines sit under the body text; the reverse walks it back and
 * only clears the negative first line once the left indent is home.
 */
function runStepHangingIndent(op: Op, env: RunEnv): OpResult {
  const { tr, ctx, sel } = env
  const delta = Number(op.delta)
  const target = targetOf(op)
  const matched = matchTarget(tr.doc, target, sel)
  const scoped = scopedRange(target, sel)
  let changed = 0
  for (const b of matched) {
    for (const p of paragraphsIn(b, scoped)) {
      const left = Number(p.node.attrs.indentLeft) || 0
      const nextLeft = Math.max(
        delta > 0
          ? Math.floor(left / INDENT_STEP) * INDENT_STEP + INDENT_STEP
          : Math.ceil(left / INDENT_STEP) * INDENT_STEP - INDENT_STEP,
        0,
      )
      const nextHanging = nextLeft > 0 ? -Math.min(nextLeft, INDENT_STEP) : 0
      if (nextLeft === left && nextHanging === (Number(p.node.attrs.indentFirstLine) || 0)) continue
      tr.setNodeMarkup(
        p.pos,
        undefined,
        changedAttrs(ctx, {
          ...p.node.attrs,
          indentLeft: nextLeft || null,
          indentFirstLine: nextHanging || null,
        }),
      )
      changed++
    }
  }
  return { op: 'stepHangingIndent', matched: matched.length, changed, skippedProtected: 0 }
}

function runSetHeadingLevel(op: Op, env: RunEnv): OpResult {
  const { tr, schema, ctx, sel } = env
  const level = Number(op.level)
  const matched = matchTarget(tr.doc, targetOf(op), sel)
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    const isHeading = b.node.type.name === 'docHeading'
    if (level === 0 && b.node.type.name === 'docParagraph') continue
    if (level > 0 && isHeading && Number(b.node.attrs.level) === level) continue
    const type = level === 0 ? schema.nodes.docParagraph : schema.nodes.docHeading
    // styleId is tied to the old block role; generate.ts picks heading styles by level
    const attrs: Record<string, unknown> = { ...b.node.attrs, styleId: null }
    if (level > 0) attrs.level = level
    tr.setNodeMarkup(b.pos, type, changedAttrs(ctx, attrs))
    changed++
  }
  return { op: 'setHeadingLevel', matched: matched.length, changed, skippedProtected }
}

function runFindReplace(op: Op, env: RunEnv): OpResult {
  const { tr, schema, ctx, sel } = env
  const find = String(op.find)
  const replace = String(op.replace ?? '')
  const matchCase = op.matchCase !== false
  const needle = matchCase ? find : find.toLowerCase()
  const blocks = op.target ? matchTarget(tr.doc, targetOf(op), sel) : topLevelBlocks(tr.doc)
  const replacements: Array<{ from: number; to: number; marks: PmDocNode['marks'] }> = []
  const touchedIndexes = new Set<number>()
  let skippedProtected = 0
  let skippedDeleted = 0

  for (const b of blocks) {
    if (b.node.type.name === 'docProtected') {
      if (blockText(b.node).includes(find)) skippedProtected++
      continue
    }
    const blockDeleted = isTrackedDeleted(b.node)
    b.node.forEach((child, offset) => {
      if (!child.isText || !child.text) return
      const struck = blockDeleted || child.marks.some((m) => m.type.name === 'del')
      const hay = matchCase ? child.text : child.text.toLowerCase()
      let at = hay.indexOf(needle)
      while (at !== -1) {
        const from = b.pos + 1 + offset + at
        if (struck) {
          skippedDeleted++
        } else if (insideClip(b, from, from + needle.length)) {
          replacements.push({ from, to: from + needle.length, marks: child.marks })
          touchedIndexes.add(b.index)
        }
        at = hay.indexOf(needle, at + needle.length)
      }
    })
  }

  // apply back-to-front so earlier positions stay valid
  for (const r of [...replacements].reverse()) {
    if (replace) tr.replaceWith(r.from, r.to, schema.text(replace, r.marks))
    else tr.delete(r.from, r.to)
  }
  if (touchedIndexes.size > 0) {
    for (const b of topLevelBlocks(tr.doc)) {
      if (touchedIndexes.has(b.index)) markChanged(tr, b.pos, ctx)
    }
  }
  return {
    op: 'findReplace',
    matched: touchedIndexes.size,
    changed: touchedIndexes.size,
    skippedProtected,
    skippedDeleted,
    detail: t('aiCmdReplacedCount', { count: replacements.length }),
  }
}

function runSetMatchedFont(op: Op, env: RunEnv): OpResult {
  const { tr, ctx, sel } = env
  const text = String(op.text)
  const matchCase = op.matchCase !== false
  const needle = matchCase ? text : text.toLowerCase()
  const blocks = op.target ? matchTarget(tr.doc, targetOf(op), sel) : topLevelBlocks(tr.doc)
  const boolKeys = present(op, Object.keys(BOOL_MARK_TYPES))
  const markPatch = buildTextMarkPatch(op)
  const touched = new Set<number>()
  let styledCount = 0
  let skippedProtected = 0
  let skippedDeleted = 0

  for (const b of blocks) {
    if (b.node.type.name === 'docProtected') {
      if (blockText(b.node).includes(text)) skippedProtected++
      continue
    }
    const blockDeleted = isTrackedDeleted(b.node)
    const contentFrom = b.pos + 1
    b.node.forEach((child, offset) => {
      if (!child.isText || !child.text) return
      const struck = blockDeleted || child.marks.some((m) => m.type.name === 'del')
      const hay = matchCase ? child.text : child.text.toLowerCase()
      let at = hay.indexOf(needle)
      while (at !== -1) {
        const from = contentFrom + offset + at
        const to = from + needle.length
        if (struck) {
          skippedDeleted++
        } else if (insideClip(b, from, to)) {
          applyFontRange(env, op, boolKeys, markPatch, from, to, b.node, contentFrom)
          styledCount++
          touched.add(b.index)
        }
        at = hay.indexOf(needle, at + needle.length)
      }
    })
  }
  // mark steps never shift positions, so block offsets are still valid here
  for (const b of topLevelBlocks(tr.doc)) {
    if (touched.has(b.index)) markChanged(tr, b.pos, ctx)
  }
  return {
    op: 'setMatchedFont',
    matched: touched.size,
    changed: touched.size,
    skippedProtected,
    skippedDeleted,
    detail: String(styledCount),
  }
}

function runDeleteBlocks(op: Op, env: RunEnv): OpResult {
  const { tr, schema, sel } = env
  const blocks = topLevelBlocks(tr.doc)
  const matchedAll = matchTarget(tr.doc, targetOf(op), sel)
  // a block that is already a pending deletion cannot be deleted again;
  // skip and report so the model does not retry forever
  const matched = matchedAll.filter((b) => !isTrackedDeleted(b.node))
  const skippedDeleted = matchedAll.length - matched.length
  if (matched.length === blocks.length && matched.length > 0) {
    // doc requires block+: deleting everything leaves one empty paragraph
    tr.replaceWith(0, tr.doc.content.size, schema.nodes.docParagraph.create())
  } else {
    for (const b of [...matched].sort((a, z) => z.pos - a.pos)) {
      tr.delete(b.pos, b.pos + b.node.nodeSize)
    }
  }
  return {
    op: 'deleteBlocks',
    matched: matchedAll.length,
    changed: matched.length,
    skippedProtected: 0,
    skippedDeleted,
  }
}

function runMoveBlocks(op: Op, env: RunEnv): OpResult {
  const { tr, ctx } = env
  const blockIndexes = op.blockIndexes as number[]
  const afterBlockIndex = Number(op.afterBlockIndex)
  const blocks = topLevelBlocks(tr.doc)
  const movedSet = new Set(blockIndexes)
  const outOfRange = blockIndexes.filter((i) => i >= blocks.length)
  if (outOfRange.length > 0) {
    throw new Error(`moveBlocks: block indexes out of range ${outOfRange.join(', ')}`)
  }
  if (afterBlockIndex >= blocks.length) {
    throw new Error(`moveBlocks: afterBlockIndex out of range ${afterBlockIndex}`)
  }
  if (movedSet.has(afterBlockIndex)) {
    throw new Error('moveBlocks: afterBlockIndex must not be one of the moved blocks')
  }

  // moved blocks keep document order; aiChanged marks the landing highlight
  const moved = [...movedSet]
    .sort((a, z) => a - z)
    .map((i) => {
      const node = blocks[i].node
      if (node.type.name === 'docProtected' || !wantsAiMark(ctx)) return node
      return node.type.create({ ...node.attrs, aiChanged: true }, node.content, node.marks)
    })
  const rest = blocks.filter((b) => !movedSet.has(b.index))
  let insertAt = 0
  if (afterBlockIndex >= 0) {
    insertAt = rest.findIndex((b) => b.index === afterBlockIndex) + 1
  }
  const children = [
    ...rest.slice(0, insertAt).map((b) => b.node),
    ...moved,
    ...rest.slice(insertAt).map((b) => b.node),
  ]
  tr.replaceWith(0, tr.doc.content.size, children)
  return { op: 'moveBlocks', matched: moved.length, changed: moved.length, skippedProtected: 0 }
}

function runSetList(op: Op, env: RunEnv): OpResult {
  const { tr, schema, ctx, sel } = env
  const matched = matchTarget(tr.doc, targetOf(op), sel)
  const kind = op.kind === 'number' ? 'ordered' : 'bullet'
  const numId = ctx.numIds?.[kind] ?? null
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    // headings keep their role; converting them to list items would drop the level
    if (b.node.type.name === 'docHeading') continue
    if (b.node.type.name === 'docListItem') {
      if (b.node.attrs.kind === kind) continue
      tr.setNodeMarkup(
        b.pos,
        undefined,
        changedAttrs(ctx, { ...b.node.attrs, kind, numId: numId ?? b.node.attrs.numId }),
      )
      changed++
      continue
    }
    tr.setNodeMarkup(
      b.pos,
      schema.nodes.docListItem,
      changedAttrs(ctx, { ...b.node.attrs, styleId: null, kind, numId, ilvl: 0 }),
    )
    changed++
  }
  return { op: 'setList', matched: matched.length, changed, skippedProtected }
}

function runClearList(op: Op, env: RunEnv): OpResult {
  const { tr, schema, ctx, sel } = env
  const matched = matchTarget(tr.doc, targetOf(op), sel)
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    if (b.node.type.name !== 'docListItem') continue
    tr.setNodeMarkup(
      b.pos,
      schema.nodes.docParagraph,
      changedAttrs(ctx, { ...b.node.attrs, styleId: null }),
    )
    changed++
  }
  return { op: 'clearList', matched: matched.length, changed, skippedProtected }
}

function runSetImageProperties(op: Op, env: RunEnv): OpResult {
  const { tr, sel } = env
  const matched = matchTarget(tr.doc, targetOf(op), sel)
  let changed = 0
  for (const b of matched) {
    if (b.node.type.name !== 'docProtected' || b.node.attrs.blockType !== 'image') continue
    const attrs: Record<string, unknown> = { ...b.node.attrs }
    if (op.align !== undefined) attrs.imageAlign = op.align ?? null
    const wantW = (op.widthPx as number | null | undefined) ?? null
    const wantH = (op.heightPx as number | null | undefined) ?? null
    const curW = Number(b.node.attrs.imageWidthPx) || null
    const curH = Number(b.node.attrs.imageHeightPx) || null
    if (wantW && wantH) {
      attrs.imageWidthPx = Math.round(wantW)
      attrs.imageHeightPx = Math.round(wantH)
    } else if (wantW && curW && curH) {
      attrs.imageWidthPx = Math.round(wantW)
      attrs.imageHeightPx = Math.round((wantW * curH) / curW)
    } else if (wantH && curW && curH) {
      attrs.imageHeightPx = Math.round(wantH)
      attrs.imageWidthPx = Math.round((wantH * curW) / curH)
    }
    // save-side patching needs both dimensions; a lone dimension with unknown
    // original size cannot be scaled and is skipped
    const dirty = Object.keys(attrs).some((k) => attrs[k] !== b.node.attrs[k])
    if (!dirty) continue
    tr.setNodeMarkup(b.pos, undefined, attrs)
    changed++
  }
  return { op: 'setImageProperties', matched: matched.length, changed, skippedProtected: 0 }
}

function runInsertToc(op: Op, env: RunEnv): OpResult {
  const { tr, schema } = env
  const afterBlockIndex = Number(op.afterBlockIndex)
  const blocks = topLevelBlocks(tr.doc)
  const entries: TocEntry[] = []
  for (const b of blocks) {
    if (b.node.type.name === 'docHeading' && b.node.textContent.trim()) {
      entries.push({ level: Number(b.node.attrs.level) || 1, text: b.node.textContent })
    }
  }
  if (entries.length === 0) {
    throw new Error(
      'insertToc: the document has no headings (levels 1-6), cannot generate a table of contents',
    )
  }
  if (afterBlockIndex >= blocks.length) {
    throw new Error(`insertToc: afterBlockIndex out of range ${afterBlockIndex}`)
  }
  // one protected block per TOC line (same shape as the ribbon's TOC button);
  // the begin fldChar is dirty so Word recalculates page numbers on open
  const nodes = generateTocFieldXml(entries).map((xml, i) =>
    schema.nodes.docProtected.create({
      docxIndex: null,
      blockType: 'passthrough',
      label: 'TOC field',
      genXml: xml,
      fieldDisplay: { kind: 'tocLine', left: entries[i].text, right: '', level: entries[i].level },
    }),
  )
  const pos =
    afterBlockIndex === -1 ? 0 : blocks[afterBlockIndex].pos + blocks[afterBlockIndex].node.nodeSize
  tr.insert(pos, nodes)
  return { op: 'insertToc', matched: entries.length, changed: nodes.length, skippedProtected: 0 }
}

// ---- op definitions ----

const FONT_SIGNATURE_FIELDS =
  'bold?, italic?, underline?, strike?, color?: "#RRGGBB"|null, highlight?, fontSize?: pt|null, fontFamily?: string|null, baseline?: "superscript"|"subscript"|"none"|null, link?: {url}|null'

register({
  name: 'setFont',
  signature: `{ op: "setFont", target, ${FONT_SIGNATURE_FIELDS} }  // run-level style of the targeted blocks (only the selected characters when the target is a partial selection); a present field is set, null clears it, absent fields stay; fontFamily applies to its own script slot (an East Asian font keeps the run's Latin font and vice versa)`,
  keys: FONT_KEYS,
  target: 'required',
  validate(op, where) {
    return validateShape(op, this, where) ?? validateFontFields(op, where)
  },
  apply: runSetFont,
})

register({
  name: 'setMatchedFont',
  signature: `{ op: "setMatchedFont", text, matchCase?, target?, ${FONT_SIGNATURE_FIELDS} }  // character-level: styles only the occurrences of text (not the whole block); target narrows the scanned blocks, omitted = whole document`,
  keys: ['text', 'matchCase', ...FONT_KEYS],
  target: 'optional',
  validate(op, where) {
    const shape = validateShape(op, this, where)
    if (shape) return shape
    if (typeof op.text !== 'string' || op.text === '') return `${where}: text must not be empty`
    if (op.matchCase !== undefined && typeof op.matchCase !== 'boolean') {
      return `${where}: matchCase must be a boolean`
    }
    return validateFontFields(op, where)
  },
  apply: runSetMatchedFont,
})

register({
  name: 'setParagraphFormat',
  signature:
    '{ op: "setParagraphFormat", target, align?: "left"|"center"|"right"|"justify"|null, lineSpacing?, indentLeft?, indentRight?, indentFirstLine?, spaceBefore?, spaceAfter?, pageBreakBefore?: boolean, shadingFill?: "#RRGGBB"|null, borders?: subset of "tblr"|null }  // paragraph format of the targeted blocks (whole blocks, even for a partial selection); lengths in twips, lineSpacing as a multiple; indentFirstLine positive = first-line indent, negative = hanging; a two-character CJK indent ≈ font size in pt × 40 twips',
  keys: PARA_KEYS,
  target: 'required',
  validate(op, where) {
    return validateShape(op, this, where) ?? validateParagraphFields(op, where)
  },
  apply: runSetParagraphFormat,
})

register({
  name: 'setHeadingLevel',
  signature: '{ op: "setHeadingLevel", target, level: 0-6 }  // 0 = demote to body paragraph',
  keys: ['level'],
  target: 'required',
  validate(op, where) {
    const shape = validateShape(op, this, where)
    if (shape) return shape
    if (!Number.isInteger(op.level) || Number(op.level) < 0 || Number(op.level) > 6) {
      return `${where}: level must be an integer between 0 and 6`
    }
    return null
  },
  apply: runSetHeadingLevel,
})

register({
  name: 'findReplace',
  signature:
    '{ op: "findReplace", find, replace, matchCase?, target? }  // plain-text find & replace inside runs (no cross-block matching); the replacement keeps the matched text\'s formatting — the right op for small in-place fixes; target narrows the scanned blocks (blockIndexes / scope:"selection" / nodeType), omitted = whole document',
  keys: ['find', 'replace', 'matchCase'],
  target: 'optional',
  validate(op, where) {
    const shape = validateShape(op, this, where)
    if (shape) return shape
    if (typeof op.find !== 'string' || op.find === '') return `${where}: find must not be empty`
    if (typeof op.replace !== 'string')
      return `${where}: replace must be a string (empty deletes the match)`
    if (op.matchCase !== undefined && typeof op.matchCase !== 'boolean') {
      return `${where}: matchCase must be a boolean`
    }
    return null
  },
  apply: runFindReplace,
})

register({
  name: 'deleteBlocks',
  signature: '{ op: "deleteBlocks", target }',
  keys: [],
  target: 'required',
  validate(op, where) {
    return validateShape(op, this, where)
  },
  apply: runDeleteBlocks,
})

register({
  name: 'moveBlocks',
  signature:
    '{ op: "moveBlocks", blockIndexes: number[], afterBlockIndex }  // move the blocks (relative order kept) after afterBlockIndex; -1 = start of document',
  keys: ['blockIndexes', 'afterBlockIndex'],
  target: 'none',
  validate(op, where) {
    const shape = validateShape(op, this, where)
    if (shape) return shape
    const indexes = op.blockIndexes
    if (
      !Array.isArray(indexes) ||
      indexes.length === 0 ||
      indexes.some((i) => !Number.isInteger(i) || Number(i) < 0)
    ) {
      return `${where}: blockIndexes must be a non-empty array of non-negative integers`
    }
    return validateAfterBlockIndex(op, where)
  },
  apply: runMoveBlocks,
})

register({
  name: 'setList',
  signature:
    '{ op: "setList", target, kind?: "bullet"|"number" }  // paragraphs to list items (default bullet); heading blocks are not converted',
  keys: ['kind'],
  target: 'required',
  validate(op, where) {
    const shape = validateShape(op, this, where)
    if (shape) return shape
    if (op.kind !== undefined && op.kind !== 'bullet' && op.kind !== 'number') {
      return `${where}: kind must be "bullet" or "number"`
    }
    return null
  },
  apply: runSetList,
})

register({
  name: 'clearList',
  signature: '{ op: "clearList", target }  // list items back to body paragraphs',
  keys: [],
  target: 'required',
  validate(op, where) {
    return validateShape(op, this, where)
  },
  apply: runClearList,
})

register({
  name: 'setImageProperties',
  signature:
    '{ op: "setImageProperties", target, widthPx?, heightPx?, align?: "left"|"center"|"right"|null }  // image blocks only (target nodeType "image"); giving one dimension scales the other proportionally',
  keys: IMAGE_KEYS,
  target: 'required',
  validate(op, where) {
    const shape = validateShape(op, this, where)
    if (shape) return shape
    for (const k of ['widthPx', 'heightPx'] as const) {
      if (op[k] !== undefined && !isPositiveOrNull(op[k])) {
        return `${where}: ${k} must be a positive number (or null)`
      }
    }
    if (
      op.align !== undefined &&
      op.align !== null &&
      !['left', 'center', 'right'].includes(String(op.align))
    ) {
      return `${where}: align must be left / center / right (or null)`
    }
    if (present(op, IMAGE_KEYS).length === 0) {
      return `${where}: give at least one of ${IMAGE_KEYS.join(', ')}`
    }
    return null
  },
  apply: runSetImageProperties,
})

register({
  name: 'insertToc',
  signature:
    '{ op: "insertToc", afterBlockIndex }  // insert a TOC field after that block (-1 = start of document); entries come from the current headings, Word computes page numbers on open; fails if the document has no headings',
  keys: ['afterBlockIndex'],
  target: 'none',
  validate(op, where) {
    return validateShape(op, this, where) ?? validateAfterBlockIndex(op, where)
  },
  apply: runInsertToc,
})

// ---- UI-only ops (same executor, same target resolution, never offered to the model) ----

const validDelta = (op: Op, where: string): string | null =>
  op.delta === 1 || op.delta === -1 ? null : `${where}: delta must be 1 or -1`

register({
  name: 'setParagraphAttrs',
  signature: '{ op: "setParagraphAttrs", target, attrs: Record<string, unknown> }',
  keys: ['attrs'],
  target: 'required',
  hidden: true,
  validate(op, where) {
    const shape = validateShape(op, this, where)
    if (shape) return shape
    if (!op.attrs || typeof op.attrs !== 'object' || Object.keys(op.attrs).length === 0) {
      return `${where}: attrs must be a non-empty object`
    }
    return null
  },
  apply: runSetParagraphAttrs,
})

register({
  name: 'stepIndent',
  signature: '{ op: "stepIndent", target, delta: 1|-1 }',
  keys: ['delta'],
  target: 'required',
  hidden: true,
  validate(op, where) {
    return validateShape(op, this, where) ?? validDelta(op, where)
  },
  apply: runStepIndent,
})

register({
  name: 'stepHangingIndent',
  signature: '{ op: "stepHangingIndent", target, delta: 1|-1 }',
  keys: ['delta'],
  target: 'required',
  hidden: true,
  validate(op, where) {
    return validateShape(op, this, where) ?? validDelta(op, where)
  },
  apply: runStepHangingIndent,
})

// ---- entry point ----

/** frozen scope -> current PM positions: exact positions when captured, else the block range (indexes clamped: AI edits earlier in the run may have shrunk the doc) */
function frozenSelectionRange(editor: Editor, frozen: OpContext['selection']): SelRange | null {
  if (!frozen) return null
  if (frozen.from !== undefined && frozen.to !== undefined) {
    const size = editor.state.doc.content.size
    const from = Math.min(Math.max(frozen.from, 0), size)
    return { from, to: Math.min(Math.max(frozen.to, from), size) }
  }
  const last = editor.state.doc.childCount - 1
  const start = Math.min(Math.max(frozen.startIndex, 0), last)
  const end = Math.min(Math.max(frozen.endIndex, start), last)
  return blockRangePositions(editor, start, end)
}

function summarize(results: OpResult[]): string {
  const total = results.reduce((sum, r) => sum + r.changed, 0)
  if (total === 0) {
    const skipped = results.reduce((sum, r) => sum + r.skippedProtected, 0)
    return skipped > 0 ? t('aiCmdNoneSkipped', { count: skipped }) : t('aiCmdNone')
  }
  const parts = results.map((r) => {
    let part: string
    switch (r.op) {
      case 'setFont':
        part = t('aiCmdTextStyle', { count: r.changed })
        break
      case 'setMatchedFont':
        part = t('aiCmdMatchedStyle', { count: r.detail ?? '0', blocks: r.changed })
        break
      case 'setParagraphFormat':
        part = t('aiCmdParaStyle', { count: r.changed })
        break
      case 'setHeadingLevel':
        part = t('aiCmdHeading', { count: r.changed })
        break
      case 'findReplace':
        part = t('aiCmdReplaceAll', {
          count: r.changed,
          detail: r.detail ?? t('aiCmdReplaceFallback'),
        })
        break
      case 'deleteBlocks':
        part = t('aiCmdDeleted', { count: r.changed })
        break
      case 'moveBlocks':
        part = t('aiCmdMoved', { count: r.changed })
        break
      case 'setList':
        part = t('aiCmdToList', { count: r.changed })
        break
      case 'clearList':
        part = t('aiCmdToBody', { count: r.changed })
        break
      case 'setImageProperties':
        part = t('aiCmdImages', { count: r.changed })
        break
      case 'insertToc':
        part = t('aiCmdToc', { count: r.matched })
        break
      default:
        part = `${r.op}: ${r.changed}`
    }
    if (r.skippedProtected > 0) part += t('aiCmdSkipped', { count: r.skippedProtected })
    return part
  })
  return parts.join(';')
}

const failure = (error: string): ExecuteOutcome => ({ ok: false, results: [], summary: '', error })

/**
 * Validate every op, then apply the batch as one transaction. Any invalid op
 * rejects the whole batch (atomic) with every error listed, each followed by
 * the op's usage line; dryRun stops after planning.
 */
export function executeOps(editor: Editor, ops: unknown, ctx: OpContext = {}): ExecuteOutcome {
  if (!Array.isArray(ops) || ops.length === 0) return failure('ops must be a non-empty array')
  const plan: string[] = []
  const errors: string[] = []
  const planned: Array<[Op, OpDef]> = []
  ops.forEach((raw, i) => {
    const where = `op #${i}`
    if (!raw || typeof raw !== 'object' || typeof (raw as Op).op !== 'string') {
      errors.push(`${where}: must be an object with an "op" name field`)
      return
    }
    const op = raw as Op
    const def = REGISTRY.get(op.op)
    if (!def || (def.hidden && ctx.source !== 'ui')) {
      errors.push(`${where}: unknown op "${op.op}". Supported ops: [${opNames().join(', ')}]`)
      return
    }
    const error = def.validate(op, `${where} ${op.op}`)
    if (error) {
      errors.push(`${error}\n  usage: ${def.signature}`)
      return
    }
    planned.push([op, def])
    plan.push(`#${i} ${op.op}${op.target ? ` ${JSON.stringify(op.target)}` : ''}`)
  })
  if (errors.length > 0) {
    return failure(
      `${errors.join('\n')}\nNothing was applied (atomic): fix the op(s) and resend the whole batch.`,
    )
  }
  if (ctx.dryRun) return { ok: true, results: [], summary: '', plan }

  const tr = editor.state.tr
  const schema = editor.schema
  const selection = frozenSelectionRange(editor, ctx.selection) ?? editor.state.selection
  const results: OpResult[] = []
  try {
    for (const [op, def] of planned) {
      // selection positions mapped through the steps applied so far
      const sel: SelRange = {
        from: tr.mapping.map(selection.from),
        to: tr.mapping.map(selection.to),
      }
      results.push(def.apply(op, { tr, schema, sel, ctx }))
    }
  } catch (e) {
    return failure(e instanceof Error ? e.message : String(e))
  }

  if (tr.steps.length > 0) {
    // intent record for consumers downstream of the Step stream (journal, analytics)
    tr.setMeta('docsOps', { source: ctx.source ?? 'ai', ops: planned.map(([op]) => op.op) })
    // route through the TrackChanges recorder so AI format/text edits become
    // reviewable revisions under the AI author (same pipeline as manual edits)
    const storage = editor.storage.trackChanges as { enabled: boolean; author: string } | undefined
    if (ctx.track && storage) {
      const prev = { enabled: storage.enabled, author: storage.author }
      storage.enabled = true
      storage.author = ctx.track.author
      try {
        editor.view.dispatch(tr)
      } finally {
        storage.enabled = prev.enabled
        storage.author = prev.author
      }
    } else {
      editor.view.dispatch(tr)
    }
  }
  return { ok: true, results, summary: summarize(results), plan }
}

/**
 * The app's own UI (ribbon, shortcuts, dialogs) issues ops through here: no AI
 * highlight, UI-only ops allowed, focus restored like a tiptap chain would.
 * Returns true when something changed. A validation failure is a programming
 * error in the caller, so it is logged rather than surfaced.
 */
export function runUiOps(editor: Editor, ops: Op[], opts: { focus?: boolean } = {}): boolean {
  const outcome = executeOps(editor, ops, { source: 'ui' })
  if (!outcome.ok) {
    console.error('runUiOps rejected', outcome.error)
    return false
  }
  if (opts.focus !== false) editor.commands.focus()
  return outcome.results.some((r) => r.changed > 0)
}
