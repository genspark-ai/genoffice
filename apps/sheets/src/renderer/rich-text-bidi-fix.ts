/**
 * Univer's rich-text layout has no bidi: glyphs are painted one by one at
 * x positions accumulated in logical order, and its ArabicHandler stores each
 * contiguous Arabic chunk REVERSED (the canvas re-reverses it at draw time),
 * so RTL rich cells (`cell.p`) render mirrored. The document model must keep
 * the logical stream — edits, sorts (journalRangeSnapshot) and saves read
 * `cell.p` back — so this fix touches layout artifacts only: after every
 * DocumentSkeleton.calculate of a 'rich-cell' document it restores each
 * glyph's logical content and re-seats glyph x positions in UAX#9 visual
 * order.
 *
 * The in-cell editor (and the formula bar) clone the body into their own
 * '__INTERNAL_EDITOR__*' documents, laid out by the same DocumentSkeleton.
 * They get the same reorder, plus the two places that turn glyph geometry
 * into caret semantics: the caret/selection rectangles
 * (NodePositionConvertToCursor) and the click-side decision
 * (DocSelectionRenderService._getNodePosition). Positions stay logical
 * (glyph index + isBack) — only their pixel mapping is mirrored for glyphs
 * that ended up at an odd (right-to-left) embedding level. Arrow keys still
 * move logically through the stream.
 */
import { DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY, DOCS_NORMAL_EDITOR_UNIT_ID_KEY } from '@univerjs/core'
import { DocSelectionRenderService, NodePositionConvertToCursor } from '@univerjs/docs-ui'
import { DocumentSkeleton, FontCache, GlyphType } from '@univerjs/engine-render'

import { resolveBidiDirection } from './rtl-text-fix'

export type BidiClass = 'L' | 'R' | 'AL' | 'EN' | 'AN' | 'ES' | 'ET' | 'CS' | 'NSM' | 'WS' | 'ON'

const HAS_RTL =
  /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}\u200F\u061C]/u
// Bidi class AL also covers three Script=Common Arabic punctuation/letter
// characters: semicolon, question mark and tatweel (the "هـ" Hijri suffix —
// a modifier letter that \p{L} would otherwise make strong LTR).
const AL_LETTER = /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\u061B\u061F\u0640]/u
const R_LETTER =
  /[\p{Script=Hebrew}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}]/u
const EN_CHAR = /[0-9\u06F0-\u06F9]/
const AN_CHAR = /[\u0660-\u0669\u066B\u066C]/
const ES_CHAR = /[+\-\u2212]/
const ET_CHAR = /[#$%\u00A2-\u00A5\u066A\u00B0\u2030\u2031\u20A0-\u20BF]/
const CS_CHAR = /[,.:/\u00A0\u060C\u2044]/

// Mirrors Univer's hasArabic() in language-ruler.ts — exactly the chunks its
// ArabicHandler stored in reversed character order. The lone U+0750 (instead
// of U+0750-U+077F) is Univer's own quirk, kept 1:1: Arabic Supplement chars
// never enter a reversed chunk, so widening this set would desync it from
// what actually needs un-reversing.
const UNIVER_ARABIC_ONLY = /^[\u0600-\u06FF\u0750\u0870-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]{2,}$/

const MIRROR: Record<string, string> = {
  '(': ')',
  ')': '(',
  '[': ']',
  ']': '[',
  '{': '}',
  '}': '{',
  '<': '>',
  '>': '<',
  '«': '»',
  '»': '«',
  '‹': '›',
  '›': '‹',
}

export function logicalGlyphContent(content: string): string {
  if (!UNIVER_ARABIC_ONLY.test(content)) return content
  return [...content].reverse().join('')
}

function weakClass(ch: string): BidiClass {
  if (ES_CHAR.test(ch)) return 'ES'
  if (ET_CHAR.test(ch)) return 'ET'
  // before \s: NBSP is a CS number separator, but JS \s matches it
  if (CS_CHAR.test(ch)) return 'CS'
  if (/\s/.test(ch)) return 'WS'
  if (/\p{M}/u.test(ch)) return 'NSM'
  return 'ON'
}

/// One class per glyph: the first strong letter wins (multi-char glyphs are
/// single-script words), otherwise the first character's weak class.
export function classifyBidiGlyph(content: string): BidiClass {
  let first: BidiClass | null = null
  for (const ch of content) {
    if (ch === '\u200E') return 'L'
    if (ch === '\u200F') return 'R'
    if (ch === '\u061C') return 'AL'
    if (AN_CHAR.test(ch)) {
      first ??= 'AN'
      continue
    }
    if (EN_CHAR.test(ch)) {
      first ??= 'EN'
      continue
    }
    if (AL_LETTER.test(ch)) return 'AL'
    if (R_LETTER.test(ch)) return 'R'
    if (/\p{L}/u.test(ch)) return 'L'
    first ??= weakClass(ch)
  }
  return first ?? 'ON'
}

/// UAX#9 W1-W7, N1-N2, I1-I2 over one line's glyph classes (the line is a
/// single level run, sos = eos = paragraph direction).
export function resolveBidiLevels(classes: readonly BidiClass[], rtlParagraph: boolean): number[] {
  const sos: BidiClass = rtlParagraph ? 'R' : 'L'
  const cls: BidiClass[] = [...classes]
  const n = cls.length
  for (let i = 0; i < n; i += 1) if (cls[i] === 'NSM') cls[i] = (i > 0 ? cls[i - 1] : sos) ?? sos
  // W2 searches back to sos, which is R (never AL) in an RTL paragraph:
  // digits before any Arabic letter stay EN ("10%" keeps its percent sign).
  let lastStrong: BidiClass = sos
  for (let i = 0; i < n; i += 1) {
    const c = cls[i]
    if (c === 'L' || c === 'R' || c === 'AL') lastStrong = c
    else if (c === 'EN' && lastStrong === 'AL') cls[i] = 'AN'
  }
  for (let i = 0; i < n; i += 1) if (cls[i] === 'AL') cls[i] = 'R'
  for (let i = 1; i < n - 1; i += 1) {
    const prev = cls[i - 1]
    if (cls[i] === 'ES' && prev === 'EN' && cls[i + 1] === 'EN') cls[i] = 'EN'
    else if (cls[i] === 'CS' && prev === cls[i + 1] && (prev === 'EN' || prev === 'AN'))
      cls[i] = prev
  }
  for (let i = 0; i < n; i += 1) {
    if (cls[i] !== 'ET') continue
    let j = i
    while (j < n && cls[j] === 'ET') j += 1
    if ((i > 0 && cls[i - 1] === 'EN') || (j < n && cls[j] === 'EN'))
      for (let k = i; k < j; k += 1) cls[k] = 'EN'
    i = j - 1
  }
  for (let i = 0; i < n; i += 1)
    if (cls[i] === 'ES' || cls[i] === 'ET' || cls[i] === 'CS') cls[i] = 'ON'
  lastStrong = sos
  for (let i = 0; i < n; i += 1) {
    const c = cls[i]
    if (c === 'L' || c === 'R') lastStrong = c
    else if (c === 'EN' && lastStrong === 'L') cls[i] = 'L'
  }
  const dirOf = (c: BidiClass | undefined): 'L' | 'R' | null =>
    c === 'L' ? 'L' : c === 'R' || c === 'EN' || c === 'AN' ? 'R' : null
  for (let i = 0; i < n; i += 1) {
    if (cls[i] !== 'WS' && cls[i] !== 'ON') continue
    let j = i
    while (j < n && (cls[j] === 'WS' || cls[j] === 'ON')) j += 1
    const before = i > 0 ? dirOf(cls[i - 1]) : sos
    const after = j < n ? dirOf(cls[j]) : sos
    const fill = before === after && before !== null ? before : sos
    for (let k = i; k < j; k += 1) cls[k] = fill
    i = j - 1
  }
  return cls.map((c) => (rtlParagraph ? (c === 'R' ? 1 : 2) : c === 'L' ? 0 : c === 'R' ? 1 : 2))
}

/// UAX#9 L2: order[visualSlot] = logical index.
export function bidiVisualOrder(levels: readonly number[]): number[] {
  const order = levels.map((_unused, i) => i)
  const lv = [...levels]
  let max = 0
  for (const level of lv) if (level > max) max = level
  for (let level = max; level >= 1; level -= 1) {
    let i = 0
    while (i < lv.length) {
      if ((lv[i] ?? 0) < level) {
        i += 1
        continue
      }
      let j = i
      while (j < lv.length && (lv[j] ?? 0) >= level) j += 1
      for (let a = i, b = j - 1; a < b; a += 1, b -= 1) {
        ;[order[a], order[b]] = [order[b]!, order[a]!]
        ;[lv[a], lv[b]] = [lv[b]!, lv[a]!]
      }
      i = j
    }
  }
  return order
}

interface SkeletonGlyphLike {
  content?: string
  width: number
  left: number
}

interface SkeletonLike {
  getViewModel?: () => {
    getDataModel?: () => {
      getSnapshot?: () => { id?: string; body?: { dataStream?: string } } | undefined
    }
  }
  getSkeletonData?: () =>
    | {
        pages?: Array<{
          sections?: Array<{
            columns?: Array<{
              lines?: Array<{
                paragraphIndex?: number | undefined
                divides?: Array<{ glyphGroup?: SkeletonGlyphLike[] }>
              }>
            }>
          }>
        }>
      }
    | null
    | undefined
}

// Memo keyed on the divide's glyphGroup ARRAY (not the line): incremental
// relayouts (DocumentSkeleton keeps a layout anchor and can retain earlier
// pages) must reprocess whenever glyphs were rebuilt. The transform itself is
// idempotent (originals map + min-left anchor), so a stale miss only costs a
// redundant pass, never a wrong paint.
const processedDivides = new WeakMap<object, { first: object | undefined; count: number }>()
const originalContents = new WeakMap<object, string>()
/** Glyphs currently seated at an odd (right-to-left) level. */
const rtlGlyphs = new WeakSet<object>()

/** Documents whose glyph geometry is reordered: sheet rich cells plus the
 *  sheet editors that lay the same text out while it is being edited. */
export const REORDERED_EDITOR_IDS: readonly string[] = [
  DOCS_NORMAL_EDITOR_UNIT_ID_KEY,
  DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY,
]

export function isReorderedDocumentId(id: string | undefined): boolean {
  return id === 'rich-cell' || REORDERED_EDITOR_IDS.includes(id ?? '')
}

/**
 * Paragraph base direction: the first strong character (matches the plain
 * cell painter). A formula being edited keeps an LTR base so its operators
 * and references read left to right whatever the first string literal is —
 * only the RTL runs inside it flip.
 */
export function paragraphIsRtl(
  dataStream: string,
  paragraphIndex: number | undefined,
  formulaAware = false,
): boolean {
  if (formulaAware && dataStream.startsWith('=')) return false
  const end =
    typeof paragraphIndex === 'number' && paragraphIndex >= 0 && paragraphIndex <= dataStream.length
      ? paragraphIndex
      : dataStream.length
  const start = dataStream.lastIndexOf('\r', end - 1) + 1
  return resolveBidiDirection(dataStream.slice(start, end)) === 'rtl'
}

/** @returns whether the divide holds any glyph now seated right-to-left */
function reorderDivide(glyphs: SkeletonGlyphLike[], rtlParagraph: boolean): boolean {
  if (glyphs.length === 0) return false
  let sawRtl = false
  const contents: string[] = []
  for (const glyph of glyphs) {
    let original = originalContents.get(glyph)
    if (original === undefined) {
      original = glyph.content ?? ''
      originalContents.set(glyph, original)
    }
    const logical = logicalGlyphContent(original)
    if (logical !== glyph.content) glyph.content = logical
    if (HAS_RTL.test(logical)) sawRtl = true
    contents.push(logical)
  }
  if (!sawRtl && !rtlParagraph) return false
  const levels = resolveBidiLevels(contents.map(classifyBidiGlyph), rtlParagraph)
  let anyRtl = false
  for (let i = 0; i < glyphs.length; i += 1) {
    const glyph = glyphs[i]!
    if ((levels[i] ?? 0) % 2 === 1) {
      anyRtl = true
      rtlGlyphs.add(glyph)
      const mirrored = MIRROR[contents[i] ?? '']
      if (mirrored !== undefined) glyph.content = mirrored
    } else {
      rtlGlyphs.delete(glyph)
    }
  }
  const order = bidiVisualOrder(levels)
  // min, not glyphs[0].left: a re-run over an already-permuted group must
  // anchor at the line start again.
  let left = glyphs.reduce((min, glyph) => Math.min(min, glyph.left), Number.POSITIVE_INFINITY)
  for (const logicalIndex of order) {
    const glyph = glyphs[logicalIndex]
    if (!glyph) continue
    glyph.left = left
    left += glyph.width
  }
  return anyRtl
}

/** Skeletons that hold (or held) right-to-left glyphs: only their caret
 *  geometry needs the bidi mapping. */
const reorderedSkeletons = new WeakSet<object>()

function eachDivide(
  skeleton: SkeletonLike,
  visit: (
    glyphs: SkeletonGlyphLike[],
    line: { paragraphIndex?: number | undefined },
    lineState: { rtl?: boolean },
  ) => void,
): void {
  const data = skeleton.getSkeletonData?.()
  for (const page of data?.pages ?? [])
    for (const section of page.sections ?? [])
      for (const column of section.columns ?? [])
        for (const line of column.lines ?? []) {
          const lineState: { rtl?: boolean } = {}
          for (const divide of line.divides ?? []) visit(divide.glyphGroup ?? [], line, lineState)
        }
}

export function reorderRichCellSkeleton(skeleton: SkeletonLike): void {
  const snapshot = skeleton.getViewModel?.()?.getDataModel?.()?.getSnapshot?.()
  if (!isReorderedDocumentId(snapshot?.id)) return
  const dataStream = snapshot?.body?.dataStream ?? ''
  if (!HAS_RTL.test(dataStream)) {
    // Text edited back to LTR-only: glyphs an incremental relayout kept from
    // the bidi layout must not keep their right-to-left caret mapping.
    if (reorderedSkeletons.has(skeleton)) {
      reorderedSkeletons.delete(skeleton)
      eachDivide(skeleton, (glyphs) => {
        for (const glyph of glyphs) rtlGlyphs.delete(glyph)
      })
    }
    return
  }
  const editor = snapshot?.id !== 'rich-cell'
  eachDivide(skeleton, (glyphs, line, lineState) => {
    const memo = processedDivides.get(glyphs)
    if (memo && memo.first === glyphs[0] && memo.count === glyphs.length) return
    processedDivides.set(glyphs, { first: glyphs[0], count: glyphs.length })
    lineState.rtl ??= paragraphIsRtl(dataStream, line.paragraphIndex, editor)
    if (reorderDivide(glyphs, lineState.rtl)) reorderedSkeletons.add(skeleton)
  })
}

interface EditorGlyphLike {
  left: number
  width: number
  glyphType?: unknown
}

/**
 * Pixel x of the logical boundary `back ? before : after` a glyph. For a glyph
 * seated right-to-left, "before" is its visual right edge.
 */
export function glyphBoundaryX(glyph: EditorGlyphLike, back: boolean, rtl: boolean): number {
  return glyph.left + (back !== rtl ? 0 : glyph.width)
}

/**
 * Horizontal extent of a caret/selection inside one divide. `startIndex` /
 * `endIndex` are logical glyph indexes with their isBack flags (Univer's
 * position model); the selected glyphs are the ones strictly between the two
 * boundaries. Their union is a single box — a logical range over mixed runs
 * is visually discontiguous in bidi text, so this is the bounding box.
 */
export function divideSelectionExtent(
  glyphs: readonly EditorGlyphLike[],
  startIndex: number,
  startBack: boolean,
  endIndex: number,
  endBack: boolean,
  isRtlGlyph: (glyph: EditorGlyphLike) => boolean,
): { startX: number; endX: number } {
  const first = startBack ? startIndex : startIndex + 1
  const last = endBack ? endIndex - 1 : endIndex
  if (first > last) {
    const glyph = glyphs[startIndex]
    if (!glyph) return { startX: 0, endX: 0 }
    const x = glyphBoundaryX(glyph, startBack, isRtlGlyph(glyph))
    return { startX: x, endX: x }
  }
  let startX = Number.POSITIVE_INFINITY
  let endX = Number.NEGATIVE_INFINITY
  for (let i = first; i <= last; i += 1) {
    const glyph = glyphs[i]
    if (!glyph) continue
    startX = Math.min(startX, glyph.left)
    endX = Math.max(endX, glyph.left + glyph.width)
  }
  return { startX, endX }
}

interface NodePositionLike {
  page: number
  section: number
  column: number
  line: number
  divide: number
  glyph: number
  isBack: boolean
}

/** Same ordering as Univer's compareNodePositionLogic (equal → a first). */
function positionIsBefore(a: NodePositionLike, b: NodePositionLike): boolean {
  for (const key of ['page', 'section', 'column', 'line', 'divide', 'glyph'] as const) {
    if (a[key] !== b[key]) return a[key] < b[key]
  }
  return true
}

interface PointGroup {
  x: number
  y: number
}

/** pushToPoints order: (s,t) (e,t) (e,b) (s,b) (s,t) */
function reseatBox(points: PointGroup[] | undefined, startX: number, endX: number): void {
  if (!points || points.length !== 5) return
  for (const index of [0, 3, 4]) points[index]!.x = startX
  for (const index of [1, 2]) points[index]!.x = endX
}

type DivideVisitor = (
  start_sp: number,
  end_sp: number,
  isFirst: boolean,
  isLast: boolean,
  divide: { glyphGroup: EditorGlyphLike[] },
  line: unknown,
) => void

/**
 * Caret and selection rectangles read `glyph.left` as the glyph's LOGICAL
 * start. Ride along the converter's own divide walk (same liquid offsets) and
 * re-seat, for each divide holding right-to-left glyphs, the box it pushed —
 * the stock walk pushes exactly one box per non-empty divide when no
 * horizontal clip is active, so boxes and visits pair up by index. Logical
 * offsets (`cursorList`) are untouched.
 */
function patchCaretGeometry(): void {
  const proto = NodePositionConvertToCursor.prototype as any
  const origGetRangePointData = proto.getRangePointData
  proto.getRangePointData = function (startOrigin: NodePositionLike, endOrigin: NodePositionLike) {
    if (!startOrigin || !endOrigin || !reorderedSkeletons.has(this._docSkeleton))
      return origGetRangePointData.call(this, startOrigin, endOrigin)
    const ordered = positionIsBefore(startOrigin, endOrigin)
    const start = ordered ? startOrigin : endOrigin
    const end = ordered ? endOrigin : startOrigin
    const fixes: Array<{ startX: number; endX: number } | null> = []
    let clipped = false
    // Shadow the prototype's walker for this one call (Univer defines it on
    // the prototype; restore an own property if a build ever moves it).
    const origIterator = this._selectionIterator
    const ownIterator = Object.prototype.hasOwnProperty.call(this, '_selectionIterator')
    this._selectionIterator = (s: NodePositionLike, e: NodePositionLike, visit: DivideVisitor) =>
      origIterator.call(this, s, e, (...args: Parameters<DivideVisitor>) => {
        const [start_sp, end_sp, isFirst, isLast, divide] = args
        const glyphs = divide.glyphGroup
        if (glyphs.length > 0) {
          if (this._horizontalClip != null) clipped = true
          const rtl =
            glyphs.some((glyph) => rtlGlyphs.has(glyph)) &&
            glyphs[start_sp]?.glyphType !== GlyphType.LIST
          if (rtl) {
            const isStartBack = start.glyph === start_sp && isFirst ? start.isBack : true
            const isEndBack = end.glyph === end_sp && isLast ? end.isBack : false
            const extent = divideSelectionExtent(
              glyphs,
              start_sp,
              isStartBack,
              end_sp,
              isEndBack,
              (glyph) => rtlGlyphs.has(glyph),
            )
            const liquidX: number = this._liquid.x
            fixes.push({ startX: liquidX + extent.startX, endX: liquidX + extent.endX })
          } else fixes.push(null)
        }
        return visit(...args)
      })
    try {
      const result = origGetRangePointData.call(this, startOrigin, endOrigin)
      if (!clipped && result.borderBoxPointGroup.length === fixes.length) {
        fixes.forEach((fix, index) => {
          if (!fix) return
          reseatBox(result.borderBoxPointGroup[index], fix.startX, fix.endX)
          reseatBox(result.contentBoxPointGroup[index], fix.startX, fix.endX)
        })
      }
      return result
    } finally {
      if (ownIterator) this._selectionIterator = origIterator
      else delete this._selectionIterator
    }
  }
}

/**
 * A click lands "before" a glyph when it hits the glyph's left half — for a
 * right-to-left glyph that half is logically AFTER it.
 */
function patchClickSide(): void {
  const proto = DocSelectionRenderService.prototype as any
  const origGetNodePosition = proto._getNodePosition
  proto._getNodePosition = function (node: { node?: EditorGlyphLike } | null | undefined) {
    const position = origGetNodePosition.call(this, node)
    const glyph = node?.node
    if (position && glyph && rtlGlyphs.has(glyph) && glyph.glyphType !== GlyphType.LIST) {
      position.isBack = !position.isBack
    }
    return position
  }
}

/** Depth of DocumentSkeleton.calculate calls for documents this module
 *  reorders (re-entrancy safe). */
let measureLogicalDepth = 0

/**
 * ArabicHandler measures each reversed chunk as it stored it, but the reorder
 * paints the logical string, whose joined letterforms are usually narrower —
 * the leftover became a gap after every Arabic word. Measure the string that
 * will be painted. Only glyph creation inside calculate() ever hands an
 * Arabic-only multi-character string to getTextSize (single letters go
 * through otherHandler), so the depth gate keeps whole-string callers such
 * as data-validation dropdown items untouched.
 */
function patchArabicChunkMeasure(): void {
  const cache = FontCache as unknown as {
    getTextSize(content: string, fontStyle: unknown): unknown
  }
  const origGetTextSize = cache.getTextSize
  cache.getTextSize = function (content: string, fontStyle: unknown) {
    const measured = measureLogicalDepth > 0 ? logicalGlyphContent(content) : content
    return origGetTextSize.call(this, measured, fontStyle)
  }
}

let installed = false

export function installRichTextBidiFix(): void {
  if (installed) return
  installed = true

  const proto = (DocumentSkeleton as unknown as { prototype: Record<string, unknown> }).prototype
  const previousCalculate = proto.calculate
  if (typeof previousCalculate !== 'function') return
  proto.calculate = function (this: SkeletonLike, bounds?: unknown): void {
    const reordered = isReorderedDocumentId(
      this.getViewModel?.()?.getDataModel?.()?.getSnapshot?.()?.id,
    )
    if (reordered) measureLogicalDepth += 1
    try {
      previousCalculate.call(this, bounds)
    } finally {
      if (reordered) measureLogicalDepth -= 1
    }
    try {
      reorderRichCellSkeleton(this)
    } catch {
      // A reorder failure must never break layout; the cell falls back to
      // Univer's logical-order paint.
    }
  }
  patchArabicChunkMeasure()
  patchCaretGeometry()
  patchClickSide()
}
