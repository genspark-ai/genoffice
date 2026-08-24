/**
 * Univer's rich-text layout has no bidi: glyphs are painted one by one at
 * x positions accumulated in logical order, and its ArabicHandler stores each
 * contiguous Arabic chunk REVERSED (the canvas re-reverses it at draw time),
 * so RTL rich cells (`cell.p`) render mirrored. The document model must keep
 * the logical stream — edits, sorts (journalRangeSnapshot) and saves read
 * `cell.p` back — so this fix touches layout artifacts only: after every
 * DocumentSkeleton.calculate of a 'rich-cell' document it restores each
 * glyph's logical content and re-seats glyph x positions in UAX#9 visual
 * order. The in-cell editor clones the body into its own '__INTERNAL_EDITOR__'
 * snapshot, so the id gate leaves editing untouched.
 */
import { DocumentSkeleton } from '@univerjs/engine-render'

import { resolveBidiDirection } from './rtl-text-fix'

export type BidiClass = 'L' | 'R' | 'AL' | 'EN' | 'AN' | 'ES' | 'ET' | 'CS' | 'NSM' | 'WS' | 'ON'

const HAS_RTL =
  /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}\u200F\u061C]/u
const AL_LETTER = /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}]/u
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

function paragraphIsRtl(dataStream: string, paragraphIndex: number | undefined): boolean {
  const end =
    typeof paragraphIndex === 'number' && paragraphIndex >= 0 && paragraphIndex <= dataStream.length
      ? paragraphIndex
      : dataStream.length
  const start = dataStream.lastIndexOf('\r', end - 1) + 1
  return resolveBidiDirection(dataStream.slice(start, end)) === 'rtl'
}

function reorderDivide(glyphs: SkeletonGlyphLike[], rtlParagraph: boolean): void {
  if (glyphs.length === 0) return
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
  if (!sawRtl && !rtlParagraph) return
  const levels = resolveBidiLevels(contents.map(classifyBidiGlyph), rtlParagraph)
  for (let i = 0; i < glyphs.length; i += 1) {
    if ((levels[i] ?? 0) % 2 === 1) {
      const mirrored = MIRROR[contents[i] ?? '']
      if (mirrored !== undefined) glyphs[i]!.content = mirrored
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
}

export function reorderRichCellSkeleton(skeleton: SkeletonLike): void {
  const snapshot = skeleton.getViewModel?.()?.getDataModel?.()?.getSnapshot?.()
  if (snapshot?.id !== 'rich-cell') return
  const dataStream = snapshot.body?.dataStream ?? ''
  if (!HAS_RTL.test(dataStream)) return
  const data = skeleton.getSkeletonData?.()
  for (const page of data?.pages ?? [])
    for (const section of page.sections ?? [])
      for (const column of section.columns ?? [])
        for (const line of column.lines ?? []) {
          let rtl: boolean | undefined
          for (const divide of line.divides ?? []) {
            const glyphs = divide.glyphGroup ?? []
            const memo = processedDivides.get(glyphs)
            if (memo && memo.first === glyphs[0] && memo.count === glyphs.length) continue
            processedDivides.set(glyphs, { first: glyphs[0], count: glyphs.length })
            rtl ??= paragraphIsRtl(dataStream, line.paragraphIndex)
            reorderDivide(glyphs, rtl)
          }
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
    previousCalculate.call(this, bounds)
    try {
      reorderRichCellSkeleton(this)
    } catch {
      // A reorder failure must never break layout; the cell falls back to
      // Univer's logical-order paint.
    }
  }
}
