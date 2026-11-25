import { Extension, type Editor } from '@tiptap/core'
import {
  isCjkLetterChar,
  isEastAsianChar,
  isEastAsianLangTag,
  isHangulChar,
  type KinsokuCustom,
} from '@genoffice/docx-engine'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { SettledParagraphCache, noteFloatTransaction } from './settled-measure'
import { sameLine } from './justify-shrink'
import {
  ANY_STANDARD_LINE_END,
  ANY_STANDARD_LINE_START,
  STRICT_JA_LINE_START,
  kinsokuLang,
} from './kinsoku-tables'
import { rangeSlot } from '../dom-range'
import { PHASED_CONTENT_SETTLED_EVENT, isPhasedContentPending } from '../phased-content'

/**
 * Word's CJK line breaking (settings characterSpacingControl =
 * compressPunctuation) pulls extra characters onto a line by compressing the
 * line's blank-half punctuation (、。, closing and opening brackets). Chromium
 * neither compresses nor pulls, so Japanese documents wrap earlier than Word,
 * fit fewer characters per line, and drift pages apart.
 *
 * Word for Mac probes (2026-09-01, MS Mincho 10.5pt, w:jc="both"):
 * - doNotCompress: no compression ever; every line stretches uniformly.
 * - compressPunctuation: a full line's compressible punctuation shrinks by a
 *   uniform per-glyph amount sized to the line's deficit; other glyphs keep
 *   their natural advance (no stretch on compressed lines); the ragged last
 *   line never compresses.
 * - Pull decisions: a voluntary pull (next char is an ordinary ideograph)
 *   compresses lightly (observed accepted up to ~0.252em, declined from ~0.286em per
 *   glyph); a kinsoku pull (the following character is a closing punctuation
 *   that cannot start a line) compresses as far as needed, observed to about
 *   half width (JIS compression floor).
 * - Alignment matters only under the Word 2013+ layout (settings
 *   compatibilityMode >= 15): there a left-aligned paragraph never pulls or
 *   compresses, a trailing 、。 that does not fit takes its predecessor down
 *   with it (Yu Mincho 11pt / Meiryo 14pt probes, 2026-09-04). Legacy modes
 *   (compatibilityMode 14 and below) pull on every alignment and the
 *   compressed line ends flush at the margin — SAS batch 2 sample 068,
 *   Meiryo 14pt, compat 14. Opening brackets lose their leading blank
 *   alongside the closing glyphs, but do not make a pull admissible (a
 *   voluntary pull was declined at 0.22em over an opening bracket + two
 *   closing glyphs, while over the two closing glyphs alone it exceeds the
 *   cap).
 * - Without compressPunctuation (doNotCompress or no characterSpacingControl;
 *   MS Mincho 10.5pt probes 2026-09-05, SAS batch 2 sample 005) nothing
 *   compresses, but a single trailing stop or closing bracket that does not
 *   fit hangs into the margin: its origin sits exactly at the text edge and
 *   the line's other glyphs stretch to the margin. Only the last glyph hangs,
 *   only when everything before it fits, on the same alignments compression
 *   applies to (justified under compat 15, every alignment under legacy).
 *   Under compressPunctuation that trailing stop is compressed to half width
 *   instead, even when it is the line's only compressible glyph. When even
 *   that does not fit, the stop still hangs at the margin provided the body
 *   before it fits once the line's own closers give up at most half a glyph
 *   in total (Yu Mincho 10.5pt probes 2026-10-13: 1.4pt and 4.5pt over one
 *   comma accepted, 7.3pt over two commas and 12.4pt over one declined -> the
 *   next line starts with the kana + stop).
 * - All of it is gated on the run's effective w:lang w:eastAsia (run rPr >
 *   character style > paragraph style chain > docDefaults; the paragraph mark
 *   does not count): under en-US / ko-KR Word neither hangs nor compresses
 *   nor applies kinsoku — the next line simply starts with the 。 (probes
 *   2026-09-05). No w:lang anywhere keeps the East Asian rules.
 *
 * compressPunctuationAndJapaneseKana (Word for Mac probes 2026-11-18 and
 * 2026-11-22, 12pt and 10.5pt, with and without a character grid) adds one
 * thing, on the same alignments:
 * - A line takes up what the next one starts with (a character with the marks
 *   that may not start a line after it, an opening bracket with what follows,
 *   a Latin word) when its kana can make the room. Hiragana U+3041-3093,
 *   katakana U+30A1-30F6 and the fullwidth digits and letters give up to a
 *   sixteenth of their advance each (the grid advance under a character
 *   grid), evenly; everything else keeps its advance.
 * - Only as far as the line would otherwise stretch: each kana gives at most
 *   about four times what each glyph of the line as it stands would be
 *   stretched by (50 kana with 1 to 11pt left over: taken up from 3pt on,
 *   from 4pt on with 19 kana). Lines filled exactly take nothing up.
 * - Punctuation goes first: a line its punctuation settles has no kana
 *   compressed; where a line-start rule forces the run up, closing punctuation
 *   is set half width and the kana give the rest.
 * - A stop that would hang comes inside when the kana can give its whole
 *   advance, is set half width at the margin when they can give half of it,
 *   and hangs otherwise; with w:overflowPunct off, half width or nothing.
 * - An opening bracket taken up this way is drawn without its blank half.
 *
 * Re-created as display-only inline decorations: the affected line's
 * compressible punctuation gets negative letter-spacing (opening brackets a
 * negative left margin, their blank half leads the glyph) so Chromium's greedy
 * breaker takes the same characters; pagination measures the decorated DOM.
 */

/** trailing-blank punctuation Word compresses (JIS stops and closing brackets) */
const COMPRESSIBLE_CLOSE = new Set('、。，．」）』】｝〕》〉〗〙')
/** leading-blank opening brackets */
const COMPRESSIBLE_OPEN = new Set('「（『【｛〔《〈〖〘')
const COMPRESSIBLE = new Set([...COMPRESSIBLE_CLOSE, ...COMPRESSIBLE_OPEN])
/** Word's default line-start kinsoku (JIS closers, ASCII closers included) */
const KINSOKU_CLOSE = new Set([
  ...COMPRESSIBLE_CLOSE,
  ...'！？：；ー々',
  ...'!%),.:;?]}\u2019\u201d',
])
const OPENERS = new Set([...COMPRESSIBLE_OPEN, ...'([{\u2018\u201c'])

/** Chromium glues these to their predecessor whatever the language: pulling
 *  the character before them drags them along */
export function forbidsLineStart(ch: string): boolean {
  return KINSOKU_CLOSE.has(ch)
}
export function forbidsLineEnd(ch: string): boolean {
  return OPENERS.has(ch)
}

export function isCompressible(ch: string): boolean {
  return COMPRESSIBLE.has(ch)
}

const KANA_CLASS = /[\u3041-\u3093\u30a1-\u30f6\uff10-\uff19\uff21-\uff3a\uff41-\uff5a]/

/** what compressPunctuationAndJapaneseKana compresses besides punctuation */
export function isCompressibleKana(ch: string): boolean {
  return KANA_CLASS.test(ch)
}

/**
 * The letters and digits that get a gap next to East Asian text. Kept here: importing
 * line-metrics changes the module order, and with it how one corpus paragraph settles.
 */
const LATIN_ALNUM = /[0-9A-Za-z\u00c0-\u00d6\u00d8-\u00f6\u00f8-\u024f]/

function autospaceGap(ch: string, autoSpace: unknown): boolean {
  if (autoSpace === false || !LATIN_ALNUM.test(ch)) return false
  return typeof autoSpace !== 'string' || (autoSpace === 'digits') === /[0-9]/.test(ch)
}

/** a character of a Latin word: a run of them does not break */
const LATIN_WORD = /[!-~\u00a1-\u024f]/

export type SpacingControl = 'none' | 'punct' | 'kana'

export function spacingControlOf(
  parsed: { compressPunctuation?: boolean; compressKana?: boolean } | null | undefined,
): SpacingControl {
  return parsed?.compressKana ? 'kana' : parsed?.compressPunctuation ? 'punct' : 'none'
}

export function setSpacingControl(storage: CjkPunctShrinkStorage, control: SpacingControl): void {
  storage.enabled = control !== 'none'
  storage.hangPunct = control === 'none'
  storage.kana = control === 'kana'
}

/** Word applies its East Asian line rules under a CJK w:lang w:eastAsia or none at all */
export function usesEastAsianRules(eastAsiaLang: string | null | undefined): boolean {
  return !eastAsiaLang || isEastAsianLangTag(eastAsiaLang, false)
}

/** the shrink decoration carries the compressed spacing: the natural base lives on its nearest other ancestor */
export function spacingHost(el: Element | null): Element | null {
  while (el && el.classList.contains('doc-cjkshrink')) el = el.parentElement
  return el
}

/** decoration style compressing one glyph by perChar px (baseLs: inherited letter-spacing it replaces) */
export function shrinkStyle(ch: string, perChar: number, baseLs: number): string {
  const px = (v: number) => Math.round(v * 100) / 100
  return COMPRESSIBLE_OPEN.has(ch)
    ? `margin-left:${px(-perChar)}px`
    : `letter-spacing:${px(baseLs - perChar)}px`
}

/** Han / kana paragraph: hangul alone never compresses (the ko-KR rule above) */
function textHasCjkNoHangul(text: string): boolean {
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    if (isEastAsianChar(cp) && !isHangulChar(cp)) return true
  }
  return false
}

/** voluntary pull cap: Word accepted 25.2% and declined 28.6% per glyph (probes) */
const VOLUNTARY_CAP = 0.27
/** kinsoku-forced pull: down to the JIS half-width floor */
const FORCED_CAP = 0.5
/** total overshoot (px) so Chromium's breaker definitely pulls */
const SHRINK_EPS = 0.5
/** measurement noise floor (px) */
const NOISE = 0.25
/** paragraphs beyond this many characters skip (per-char measurement cost) */
const PARA_CHAR_BUDGET = 4000

const MEASURE_RETRY_MAX = 10
const MEASURE_SIGS_MAX = 12

export interface CjkPunctShrinkStorage {
  /** settings.xml characterSpacingControl compresses punctuation */
  enabled: boolean
  /** no compressPunctuation: a trailing stop that does not fit hangs into the margin */
  hangPunct: boolean
  /** compressPunctuationAndJapaneseKana: kana give way too (only read when enabled) */
  kana: boolean
  /** settings.xml compatibilityMode < 15: Word 2010 layout also pulls on unjustified lines */
  legacyLayout: boolean
  /** docDefaults w:lang w:eastAsia: the fallback of the run / paragraph-style chain */
  docEastAsiaLang: string | null
  /** settings.xml w:strictFirstAndLastChars: Word's Strict kinsoku level */
  strictKinsoku: boolean
  /** settings.xml w:noLineBreaksBefore / After: Word's Custom kinsoku level */
  kinsokuCustom: KinsokuCustom | null
}

/** dispatched on document after the doc-scoped <style> commits (App) */
export const DOC_CSS_COMMITTED_EVENT = 'genoffice:doc-css'
/** the paginator applied per-block wrap widths (section widths, columns,
 *  letter-spacing): the line breaks the shrink loops measured are stale */
export const WRAP_WIDTHS_COMMITTED_EVENT = 'genoffice:wrap-widths'

export function measuresAlignment(textAlign: string, legacyLayout: boolean): boolean {
  return textAlign === 'justify' || legacyLayout
}

declare module '@tiptap/core' {
  interface Storage {
    cjkPunctShrink: CjkPunctShrinkStorage
  }
}

export const cjkPunctShrinkPluginKey = new PluginKey<DecorationSet>('cjkPunctShrink')

interface CharBox {
  ch: string
  from: number
  /** the run's effective East Asian language enables Word's line rules for this glyph */
  ea: boolean
  /** natural advance (layout px), independent of active decorations */
  width: number
  /** inherited letter-spacing (px) included in width */
  ls: number
  top: number
  bottom: number
  left: number
  right: number
  /** rendered right edge (screen px) when measured straight off the DOM text */
  rendRight?: number
}

const shiftShrinks = (r: MeasuredShrink[], d: number) => r.map((s) => ({ ...s, from: s.from + d }))

const isPunct = (c: CharBox) => c.ea && COMPRESSIBLE.has(c.ch)
const isClose = (c: CharBox) => c.ea && COMPRESSIBLE_CLOSE.has(c.ch)
const isKinsokuClose = (c: CharBox) => c.ea && KINSOKU_CLOSE.has(c.ch)

interface MeasuredShrink {
  from: number
  ch: string
  perChar: number
  /** paragraph base letter-spacing (px, docGrid charSpace): the decoration's
   *  own letter-spacing replaces the inherited value, so it must fold it in */
  baseLs: number
  /** the glyph hangs at zero advance (a trailing stop at the margin) */
  hang?: boolean
  /** w:topLinePunct: an opening bracket that starts a line gives up its blank half */
  lineStart?: boolean
  /** the narrower bracket would fit back on the line above: a break in front keeps it where Word has it */
  pin?: boolean
}

export interface ShrinkLineChars {
  /** natural width of the line's characters */
  natural: number
  /** rendered (justified) width the line fills */
  avail: number
  /** closing compressible glyphs on the line (Word's admissibility pool) */
  punctCount: number
  /** opening brackets on the line: they share the compression but not the pool */
  openCount?: number
  /** average natural advance of the closing glyphs (cap base; one body size per line in practice) */
  avgPunctW: number
  /** natural widths of the pull candidate chain (next line's leading chars); empty = no candidate */
  candWidths: number[]
  /** closing glyphs inside the chain (they join the line's pool once pulled) */
  candPunctCount: number
  /** the chain drags a kinsoku-close character that cannot start a line */
  forced: boolean
  /** natural advance of the line's last glyph when it already hangs (decorated) */
  hungWidth?: number | null
  /** the chain ends with a stop / closing bracket that may hang */
  candEndsWithStop?: boolean
}

export interface HangLineChars {
  natural: number
  avail: number
  /** natural advance of the line's last glyph when it already hangs (decorated) */
  hungWidth: number | null
  /** natural widths of the pull candidate chain; empty = no candidate */
  candWidths: number[]
  /** the chain ends with a stop / closing bracket that may hang */
  candEndsWithStop: boolean
}

/** Word's hang rule (no compressPunctuation): the chain's trailing stop hangs
 *  at zero advance once everything before it fits the line. */
export function decideCjkHang(line: HangLineChars): 'keep' | 'pull' | null {
  if (line.hungWidth !== null) {
    return line.natural - line.hungWidth <= line.avail + NOISE ? 'keep' : null
  }
  if (!line.candEndsWithStop || line.candWidths.length === 0) return null
  const body = line.candWidths.slice(0, -1).reduce((s, w) => s + w, 0)
  return line.natural + body <= line.avail + NOISE ? 'pull' : null
}

/**
 * Word's hang rule under compressPunctuation, tried once decideCjkShrinks
 * declined: the chain's trailing stop hangs at zero advance when the body
 * before it fits after the line's own closers (the hung stop excluded) give
 * up at most half a glyph in total. Returns the per-closer compression (0 =
 * hang only) or null.
 */
export function decideCjkHangUnderCompression(line: ShrinkLineChars): number | null {
  const hung = line.hungWidth != null
  const body = hung
    ? line.natural - line.hungWidth!
    : line.natural + line.candWidths.slice(0, -1).reduce((s, w) => s + w, 0)
  if (!hung && (!line.candEndsWithStop || line.candWidths.length === 0)) return null
  const deficit = body - line.avail
  if (deficit <= NOISE) return 0
  const pool = line.punctCount - (hung ? 1 : 0)
  if (pool <= 0 || line.avgPunctW <= 0 || deficit > FORCED_CAP * line.avgPunctW) return null
  return (deficit + SHRINK_EPS) / pool
}

/** the run a line would take up: what the next line starts with and cannot be parted from */
export interface KanaRun {
  width: number
  kanaWidth: number
  kanaCount: number
  /** half the advance of its opening brackets: Word draws them without the blank half */
  openerCredit: number
  /** advance of the stop or closing bracket it ends with, 0 when it ends otherwise */
  stopWidth: number
}

export interface KanaLineChars {
  natural: number
  avail: number
  chars: number
  kanaWidth: number
  kanaCount: number
  /** what the line's closing punctuation can give: half its advance, for a run a line-start rule forces up */
  punctCredit: number
  chain: KanaRun
  hangs: boolean
}

export interface KanaShrink {
  perKana: number
  halfStop: boolean
  /** the share of punctCredit the closing punctuation gives, 0 to 1 */
  punct: number
}

/** each kana gives at most this many times what each glyph of the line would stretch (probes: 3.9 to 4.7) */
const KANA_PER_STRETCH = 4

/** Word's kana rule over one measured line; null when the kana cannot make the room */
export function decideKanaShrink(line: KanaLineChars): KanaShrink | null {
  const { chain } = line
  const room = line.avail - line.natural
  const total = chain.width - chain.openerCredit - room
  const count = line.kanaCount + chain.kanaCount
  if (room < -NOISE) return null
  if (total <= NOISE) {
    // the run fits once its opening bracket gives up the blank half: the kana only carry the overshoot
    const byBracket = chain.openerCredit > 0 && total + chain.openerCredit > NOISE
    if (!byBracket || (count === 0 && total > -SHRINK_EPS)) return null
    const perKana = count > 0 ? Math.max(0, total + SHRINK_EPS) / count : 0
    return { perKana, halfStop: false, punct: 0 }
  }
  if (count === 0) return null
  const cap = (line.kanaWidth + chain.kanaWidth) / 16 + NOISE
  const perKanaLimit = (KANA_PER_STRETCH * room) / Math.max(1, line.chars - 1)
  const plan = (need: number, halfStop: boolean): KanaShrink | null => {
    const punct = Math.min(line.punctCredit, Math.max(0, need))
    const kana = Math.max(0, need) - punct
    if (kana > cap || kana / count > perKanaLimit) return null
    return {
      perKana: (kana + SHRINK_EPS) / count,
      halfStop,
      punct: punct > 0 ? punct / line.punctCredit : 0,
    }
  }
  const stop = chain.stopWidth
  return (
    (stop === 0 || line.hangs ? plan(total, false) : null) ??
    (stop > 0 ? plan(total - stop / 2, true) : null)
  )
}

/** The glyphs a decided amount lands on: the line's own punctuation (closers
 *  and openers) when it has a closer to compress, otherwise the pulled chain's
 *  stops alone — decideCjkShrinks sized the amount for exactly those. */
export function shrinkTargets<T>(line: ShrinkLineChars, own: T[], cand: T[]): T[] {
  return line.punctCount === 0 ? cand : own
}

/**
 * Word's pull/keep rule over measured lines. Returns the per-glyph compression
 * (px) to decorate onto the line's current compressible glyphs, or null.
 * Admissibility is judged over the post-pull pool of closing glyphs (pulled
 * punctuation joins the line), but the emitted amount spreads the whole
 * deficit over the compressible glyphs that exist now, opening brackets
 * included — once Chromium re-breaks, the keep path re-balances.
 */
export function decideCjkShrinks(lines: ShrinkLineChars[]): Array<number | null> {
  return lines.map((line) => {
    // a hung stop is decideCjkHangUnderCompression's keep state: its natural
    // width is not line overflow
    if (line.hungWidth != null) return null
    if (line.avgPunctW <= 0) return null
    const overflow = line.natural - line.avail
    if (line.punctCount === 0) {
      // the pulled stop is the line's only compressible glyph: it absorbs the
      // whole deficit itself (Word halves it), so the amount lands on the chain
      if (overflow > NOISE || !line.forced || line.candPunctCount === 0) return null
      const deficit = overflow + line.candWidths.reduce((s, w) => s + w, 0)
      if (deficit <= NOISE) return null
      const perGlyph = (deficit + SHRINK_EPS) / line.candPunctCount
      return perGlyph <= FORCED_CAP * line.avgPunctW ? perGlyph : null
    }
    const spread = line.punctCount + (line.openCount ?? 0)
    if (overflow > NOISE) {
      // the line already holds characters pulled by a previous round: keep the
      // compression that fits them (recomputed fresh from natural advances)
      const total = overflow + SHRINK_EPS
      return total / line.punctCount <= FORCED_CAP * line.avgPunctW ? total / spread : null
    }
    if (line.candWidths.length === 0) return null
    const deficit = overflow + line.candWidths.reduce((s, w) => s + w, 0)
    if (deficit <= NOISE) return null
    const pool = line.punctCount + line.candPunctCount
    const capFrac = line.forced ? FORCED_CAP : VOLUNTARY_CAP
    if ((deficit + SHRINK_EPS) / pool > capFrac * line.avgPunctW) return null
    return (deficit + SHRINK_EPS) / spread
  })
}

interface TextSlot {
  dom: Text
  from: number
  ea: boolean
}

/**
 * The paragraph's DOM text nodes mapped to document positions. Valid only when
 * the DOM text is exactly the node's text in order (decorations split text
 * nodes but never add characters); widget text, cursor wrappers or inline
 * atoms make it null and the caller falls back to view.coordsAtPos. The direct
 * map avoids ProseMirror's per-position descent through every top-level
 * block, which made per-character measurement scale with document size.
 */
export function mapTextNodes(
  el: HTMLElement,
  node: ProseMirrorNode,
  pos: number,
  eaOf: (child: ProseMirrorNode) => boolean,
): { slots: TextSlot[]; breaks: number[] } | null {
  const pieces: Array<{ text: string; from: number; ea: boolean }> = []
  const breaks: number[] = []
  let atom = false
  node.forEach((child, offset) => {
    if (child.isText && child.text)
      pieces.push({ text: child.text, from: pos + 1 + offset, ea: eaOf(child) })
    else if (child.type.name === 'hardBreak') breaks.push(pos + 1 + offset)
    else atom = true
  })
  if (atom) return null
  const slots: TextSlot[] = []
  let pi = 0
  let off = 0
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
  for (let tn = walker.nextNode() as Text | null; tn; tn = walker.nextNode() as Text | null) {
    const data = tn.data
    if (!data) continue
    const piece = pieces[pi]
    if (!piece || off + data.length > piece.text.length) return null
    if (piece.text.substr(off, data.length) !== data) return null
    slots.push({ dom: tn, from: piece.from + off, ea: piece.ea })
    off += data.length
    if (off === piece.text.length) {
      pi++
      off = 0
    }
  }
  return pi === pieces.length ? { slots, breaks } : null
}

const charRange = rangeSlot()

const candidateCache = new WeakMap<ProseMirrorNode, boolean>()
const kanaCandidateCache = new WeakMap<ProseMirrorNode, boolean>()
/** CJK paragraph with something to compress, within the per-char measurement budget */
function isCandidate(node: ProseMirrorNode, kana: boolean): boolean {
  const cache = kana ? kanaCandidateCache : candidateCache
  let is = cache.get(node)
  if (is === undefined) {
    const text = node.textContent
    is = text.length <= PARA_CHAR_BUDGET && textHasCjkNoHangul(text)
    if (is) {
      is = false
      for (const ch of text) {
        if (COMPRESSIBLE.has(ch) || (kana && KANA_CLASS.test(ch))) {
          is = true
          break
        }
      }
    }
    cache.set(node, is)
  }
  return is
}

let measureCtx: CanvasRenderingContext2D | null | undefined
const advanceCache = new Map<string, Map<string, number>>()

function fontOf(cs: CSSStyleDeclaration): string {
  return `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`
}

function charAdvancePx(font: string, ch: string): number {
  if (measureCtx === undefined) measureCtx = document.createElement('canvas').getContext('2d')
  if (!measureCtx) return 0
  let perFont = advanceCache.get(font)
  if (!perFont) advanceCache.set(font, (perFont = new Map()))
  let w = perFont.get(ch)
  if (w === undefined) {
    measureCtx.font = font
    // Blink snaps each advance up to a LayoutUnit; raw floats undercount a
    // 40-glyph line by ~0.5px, enough to miss a pull
    w = Math.ceil(measureCtx.measureText(ch).width * 64) / 64
    perFont.set(ch, w)
  }
  return w
}

class CjkPunctShrinkView {
  private lastSig = ''
  private results = new SettledParagraphCache<MeasuredShrink[]>(shiftShrinks, (r) => r.length === 0)
  private seenSigs = new Set<string>()
  /** per paragraph: the offsets of the line-start brackets trimmed last pass, and of those pinned */
  private trimmed = new WeakMap<ProseMirrorNode, Set<number>>()
  private pinned = new WeakMap<ProseMirrorNode, Set<number>>()
  private frozen = false
  private retryRaf = 0
  private retries = 0
  private resizeObserver?: ResizeObserver
  private lastDomWidth = -1
  private onFontsLoaded = () => {
    // canvas advances measured before a @font-face finished loading are stale
    advanceCache.clear()
    this.invalidate()
    this.measure()
  }
  // style-level w:jc arrives with the doc stylesheet after setContent measured
  private onDocCss = () => {
    this.invalidate()
    this.measure()
  }
  private onPhasedSettled = () => {
    this.invalidate()
    this.measure()
  }

  constructor(
    private view: EditorView,
    private storage: CjkPunctShrinkStorage,
  ) {
    this.measure()
    document.fonts?.addEventListener('loadingdone', this.onFontsLoaded)
    document.addEventListener(DOC_CSS_COMMITTED_EVENT, this.onDocCss)
    document.addEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.onDocCss)
    document.addEventListener(PHASED_CONTENT_SETTLED_EVENT, this.onPhasedSettled)
    if (typeof ResizeObserver !== 'undefined') {
      this.resizeObserver = new ResizeObserver(() => {
        const w = this.view.dom.offsetWidth
        if (w === this.lastDomWidth) return
        this.lastDomWidth = w
        this.invalidate()
        this.measure()
      })
      this.resizeObserver.observe(view.dom)
    }
  }

  /** layout input changed (resize, fonts, stylesheet): every paragraph re-measures */
  private invalidate() {
    this.results.clear()
    this.trimmed = new WeakMap()
    this.pinned = new WeakMap()
    this.restartConvergence()
  }

  private restartConvergence() {
    this.seenSigs.clear()
    this.frozen = false
    this.lastSig = ''
  }

  update(view: EditorView, prevState: EditorState) {
    if (view.state.doc !== prevState.doc) {
      // untouched paragraphs keep their settled results (keyed on node identity)
      this.restartConvergence()
    } else if (
      cjkPunctShrinkPluginKey.getState(view.state) === cjkPunctShrinkPluginKey.getState(prevState)
    ) {
      return
    }
    this.measure()
  }

  destroy() {
    document.fonts?.removeEventListener('loadingdone', this.onFontsLoaded)
    document.removeEventListener(DOC_CSS_COMMITTED_EVENT, this.onDocCss)
    document.removeEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.onDocCss)
    document.removeEventListener(PHASED_CONTENT_SETTLED_EVENT, this.onPhasedSettled)
    this.resizeObserver?.disconnect()
    if (this.retryRaf) cancelAnimationFrame(this.retryRaf)
  }

  private scheduleRetry() {
    if (this.retryRaf || this.retries >= MEASURE_RETRY_MAX) return
    this.retries++
    this.retryRaf = requestAnimationFrame(() => {
      this.retryRaf = 0
      this.measure()
    })
  }

  private measure() {
    if (this.retryRaf) {
      cancelAnimationFrame(this.retryRaf)
      this.retryRaf = 0
    }
    const { view } = this
    // a streamed tail is still landing: measured once, when it has (settled event)
    if (isPhasedContentPending()) return
    // PDF export parks the editor subtree (.app.pv-exporting); any layout
    // read here would force the parked document to re-lay out per print chunk
    if (view.dom.closest('.app.pv-exporting')) {
      this.retries = 0
      this.scheduleRetry()
      return
    }
    const old = cjkPunctShrinkPluginKey.getState(view.state)
    if (!this.storage.enabled && !this.storage.hangPunct) {
      if (old && old !== DecorationSet.empty)
        view.dispatch(
          view.state.tr.setMeta(cjkPunctShrinkPluginKey, []).setMeta('addToHistory', false),
        )
      return
    }
    if (!view.dom.isConnected) {
      this.scheduleRetry()
      return
    }

    const paras: Array<{ node: ProseMirrorNode; pos: number }> = []
    view.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true
      // justification usually comes from the paragraph style, not a direct
      // attr: filter on the rendered alignment in measureParagraph instead
      if (isCandidate(node, this.storage.enabled && this.storage.kana)) paras.push({ node, pos })
      return false
    })

    const shrinks: MeasuredShrink[] = []
    let measurable = paras.length === 0
    this.results.beginPass(view)
    const topLevel = SettledParagraphCache.topLevelDom(view)
    for (const para of paras) {
      const measured = this.results.measure(
        view,
        para.node,
        para.pos,
        (el) => this.measureParagraph(para.node, para.pos, el),
        topLevel.get(para.node),
      )
      if (!measured) continue
      measurable = true
      shrinks.push(...measured)
    }
    if (!measurable) {
      this.scheduleRetry()
      return
    }
    this.retries = 0

    const sig = JSON.stringify(
      shrinks.map((s) => [s.from, s.perChar, s.baseLs, s.lineStart, s.pin]),
    )
    if (sig === this.lastSig) return
    if (this.frozen) return
    if (this.seenSigs.has(sig) || this.seenSigs.size >= MEASURE_SIGS_MAX) {
      this.frozen = true
      console.warn('[docs] cjk-punct-shrink layout did not converge; keeping current decorations')
      return
    }
    this.seenSigs.add(sig)
    this.lastSig = sig

    if (shrinks.length === 0 && (!old || old === DecorationSet.empty)) return
    const pins = shrinks
      .filter((s) => s.pin)
      .map((s) =>
        Decoration.widget(s.from, () => document.createElement('br'), {
          side: -1,
          key: 'topline-pin',
        }),
      )
    const decos = shrinks.map((s) =>
      Decoration.inline(
        s.from,
        s.from + 1,
        {
          class: 'doc-cjkshrink',
          // Word probe 2026-11-18: half an em, whatever the alignment or the document's compression
          style: s.lineStart ? 'margin-left:-0.5em' : shrinkStyle(s.ch, s.perChar, s.baseLs),
        },
        { hang: s.hang === true },
      ),
    )
    view.dispatch(
      view.state.tr
        .setMeta(cjkPunctShrinkPluginKey, [...pins, ...decos])
        .setMeta('addToHistory', false),
    )
  }

  /** null = not measurable right now (hidden / not mounted) → retry */
  private measureParagraph(
    node: ProseMirrorNode,
    pos: number,
    el: HTMLElement,
  ): MeasuredShrink[] | null {
    // Word probe 2026-11-18: with w:kinsoku off nothing is pulled, hung or compressed
    const shrinks = node.attrs.kinsoku === false ? [] : this.measureShrinks(node, pos, el)
    if (shrinks === null || node.attrs.topLinePunct !== true) return shrinks
    if (el.offsetWidth === 0) return null
    const starts = this.lineStartOpeners(node, pos)
    const trimmed = new Set(starts.map((s) => s.from))
    return [...starts, ...shrinks.filter((s) => !trimmed.has(s.from))]
  }

  /** the opening brackets that are the first character of a line */
  private lineStartOpeners(node: ProseMirrorNode, pos: number): MeasuredShrink[] {
    const { view } = this
    const out: MeasuredShrink[] = []
    const was = this.trimmed.get(node)
    const pins = this.pinned.get(node) ?? new Set<number>()
    const now = new Set<number>()
    node.forEach((child, offset) => {
      if (!child.isText) return
      const text = child.text ?? ''
      for (let i = 0; i < text.length; i++) {
        if (!COMPRESSIBLE_OPEN.has(text[i])) continue
        const at = pos + 1 + offset + i
        const first =
          offset + i === 0 ||
          view.coordsAtPos(at - 1, 1).bottom <= view.coordsAtPos(at, 1).top + NOISE
        // trimmed last pass and gone from the line start: half an em narrower, it fitted
        // back on the line above. Word keeps it on the new line, so it is pinned there
        if (!first && was?.has(offset + i)) pins.add(offset + i)
        const pin = pins.has(offset + i)
        if (!first && !pin) continue
        now.add(offset + i)
        out.push({ from: at, ch: text[i], perChar: 0, baseLs: 0, lineStart: true, pin })
      }
    })
    this.trimmed.set(node, now)
    this.pinned.set(node, pins)
    return out
  }

  private measureShrinks(
    node: ProseMirrorNode,
    pos: number,
    el: HTMLElement,
  ): MeasuredShrink[] | null {
    const { view } = this
    if (el.offsetWidth === 0) return null
    const rect = el.getBoundingClientRect()
    if (rect.width === 0) return null
    const zoom = rect.width / el.offsetWidth
    const cs = window.getComputedStyle(el)
    if (cs.direction === 'rtl') return []
    if (!measuresAlignment(cs.textAlign, this.storage.legacyLayout)) return []
    if (!this.storage.enabled && node.attrs.overflowPunct === false) return []
    const justified = cs.textAlign === 'justify'
    // docGrid charSpace letter-spacing (inherited): canvas advances don't see
    // it, so natural per-char advances add it back (Word compresses relative
    // to the grid advance — the caps below scale with it automatically)
    const baseLs = parseFloat(cs.letterSpacing) || 0
    // content-box width: capacity reference for the ragged last line, whose
    // rendered extent shrinks with its own compression (using it would ratchet
    // the keep amount by EPS every round)
    const contentW =
      el.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0)
    const textIndent = parseFloat(cs.textIndent) || 0
    const paraLang = (node.attrs.eaLang as string | null) ?? this.storage.docEastAsiaLang

    // per-character boxes: character rects give the line geometry (straight
    // off the DOM text when it maps cleanly, else via caret coords, both
    // surviving decoration-split text nodes); natural advances come from
    // canvas metrics so active shrink decorations do not feed back into the
    // next round
    const chars: CharBox[] = []
    let breakPositions: number[] = []
    const paraFont = fontOf(cs)
    // balanced grids give full-width spans a larger letter-spacing than the paragraph
    const styleCache = new Map<Element, { font: string; ls: number }>()
    const styleAt = (el: Element | null): { font: string; ls: number } => {
      const parent = spacingHost(el)
      if (!parent) return { font: paraFont, ls: baseLs }
      let st = styleCache.get(parent)
      if (!st) {
        const pcs = window.getComputedStyle(parent)
        st = { font: fontOf(pcs), ls: parseFloat(pcs.letterSpacing) || 0 }
        styleCache.set(parent, st)
      }
      return st
    }
    const eaOf = (child: ProseMirrorNode): boolean => {
      const runLang = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.eaLang as
        string | null | undefined
      return usesEastAsianRules(runLang ?? paraLang)
    }
    let bail = false
    const mapped = mapTextNodes(el, node, pos, eaOf)
    if (mapped) {
      breakPositions = mapped.breaks
      const range = charRange()
      for (const slot of mapped.slots) {
        const { font, ls } = styleAt(slot.dom.parentElement)
        const data = slot.dom.data
        for (let i = 0; i < data.length; i++) {
          range.setStart(slot.dom, i)
          range.setEnd(slot.dom, i + 1)
          const r = range.getBoundingClientRect()
          const width = charAdvancePx(font, data[i]) + ls
          chars.push({
            ch: data[i],
            from: slot.from + i,
            ea: slot.ea,
            width,
            ls,
            top: r.top,
            bottom: r.bottom,
            left: r.left,
            right: r.left + width * zoom,
            rendRight: r.right,
          })
        }
      }
    }
    // positions map linearly inside one DOM text node: only re-resolve at its end
    let dp: { node: Node; offset: number; pos: number } | null = null
    if (!mapped)
      node.forEach((child, offset) => {
        if (bail || !child.isText || !child.text) {
          // atoms (images, tabs, fields) break the char model: skip the paragraph
          if (!child.isText && child.type.name !== 'hardBreak') bail = true
          else if (!child.isText) breakPositions.push(pos + 1 + offset)
          return
        }
        const base = pos + 1 + offset
        const ea = eaOf(child)
        for (let i = 0; i < child.text.length; i++) {
          const from = base + i
          let a: { top: number; bottom: number; left: number; right: number }
          try {
            a = view.coordsAtPos(from, 1)
          } catch {
            bail = true
            return
          }
          dp =
            dp &&
            dp.pos + 1 === from &&
            dp.node.nodeType === Node.TEXT_NODE &&
            dp.offset + 1 < (dp.node as Text).length
              ? { node: dp.node, offset: dp.offset + 1, pos: from }
              : { ...view.domAtPos(from, 1), pos: from }
          const parent =
            dp.node.nodeType === Node.TEXT_NODE
              ? dp.node.parentElement
              : (dp.node as Element | null)
          const st = styleAt(parent)
          const width = charAdvancePx(st.font, child.text[i]) + st.ls
          chars.push({
            ch: child.text[i],
            from,
            ea,
            width,
            ls: st.ls,
            top: a.top,
            bottom: a.bottom,
            left: a.left,
            right: a.left + width * zoom,
          })
        }
      })
    if (bail || chars.length === 0) return []

    // group into rendered lines
    const lines: CharBox[][] = []
    let cur: CharBox[] | null = null
    for (const c of chars) {
      if (cur && sameLine(cur[cur.length - 1], c) && c.left >= cur[0].left - 1) {
        cur.push(c)
      } else {
        cur = [c]
        lines.push(cur)
      }
    }
    // a paragraph a previous round collapsed onto a single line must keep its
    // compression (keep path against the content box), so only empty bails
    if (lines.length === 0) return []

    const active = cjkPunctShrinkPluginKey.getState(view.state)
    const decorated = (c: CharBox) => (active?.find(c.from, c.from + 1).length ?? 0) > 0
    const hungDecorated = (c: CharBox) =>
      (active?.find(c.from, c.from + 1) ?? []).some((d) => d.spec.hang === true)
    const punctBoxesPerLine: CharBox[][] = []
    const candPunctBoxesPerLine: CharBox[][] = []
    const hangModels: HangLineChars[] = []
    const lineModels: ShrinkLineChars[] = lines.map((line, k) => {
      const natural = line.reduce((s, c) => s + c.width, 0)
      const left = Math.min(...line.map((c) => c.left))
      // rendered right edge: the caret after the line's last character
      const lastChar = line[line.length - 1]
      let right = lastChar.rendRight ?? Math.max(...line.map((c) => c.right))
      if (lastChar.rendRight === undefined) {
        try {
          right = view.coordsAtPos(lastChar.from + 1, -1).left
        } catch {
          /* keep the advance-based fallback */
        }
      }
      const punctBoxes = line.filter(isPunct)
      const closeBoxes = punctBoxes.filter(isClose)
      punctBoxesPerLine.push(punctBoxes)
      // pull candidate: the next line's first char plus any kinsoku-close run
      // it would drag along (those cannot start the shortened next line) and
      // the character after an opening bracket (which cannot end a line);
      // a hard break between the lines forbids pulling entirely
      const next = lines[k + 1]
      const lastFrom = line[line.length - 1].from
      const brBetween = next && breakPositions.some((p) => p > lastFrom && p < next[0].from)
      const candWidths: number[] = []
      const candPunctBoxes: CharBox[] = []
      let forced = false
      let candLast: CharBox | null = null
      if (k < lines.length - 1 && next && next.length > 0 && !brBetween) {
        candWidths.push(next[0].width)
        candLast = next[0]
        let j = 1
        while (
          j < next.length &&
          (forbidsLineStart(next[j].ch) || forbidsLineEnd(next[j - 1].ch))
        ) {
          candWidths.push(next[j].width)
          candLast = next[j]
          if (isClose(next[j])) candPunctBoxes.push(next[j])
          if (isKinsokuClose(next[j])) forced = true
          j++
        }
      }
      candPunctBoxesPerLine.push(candPunctBoxes)
      const capBoxes = closeBoxes.length > 0 ? closeBoxes : candPunctBoxes
      // ragged lines (unjustified, the last one, or one ended by a hard break)
      // report their own compressed extent, which would ratchet the keep
      // amount — measure those against the paragraph content box instead
      const ragged = !justified || k === lines.length - 1 || brBetween
      const avail = ragged ? contentW - (k === 0 ? textIndent : 0) : (right - left) / zoom
      // a hung stop keeps its trailing spaces on the line: look past them
      let lastIx = line.length - 1
      while (lastIx > 0 && line[lastIx].ch === ' ') lastIx--
      const last = line[lastIx]
      hangModels.push({
        natural,
        avail,
        hungWidth: isClose(last) && decorated(last) ? last.width : null,
        candWidths: k === lines.length - 1 ? [] : candWidths,
        candEndsWithStop: candLast !== null && isClose(candLast),
      })
      return {
        natural,
        avail,
        punctCount: closeBoxes.length,
        openCount: punctBoxes.length - closeBoxes.length,
        avgPunctW:
          capBoxes.length > 0 ? capBoxes.reduce((s, c) => s + c.width, 0) / capBoxes.length : 0,
        candWidths: k === lines.length - 1 ? [] : candWidths,
        candPunctCount: candPunctBoxes.length,
        forced,
        hungWidth: isClose(last) && hungDecorated(last) ? last.width : null,
        candEndsWithStop: candLast !== null && isClose(candLast),
      }
    })

    const out: MeasuredShrink[] = []
    if (!this.storage.enabled) {
      lines.forEach((line, k) => {
        const decision = decideCjkHang(hangModels[k])
        if (decision === null) return
        let lastIx = line.length - 1
        while (lastIx > 0 && line[lastIx].ch === ' ') lastIx--
        const stop =
          decision === 'keep' ? line[lastIx] : lines[k + 1][hangModels[k].candWidths.length - 1]
        out.push({ from: stop.from, ch: stop.ch, perChar: stop.width, baseLs: stop.ls, hang: true })
      })
      return out
    }
    const decisions = decideCjkShrinks(lineModels)
    // a stop hung from line k is also a closer on line k+1: one decoration only
    const hungFrom = new Set<number>()
    for (let k = 0; k < decisions.length; k++) {
      const perGlyph = decisions[k]
      const lm = lineModels[k]
      if (this.storage.kana && perGlyph === null) {
        const brBetween =
          k + 1 < lines.length &&
          breakPositions.some(
            (p) => p > lines[k][lines[k].length - 1].from && p < lines[k + 1][0].from,
          )
        const kana = kanaShrinks(
          lines[k],
          brBetween ? undefined : lines[k + 1],
          lm,
          node.attrs.overflowPunct !== false,
          node.attrs.autoSpace,
        )
        if (kana) {
          for (const s of kana) {
            if (hungFrom.has(s.from)) continue
            hungFrom.add(s.from)
            out.push(s)
          }
          continue
        }
      }
      // a compression emitted last round that did not pull its chain (own
      // closers decorated, line still short, chain still on the next line):
      // Chromium's breaker rejected the marginal fit, so hang instead
      const pullFailed =
        lm.natural - lm.avail <= NOISE &&
        lm.candWidths.length > 0 &&
        punctBoxesPerLine[k].some((c) => isClose(c) && decorated(c))
      const own = perGlyph === null || pullFailed ? decideCjkHangUnderCompression(lm) : null
      if (own !== null || perGlyph === null) {
        if (own === null) continue
        const line = lines[k]
        const hung = lineModels[k].hungWidth != null
        let lastIx = line.length - 1
        while (lastIx > 0 && line[lastIx].ch === ' ') lastIx--
        const stop = hung ? line[lastIx] : lines[k + 1][lineModels[k].candWidths.length - 1]
        if (!hungFrom.has(stop.from)) {
          hungFrom.add(stop.from)
          out.push({
            from: stop.from,
            ch: stop.ch,
            perChar: stop.width,
            baseLs: stop.ls,
            hang: true,
          })
        }
        const perChar = Math.round(own * 100) / 100
        if (perChar > 0) {
          for (const c of punctBoxesPerLine[k]) {
            if (isClose(c) && !hungFrom.has(c.from))
              out.push({ from: c.from, ch: c.ch, perChar, baseLs: c.ls })
          }
        }
        continue
      }
      const perChar = Math.round(perGlyph * 100) / 100
      if (perChar <= 0) continue
      const boxes = shrinkTargets(lineModels[k], punctBoxesPerLine[k], candPunctBoxesPerLine[k])
      for (const c of boxes) {
        if (!hungFrom.has(c.from)) out.push({ from: c.from, ch: c.ch, perChar, baseLs: c.ls })
      }
    }
    return out
  }
}

const isKana = (c: CharBox) => c.ea && KANA_CLASS.test(c.ch)
const isOpen = (c: CharBox) => c.ea && COMPRESSIBLE_OPEN.has(c.ch)

const glued = (a: CharBox, b: CharBox) =>
  forbidsLineStart(b.ch) || forbidsLineEnd(a.ch) || (LATIN_WORD.test(a.ch) && LATIN_WORD.test(b.ch))

function kanaShrinks(
  line: CharBox[],
  next: CharBox[] | undefined,
  lm: ShrinkLineChars,
  hangs: boolean,
  autoSpace: unknown,
): MeasuredShrink[] | null {
  const sum = (boxes: CharBox[]) => boxes.reduce((s, c) => s + c.width, 0)
  // the quarter-em gaps between East Asian and Latin text: canvas advances do not carry them
  const gaps = (boxes: CharBox[]) => {
    let w = 0
    for (let i = 1; i < boxes.length; i++) {
      const [a, b] = [boxes[i - 1], boxes[i]]
      const wide = isCjkLetterChar(a.ch.codePointAt(0)!) ? a : b
      const latin = wide === a ? b : a
      if (isCjkLetterChar(wide.ch.codePointAt(0)!) && autospaceGap(latin.ch, autoSpace))
        w += (wide.width - wide.ls) / 4
    }
    return w
  }
  let end = line.length
  while (end > 1 && line[end - 1].ch === ' ') end--
  let body = line.slice(0, end)
  let run: CharBox[]
  if (sum(body) + gaps(body) - lm.avail > NOISE) {
    // the line holds a run an earlier round took up (or a hung stop): it is weighed
    // again as if still on the line below, so the answer does not depend on the order
    let at = end - 1
    while (at > 0 && glued(body[at - 1], body[at])) at--
    if (at === 0) return null
    run = body.slice(at)
    body = body.slice(0, at)
  } else {
    // a word broken in the middle has no break point to mend; behind a full stop or a bracket it has
    if (!next || (LATIN_ALNUM.test(body[end - 1].ch) && LATIN_ALNUM.test(next[0].ch))) return null
    let to = 1
    while (to < next.length && glued(next[to - 1], next[to])) to++
    run = next.slice(0, to)
  }
  const last = body[body.length - 1]
  const kana = body.filter(isKana)
  const closers = body.filter(isClose)
  const runKana = run.filter(isKana)
  const openers = run.filter(isOpen)
  const stop = run[run.length - 1]
  const natural = sum(body) + gaps(body)
  // forced up: its first glyph fits, and only what follows it keeps it off the line
  const forced =
    run.slice(1).some(isKinsokuClose) &&
    natural + run[0].width + gaps([last, run[0]]) <= lm.avail + NOISE
  const decision = decideKanaShrink({
    natural,
    avail: lm.avail,
    chars: body.length,
    kanaWidth: sum(kana),
    kanaCount: kana.length,
    punctCredit: forced ? sum(closers) / 2 : 0,
    chain: {
      width: sum(run) + gaps([last, ...run]),
      kanaWidth: sum(runKana),
      kanaCount: runKana.length,
      openerCredit: sum(openers) / 2,
      stopWidth: isClose(stop) ? stop.width : 0,
    },
    hangs,
  })
  if (!decision) return null
  const out: MeasuredShrink[] = []
  // Blink snaps each letter-spacing: over forty kana the half-pixel overshoot would be lost
  const perChar = Math.ceil(decision.perKana * 100) / 100 + 0.02
  if (decision.perKana > 0)
    for (const c of [...kana, ...runKana])
      out.push({ from: c.from, ch: c.ch, perChar, baseLs: c.ls })
  if (decision.punct > 0)
    for (const c of closers)
      out.push({ from: c.from, ch: c.ch, perChar: (decision.punct * c.width) / 2, baseLs: c.ls })
  for (const c of openers) out.push({ from: c.from, ch: c.ch, perChar: c.width / 2, baseLs: c.ls })
  if (decision.halfStop)
    out.push({ from: stop.from, ch: stop.ch, perChar: stop.width / 2, baseLs: stop.ls })
  return out
}

export const kinsokuOffPluginKey = new PluginKey<DecorationSet>('kinsokuOff')

/** letters, digits, marks, format characters, spaces and the non-breaking hyphen: never set free by a list */
const WORDISH = /[\p{L}\p{N}\p{M}\p{Cf}\s\u2011\ud800-\udfff]/u

/**
 * Under a custom list a character it leaves out breaks freely: the ones a
 * standard list names, and any other mark or symbol the browser would hold.
 */
const frees = (ch: string, standard: Set<string>) => standard.has(ch) || !WORDISH.test(ch)

function joiner(): HTMLElement {
  const span = document.createElement('span')
  span.className = 'doc-kr-wj'
  span.textContent = '\u2060'
  return span
}

/**
 * The line-break opportunities that differ from Chromium's own rules.
 * w:kinsoku off (Word probe 2026-11-18): a closing mark may start a line and an
 * opening bracket may end one, so the paragraph gets a break opportunity in
 * front of each closer and behind each opener. The Strict level: Japanese text
 * gets a joiner in front of each character it adds to the line-start set.
 * The Custom level (same probe): the two lists replace the standard ones for
 * text of their language, so a listed character gets a joiner and a standard
 * one that is not listed a break opportunity; Strict wins over a Japanese list.
 */
function kinsokuBreaks(doc: ProseMirrorNode, storage: CjkPunctShrinkStorage): DecorationSet {
  const decos: Decoration[] = []
  const wbr = () => document.createElement('wbr')
  const custom = storage.kinsokuCustom
  const customLang = custom ? kinsokuLang(custom.lang || 'x') : null
  const before = new Set(customLang ? custom!.before : '')
  const after = new Set(customLang ? custom!.after : '')
  const ea = (ch: string | undefined) => !!ch && isEastAsianChar(ch.codePointAt(0)!)
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    const off = node.attrs.kinsoku === false
    if (!off && !storage.strictKinsoku && !customLang) return false
    const paraLang = (node.attrs.eaLang as string | null) ?? storage.docEastAsiaLang
    let prev: string | undefined
    node.forEach((child, offset) => {
      if (!child.isText) {
        prev = undefined
        return
      }
      const text = child.text ?? ''
      const lang =
        (child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.eaLang as
          string | null | undefined) ?? paraLang
      const strict = !off && storage.strictKinsoku && /^ja\b/i.test(lang ?? '')
      const listed = !off && !strict && customLang !== null && kinsokuLang(lang) === customLang
      for (let i = 0; i < text.length; i++) {
        const at = pos + 1 + offset + i
        const ch = text[i]
        const last = i > 0 ? text[i - 1] : prev
        if (strict) {
          if (STRICT_JA_LINE_START.has(ch) && offset + i > 0)
            decos.push(Decoration.widget(at, joiner, { side: -1, key: 'kinsoku-strict' }))
        } else if (listed) {
          if (offset + i === 0) continue
          // a Latin mark is under these rules only where it touches East Asian text
          const wide = (c: string) => c.charCodeAt(0) >= 0x3000
          // `last` is missing behind a picture or a note mark: the listed character still holds on
          if ((last !== undefined && after.has(last)) || before.has(ch))
            decos.push(Decoration.widget(at, joiner, { side: -1, key: 'kinsoku-custom' }))
          else if (
            last !== undefined &&
            ((frees(ch, ANY_STANDARD_LINE_START) && (wide(ch) || ea(last))) ||
              (frees(last, ANY_STANDARD_LINE_END) && (wide(last) || ea(ch))))
          )
            decos.push(Decoration.widget(at, wbr, { side: -1, key: 'kinsoku-free' }))
        } else if (off && ch.charCodeAt(0) >= 0x3000) {
          // Latin punctuation was never under these rules: a break there would split "example.com"
          if (forbidsLineStart(ch) && offset + i > 0)
            decos.push(Decoration.widget(at, wbr, { side: -1, key: 'kinsoku-off' }))
          if (forbidsLineEnd(ch))
            decos.push(Decoration.widget(at + 1, wbr, { side: -1, key: 'kinsoku-off' }))
        }
      }
      prev = text[text.length - 1]
    })
    return false
  })
  return decos.length > 0 ? DecorationSet.create(doc, decos) : DecorationSet.empty
}

/** Typography > Kinsoku level changed: the page follows at once */
export function setKinsokuLevel(
  editor: Editor,
  strict: boolean,
  custom: KinsokuCustom | null,
): void {
  editor.storage.cjkPunctShrink.strictKinsoku = strict
  editor.storage.cjkPunctShrink.kinsokuCustom = custom
  editor.view.dispatch(
    editor.state.tr.setMeta(kinsokuOffPluginKey, true).setMeta('addToHistory', false),
  )
}

export const CjkPunctShrinkExtension = Extension.create({
  name: 'cjkPunctShrink',
  addStorage(): CjkPunctShrinkStorage {
    return {
      enabled: false,
      hangPunct: false,
      kana: false,
      legacyLayout: false,
      docEastAsiaLang: null,
      strictKinsoku: false,
      kinsokuCustom: null,
    }
  },
  addProseMirrorPlugins() {
    const storage = this.storage as CjkPunctShrinkStorage
    return [
      new Plugin({
        key: cjkPunctShrinkPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            noteFloatTransaction(tr)
            const meta = tr.getMeta(cjkPunctShrinkPluginKey) as Decoration[] | undefined
            if (meta)
              return meta.length > 0 ? DecorationSet.create(tr.doc, meta) : DecorationSet.empty
            return old.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
        view: (editorView) => new CjkPunctShrinkView(editorView, storage),
      }),
      new Plugin<DecorationSet>({
        key: kinsokuOffPluginKey,
        state: {
          init: (_, state) => kinsokuBreaks(state.doc, storage),
          apply: (tr, old) =>
            tr.docChanged || tr.getMeta(kinsokuOffPluginKey) ? kinsokuBreaks(tr.doc, storage) : old,
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
      }),
    ]
  },
})
