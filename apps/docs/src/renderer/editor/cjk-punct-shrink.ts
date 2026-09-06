import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { SettledParagraphCache } from './settled-measure'

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
 *   instead, even when it is the line's only compressible glyph.
 * - All of it is gated on the run's effective w:lang w:eastAsia (run rPr >
 *   character style > paragraph style chain > docDefaults; the paragraph mark
 *   does not count): under en-US / ko-KR Word neither hangs nor compresses
 *   nor applies kinsoku — the next line simply starts with the 。 (probes
 *   2026-09-05). No w:lang anywhere keeps the East Asian rules.
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

/** Word applies its East Asian line rules under a CJK w:lang w:eastAsia or none at all */
export function usesEastAsianRules(eastAsiaLang: string | null | undefined): boolean {
  return !eastAsiaLang || /^(ja|zh)(-|$)/i.test(eastAsiaLang)
}

/** decoration style compressing one glyph by perChar px (baseLs: inherited letter-spacing it replaces) */
export function shrinkStyle(ch: string, perChar: number, baseLs: number): string {
  const px = (v: number) => Math.round(v * 100) / 100
  return COMPRESSIBLE_OPEN.has(ch)
    ? `margin-left:${px(-perChar)}px`
    : `letter-spacing:${px(baseLs - perChar)}px`
}
const CJK_RE = /[⺀-〿぀-ヿㇰ-䶿一-鿿豈-﫿＀-￯]/

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
  /** settings.xml compatibilityMode < 15: Word 2010 layout also pulls on unjustified lines */
  legacyLayout: boolean
  /** docDefaults w:lang w:eastAsia: the fallback of the run / paragraph-style chain */
  docEastAsiaLang: string | null
}

/** dispatched on document after the doc-scoped <style> commits (App) */
export const DOC_CSS_COMMITTED_EVENT = 'genoffice:doc-css'

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
  top: number
  bottom: number
  left: number
  right: number
}

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

let measureCtx: CanvasRenderingContext2D | null | undefined
const advanceCache = new Map<string, number>()

function charAdvancePx(cs: CSSStyleDeclaration, ch: string): number {
  if (measureCtx === undefined) measureCtx = document.createElement('canvas').getContext('2d')
  if (!measureCtx) return 0
  const font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`
  const key = `${font}|${ch}`
  let w = advanceCache.get(key)
  if (w === undefined) {
    measureCtx.font = font
    // Blink snaps each advance up to a LayoutUnit; raw floats undercount a
    // 40-glyph line by ~0.5px, enough to miss a pull
    w = Math.ceil(measureCtx.measureText(ch).width * 64) / 64
    advanceCache.set(key, w)
  }
  return w
}

function sameLine(a: { top: number; bottom: number }, b: { top: number; bottom: number }): boolean {
  return a.top < b.bottom - 1 && a.bottom > b.top + 1
}

class CjkPunctShrinkView {
  private lastSig = ''
  private results = new SettledParagraphCache<MeasuredShrink[]>()
  private seenSigs = new Set<string>()
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

  constructor(
    private view: EditorView,
    private storage: CjkPunctShrinkStorage,
  ) {
    this.measure()
    document.fonts?.addEventListener('loadingdone', this.onFontsLoaded)
    document.addEventListener(DOC_CSS_COMMITTED_EVENT, this.onDocCss)
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

  private invalidate() {
    this.seenSigs.clear()
    this.results.clear()
    this.frozen = false
    this.lastSig = ''
  }

  update(view: EditorView, prevState: EditorState) {
    if (view.state.doc !== prevState.doc) {
      this.invalidate()
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
      const text = node.textContent
      if (text.length > PARA_CHAR_BUDGET) return false
      if (!CJK_RE.test(text)) return false
      let has = false
      for (const ch of text) if (COMPRESSIBLE.has(ch)) has = true
      if (!has) return false
      paras.push({ node, pos })
      return false
    })

    const shrinks: MeasuredShrink[] = []
    let measurable = paras.length === 0
    this.results.beginPass(view)
    for (const para of paras) {
      const measured = this.results.measure(view, para.node, para.pos, () =>
        this.measureParagraph(para.node, para.pos),
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

    const sig = JSON.stringify(shrinks.map((s) => [s.from, s.perChar, s.baseLs]))
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
    const decos = shrinks.map((s) =>
      Decoration.inline(s.from, s.from + 1, {
        class: 'doc-cjkshrink',
        style: shrinkStyle(s.ch, s.perChar, s.baseLs),
      }),
    )
    view.dispatch(
      view.state.tr.setMeta(cjkPunctShrinkPluginKey, decos).setMeta('addToHistory', false),
    )
  }

  /** null = not measurable right now (hidden / not mounted) → retry */
  private measureParagraph(node: ProseMirrorNode, pos: number): MeasuredShrink[] | null {
    const { view } = this
    const el = view.nodeDOM(pos)
    if (!(el instanceof HTMLElement)) return null
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

    // per-character boxes: caret coords give the line geometry (they survive
    // decoration-split text nodes); natural advances come from canvas metrics
    // so active shrink decorations do not feed back into the next round
    const chars: CharBox[] = []
    const breakPositions: number[] = []
    const styleCache = new Map<Element, CSSStyleDeclaration>()
    let bail = false
    // positions map linearly inside one DOM text node: only re-resolve at its end
    let dp: { node: Node; offset: number; pos: number } | null = null
    node.forEach((child, offset) => {
      if (bail || !child.isText || !child.text) {
        // atoms (images, tabs, fields) break the char model: skip the paragraph
        if (!child.isText && child.type.name !== 'hardBreak') bail = true
        else if (!child.isText) breakPositions.push(pos + 1 + offset)
        return
      }
      const base = pos + 1 + offset
      const runLang = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.eaLang as
        string | null | undefined
      const ea = usesEastAsianRules(runLang ?? paraLang)
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
          dp.node.nodeType === Node.TEXT_NODE ? dp.node.parentElement : (dp.node as Element | null)
        let pcs = parent ? styleCache.get(parent) : undefined
        if (!pcs && parent) {
          pcs = window.getComputedStyle(parent)
          styleCache.set(parent, pcs)
        }
        const width = charAdvancePx(pcs ?? cs, child.text[i]) + baseLs
        chars.push({
          ch: child.text[i],
          from,
          ea,
          width,
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
    const punctBoxesPerLine: CharBox[][] = []
    const candPunctBoxesPerLine: CharBox[][] = []
    const hangModels: HangLineChars[] = []
    const lineModels: ShrinkLineChars[] = lines.map((line, k) => {
      const natural = line.reduce((s, c) => s + c.width, 0)
      const left = Math.min(...line.map((c) => c.left))
      // rendered right edge: the caret after the line's last character
      let right = Math.max(...line.map((c) => c.right))
      try {
        right = view.coordsAtPos(line[line.length - 1].from + 1, -1).left
      } catch {
        /* keep the advance-based fallback */
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
      const last = line[line.length - 1]
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
      }
    })

    const out: MeasuredShrink[] = []
    if (!this.storage.enabled) {
      lines.forEach((line, k) => {
        const decision = decideCjkHang(hangModels[k])
        if (decision === null) return
        const stop =
          decision === 'keep'
            ? line[line.length - 1]
            : lines[k + 1][hangModels[k].candWidths.length - 1]
        out.push({ from: stop.from, ch: stop.ch, perChar: stop.width, baseLs })
      })
      return out
    }
    const decisions = decideCjkShrinks(lineModels)
    for (let k = 0; k < decisions.length; k++) {
      const perGlyph = decisions[k]
      if (perGlyph === null) continue
      const perChar = Math.round(perGlyph * 100) / 100
      if (perChar <= 0) continue
      const boxes = shrinkTargets(lineModels[k], punctBoxesPerLine[k], candPunctBoxesPerLine[k])
      for (const c of boxes) out.push({ from: c.from, ch: c.ch, perChar, baseLs })
    }
    return out
  }
}

export const CjkPunctShrinkExtension = Extension.create({
  name: 'cjkPunctShrink',
  addStorage(): CjkPunctShrinkStorage {
    return { enabled: false, hangPunct: false, legacyLayout: false, docEastAsiaLang: null }
  },
  addProseMirrorPlugins() {
    const storage = this.storage as CjkPunctShrinkStorage
    return [
      new Plugin({
        key: cjkPunctShrinkPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
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
    ]
  },
})
