import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { rangeSlot } from '../dom-range'
import { SettledParagraphCache, noteFloatTransaction } from './settled-measure'
import { PHASED_CONTENT_SETTLED_EVENT, isPhasedContentPending } from '../phased-content'
import { DOC_CSS_COMMITTED_EVENT } from './cjk-punct-shrink'

/**
 * Word 2013+ (settings compatibilityMode >= 15) justified line breaking pulls
 * an extra word onto a line by shrinking the line's spaces. CSS justification
 * only stretches, so Chromium wraps earlier than Word and long justified
 * documents drift pages apart. This extension re-creates Word's pull rule as
 * display-only inline decorations: the affected line's space characters get a
 * negative word-spacing sized so Chromium's greedy breaker takes the same
 * word, and text-align:justify re-distributes whatever slack remains.
 *
 * Pull rule (Word for Mac probes, 2026-08-27): with the candidate word pulled
 * the line overflows its column by delta and holds S space chars; Word pulls iff
 *   1. delta <= 25% of the line's total space width (spaces shrink to >= 75%), and
 *   2. delta/S <= 1/2 * (w + gap - delta)/(S - gap chars) — shrink per space
 *      may cost at most half the per-space stretch that pulling avoids.
 * Verified across 5/10/15/20 spaces, 14/28pt, candidate widths 6-71pt;
 * compatibilityMode 14 documents never shrink.
 */

export interface ShrinkGap {
  /** natural (unshrunk) width, layout px */
  width: number
  /** number of space characters */
  chars: number
  from: number
  to: number
}

export interface ShrinkLine {
  wordWidths: number[]
  /** gaps between words on this line (zero-char entries join mark-split words) */
  gaps: ShrinkGap[]
  /** the wrap-point gap after the last word; null = hard break / paragraph end */
  boundary: ShrinkGap | null
  /** rendered (justify-stretched) width the line fills, layout px */
  avail: number
  /** natural width of the next line's first word; null = no pull candidate */
  nextWordWidth: number | null
  /** natural width of that word's head up to its first break opportunity (a
   *  hyphen fragment); null = the word breaks nowhere */
  nextFragmentWidth?: number | null
}

export interface ShrinkDecision {
  gaps: ShrinkGap[]
  /** negative word-spacing per space char, px */
  perChar: number
}

const SPACE_SHRINK_MAX = 0.25
const SHRINK_VS_STRETCH = 0.5
/** total overshoot (px) so Chromium's breaker definitely pulls the word */
const SHRINK_EPS = 0.5
/** imbalance floor (px): Blink breaks lines in 1/64px LayoutUnits, so any
 *  overflow of one unit or more is a real wrap (a Cyrillic Times line that
 *  Word fits by twip rounding overflows Chromium by ~1/64px) */
const NOISE = 1 / 128
/** Blink's LayoutUnit: a cap missed by less than one is a measurement artefact */
const LAYOUT_UNIT = 1 / 64

export function decideLineShrinks(lines: ShrinkLine[]): Array<ShrinkDecision | null> {
  return lines.map((line) => {
    const spaceW = line.gaps.reduce((s, g) => s + g.width, 0)
    const spaceChars = line.gaps.reduce((s, g) => s + g.chars, 0)
    const natural = line.wordWidths.reduce((s, w) => s + w, 0) + spaceW
    const needed = natural - line.avail
    if (needed > NOISE) {
      // the line already holds a word pulled by a previous round: keep the
      // compression that fits it (recomputed fresh from natural widths)
      if (spaceChars === 0) return null
      return { gaps: line.gaps, perChar: (needed + SHRINK_EPS) / spaceChars }
    }
    // a zero-char boundary is a hyphen break: the line's own spaces carry the pull
    const b = line.boundary
    if (!b || line.nextWordWidth == null) return null
    const S = spaceChars + b.chars
    const sPrev = spaceChars
    if (sPrev < 1) return null
    // the whole word first (Word takes the longest candidate that fits), then
    // the fragment before its first break opportunity
    const candidates = [line.nextWordWidth]
    if (line.nextFragmentWidth != null && line.nextFragmentWidth < line.nextWordWidth)
      candidates.push(line.nextFragmentWidth)
    for (const w of candidates) {
      const delta = natural + b.width + w - line.avail
      if (delta <= NOISE) return null
      if (delta > SPACE_SHRINK_MAX * (spaceW + b.width) + LAYOUT_UNIT) continue
      if (delta / S > (SHRINK_VS_STRETCH * (w + b.width - delta)) / sPrev) continue
      return { gaps: [...line.gaps, b], perChar: (delta + SHRINK_EPS) / S }
    }
    return null
  })
}

const BREAK_HYPHENS = new Set(['-', '\u2010', '\u2013'])

/** length of the word's head through its first hyphen Chromium breaks after
 *  (anything but whitespace or another hyphen following; digits on either
 *  side break too, "53:4-" | "12"), or 0 */
export function breakOpportunity(text: string): number {
  for (let i = 1; i < text.length - 1; i++) {
    if (!BREAK_HYPHENS.has(text[i])) continue
    const next = text[i + 1]
    if (/\s/.test(next) || BREAK_HYPHENS.has(next)) continue
    return i + 1
  }
  return 0
}

// ── DOM measurement / decoration plumbing ──────────────────────────────────

export interface JustifyShrinkStorage {
  /** settings.xml compatibilityMode >= 15 (legacy modes never shrink) */
  enabled: boolean
}

declare module '@tiptap/core' {
  interface Storage {
    justifyShrink: JustifyShrinkStorage
  }
}

export const justifyShrinkPluginKey = new PluginKey<DecorationSet>('justifyShrink')

/** scripts that break lines without spaces or reorder visually (bidi) —
 *  the word-token model does not hold, leave those paragraphs untouched */
const SKIP_SCRIPT_RE = new RegExp(
  '[\\u0590-\\u08FF\\u200F\\uFB1D-\\uFDFF\\uFE70-\\uFEFF' + // RTL scripts
    '\\u1100-\\u11FF\\u2E80-\\u303F\\u3040-\\u30FF\\u3130-\\u318F' + // jamo, CJK punct, kana
    '\\u31F0-\\u4DBF\\u4E00-\\u9FFF\\uA960-\\uA97F\\uAC00-\\uD7FF' + // CJK, hangul
    '\\uF900-\\uFAFF\\uFE30-\\uFE4F\\uFF00-\\uFFEF]', // compat ideographs, fullwidth forms
)

/** Justification usually comes from the paragraph style (Normal w:jc=both) and
 *  then is not a node attr: an inherited (null) alignment stays a candidate and
 *  the rendered text-align decides in measureParagraph. */
export function isShrinkCandidate(node: {
  attrs?: Record<string, unknown>
  textContent: string
}): boolean {
  const align = node.attrs?.align
  if (align != null && align !== 'justify') return false
  const text = node.textContent
  if (!text.includes(' ') || SKIP_SCRIPT_RE.test(text)) return false
  // a tab after the first space absorbs any shrink of the spaces before it
  // (the segment re-anchors at its stop), so the word model only holds for
  // leading tabs ("5.<tab>The claim...", "<tab>First line indent")
  if (text.lastIndexOf('\t') > text.indexOf(' ')) return false
  return true
}

const MEASURE_RETRY_MAX = 10
const MEASURE_SIGS_MAX = 12

interface MeasuredShrink {
  from: number
  to: number
  perChar: number
}

type Token =
  | { kind: 'word'; from: number; to: number; atom: boolean; text?: string }
  | { kind: 'space'; from: number; to: number; chars: number }
  | { kind: 'break' }

interface WordBox {
  width: number
  top: number
  bottom: number
  left: number
  right: number
  /** width of the head before the first break opportunity (hyphen fragment) */
  frag?: number
}

/** a word Chromium broke at a hyphen: head on one line, tail on the next */
interface SplitWord {
  head: WordBox
  tail: WordBox
}

const HYPHEN_GAP: ShrinkGap = { width: 0, chars: 0, from: 0, to: 0 }

interface LineAcc {
  words: WordBox[]
  gaps: ShrinkGap[]
  boundary: ShrinkGap | null
  left: number
  right: number
}

let spaceCtx: CanvasRenderingContext2D | null | undefined
const wordRange = rangeSlot()

function spaceAdvancePx(cs: CSSStyleDeclaration): number {
  if (spaceCtx === undefined) spaceCtx = document.createElement('canvas').getContext('2d')
  if (!spaceCtx) return 4
  spaceCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`
  return spaceCtx.measureText(' ').width + (parseFloat(cs.letterSpacing) || 0)
}

/** same rendered line = the boxes share more than half the shorter box's height.
 *  Glyph boxes (ascent+descent) are taller than tight line-heights (w:line < 240
 *  auto, or a 1.0 face factor under Blink's Mac ascent bump), so consecutive
 *  lines overlap by a few px and a hairline test would chain the whole
 *  paragraph into one line */
export function sameLine(
  a: { top: number; bottom: number },
  b: { top: number; bottom: number },
): boolean {
  const overlap = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top)
  return overlap > Math.min(a.bottom - a.top, b.bottom - b.top) / 2
}

class JustifyShrinkView {
  private lastSig = ''
  private seenSigs = new Set<string>()
  private frozen = false
  private retryRaf = 0
  private retries = 0
  private resizeObserver?: ResizeObserver
  private lastDomWidth = -1
  private results = new SettledParagraphCache<MeasuredShrink[]>(
    (r, d) => r.map((s) => ({ ...s, from: s.from + d, to: s.to + d })),
    (r) => r.length === 0,
  )
  private onFontsLoaded = () => {
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
    private storage: JustifyShrinkStorage,
  ) {
    this.measure()
    document.fonts?.addEventListener('loadingdone', this.onFontsLoaded)
    document.addEventListener(DOC_CSS_COMMITTED_EVENT, this.onDocCss)
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

  /** layout input changed (resize, fonts): every paragraph re-measures */
  private invalidate() {
    this.results.clear()
    this.restartConvergence()
  }

  private restartConvergence() {
    this.seenSigs.clear()
    this.frozen = false
    // decorations may have been dropped with the old doc (setContent/reload):
    // an unchanged shrink list must still re-dispatch
    this.lastSig = ''
  }

  update(view: EditorView, prevState: EditorState) {
    if (view.state.doc !== prevState.doc) {
      // untouched paragraphs keep their settled results (keyed on node identity)
      this.restartConvergence()
    } else if (
      justifyShrinkPluginKey.getState(view.state) === justifyShrinkPluginKey.getState(prevState)
    ) {
      return
    }
    this.measure()
  }

  destroy() {
    document.fonts?.removeEventListener('loadingdone', this.onFontsLoaded)
    document.removeEventListener(DOC_CSS_COMMITTED_EVENT, this.onDocCss)
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
    const old = justifyShrinkPluginKey.getState(view.state)
    if (!this.storage.enabled) {
      if (old && old !== DecorationSet.empty)
        view.dispatch(
          view.state.tr.setMeta(justifyShrinkPluginKey, []).setMeta('addToHistory', false),
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
      if (isShrinkCandidate(node)) paras.push({ node, pos })
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

    const sig = JSON.stringify(shrinks.map((s) => [s.from, s.to, s.perChar]))
    if (sig === this.lastSig) return
    if (this.frozen) return
    if (this.seenSigs.has(sig) || this.seenSigs.size >= MEASURE_SIGS_MAX) {
      this.frozen = true
      console.warn('[docs] justify-shrink layout did not converge; keeping current decorations')
      return
    }
    this.seenSigs.add(sig)
    this.lastSig = sig

    if (shrinks.length === 0 && (!old || old === DecorationSet.empty)) return
    const decos = shrinks.map((s) =>
      Decoration.inline(s.from, s.to, {
        class: 'doc-jshrink',
        style: `word-spacing:-${s.perChar}px`,
      }),
    )
    view.dispatch(
      view.state.tr.setMeta(justifyShrinkPluginKey, decos).setMeta('addToHistory', false),
    )
  }

  /** null = not measurable right now (hidden / not mounted) → retry */
  private measureParagraph(
    node: ProseMirrorNode,
    pos: number,
    el: HTMLElement,
  ): MeasuredShrink[] | null {
    const { view } = this
    if (el.offsetWidth === 0) return null
    const rect = el.getBoundingClientRect()
    if (rect.width === 0) return null
    // rects are screen px (page zoom transform); emitted widths are layout px
    const zoom = rect.width / el.offsetWidth
    const cs = window.getComputedStyle(el)
    if (cs.direction === 'rtl' || cs.textAlign !== 'justify') return []
    // content-box width (layout px): capacity reference for the ragged last
    // line, whose rendered extent shrinks with its own compression
    const contentW =
      el.clientWidth - (parseFloat(cs.paddingLeft) || 0) - (parseFloat(cs.paddingRight) || 0)
    const textIndent = parseFloat(cs.textIndent) || 0

    const tokens: Token[] = []
    node.forEach((child, offset) => {
      const base = pos + 1 + offset
      if (child.isText && child.text) {
        const re = /( +)|[^ ]+/g
        let m: RegExpExecArray | null
        while ((m = re.exec(child.text))) {
          if (m[1])
            tokens.push({
              kind: 'space',
              from: base + m.index,
              to: base + m.index + m[1].length,
              chars: m[1].length,
            })
          else
            tokens.push({
              kind: 'word',
              from: base + m.index,
              to: base + m.index + m[0].length,
              atom: false,
              text: m[0],
            })
        }
      } else if (child.type.name === 'hardBreak') {
        tokens.push({ kind: 'break' })
      } else {
        tokens.push({ kind: 'word', from: base, to: base + child.nodeSize, atom: true })
      }
    })

    const styleCache = new Map<Element, number>()
    // a space starts where the previous word ended: the DOM is static within a pass
    const domCache = new Map<number, { node: Node; offset: number }>()
    const domAt = (p: number) => {
      let d = domCache.get(p)
      if (!d) domCache.set(p, (d = view.domAtPos(p)))
      return d
    }
    const spaceGap = (t: Extract<Token, { kind: 'space' }>): ShrinkGap => {
      const dp = domAt(t.from)
      const parent =
        dp.node.nodeType === Node.TEXT_NODE ? dp.node.parentElement : (dp.node as Element)
      let adv = styleCache.get(parent ?? el)
      if (adv === undefined) {
        adv = spaceAdvancePx(window.getComputedStyle(parent ?? el))
        styleCache.set(parent ?? el, adv)
      }
      return { width: adv * t.chars, chars: t.chars, from: t.from, to: t.to }
    }

    const rangeRects = (from: number, to: number): DOMRect[] | null => {
      const a = domAt(from)
      const b = domAt(to)
      const range = wordRange()
      try {
        range.setStart(a.node, a.offset)
        range.setEnd(b.node, b.offset)
      } catch {
        return null
      }
      return Array.from(range.getClientRects()).filter((r) => r.width > 0.01)
    }
    const boxOf = (rects: DOMRect[]): WordBox => {
      const box: WordBox = {
        width: 0,
        top: Infinity,
        bottom: -Infinity,
        left: Infinity,
        right: -Infinity,
      }
      for (const r of rects) {
        box.top = Math.min(box.top, r.top)
        box.bottom = Math.max(box.bottom, r.bottom)
        box.left = Math.min(box.left, r.left)
        box.right = Math.max(box.right, r.right)
      }
      box.width = (box.right - box.left) / zoom
      return box
    }
    // 'wrapped' = the token spans two rendered lines other than at a hyphen:
    // the word model breaks down
    const measureWord = (
      t: Extract<Token, { kind: 'word' }>,
    ): WordBox | SplitWord | null | 'wrapped' => {
      let rects: DOMRect[] | null
      const dom = t.atom ? view.nodeDOM(t.from) : null
      if (dom instanceof HTMLElement) {
        rects = [dom.getBoundingClientRect()]
      } else {
        rects = rangeRects(t.from, t.to)
        if (!rects) return 'wrapped'
      }
      if (rects.length === 0) return null // zero-width (hidden run): ignore
      const top = Math.min(...rects.map((r) => r.top))
      const first = rects.filter((r) => r.top - top <= r.height / 2)
      const rest = rects.filter((r) => r.top - top > r.height / 2)
      const headLen = t.text ? breakOpportunity(t.text) : 0
      if (rest.length > 0) {
        if (!headLen) return 'wrapped'
        const tail = boxOf(rest)
        for (const r of rest) if (r.top - tail.top > r.height / 2) return 'wrapped'
        return { head: boxOf(first), tail }
      }
      const box = boxOf(first)
      if (headLen) {
        const head = rangeRects(t.from, t.from + headLen)
        if (head && head.length > 0) {
          const frag = boxOf(head).width
          if (frag > 0 && frag < box.width) box.frag = frag
        }
      }
      return box
    }

    const lines: LineAcc[] = []
    let cur: LineAcc | null = null
    let lastWord: WordBox | null = null
    let pendingSpaces: Array<Extract<Token, { kind: 'space' }>> = []
    let pendingBreak = false
    let pendingHyphen = false
    let pendingWord: WordBox | null = null // mark-split word pieces merge until a space

    const mergedGap = (): ShrinkGap | null => {
      if (pendingSpaces.length === 0) return null
      const gs = pendingSpaces.map(spaceGap)
      return {
        width: gs.reduce((s, g) => s + g.width, 0),
        chars: gs.reduce((s, g) => s + g.chars, 0),
        from: gs[0].from,
        to: gs[gs.length - 1].to,
      }
    }

    const flushWord = (): void => {
      if (!pendingWord) return
      const w = pendingWord
      pendingWord = null
      if (cur && lastWord && sameLine(w, lastWord) && !pendingBreak && !pendingHyphen) {
        cur.gaps.push(mergedGap() ?? HYPHEN_GAP)
        cur.words.push(w)
        cur.left = Math.min(cur.left, w.left)
        cur.right = Math.max(cur.right, w.right)
      } else {
        if (cur) cur.boundary = pendingBreak ? null : pendingHyphen ? HYPHEN_GAP : mergedGap()
        cur = { words: [w], gaps: [], boundary: null, left: w.left, right: w.right }
        lines.push(cur)
      }
      lastWord = w
      pendingSpaces = []
      pendingBreak = false
      pendingHyphen = false
    }

    // pieces of one visual word split by mark boundaries: merge (a line
    // mismatch means the compound wrapped mid-word — bail)
    const joinPiece = (w: WordBox): boolean => {
      if (!pendingWord) {
        pendingWord = w
        return true
      }
      if (!sameLine(w, pendingWord)) return false
      pendingWord = {
        width: pendingWord.width + w.width,
        top: Math.min(pendingWord.top, w.top),
        bottom: Math.max(pendingWord.bottom, w.bottom),
        left: Math.min(pendingWord.left, w.left),
        right: Math.max(pendingWord.right, w.right),
        ...(pendingWord.frag != null
          ? { frag: pendingWord.frag }
          : w.frag != null
            ? { frag: pendingWord.width + w.frag }
            : {}),
      }
      return true
    }

    for (const t of tokens) {
      if (t.kind === 'space') {
        flushWord()
        pendingSpaces.push(t)
      } else if (t.kind === 'break') {
        flushWord()
        pendingBreak = true
        pendingSpaces = []
      } else {
        const w = measureWord(t)
        if (w === 'wrapped') return []
        if (w === null) continue
        if ('tail' in w) {
          // the head ends its line at the hyphen; the tail opens the next one
          if (!joinPiece(w.head)) return []
          flushWord()
          pendingHyphen = true
          pendingWord = w.tail
        } else if (!joinPiece(w)) return []
      }
    }
    flushWord()
    // a paragraph a previous round compressed onto a single line must keep its
    // compression (needed > 0 path), so only a lineless paragraph bails
    if (lines.length === 0) return []

    const shrinkLines: ShrinkLine[] = lines.map((l, k) => ({
      wordWidths: l.words.map((w) => w.width),
      gaps: l.gaps,
      boundary: l.boundary,
      // justified lines stretch to their true capacity (float-aware); the
      // ragged last line reports its own compressed extent instead, which
      // would ratchet the keep-compression by EPS every round — use the
      // paragraph content box (minus the first line's text-indent) there
      avail:
        k === lines.length - 1 ? contentW - (k === 0 ? textIndent : 0) : (l.right - l.left) / zoom,
      nextWordWidth: l.boundary ? (lines[k + 1]?.words[0]?.width ?? null) : null,
      nextFragmentWidth: l.boundary ? (lines[k + 1]?.words[0]?.frag ?? null) : null,
    }))

    const out: MeasuredShrink[] = []
    for (const d of decideLineShrinks(shrinkLines)) {
      if (!d) continue
      const perChar = Math.round(d.perChar * 100) / 100
      if (perChar <= 0) continue
      for (const g of d.gaps) {
        if (g.chars === 0) continue
        out.push({ from: g.from, to: g.to, perChar })
      }
    }
    return out
  }
}

export const JustifyShrinkExtension = Extension.create({
  name: 'justifyShrink',
  addStorage(): JustifyShrinkStorage {
    return { enabled: false }
  },
  addProseMirrorPlugins() {
    const storage = this.storage as JustifyShrinkStorage
    return [
      new Plugin({
        key: justifyShrinkPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            noteFloatTransaction(tr)
            const meta = tr.getMeta(justifyShrinkPluginKey) as Decoration[] | undefined
            if (meta)
              // create() consumes (nulls out) entries of the array it is given;
              // the meta array must survive for the App-level 'transaction'
              // listener that re-anchors the pagination pass on it
              return meta.length > 0 ? DecorationSet.create(tr.doc, [...meta]) : DecorationSet.empty
            return old.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
        view: (editorView) => new JustifyShrinkView(editorView, storage),
      }),
    ]
  },
})
