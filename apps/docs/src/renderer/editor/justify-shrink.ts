import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

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
/** measurement noise floor (px): smaller imbalances are not real shrink state */
const NOISE = 0.25

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
    const b = line.boundary
    const w = line.nextWordWidth
    if (!b || b.chars === 0 || w == null) return null
    const delta = natural + b.width + w - line.avail
    if (delta <= NOISE) return null
    const S = spaceChars + b.chars
    const sPrev = S - b.chars
    if (sPrev < 1) return null
    if (delta > SPACE_SHRINK_MAX * (spaceW + b.width)) return null
    if (delta / S > (SHRINK_VS_STRETCH * (w + b.width - delta)) / sPrev) return null
    return { gaps: [...line.gaps, b], perChar: (delta + SHRINK_EPS) / S }
  })
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

const MEASURE_RETRY_MAX = 10
const MEASURE_SIGS_MAX = 12

interface MeasuredShrink {
  from: number
  to: number
  perChar: number
}

type Token =
  | { kind: 'word'; from: number; to: number; atom: boolean }
  | { kind: 'space'; from: number; to: number; chars: number }
  | { kind: 'break' }

interface WordBox {
  width: number
  top: number
  bottom: number
  left: number
  right: number
}

interface LineAcc {
  words: WordBox[]
  gaps: ShrinkGap[]
  boundary: ShrinkGap | null
  left: number
  right: number
}

let spaceCtx: CanvasRenderingContext2D | null | undefined

function spaceAdvancePx(cs: CSSStyleDeclaration): number {
  if (spaceCtx === undefined) spaceCtx = document.createElement('canvas').getContext('2d')
  if (!spaceCtx) return 4
  spaceCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`
  return spaceCtx.measureText(' ').width + (parseFloat(cs.letterSpacing) || 0)
}

/** same rendered line = the boxes overlap vertically by more than a hairline */
function sameLine(a: WordBox, b: WordBox): boolean {
  return a.top < b.bottom - 1 && a.bottom > b.top + 1
}

class JustifyShrinkView {
  private lastSig = ''
  private seenSigs = new Set<string>()
  private frozen = false
  private retryRaf = 0
  private retries = 0
  private resizeObserver?: ResizeObserver
  private lastDomWidth = -1
  private onFontsLoaded = () => {
    this.invalidate()
    this.measure()
  }

  constructor(
    private view: EditorView,
    private storage: JustifyShrinkStorage,
  ) {
    this.measure()
    document.fonts?.addEventListener('loadingdone', this.onFontsLoaded)
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
    this.frozen = false
    // decorations may have been dropped with the old doc (setContent/reload):
    // an unchanged shrink list must still re-dispatch
    this.lastSig = ''
  }

  update(view: EditorView, prevState: EditorState) {
    if (view.state.doc !== prevState.doc) {
      this.invalidate()
    } else if (
      justifyShrinkPluginKey.getState(view.state) === justifyShrinkPluginKey.getState(prevState)
    ) {
      return
    }
    this.measure()
  }

  destroy() {
    document.fonts?.removeEventListener('loadingdone', this.onFontsLoaded)
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
      if (node.attrs?.align !== 'justify') return false
      const text = node.textContent
      if (!text.includes(' ') || text.includes('\t') || SKIP_SCRIPT_RE.test(text)) return false
      paras.push({ node, pos })
      return false
    })

    const shrinks: MeasuredShrink[] = []
    let measurable = paras.length === 0
    for (const para of paras) {
      const measured = this.measureParagraph(para.node, para.pos)
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
  private measureParagraph(node: ProseMirrorNode, pos: number): MeasuredShrink[] | null {
    const { view } = this
    const el = view.nodeDOM(pos)
    if (!(el instanceof HTMLElement)) return null
    if (el.offsetWidth === 0) return null
    const rect = el.getBoundingClientRect()
    if (rect.width === 0) return null
    // rects are screen px (page zoom transform); emitted widths are layout px
    const zoom = rect.width / el.offsetWidth
    const cs = window.getComputedStyle(el)
    if (cs.direction === 'rtl') return []
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
            })
        }
      } else if (child.type.name === 'hardBreak') {
        tokens.push({ kind: 'break' })
      } else {
        tokens.push({ kind: 'word', from: base, to: base + child.nodeSize, atom: true })
      }
    })

    const styleCache = new Map<Element, number>()
    const spaceGap = (t: Extract<Token, { kind: 'space' }>): ShrinkGap => {
      const dp = view.domAtPos(t.from)
      const parent =
        dp.node.nodeType === Node.TEXT_NODE ? dp.node.parentElement : (dp.node as Element)
      let adv = styleCache.get(parent ?? el)
      if (adv === undefined) {
        adv = spaceAdvancePx(window.getComputedStyle(parent ?? el))
        styleCache.set(parent ?? el, adv)
      }
      return { width: adv * t.chars, chars: t.chars, from: t.from, to: t.to }
    }

    // 'wrapped' = the token spans two rendered lines: the word model breaks down
    const measureWord = (t: Extract<Token, { kind: 'word' }>): WordBox | null | 'wrapped' => {
      let rects: DOMRect[]
      const dom = t.atom ? view.nodeDOM(t.from) : null
      if (dom instanceof HTMLElement) {
        rects = [dom.getBoundingClientRect()]
      } else {
        const a = view.domAtPos(t.from)
        const b = view.domAtPos(t.to)
        const range = document.createRange()
        try {
          range.setStart(a.node, a.offset)
          range.setEnd(b.node, b.offset)
        } catch {
          return 'wrapped'
        }
        rects = Array.from(range.getClientRects()).filter((r) => r.width > 0.01)
      }
      if (rects.length === 0) return null // zero-width (hidden run): ignore
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
      for (const r of rects) if (r.top - box.top > r.height / 2) return 'wrapped'
      box.width = (box.right - box.left) / zoom
      return box
    }

    const lines: LineAcc[] = []
    let cur: LineAcc | null = null
    let lastWord: WordBox | null = null
    let pendingSpaces: Array<Extract<Token, { kind: 'space' }>> = []
    let pendingBreak = false
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
      if (cur && lastWord && sameLine(w, lastWord) && !pendingBreak) {
        cur.gaps.push(mergedGap() ?? { width: 0, chars: 0, from: 0, to: 0 })
        cur.words.push(w)
        cur.left = Math.min(cur.left, w.left)
        cur.right = Math.max(cur.right, w.right)
      } else {
        if (cur) cur.boundary = pendingBreak ? null : mergedGap()
        cur = { words: [w], gaps: [], boundary: null, left: w.left, right: w.right }
        lines.push(cur)
      }
      lastWord = w
      pendingSpaces = []
      pendingBreak = false
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
        if (pendingWord) {
          // pieces of one visual word split by mark boundaries: merge (a line
          // mismatch means the compound wrapped mid-word — bail)
          if (!sameLine(w, pendingWord)) return []
          pendingWord = {
            width: pendingWord.width + w.width,
            top: Math.min(pendingWord.top, w.top),
            bottom: Math.max(pendingWord.bottom, w.bottom),
            left: Math.min(pendingWord.left, w.left),
            right: Math.max(pendingWord.right, w.right),
          }
        } else {
          pendingWord = w
        }
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
            const meta = tr.getMeta(justifyShrinkPluginKey) as Decoration[] | undefined
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
        view: (editorView) => new JustifyShrinkView(editorView, storage),
      }),
    ]
  },
})
