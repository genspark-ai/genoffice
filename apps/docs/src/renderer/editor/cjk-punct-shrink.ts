import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

/**
 * Word's CJK justified line breaking (settings characterSpacingControl =
 * compressPunctuation) pulls extra characters onto a line by compressing the
 * line's trailing-blank punctuation (、。 and closing brackets). Chromium
 * neither compresses nor pulls, so justified Japanese documents wrap earlier
 * than Word, fit fewer characters per line, and drift pages apart.
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
 *
 * Re-created as display-only inline decorations: the affected line's
 * compressible punctuation gets negative letter-spacing so Chromium's greedy
 * breaker takes the same characters; pagination measures the decorated DOM.
 */

/** trailing-blank punctuation Word compresses (JIS stops and closing brackets) */
const COMPRESSIBLE = new Set('、。，．」）』】｝〕》〉〗〙')
/** characters forbidden at a line start: pulling their predecessor drags them along */
const KINSOKU_CLOSE = new Set([...COMPRESSIBLE, '！', '？', '：', '；', 'ー', '々'])
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
  /** natural advance (layout px), independent of active decorations */
  width: number
  top: number
  bottom: number
  left: number
  right: number
}

interface MeasuredShrink {
  from: number
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
  /** compressible glyphs on the line */
  punctCount: number
  /** average natural advance of those glyphs (cap base; one body size per line in practice) */
  avgPunctW: number
  /** natural widths of the pull candidate chain (next line's leading chars); empty = no candidate */
  candWidths: number[]
  /** compressible glyphs inside the chain (they join the line's pool once pulled) */
  candPunctCount: number
  /** the chain drags a kinsoku-close character that cannot start a line */
  forced: boolean
}

/**
 * Word's pull/keep rule over measured lines. Returns the per-glyph compression
 * (px) to decorate onto the line's current compressible glyphs, or null.
 * Admissibility is judged over the post-pull pool (pulled punctuation joins
 * the line), but the emitted amount spreads the whole deficit over the glyphs
 * that exist now — once Chromium re-breaks, the keep path re-balances.
 */
export function decideCjkShrinks(lines: ShrinkLineChars[]): Array<number | null> {
  return lines.map((line) => {
    if (line.punctCount === 0 || line.avgPunctW <= 0) return null
    const overflow = line.natural - line.avail
    if (overflow > NOISE) {
      // the line already holds characters pulled by a previous round: keep the
      // compression that fits them (recomputed fresh from natural advances)
      const perGlyph = (overflow + SHRINK_EPS) / line.punctCount
      return perGlyph <= FORCED_CAP * line.avgPunctW ? perGlyph : null
    }
    if (line.candWidths.length === 0) return null
    const deficit = overflow + line.candWidths.reduce((s, w) => s + w, 0)
    if (deficit <= NOISE) return null
    const pool = line.punctCount + line.candPunctCount
    const capFrac = line.forced ? FORCED_CAP : VOLUNTARY_CAP
    if ((deficit + SHRINK_EPS) / pool > capFrac * line.avgPunctW) return null
    return (deficit + SHRINK_EPS) / line.punctCount
  })
}

let measureCtx: CanvasRenderingContext2D | null | undefined

function charAdvancePx(cs: CSSStyleDeclaration, ch: string): number {
  if (measureCtx === undefined) measureCtx = document.createElement('canvas').getContext('2d')
  if (!measureCtx) return 0
  measureCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`
  return measureCtx.measureText(ch).width
}

function sameLine(a: { top: number; bottom: number }, b: { top: number; bottom: number }): boolean {
  return a.top < b.bottom - 1 && a.bottom > b.top + 1
}

class CjkPunctShrinkView {
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
    private storage: CjkPunctShrinkStorage,
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
    if (!this.storage.enabled) {
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
        style: `letter-spacing:${Math.round((s.baseLs - s.perChar) * 100) / 100}px`,
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
    if (cs.textAlign !== 'justify') return []
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

    // per-character boxes: caret coords give the line geometry (they survive
    // decoration-split text nodes); natural advances come from canvas metrics
    // so active shrink decorations do not feed back into the next round
    const chars: CharBox[] = []
    const breakPositions: number[] = []
    const styleCache = new Map<Element, CSSStyleDeclaration>()
    let bail = false
    node.forEach((child, offset) => {
      if (bail || !child.isText || !child.text) {
        // atoms (images, tabs, fields) break the char model: skip the paragraph
        if (!child.isText && child.type.name !== 'hardBreak') bail = true
        else if (!child.isText) breakPositions.push(pos + 1 + offset)
        return
      }
      const base = pos + 1 + offset
      for (let i = 0; i < child.text.length; i++) {
        const from = base + i
        let a: { top: number; bottom: number; left: number; right: number }
        try {
          a = view.coordsAtPos(from, 1)
        } catch {
          bail = true
          return
        }
        const dp = view.domAtPos(from, 1)
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

    const punctBoxesPerLine: CharBox[][] = []
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
      const punctBoxes = line.filter((c) => COMPRESSIBLE.has(c.ch))
      punctBoxesPerLine.push(punctBoxes)
      // pull candidate: the next line's first char plus any kinsoku-close run
      // it would drag along (those cannot start the shortened next line);
      // a hard break between the lines forbids pulling entirely
      const next = lines[k + 1]
      const lastFrom = line[line.length - 1].from
      const brBetween = next && breakPositions.some((p) => p > lastFrom && p < next[0].from)
      const candWidths: number[] = []
      let candPunctCount = 0
      let forced = false
      if (k < lines.length - 1 && next && next.length > 0 && !brBetween) {
        candWidths.push(next[0].width)
        let j = 1
        while (j < next.length && KINSOKU_CLOSE.has(next[j].ch)) {
          candWidths.push(next[j].width)
          if (COMPRESSIBLE.has(next[j].ch)) candPunctCount++
          forced = true
          j++
        }
      }
      // ragged lines (the last one, or one ended by a hard break) report their
      // own compressed extent, which would ratchet the keep amount — measure
      // those against the paragraph content box instead
      const ragged = k === lines.length - 1 || brBetween
      return {
        natural,
        avail: ragged ? contentW - (k === 0 ? textIndent : 0) : (right - left) / zoom,
        punctCount: punctBoxes.length,
        avgPunctW:
          punctBoxes.length > 0
            ? punctBoxes.reduce((s, c) => s + c.width, 0) / punctBoxes.length
            : 0,
        candWidths: k === lines.length - 1 ? [] : candWidths,
        candPunctCount,
        forced,
      }
    })

    const out: MeasuredShrink[] = []
    const decisions = decideCjkShrinks(lineModels)
    for (let k = 0; k < decisions.length; k++) {
      const perGlyph = decisions[k]
      if (perGlyph === null) continue
      const perChar = Math.round(perGlyph * 100) / 100
      if (perChar <= 0) continue
      for (const c of punctBoxesPerLine[k]) out.push({ from: c.from, perChar, baseLs })
    }
    return out
  }
}

export const CjkPunctShrinkExtension = Extension.create({
  name: 'cjkPunctShrink',
  addStorage(): CjkPunctShrinkStorage {
    return { enabled: false }
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
