import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { isInTable } from '@tiptap/pm/tables'
import { t } from '../i18n/locale'
import { type TabStop } from '@genoffice/docx-engine'

/**
 * Custom schema mirroring the docx-engine Block model 1:1.
 * Every top-level node carries `docxIndex` (patch anchor, null = new) and
 * `aiChanged` (diff highlighting for AI edits).
 */

import { SearchHighlight } from './extensions'
import { revisionDisplayState } from './marks'
import { borderMergeFlags, type ParaBorderAttrs } from './para-border-merge'

export const searchPluginKey = new PluginKey<DecorationSet>('docSearch')

/** decorates search hits; FindPanel pushes ranges via setMeta(searchPluginKey) */
export const SearchHighlightExtension = Extension.create({
  name: 'docSearch',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: searchPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            const meta = tr.getMeta(searchPluginKey) as SearchHighlight | undefined
            if (meta) {
              return DecorationSet.create(
                tr.doc,
                meta.ranges.map((r, i) =>
                  Decoration.inline(r.from, r.to, {
                    class: i === meta.activeIndex ? 'search-hit search-hit-active' : 'search-hit',
                  }),
                ),
              )
            }
            return old.map(tr.mapping, tr.doc)
          },
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

export const pendingCommentPluginKey = new PluginKey<DecorationSet>('pendingComment')

/**
 * keeps the to-be-commented range visibly highlighted while the comment
 * composer holds focus (the native selection highlight disappears on blur)
 */
export const PendingCommentHighlightExtension = Extension.create({
  name: 'pendingComment',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: pendingCommentPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            const meta = tr.getMeta(pendingCommentPluginKey) as
              { from: number; to: number } | null | undefined
            if (meta === null) return DecorationSet.empty
            if (meta) {
              return DecorationSet.create(tr.doc, [
                Decoration.inline(meta.from, meta.to, { class: 'comment-pending' }),
              ])
            }
            return old.map(tr.mapping, tr.doc)
          },
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

export const resolvedCommentsPluginKey = new PluginKey<Set<string>>('resolvedComments')

/**
 * Word hides the in-text shading of resolved comment threads. App pushes the
 * resolved id set via meta; spans whose ids are all resolved get un-highlighted.
 */
export const ResolvedCommentsExtension = Extension.create({
  name: 'resolvedComments',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: resolvedCommentsPluginKey,
        state: {
          init: () => new Set<string>(),
          apply(tr, old) {
            const meta = tr.getMeta(resolvedCommentsPluginKey) as Set<string> | undefined
            return meta ?? old
          },
        },
        props: {
          decorations(state) {
            const done = resolvedCommentsPluginKey.getState(state)
            if (!done || done.size === 0) return null
            const markType = state.schema.marks.comment
            const decos: Decoration[] = []
            state.doc.descendants((node, pos) => {
              if (!node.isText) return
              const mark = node.marks.find((m) => m.type === markType)
              if (!mark) return
              const ids = String(mark.attrs.ids ?? '')
                .split(' ')
                .filter(Boolean)
              if (ids.length === 0 || !ids.every((id) => done.has(id))) return
              decos.push(
                Decoration.inline(pos, pos + node.nodeSize, { class: 'doc-comment-resolved' }),
              )
            })
            return DecorationSet.create(state.doc, decos)
          },
        },
      }),
    ]
  },
})

// ---- tab stop rendering extension ----

const tabStopPluginKey = new PluginKey<DecorationSet>('tabStops')

const TWIPS_PER_PX = 15
/** fallback when the document's settings.xml declares no w:defaultTabStop (Word: 0.5") */
const DEFAULT_TAB_TWIPS = 720

export interface TabStopStorage {
  /** settings.xml w:defaultTabStop (twips); 0 = zero-width default tabs (Word barely advances) */
  defaultTabStopTwips: number | null
}

declare module '@tiptap/core' {
  interface Storage {
    tabStops: TabStopStorage
  }
}

interface MeasuredTab {
  pos: number
  /** CSS tab-size (px) for this tab's decoration span */
  cssSize: number
  leader?: string
  /** run carries the underline mark: browsers do not draw text-decoration
   *  across a tab advance, so the gap gets a border-bottom line instead */
  underlined?: boolean
  /** no room left on the line: render at a hair's width (font-size 0 lifts
   *  Chromium's one-space minimum tab advance) */
  collapsed?: boolean
}

/**
 * Measures every tab character (in any textblock, including table cells) and
 * turns it into a tab-stop jump: the paragraph gets `tab-size: 0` (class
 * has-tab-stops) and each tab char gets an inline decoration with a per-span
 * `tab-size: <length>`. Chromium ignores letter-/word-spacing on tab chars,
 * but honors a length tab-size: the tab advances to the next multiple of it
 * counted from the paragraph's content edge — with size = stop position and
 * the tab left of the stop, it lands exactly on the stop. The tab stays a
 * plain text character, so an ancestor <u>'s underline is drawn across the
 * gap — signature blanks like "By: ____" depend on this.
 *
 * Word measures stops from the text column edge (page margin / cell content
 * edge) while CSS measures from the paragraph content edge, so the paragraph
 * indent (margin/padding-left) is subtracted when converting.
 *
 * Measure → decorate → re-measure converges (a tab's target is absolute, not
 * cumulative) and is guarded by a signature check to avoid dispatch loops.
 */
/** nodes are immutable, so the has-tab verdict per textblock never goes stale */
const paraHasTabCache = new WeakMap<ProseMirrorNode, boolean>()

let spaceMeasureCtx: CanvasRenderingContext2D | null | undefined
/** width of a space glyph in the paragraph's computed font (px, layout space) */
function spaceWidthPx(cs: CSSStyleDeclaration): number {
  if (spaceMeasureCtx === undefined)
    spaceMeasureCtx = document.createElement('canvas').getContext('2d')
  if (!spaceMeasureCtx) return 4
  spaceMeasureCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`
  const w = spaceMeasureCtx.measureText(' ').width + (parseFloat(cs.letterSpacing) || 0)
  // 20% headroom: run-level fonts/sizes inside the paragraph may shape wider
  return Math.max(1, w * 1.2)
}

const MEASURE_RETRY_MAX = 10
/**
 * Distinct decoration signatures tolerated between external invalidations
 * (edit / resize / font load). A convergent measure→decorate→re-measure
 * settles within a couple of dispatches; seeing a signature twice (or an
 * ever-drifting stream of new ones) means decorating keeps changing the very
 * layout being measured. Each dispatch re-enters update() synchronously, so
 * an unguarded oscillation recurses until the stack overflows (observed on
 * tab-heavy government forms: the renderer hangs at 100% CPU).
 */
const MEASURE_SIGS_MAX = 8

class TabLayoutView {
  private lastSig = ''
  private seenSigs = new Set<string>()
  private frozen = false
  private retryRaf = 0
  private retries = 0
  private resizeObserver?: ResizeObserver
  private lastDomWidth = -1
  // font swaps change text metrics, which shifts every tab's start x
  private onFontsLoaded = () => {
    this.invalidate()
    this.measure()
  }

  constructor(
    private view: EditorView,
    private storage: TabStopStorage,
  ) {
    this.measure()
    document.fonts?.addEventListener('loadingdone', this.onFontsLoaded)
    if (typeof ResizeObserver !== 'undefined') {
      // width-only trigger: height changes on every keystroke
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

  /** external layout input changed: start a fresh convergence run */
  private invalidate() {
    this.seenSigs.clear()
    this.frozen = false
  }

  update(view: EditorView, prevState: EditorState) {
    if (view.state.doc !== prevState.doc) {
      this.invalidate()
    } else if (
      // selection-only transactions change neither the doc nor tab layout; the
      // plugin-state check keeps the measure→decorate→re-measure convergence alive
      tabStopPluginKey.getState(view.state) === tabStopPluginKey.getState(prevState)
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

  /** initial-load dispatches measure before the DOM is measurable; retry on rAF (bounded) */
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
    if (!view.dom.isConnected) {
      this.scheduleRetry()
      return
    }

    const paras: Array<{ node: ProseMirrorNode; pos: number }> = []
    view.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true
      let hasTab = paraHasTabCache.get(node)
      if (hasTab === undefined) {
        hasTab = node.textContent.includes('\t')
        paraHasTabCache.set(node, hasTab)
      }
      if (hasTab) paras.push({ node, pos })
      return false
    })

    const paraRanges: Array<{ from: number; to: number; flattenJustify?: boolean }> = []
    const tabs: MeasuredTab[] = []
    let measurable = false
    for (const para of paras) {
      const measured = this.measureParagraph(para.node, para.pos)
      if (!measured) continue
      measurable = true
      paraRanges.push({
        from: para.pos,
        to: para.pos + para.node.nodeSize,
        ...(measured.flattenJustify ? { flattenJustify: true } : {}),
      })
      tabs.push(...measured.tabs)
    }

    if (paras.length > 0 && !measurable) {
      this.scheduleRetry()
      return
    }
    this.retries = 0

    const sig = JSON.stringify([paraRanges, tabs])
    if (sig === this.lastSig) return
    if (this.frozen) return
    if (this.seenSigs.has(sig) || this.seenSigs.size >= MEASURE_SIGS_MAX) {
      // oscillation: keep the current decorations until an edit/resize/font
      // load invalidates, instead of dispatching (and recursing) forever
      this.frozen = true
      console.warn('[docs] tab-stop layout did not converge; keeping current tab decorations')
      return
    }
    this.seenSigs.add(sig)
    this.lastSig = sig

    const old = tabStopPluginKey.getState(view.state)
    if (tabs.length === 0 && paraRanges.length === 0 && (!old || old === DecorationSet.empty))
      return

    const decos: Decoration[] = []
    for (const r of paraRanges)
      decos.push(
        Decoration.node(r.from, r.to, {
          class: r.flattenJustify ? 'has-tab-stops tab-stops-no-justify' : 'has-tab-stops',
        }),
      )
    for (const t of tabs) {
      const leader = t.leader && t.leader !== 'none' ? ` doc-tab-leader-${t.leader}` : ''
      const underline = t.underlined ? ' doc-tab-underline' : ''
      const collapse = t.collapsed ? ' doc-tab-collapse' : ''
      decos.push(
        Decoration.inline(t.pos, t.pos + 1, {
          class: `doc-tab${leader}${underline}${collapse}`,
          style: `tab-size:${t.cssSize}px`,
        }),
      )
    }
    view.dispatch(view.state.tr.setMeta(tabStopPluginKey, decos).setMeta('addToHistory', false))
  }

  /** null = paragraph not measurable right now (hidden, not mounted...) */
  private measureParagraph(
    node: ProseMirrorNode,
    pos: number,
  ): { tabs: MeasuredTab[]; flattenJustify: boolean } | null {
    const { view } = this
    const el = view.nodeDOM(pos)
    if (!(el instanceof HTMLElement) || el.offsetWidth === 0) return null
    const rect = el.getBoundingClientRect()
    if (rect.width === 0) return null
    // rects are in screen px (affected by the page zoom transform); widths we
    // emit are layout px, so divide all measured distances by the scale factor
    const zoom = rect.width / el.offsetWidth
    const cs = window.getComputedStyle(el)
    if (cs.direction === 'rtl') return null
    // Word's tab-stop origin: the text column edge (excludes paragraph indent,
    // whether it is margin-left on docParagraph or padding-left on list items)
    const marginLeft = parseFloat(cs.marginLeft) || 0
    const paddingLeft = parseFloat(cs.paddingLeft) || 0
    const paddingRight = parseFloat(cs.paddingRight) || 0
    const originX = rect.left - marginLeft * zoom
    // paragraph right content edge in tab-origin space
    const paraW = marginLeft + el.clientWidth - paddingRight
    // CSS tab-size origin: the paragraph content edge, offset from the column
    // edge by the full indent. Known limit: a hanging list marker that escapes
    // its hang box (--li-tab) shifts Chromium's real anchor by amounts that
    // depend on the rendered marker width; such lines can settle one grid
    // cell off (rare TOC-style list entries with wide roman markers).
    const contentEdge = marginLeft + paddingLeft

    // Chromium expands tabs in left-aligned space and only then shifts the
    // line by text-align (same order as Word). coordsAtPos returns post-shift
    // positions, so measuring a centered/right line inflates every tab's x by
    // the shift — targets grow, the line re-shifts, and the loop never
    // converges (TOC page numbers scatter). Subtract the line's shift: visual
    // line start minus the layout-space start (content edge, plus the first
    // line's text-indent).
    // Intended alignment: the tab-stops-no-justify decoration forces computed
    // textAlign to left, so reading it back would clear the flatten flag and
    // oscillate. The paragraph attr wins; an already-flattened element with no
    // attr keeps counting as justified (its pre-flatten computed value).
    const attrAlign = node.attrs?.align as string | null
    const align =
      attrAlign ?? (el.classList.contains('tab-stops-no-justify') ? 'justify' : cs.textAlign)
    let lineRects: DOMRect[] | null = null
    let firstLineTop = Infinity
    if (align === 'center' || align === 'right' || align === 'end') {
      const range = document.createRange()
      range.selectNodeContents(el)
      lineRects = Array.from(range.getClientRects())
      for (const r of lineRects) if (r.height > 0) firstLineTop = Math.min(firstLineTop, r.top)
    }
    const contentLeft = rect.left + ((parseFloat(cs.borderLeftWidth) || 0) + paddingLeft) * zoom
    const textIndent = parseFloat(cs.textIndent) || 0
    const alignShiftAt = (coords: { top: number; bottom: number }): number => {
      if (!lineRects) return 0
      const mid = (coords.top + coords.bottom) / 2
      let lineLeft = Infinity
      let lineTop = Infinity
      for (const r of lineRects) {
        if (r.top > mid || r.bottom < mid) continue
        lineLeft = Math.min(lineLeft, r.left)
        lineTop = Math.min(lineTop, r.top)
      }
      if (!Number.isFinite(lineLeft)) return 0
      const expected = contentLeft + (lineTop <= firstLineTop + 1 ? textIndent * zoom : 0)
      return Math.max(0, lineLeft - expected)
    }

    // Chromium renders a tab whose distance to the next multiple is smaller
    // than a space glyph by skipping to the following multiple (WebKit rule).
    // Every emitted target must clear that minimum or the rendered advance
    // doubles, the line overflows, and re-measures diverge.
    const minAdv = spaceWidthPx(cs) + 1
    let stops: TabStop[] = []
    const raw = node.attrs?.tabStops as string | null
    if (raw) {
      try {
        const parsed = JSON.parse(raw)
        if (Array.isArray(parsed)) stops = parsed
      } catch {
        /* ignore malformed */
      }
    }
    const stopsPx = stops
      .filter((s) => s.val !== 'clear' && s.val !== 'bar' && Number.isFinite(s.pos))
      // rel stops mirror w:ptab: pos is a percent of the column width
      .map((s) => ({
        x: s.rel === 'margin' ? (s.pos / 100) * paraW : s.pos / TWIPS_PER_PX,
        val: s.val,
        leader: s.leader,
      }))
      .sort((a, b) => a.x - b.x)

    // doc positions of every tab char in this paragraph
    const tabPositions: number[] = []
    const tabUnderlined = new Map<number, boolean>()
    node.forEach((child, offset) => {
      if (!child.isText || !child.text) return
      const underlined = child.marks.some((m) => m.type.name === 'underline')
      for (let i = 0; i < child.text.length; i++) {
        if (child.text[i] !== '\t') continue
        const p = pos + 1 + offset + i
        tabPositions.push(p)
        if (underlined) tabUnderlined.set(p, true)
      }
    })

    const paraEnd = pos + node.nodeSize - 1
    // Width of the text between each tab and the next tab (or paragraph end),
    // measured from the tab's END so it excludes the tab's current advance —
    // otherwise the computed target depends on the layout being measured and
    // re-measure never reaches a fixed point.
    const segWidths = tabPositions.map((tabPos, i) => {
      const segStart = tabPos + 1
      const segEnd = i + 1 < tabPositions.length ? tabPositions[i + 1] : paraEnd
      if (segEnd <= segStart) return 0
      try {
        const startCoords = view.coordsAtPos(segStart, 1)
        const endCoords = view.coordsAtPos(segEnd, -1)
        return Math.max(0, (endCoords.left - startCoords.left) / zoom)
      } catch {
        return 0
      }
    })
    // total width of this tab's segment plus everything after it
    const restWidths = [...segWidths]
    for (let i = restWidths.length - 2; i >= 0; i--) restWidths[i] += restWidths[i + 1]
    const out: MeasuredTab[] = []
    // analytic chain: x of a tab following another tab on the same visual line
    // is the previous target + segment width. DOM positions of later tabs
    // depend on the very decorations being measured (a tab measured on a
    // wrapped line yields a line-relative x whose tab-size then overshoots to
    // a higher multiple once the line unwraps — a wrap/unwrap 2-cycle the
    // signature guard then freezes mid-flight). The chain removes that
    // dependence; a real line break between tabs resets to the measured x.
    let prevLine: { top: number; bottom: number } | null = null
    let prevEnd = 0
    for (let i = 0; i < tabPositions.length; i++) {
      const tabPos = tabPositions[i]
      let coords: { left: number; top: number; bottom: number }
      try {
        coords = view.coordsAtPos(tabPos, 1)
      } catch {
        continue
      }
      const measuredX = (coords.left - alignShiftAt(coords) - originX) / zoom
      const sameLine =
        prevLine != null && coords.top < prevLine.bottom && coords.bottom > prevLine.top
      const x = sameLine ? prevEnd : measuredX
      prevLine = { top: coords.top, bottom: coords.bottom }

      const next = stopsPx.find((s) => s.x > x + minAdv)
      let target: number
      let val: TabStop['val'] = 'left'
      let leader: string | undefined
      if (next) {
        target = next.x
        val = next.val
        leader = next.leader
      } else {
        const gridTwips = this.storage.defaultTabStopTwips ?? DEFAULT_TAB_TWIPS
        if (gridTwips > 0) {
          const grid = gridTwips / TWIPS_PER_PX
          target = (Math.floor((x + minAdv) / grid) + 1) * grid
        } else {
          // defaultTabStop 0: Word advances the caret imperceptibly (tdf#168607)
          target = x + minAdv
        }
      }

      const segWidth = segWidths[i]
      // right/decimal/center align the segment at the stop; decimal is
      // approximated as right (no '.'-splitting)
      if (val === 'right' || val === 'decimal') target -= segWidth
      else if (val === 'center') target -= segWidth / 2
      // A right/center/decimal stop that would push the segment past the
      // paragraph width pins it flush to the right edge (Word never wraps such
      // TOC-style lines; Chromium would). In-column left stops advance to the
      // stop and let a segment too wide for the trailing space wrap naturally —
      // but a left stop past the right edge pins too (TOC page numbers), and so
      // does a short segment that still fits flush right of the tab: wrapping
      // it makes non-last lines justify-stretch, whose inflated measurements
      // feed back into ever-larger targets (three-column signature rows).
      // Pinning reserves room for everything after this tab (restWidths), so a
      // run of trailing tabs packs against the edge instead of spilling over.
      // The 1px slack keeps the 0.5px cssSize round-up from re-triggering wrap.
      if (
        target + segWidth > paraW - 1 &&
        (val !== 'left' || target > paraW - 1 || paraW - segWidth - 1 > x)
      )
        target = paraW - 1 - restWidths[i]
      let collapsed = false
      if (target < x + minAdv) {
        // no room before the edge: a normal tab would still advance a space
        // width (Chromium's minimum), overflowing the line — collapse it to a
        // hair's width instead (font-size 0 lifts the minimum)
        collapsed = true
        target = x + 0.6
      }
      // convert the Word-space target to a CSS tab-size: the next multiple of
      // it past the tab's position must be the target itself, so it needs to
      // stay greater than the tab's content-edge-relative x (by the minimum
      // rendered advance, or Chromium skips to the following multiple)
      const cssSize = Math.max(target - contentEdge, x - contentEdge + 0.6, 0.5)
      prevEnd = target + segWidth
      // 0.5px rounding damps measure→decorate→re-measure oscillation
      out.push({
        pos: tabPos,
        cssSize: Math.round(cssSize * 2) / 2,
        leader,
        underlined: tabUnderlined.get(tabPos),
        collapsed,
      })
    }
    // Word keeps tab segments at their stops on justified lines; Chromium
    // justify-stretches them (and the stretched positions would feed back into
    // the measurements), so tabbed justified paragraphs lay out left-aligned
    return { tabs: out, flattenJustify: align === 'justify' || align === 'distribute' }
  }
}

export const TabStopExtension = Extension.create({
  name: 'tabStops',
  addStorage() {
    return { defaultTabStopTwips: null } as TabStopStorage
  },
  addKeyboardShortcuts() {
    // Word: Tab in a body paragraph inserts a tab character (default 0.5"
    // stops, or the paragraph's custom w:tabs). Lists indent and tables move
    // to the next cell — those handlers live on DocListItem / NativeTableSupport
    // and run after this one returns false. An unhandled Tab would leave the
    // editor and cycle the ribbon buttons (github.com/genspark-ai/genoffice/issues/101).
    const insertTab = () => {
      if (!this.editor.isEditable) return false
      if (this.editor.isActive('docListItem')) return false
      if (isInTable(this.editor.state)) return false
      const { state, view } = this.editor
      if (!state.selection.$from.parent.isTextblock) return true
      view.dispatch(state.tr.insertText('\t').scrollIntoView())
      return true
    }
    const swallowShiftTab = () => {
      if (!this.editor.isEditable) return false
      if (this.editor.isActive('docListItem')) return false
      if (isInTable(this.editor.state)) return false
      return true
    }
    return {
      Tab: insertTab,
      'Shift-Tab': swallowShiftTab,
    }
  },
  addProseMirrorPlugins() {
    const storage = this.storage as TabStopStorage
    return [
      new Plugin({
        key: tabStopPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            const meta = tr.getMeta(tabStopPluginKey) as Decoration[] | undefined
            if (meta) return DecorationSet.create(tr.doc, meta)
            return old.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
        view: (view) => new TabLayoutView(view, storage),
      }),
    ]
  },
})

// ---- drop cap rendering extension ----

const dropCapPluginKey = new PluginKey<DecorationSet>('dropCap')

export const DropCapExtension = Extension.create({
  name: 'dropCap',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: dropCapPluginKey,
        props: {
          decorations(state) {
            const decos: Decoration[] = []
            state.doc.forEach((node, offset) => {
              const raw = node.attrs?.dropCap as string | null
              if (!raw) return
              // Word-imported caps store the enlarged size on the run itself
              // (a 52pt letter for lines=3): the 3em first-letter scale is only
              // for editor-created caps whose run keeps the body size —
              // stacking both painted an 8-line giant (real_run2/44). "Own
              // size" means the first run is already enlarged: clearly larger
              // than the paragraph's other text (or 24pt+ when the cap letter
              // is the whole paragraph) — an ordinary explicit body size on
              // the letter still gets the 3em scale (Bugbot).
              const szOf = (n: ProseMirrorNode | null): number | undefined =>
                n?.isText
                  ? ((n.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.sizeHalfPoints as
                      number | undefined) ?? undefined)
                  : undefined
              const firstSz = szOf(node.firstChild)
              let restSz: number | undefined
              node.forEach((child, _off, i) => {
                if (i === 0 || restSz !== undefined || !child.isText) return
                restSz = szOf(child) ?? 0
              })
              const ownSize =
                firstSz !== undefined &&
                (restSz !== undefined ? firstSz >= Math.max(restSz, 20) * 1.8 : firstSz >= 48)
              decos.push(
                Decoration.node(offset, offset + node.nodeSize, {
                  'data-drop-cap': raw,
                  class: ownSize ? 'has-drop-cap drop-cap-own-size' : 'has-drop-cap',
                }),
              )
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

// ---- whitespace-only run line height (Word ignores their font size) ----

const wsRunPluginKey = new PluginKey<DecorationSet>('wsRunLineHeight')

/**
 * Word ignores the font size of whitespace-only runs (spaces/tabs) for line
 * height (tdf#137335): the oversized whitespace keeps its advance but the line
 * box stays governed by the text portions (strut on blank lines). Recomputed
 * from the live text so editing into such a run drops the suppression, and a
 * character style's font size is covered the same as a direct w:sz.
 */
export const WsRunLineHeightExtension = Extension.create({
  name: 'wsRunLineHeight',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: wsRunPluginKey,
        props: {
          decorations(state) {
            const decos: Decoration[] = []
            state.doc.descendants((node, pos) => {
              if (!node.isText || !node.text) return
              if (!/^[ \t]+$/.test(node.text)) return
              // only runs with a font-size source can inflate the line
              const styled = node.marks.some(
                (m) =>
                  m.type.name === 'docTextStyle' && (m.attrs.sizeHalfPoints || m.attrs.styleId),
              )
              if (!styled) return
              decos.push(Decoration.inline(pos, pos + node.nodeSize, { class: 'doc-ws-run' }))
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

// ---- hinted East Asian curly quotes (w:hint="eastAsia") ----

const eaHintQuotesPluginKey = new PluginKey<DecorationSet>('eaHintQuotes')

/** the run's own w:rFonts carries w:hint="eastAsia" (revision snapshots inside w:rPrChange don't count) */
const EA_HINT_RE = /<w:rFonts\b[^>]*\bw:hint="eastAsia"/

const EA_QUOTES_RE = /[‘’“”]+/g

/** quote ranges of `text` that Word renders with the East Asian font; exported for tests */
export function eaHintQuoteRanges(
  rawRPr: string | null | undefined,
  text: string,
): Array<{ from: number; to: number }> {
  if (!rawRPr || !text || !/[‘’“”]/.test(text)) return []
  if (!EA_HINT_RE.test(rawRPr.replace(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/, ''))) return []
  const ranges: Array<{ from: number; to: number }> = []
  for (const m of text.matchAll(EA_QUOTES_RE)) {
    ranges.push({ from: m.index, to: m.index + m[0].length })
  }
  return ranges
}

/**
 * Curly quotes U+2018/2019/201C/201D are dual-width codepoints: Word renders
 * them with the run's East Asian font (fullwidth advance) iff the run's
 * w:rFonts carries w:hint="eastAsia" — surrounding text and w:lang play no
 * part (Word for Mac probe 2026-09-01). Chromium resolves them through the
 * Latin head of the font chain (halfwidth), so hinted quotes get a
 * display-only decoration whose bundled CJK face restores the 1em advance.
 */
export const EaHintQuotesExtension = Extension.create({
  name: 'eaHintQuotes',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: eaHintQuotesPluginKey,
        props: {
          decorations(state) {
            const decos: Decoration[] = []
            state.doc.descendants((node, pos) => {
              if (!node.isText || !node.text) return
              const style = node.marks.find((m) => m.type.name === 'docTextStyle')
              const raw = style?.attrs.rawRPr as string | null | undefined
              for (const r of eaHintQuoteRanges(raw, node.text)) {
                decos.push(Decoration.inline(pos + r.from, pos + r.to, { class: 'doc-ea-quotes' }))
              }
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

// ---- adjacent-paragraph border merging (Word border groups) ----

const paraBorderMergePluginKey = new PluginKey<DecorationSet>('paraBorderMerge')

/**
 * Word border groups (ECMA-376 §17.3.1.24): adjacent top-level paragraphs with
 * identical borders/shading draw top/bottom lines only at the group edges.
 * Display-only classes; the borders attrs still serialize unchanged on save.
 */
export const ParaBorderMergeExtension = Extension.create({
  name: 'paraBorderMerge',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: paraBorderMergePluginKey,
        props: {
          decorations(state) {
            const items: ParaBorderAttrs[] = []
            const spans: Array<{ from: number; to: number }> = []
            state.doc.forEach((node, offset) => {
              // non-paragraph blocks (tables...) enter as border-less entries: they break adjacency
              items.push(node.isTextblock ? (node.attrs as ParaBorderAttrs) : {})
              spans.push({ from: offset, to: offset + node.nodeSize })
            })
            const decos: Decoration[] = []
            borderMergeFlags(items).forEach((f, i) => {
              if (!f.suppressTop && !f.suppressBottom) return
              const cls = [
                f.suppressTop ? 'pbdr-suppress-top' : '',
                f.suppressBottom ? 'pbdr-suppress-bottom' : '',
              ]
                .filter(Boolean)
                .join(' ')
              decos.push(Decoration.node(spans[i].from, spans[i].to, { class: cls }))
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

// ---- SDT (structured document tag) rendering extension ----

const sdtPluginKey = new PluginKey<boolean>('sdt')

export const SdtExtension = Extension.create({
  name: 'sdt',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: sdtPluginKey,
        // plugin state = "editor focused"; decorations() can't reach the view,
        // and focus changes don't dispatch transactions on their own
        state: {
          init: () => false,
          apply: (tr, focused) => (tr.getMeta(sdtPluginKey) as boolean | undefined) ?? focused,
        },
        // after a ribbon click the editor's own blur already fired (with a
        // relatedTarget), so a later alt-tab only blurs the ribbon control;
        // the window blur is the reliable "chrome must disappear" signal
        view: (editorView) => {
          const onWindowBlur = () => {
            if (sdtPluginKey.getState(editorView.state)) {
              editorView.dispatch(editorView.state.tr.setMeta(sdtPluginKey, false))
            }
          }
          window.addEventListener('blur', onWindowBlur)
          return { destroy: () => window.removeEventListener('blur', onWindowBlur) }
        },
        props: {
          handleDOMEvents: {
            focus: (view) => {
              view.dispatch(view.state.tr.setMeta(sdtPluginKey, true))
              return false
            },
            blur: (view, event) => {
              // Word keeps control chrome while formatting from the toolbar:
              // only clear when focus leaves the document (relatedTarget null —
              // window switch / screenshot), not on ribbon clicks
              if ((event as FocusEvent).relatedTarget) return false
              view.dispatch(view.state.tr.setMeta(sdtPluginKey, false))
              return false
            },
          },
          decorations(state) {
            // Word shows content-control chrome only while the editor has focus
            // and the cursor is inside the control; a static document view
            // (export, screenshot, unfocused window) renders the content alone.
            // "Inside" means fully contained — a doc-wide selection (initial
            // AllSelection, ctrl-A) must not light up every control.
            if (!sdtPluginKey.getState(state)) return DecorationSet.empty
            const { from, to } = state.selection
            // A multi-paragraph control (blocks sharing sdtShell.group) is one
            // control: the caret anywhere inside lights up every member block.
            type Entry = { offset: number; end: number; alias: string; group?: number }
            const entries: Entry[] = []
            state.doc.forEach((node, offset) => {
              const raw = node.attrs?.sdtShell as string | null
              if (!raw) return
              let alias = ''
              let group: number | undefined
              try {
                const parsed = JSON.parse(raw)
                alias = parsed?.alias || parsed?.tag || t('editorContentControl')
                if (typeof parsed?.group === 'number') group = parsed.group
              } catch {
                /* ignore */
              }
              entries.push({ offset, end: offset + node.nodeSize, alias, group })
            })
            const rangeOf = (e: Entry): [number, number] => {
              if (e.group === undefined) return [e.offset, e.end]
              const members = entries.filter((m) => m.group === e.group)
              return [members[0].offset, members[members.length - 1].end]
            }
            const hit = entries.find((e) => {
              const [start, end] = rangeOf(e)
              return from >= start && to <= end
            })
            if (!hit) return DecorationSet.empty
            const members =
              hit.group === undefined ? [hit] : entries.filter((m) => m.group === hit.group)
            const decos = members.map((m, i) =>
              Decoration.node(m.offset, m.end, {
                // one control-wide box: the alias tab renders once, on the first member
                ...(i === 0 ? { 'data-sdt-alias': m.alias } : {}),
                class: i === 0 ? 'has-sdt' : 'has-sdt has-sdt-follow',
              }),
            )
            return DecorationSet.create(state.doc, decos)
          },
        },
      }),
    ]
  },
})

// ---- move revision rendering extension ----

const moveRevisionPluginKey = new PluginKey<DecorationSet>('moveRevision')

export const MoveRevisionExtension = Extension.create({
  name: 'moveRevision',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: moveRevisionPluginKey,
        props: {
          decorations(state) {
            const decos: Decoration[] = []
            state.doc.forEach((node, offset) => {
              const rev = node.attrs?.moveRevision as string | null
              if (!rev) return
              decos.push(
                Decoration.node(offset, offset + node.nodeSize, {
                  'data-move-revision': rev,
                  class: rev === 'from' ? 'has-move-from' : 'has-move-to',
                }),
              )
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

// ---- deleted paragraph mark (w:pPr/w:rPr/w:del) collapse extension ----

const paraMarkDelPluginKey = new PluginKey<DecorationSet>('paraMarkDel')

/**
 * A paragraph whose mark is a tracked deletion and whose inline content is all
 * deleted leaves no trace in Word's All Markup view (the text lives in a margin
 * balloon); collapse it so pagination matches. Partially deleted paragraphs
 * (Word joins them into the next one) keep their height — approximate, but the
 * remaining text still occupies a line either way.
 */
export const ParaMarkDelExtension = Extension.create({
  name: 'paraMarkDel',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: paraMarkDelPluginKey,
        props: {
          decorations(state) {
            const decos: Decoration[] = []
            state.doc.forEach((node, offset) => {
              if (!node.attrs?.paraMarkDel) return
              let visible = false
              node.forEach((child) => {
                if (!child.marks.some((m) => m.type.name === 'del')) visible = true
              })
              decos.push(
                Decoration.node(offset, offset + node.nodeSize, {
                  class: visible ? 'doc-para-mark-del' : 'doc-para-mark-del doc-para-del-collapse',
                }),
              )
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})

// ---- pPrChange (tracked paragraph format) rendering extension ----

const pPrChangePluginKey = new PluginKey<DecorationSet>('pPrChange')

export const PPrChangeExtension = Extension.create({
  name: 'pPrChange',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: pPrChangePluginKey,
        props: {
          decorations(state) {
            const decos: Decoration[] = []
            state.doc.forEach((node, offset) => {
              const raw = node.attrs?.pPrChange as string | null
              if (!raw) return
              let author = ''
              let old: Record<string, unknown> = {}
              try {
                const parsed = JSON.parse(raw)
                author = parsed?.author || ''
                old = (parsed?.old ?? {}) as Record<string, unknown>
              } catch {
                /* ignore */
              }
              // original view: override with the pre-revision paragraph format (mirroring the restore subset of revisions.ts reject)
              let style = ''
              if (revisionDisplayState.mode === 'original') {
                const styles = [
                  `text-align:${(old.align as string) ?? 'start'}`,
                  `margin-left:${old.indentLeft ? Number(old.indentLeft) / 20 : 0}pt`,
                  `margin-right:${old.indentRight ? Number(old.indentRight) / 20 : 0}pt`,
                  `text-indent:${old.indentFirstLine ? Number(old.indentFirstLine) / 20 : 0}pt`,
                  `margin-top:${old.spaceBefore ? Number(old.spaceBefore) / 20 : 0}pt`,
                  `margin-bottom:${old.spaceAfter ? Number(old.spaceAfter) / 20 : 0}pt`,
                  `background:${old.shadingFill ? `#${old.shadingFill}` : 'transparent'}`,
                ]
                if (old.lineRule === 'exact' || old.lineRule === 'atLeast') {
                  styles.push(`line-height:${Number(old.lineRawTwips || 240) / 20}pt`)
                } else if (old.lineSpacing) {
                  styles.push(`line-height:${Number(old.lineSpacing) * 1.2}`)
                } else {
                  styles.push('line-height:inherit')
                }
                style = styles.join(';')
              }
              decos.push(
                Decoration.node(offset, offset + node.nodeSize, {
                  'data-ppr-change-author': author,
                  'data-ppr-change-label': t('editorFormatRevision'),
                  class: 'has-ppr-change',
                  ...(style ? { style } : {}),
                }),
              )
            })
            return decos.length > 0 ? DecorationSet.create(state.doc, decos) : DecorationSet.empty
          },
        },
      }),
    ]
  },
})
