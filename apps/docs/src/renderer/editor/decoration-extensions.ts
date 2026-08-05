import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import {} from '@tiptap/pm/tables'
import { t } from '../i18n/locale'
import { type TabStop } from '@genoffice/docx-engine'

/**
 * Custom schema mirroring the docx-engine Block model 1:1.
 * Every top-level node carries `docxIndex` (patch anchor, null = new) and
 * `aiChanged` (diff highlighting for AI edits).
 */

import { SearchHighlight } from './extensions'
import { revisionDisplayState } from './marks'

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
/** w:defaultTabStop; 720 twips (0.5") is Word's default (settings.xml is not parsed yet) */
const DEFAULT_TAB_TWIPS = 720

interface MeasuredTab {
  pos: number
  /** CSS tab-size (px) for this tab's decoration span */
  cssSize: number
  leader?: string
  /** run carries the underline mark: browsers do not draw text-decoration
   *  across a tab advance, so the gap gets a border-bottom line instead */
  underlined?: boolean
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

class TabLayoutView {
  private lastSig = ''

  constructor(private view: EditorView) {
    this.measure()
    // font swaps change text metrics, which shifts every tab's start x
    document.fonts?.ready.then(() => this.measure()).catch(() => {})
  }

  update(view: EditorView, prevState: EditorState) {
    // selection-only transactions change neither the doc nor tab layout; the
    // plugin-state check keeps the measure→decorate→re-measure convergence alive
    if (
      view.state.doc === prevState.doc &&
      tabStopPluginKey.getState(view.state) === tabStopPluginKey.getState(prevState)
    )
      return
    this.measure()
  }

  destroy() {}

  private measure() {
    const { view } = this
    if (!view.dom.isConnected) return

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

    const paraRanges: Array<{ from: number; to: number }> = []
    const tabs: MeasuredTab[] = []
    for (const para of paras) {
      const measured = this.measureParagraph(para.node, para.pos)
      if (!measured) continue
      paraRanges.push({ from: para.pos, to: para.pos + para.node.nodeSize })
      tabs.push(...measured)
    }

    const sig = JSON.stringify([paraRanges, tabs])
    if (sig === this.lastSig) return
    this.lastSig = sig

    const old = tabStopPluginKey.getState(view.state)
    if (tabs.length === 0 && paraRanges.length === 0 && (!old || old === DecorationSet.empty))
      return

    const decos: Decoration[] = []
    for (const r of paraRanges)
      decos.push(Decoration.node(r.from, r.to, { class: 'has-tab-stops' }))
    for (const t of tabs) {
      const leader = t.leader && t.leader !== 'none' ? ` doc-tab-leader-${t.leader}` : ''
      const underline = t.underlined ? ' doc-tab-underline' : ''
      decos.push(
        Decoration.inline(t.pos, t.pos + 1, {
          class: `doc-tab${leader}${underline}`,
          style: `tab-size:${t.cssSize}px`,
        }),
      )
    }
    view.dispatch(view.state.tr.setMeta(tabStopPluginKey, decos).setMeta('addToHistory', false))
  }

  /** null = paragraph not measurable right now (hidden, not mounted...) */
  private measureParagraph(node: ProseMirrorNode, pos: number): MeasuredTab[] | null {
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
    const originX = rect.left - marginLeft * zoom
    // CSS tab-size origin: the paragraph content edge, offset from the column
    // edge by the full indent
    const contentEdge = marginLeft + paddingLeft

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
      .map((s) => ({ x: s.pos / TWIPS_PER_PX, val: s.val, leader: s.leader }))
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
    const out: MeasuredTab[] = []
    for (let i = 0; i < tabPositions.length; i++) {
      const tabPos = tabPositions[i]
      let coords: { left: number }
      try {
        coords = view.coordsAtPos(tabPos, 1)
      } catch {
        continue
      }
      const x = (coords.left - originX) / zoom

      const next = stopsPx.find((s) => s.x > x + 0.5)
      let target: number
      let val: TabStop['val'] = 'left'
      let leader: string | undefined
      if (next) {
        target = next.x
        val = next.val
        leader = next.leader
      } else {
        const grid = DEFAULT_TAB_TWIPS / TWIPS_PER_PX
        target = (Math.floor((x + 0.5) / grid) + 1) * grid
      }

      if (val === 'right' || val === 'decimal' || val === 'center') {
        // text between this tab and the next tab (or paragraph end) is aligned
        // at the stop; decimal is approximated as right (no '.'-splitting)
        const segStart = tabPos + 1
        const segEnd = i + 1 < tabPositions.length ? tabPositions[i + 1] : paraEnd
        let segWidth = 0
        if (segEnd > segStart) {
          try {
            const endCoords = view.coordsAtPos(segEnd, -1)
            segWidth = Math.max(0, (endCoords.left - coords.left) / zoom)
          } catch {
            /* keep 0 */
          }
        }
        target -= val === 'center' ? segWidth / 2 : segWidth
      }
      // convert the Word-space target to a CSS tab-size: the next multiple of
      // it past the tab's position must be the target itself, so it needs to
      // stay greater than the tab's content-edge-relative x
      const cssSize = Math.max(target - contentEdge, x - contentEdge + 0.5, 0.5)
      // 0.5px rounding damps measure→decorate→re-measure oscillation
      out.push({
        pos: tabPos,
        cssSize: Math.round(cssSize * 2) / 2,
        leader,
        underlined: tabUnderlined.get(tabPos),
      })
    }
    return out
  }
}

export const TabStopExtension = Extension.create({
  name: 'tabStops',
  addProseMirrorPlugins() {
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
        view: (view) => new TabLayoutView(view),
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
              decos.push(
                Decoration.node(offset, offset + node.nodeSize, {
                  'data-drop-cap': raw,
                  class: 'has-drop-cap',
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

// ---- SDT (structured document tag) rendering extension ----

const sdtPluginKey = new PluginKey<DecorationSet>('sdt')

export const SdtExtension = Extension.create({
  name: 'sdt',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: sdtPluginKey,
        props: {
          decorations(state) {
            const decos: Decoration[] = []
            state.doc.forEach((node, offset) => {
              const raw = node.attrs?.sdtShell as string | null
              if (!raw) return
              let alias = ''
              try {
                const parsed = JSON.parse(raw)
                alias = parsed?.alias || parsed?.tag || t('editorContentControl')
              } catch {
                /* ignore */
              }
              decos.push(
                Decoration.node(offset, offset + node.nodeSize, {
                  'data-sdt-alias': alias,
                  class: 'has-sdt',
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
