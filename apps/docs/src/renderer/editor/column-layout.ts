import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'

const key = new PluginKey<DecorationSet>('columnLayout')

/**
 * Mixed-column canvas layout (documents whose sections disagree on column
 * count/width, so the uniform whole-page CSS multicol path can't apply).
 *
 * The engine keeps its 1-D single-flow virtual coordinates; this extension only
 * *paints* the flow into columns: each block in a multi-column region gets its
 * section's column width plus a constant per-column translate, and blocks of
 * later regions on the same page translate up over the space the columns
 * vacated (the page gap below pulls the next page up by the same amount).
 * Decorations are visual-only; measurement runs in the `.measuring-columns`
 * state, whose CSS neutralizes the transforms (widths stay — the engine's line
 * boxes must reflect column wrapping).
 *
 * v1 granularity: a block is placed whole into the column containing its top;
 * a block the engine split across columns paints overflowing its column bottom.
 */
export const ColumnLayoutExtension = Extension.create({
  name: 'columnLayout',

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(key) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})

export interface ColumnBlockSpec {
  /** top-level block element (same anchoring as page gaps) */
  el: HTMLElement
  /** column width (px); absent for single-column regions (natural width) */
  widthPx?: number
  /** owning section's content width (--doc-content-w): tables resolve their spill/centering caps against it */
  contentWPx?: number
  /** owning section's side margins (--doc-margin-left/right overrides) */
  marginLeftPx?: number
  marginRightPx?: number
  /** owning section's typed docGrid pitch (pt) in mixed-grid docs; 0 = untyped
   *  section (opts out like snapToGrid=0 via doc-grid-nosnap) */
  gridPitchPt?: number
  /** owning section's character-grid letter-spacing delta (pt) in docs whose
   *  sections disagree on docGrid charSpace (doc-charspace-block) */
  charSpacePt?: number
  /** translate; dy < 0 pulls content up over vacated column space */
  dx: number
  dy: number
}

const PATCH_PROPS = [
  '--col-w',
  '--col-dx',
  '--col-dy',
  '--doc-content-w',
  '--doc-margin-left',
  '--doc-margin-right',
  '--doc-grid-pitch',
  '--doc-char-space',
]

const PATCH_CLASSES = ['doc-col-block', 'doc-grid-block', 'doc-grid-nosnap', 'doc-charspace-block']

/** pitch/charSpace-only specs get an inert marker class: .doc-col-block's
 *  width/transform rules must not touch blocks that only carry their section's
 *  grid pitch or character spacing. Pitch-0 opt-outs use their own class (same
 *  CSS as doc-nosnap): reusing doc-nosnap would trip sectionGridPitchSpecs'
 *  own-opt-out skip on the next pass and let patch cleanup strip a node's own
 *  snapToGrid=0 class. */
function specClass(spec: ColumnBlockSpec): string {
  const varsOnly =
    (spec.gridPitchPt !== undefined || spec.charSpacePt !== undefined) &&
    spec.widthPx === undefined &&
    spec.contentWPx === undefined &&
    spec.dx === 0 &&
    spec.dy === 0
  let cls = varsOnly ? 'doc-grid-block' : 'doc-col-block'
  if (spec.gridPitchPt === 0) cls += ' doc-grid-nosnap'
  if (spec.charSpacePt !== undefined) cls += ' doc-charspace-block'
  return cls
}

/** Rebuild the column-layout decorations (an empty list clears them). */
export function setColumnLayout(view: EditorView, specs: ColumnBlockSpec[]): void {
  const decos: Decoration[] = []
  const styleOf = new Map<HTMLElement, { style: string; cls: string }>()
  for (const spec of specs) {
    const style =
      (spec.widthPx !== undefined ? `--col-w:${round2(spec.widthPx)}px;` : '') +
      (spec.contentWPx !== undefined
        ? `--doc-content-w:${round2(spec.contentWPx)}px;--doc-margin-left:${round2(spec.marginLeftPx ?? 0)}px;--doc-margin-right:${round2(spec.marginRightPx ?? 0)}px;`
        : '') +
      (spec.gridPitchPt !== undefined
        ? `--doc-grid-pitch:${spec.gridPitchPt > 0 ? `${round2(spec.gridPitchPt)}pt` : '0.0001px'};`
        : '') +
      (spec.charSpacePt !== undefined ? `--doc-char-space:${round4(spec.charSpacePt)}pt;` : '') +
      `--col-dx:${round2(spec.dx)}px;--col-dy:${round2(spec.dy)}px`
    const cls = specClass(spec)
    styleOf.set(spec.el, { style, cls })
    let from: number
    let to: number
    try {
      const $inside = view.state.doc.resolve(view.posAtDOM(spec.el, 0))
      from = $inside.before(1)
      to = $inside.after(1)
    } catch {
      continue
    }
    decos.push(Decoration.node(from, to, { class: cls, style }, { key: `col-${cls}-${style}` }))
  }
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = key.getState(view.state)
  if (!prev || !sameCols(prev, next))
    view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false))
  // custom NodeViews (protected blocks etc.) don't apply node decorations — patch
  // their DOM directly, observer paused so PM never re-parses the mutation (same
  // technique as the phantom-rowspan sync). Re-applied by every remeasure pass.
  const obs = (view as unknown as { domObserver?: { stop(): void; start(): void } }).domObserver
  obs?.stop()
  try {
    for (const el of Array.from(
      view.dom.querySelectorAll<HTMLElement>('[data-col-patch]'),
    ) as HTMLElement[]) {
      if (styleOf.has(el)) continue
      el.removeAttribute('data-col-patch')
      el.classList.remove(...PATCH_CLASSES)
      for (const p of PATCH_PROPS) el.style.removeProperty(p)
    }
    for (const [el, { style, cls }] of styleOf) {
      const classes = cls.split(' ')
      if (classes.every((c) => el.classList.contains(c)) && !el.hasAttribute('data-col-patch'))
        continue
      el.setAttribute('data-col-patch', '1')
      el.classList.add(...classes)
      const parts = style.split(';')
      for (const p of PATCH_PROPS) el.style.removeProperty(p)
      for (const part of parts) {
        const i = part.indexOf(':')
        if (i > 0) el.style.setProperty(part.slice(0, i), part.slice(i + 1))
      }
    }
  } finally {
    obs?.start()
  }
}

const round2 = (n: number) => Math.round(n * 100) / 100
const round4 = (n: number) => Math.round(n * 10000) / 10000

function sameCols(a: DecorationSet, b: DecorationSet): boolean {
  const fa = a.find()
  const fb = b.find()
  if (fa.length !== fb.length) return false
  const sig = (d: Decoration) => `${d.from}-${d.to}-${(d.spec as { key?: string }).key ?? ''}`
  const sa = fa.map(sig).sort()
  const sb = fb.map(sig).sort()
  return sa.every((s, i) => s === sb[i])
}
