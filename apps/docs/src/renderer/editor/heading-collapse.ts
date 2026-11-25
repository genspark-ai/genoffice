/**
 * Word's Expand / Collapse of the body under a heading in the page views.
 *
 * A collapsed heading hides every block down to the next heading of the same or
 * a higher level; Word stops earlier at the paragraph that ends the section, and
 * a heading inside a table neither folds nor ends a fold. The folds are view
 * state (positions mapped through edits); "Collapsed by default" (w15:collapsed
 * on the paragraph or its style) is what a document opens with.
 */
import { Extension, type Editor } from '@tiptap/core'
import type { Node as PmNode, Slice } from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import { ReplaceStep } from '@tiptap/pm/transform'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { StyleInfo } from '@genoffice/docx-engine'
import { collectHeadings, type HeadingRef } from './headings'
import { headingAtPos, headingKey, subtreeEnd } from './nav-outline'
import { TRACK_IGNORE } from './revisions'
import { PHASED_APPEND } from './streaming-tail-guard'
import { searchPluginKey } from './decoration-extensions'

declare module '@tiptap/core' {
  interface Storage {
    headingCollapse: CollapseStorage
  }
}

export type CollapseStyles = ReadonlyMap<
  string,
  Pick<StyleInfo, 'headingLevel' | 'headingOutlineOff' | 'collapsed'>
>

/** a whole-document replacement that shows the same document again (the reload after a save) */
export const HEADING_COLLAPSE_KEEP = 'headingCollapseKeep'

/** an edit that takes a collapsed heading's text and leaves the body it hides (Replace) */
export const HEADING_BODY_KEEP = 'headingBodyKeep'

const SECTION_BREAK_LABEL = 'Section break paragraph'

/** whether the block is the last paragraph of a section that is not the document's last */
export type SectionEndTest = (node: PmNode) => boolean

export function sectionEndTest(sectionEnds?: ReadonlySet<number>): SectionEndTest {
  return (node) =>
    (node.type.name === 'docProtected' && node.attrs.label === SECTION_BREAK_LABEL) ||
    (typeof node.attrs.docxIndex === 'number' && !!sectionEnds?.has(node.attrs.docxIndex))
}

/** the blocks a collapsed headings[index] hides; null when it has none */
export function headingBodyRange(
  doc: PmNode,
  headings: HeadingRef[],
  index: number,
  isSectionEnd: SectionEndTest,
): { from: number; to: number } | null {
  const heading = doc.nodeAt(headings[index].pos)
  if (!heading || isSectionEnd(heading)) return null
  const from = headings[index].pos + heading.nodeSize
  const limit = subtreeEnd(headings, index, doc.content.size)
  let to = from
  while (to < limit) {
    const node = doc.nodeAt(to)
    if (!node || isSectionEnd(node)) break
    to += node.nodeSize
  }
  return to > from ? { from, to } : null
}

/** where a fold stops short of its natural end: heading position -> end position */
export type FoldLimits = ReadonlyMap<number, number>

/** what the collapsed heading at `pos` hides by itself, whether or not an outer fold hides it too */
function foldRange(
  doc: PmNode,
  headings: HeadingRef[],
  pos: number,
  isSectionEnd: SectionEndTest,
  limits?: FoldLimits,
): { heading: number; from: number; to: number } | null {
  const index = headings.findIndex((h) => h.pos === pos)
  const range = index < 0 ? null : headingBodyRange(doc, headings, index, isSectionEnd)
  if (!range) return null
  const to = Math.min(range.to, limits?.get(pos) ?? range.to)
  return to > range.from ? { heading: pos, from: range.from, to } : null
}

/** hidden ranges of the collapsed headings, outer folds swallowing the ones inside them */
export function hiddenRanges(
  doc: PmNode,
  headings: HeadingRef[],
  collapsed: ReadonlyArray<number>,
  isSectionEnd: SectionEndTest,
  limits?: FoldLimits,
): Array<{ heading: number; from: number; to: number }> {
  const out: Array<{ heading: number; from: number; to: number }> = []
  for (const pos of [...collapsed].sort((a, b) => a - b)) {
    const last = out[out.length - 1]
    if (last && pos < last.to) continue
    const range = foldRange(doc, headings, pos, isSectionEnd, limits)
    if (range) out.push(range)
  }
  return out
}

/** "Collapsed by default" of a heading: its own w15:collapsed, else its style's */
export function collapsedByDefault(node: PmNode, styles?: CollapseStyles): boolean {
  const own = node.attrs.collapsed
  if (typeof own === 'boolean') return own
  const styleId = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : ''
  return (styleId && styles?.get(styleId)?.collapsed) === true
}

export function defaultCollapsed(
  doc: PmNode,
  headings: HeadingRef[],
  styles: CollapseStyles | undefined,
  fromPos = 0,
): number[] {
  const out: number[] = []
  headings.forEach((ref) => {
    if (ref.pos < fromPos) return
    const node = doc.nodeAt(ref.pos)
    if (node && collapsedByDefault(node, styles)) out.push(ref.pos)
  })
  return out
}

export interface HeadingCollapseState {
  /** top-level positions of the collapsed headings, ascending */
  collapsed: readonly number[]
  /** a paragraph opened below a hidden body stays in view until the heading is expanded */
  limits: FoldLimits
  ranges: ReadonlyArray<{ heading: number; from: number; to: number }>
  decorations: DecorationSet
}

export const headingCollapseKey = new PluginKey<HeadingCollapseState>('docsHeadingCollapse')

interface CollapseMeta {
  collapse?: number[]
  expand?: number[]
  /** replace the whole set */
  set?: number[]
  /** rebuild after the section list changed */
  refresh?: true
  limit?: { heading: number; end: number }
}

const NO_LIMITS: FoldLimits = new Map()
const EMPTY: HeadingCollapseState = {
  collapsed: [],
  limits: NO_LIMITS,
  ranges: [],
  decorations: DecorationSet.empty,
}

function isFullReplace(tr: Transaction, before: PmNode): boolean {
  const step = tr.steps[0]
  return step instanceof ReplaceStep && step.from === 0 && step.to === before.content.size
}

function toggleDom(label: string): HTMLElement {
  const el = document.createElement('span')
  el.className = 'heading-toggle is-collapsed'
  el.contentEditable = 'false'
  el.setAttribute('role', 'button')
  el.setAttribute('aria-expanded', 'false')
  el.setAttribute('aria-label', label)
  el.title = label
  return el
}

export interface CollapseStorage {
  sectionEnds: ReadonlySet<number>
  /** runs the collapse command; unset in editors without a command host */
  onToggle: ((pos: number, collapse: boolean) => void) | null
  labels: { expand: string; collapse: string }
}

export function collapseHeadingsOf(doc: PmNode, styles: CollapseStyles | undefined): HeadingRef[] {
  return collectHeadings(doc, styles, true)
}

/** collapsed headings and their hidden ranges ([] when nothing is folded) */
export function collapsedHeadingRanges(state: EditorState): HeadingCollapseState['ranges'] {
  return headingCollapseKey.getState(state)?.ranges ?? []
}

const ACTIONS = ['collapse', 'expand', 'toggle', 'collapse-all', 'expand-all'] as const
export type HeadingCollapseAction = (typeof ACTIONS)[number]

export function isHeadingCollapseAction(value: string): value is HeadingCollapseAction {
  return (ACTIONS as readonly string[]).includes(value)
}

/** runHeadingCollapse with the editor's own styles and section ends */
export function foldHeading(editor: Editor, action: HeadingCollapseAction, pos?: number): boolean {
  const storage = editor.storage.headingCollapse
  return runHeadingCollapse(
    editor.view,
    action,
    editor.storage.listNumbering?.styles,
    sectionEndTest(storage?.sectionEnds),
    pos,
  )
}

/**
 * Expand / collapse the heading at `pos` (default: the heading the caret is in or
 * under), or every heading. A caret left inside a body that folds moves to the
 * end of its heading.
 */
export function runHeadingCollapse(
  view: EditorView,
  action: HeadingCollapseAction,
  styles: CollapseStyles | undefined,
  isSectionEnd: SectionEndTest,
  pos?: number,
): boolean {
  const { state } = view
  const plugin = headingCollapseKey.getState(state)
  if (!plugin) return false
  const headings = collapseHeadingsOf(state.doc, styles)
  const foldable = (index: number) => !!headingBodyRange(state.doc, headings, index, isSectionEnd)
  let meta: CollapseMeta
  if (action === 'collapse-all') {
    meta = { set: headings.filter((_, i) => foldable(i)).map((h) => h.pos) }
  } else if (action === 'expand-all') {
    meta = { set: [] }
  } else {
    const $head = state.selection.$head
    const at = pos ?? ($head.depth ? $head.before(1) : $head.pos)
    let index = headingAtPos(headings, at)
    if (index < 0) return false
    const isCollapsed = (i: number) => plugin.collapsed.includes(headings[i].pos)
    // under a collapsed ancestor the fold that shows is the ancestor's
    if (action !== 'collapse') {
      const hiding = plugin.ranges.find((r) => at >= r.from && at < r.to)
      if (hiding) index = headings.findIndex((h) => h.pos === hiding.heading)
    }
    const collapse = action === 'toggle' ? !isCollapsed(index) : action === 'collapse'
    if (collapse === isCollapsed(index) || !foldable(index)) return false
    meta = collapse ? { collapse: [headings[index].pos] } : { expand: [headings[index].pos] }
  }
  const tr = state.tr.setMeta(headingCollapseKey, meta).setMeta('addToHistory', false)
  const next = meta.set ?? [...plugin.collapsed, ...(meta.collapse ?? [])]
  const hidden = hiddenRanges(state.doc, headings, next, isSectionEnd)
  const caret = state.selection.head
  const over = hidden.find((r) => caret >= r.from && caret < r.to)
  if (over) tr.setSelection(TextSelection.near(state.doc.resolve(over.from - 1), -1))
  view.dispatch(tr)
  return true
}

export const HeadingCollapseExtension = Extension.create<object, CollapseStorage>({
  name: 'headingCollapse',

  addStorage() {
    return {
      sectionEnds: new Set<number>(),
      onToggle: null,
      labels: { expand: 'Expand', collapse: 'Collapse' },
    }
  },

  addProseMirrorPlugins() {
    const editor = this.editor
    const storage = this.storage
    const styles = (): CollapseStyles | undefined => editor.storage.listNumbering?.styles
    const isSectionEnd: SectionEndTest = (node) => sectionEndTest(storage.sectionEnds)(node)

    const build = (
      doc: PmNode,
      wanted: readonly number[],
      wantedLimits: FoldLimits,
    ): HeadingCollapseState => {
      if (wanted.length === 0) return EMPTY
      const headings = collapseHeadingsOf(doc, styles())
      const limits = new Map([...wantedLimits].filter(([pos]) => wanted.includes(pos)))
      const ranges = hiddenRanges(doc, headings, wanted, isSectionEnd, limits)
      // a fold inside a hidden range stays in the set: it shows again when the outer one opens
      const at = new Set(headings.map((h) => h.pos))
      const collapsed = [...new Set(wanted)].filter((pos) => at.has(pos)).sort((a, b) => a - b)
      if (collapsed.length === 0) return EMPTY
      const decos: Decoration[] = []
      for (const range of ranges) {
        const heading = doc.nodeAt(range.heading)!
        decos.push(
          Decoration.node(range.heading, range.heading + heading.nodeSize, {
            class: 'doc-heading-collapsed',
          }),
          Decoration.widget(range.heading + 1, () => toggleDom(storage.labels.expand), {
            side: -1,
            key: 'heading-toggle',
            ignoreSelection: true,
          }),
        )
        doc.nodesBetween(range.from, range.to, (node, pos) => {
          decos.push(Decoration.node(pos, pos + node.nodeSize, { class: 'doc-collapsed-body' }))
          return false
        })
        if (range.to < doc.content.size)
          decos.push(
            Decoration.widget(range.to, foldGapDom, {
              side: -1,
              key: 'fold-gap',
              ignoreSelection: true,
            }),
          )
      }
      return { collapsed, limits, ranges, decorations: DecorationSet.create(doc, decos) }
    }

    const toggleAt = (view: EditorView, pos: number, collapse: boolean) => {
      if (storage.onToggle) storage.onToggle(pos, collapse)
      else runHeadingCollapse(view, collapse ? 'collapse' : 'expand', styles(), isSectionEnd, pos)
    }

    return [
      new Plugin<HeadingCollapseState>({
        key: headingCollapseKey,
        state: {
          init: () => EMPTY,
          apply(tr, value, oldState) {
            const meta = tr.getMeta(headingCollapseKey) as CollapseMeta | undefined
            let wanted: readonly number[] = value.collapsed
            let limits = value.limits
            if (tr.docChanged) {
              const fresh = (fromPos: number) => {
                const headings = collapseHeadingsOf(tr.doc, styles())
                return defaultCollapsed(tr.doc, headings, styles(), fromPos)
              }
              // through the heading's text, so a split at its start keeps the fold on the text
              const follow = (pos: number): number | null => {
                if (tr.mapping.mapResult(pos, 1).deleted) return null
                const $inside = tr.doc.resolve(tr.mapping.map(pos + 1, 1))
                return $inside.depth ? $inside.before(1) : $inside.pos
              }
              if (isFullReplace(tr, oldState.doc)) {
                if (!tr.getMeta(HEADING_COLLAPSE_KEEP)) {
                  wanted = fresh(0)
                  limits = NO_LIMITS
                } else if (wanted.length > 0) {
                  // the same document again: folds follow their heading by order and text
                  const before = collapseHeadingsOf(oldState.doc, styles())
                  const after = collapseHeadingsOf(tr.doc, styles())
                  const twinOf = (pos: number): number | null => {
                    const i = before.findIndex((h) => h.pos === pos)
                    const twin = i < 0 ? undefined : after[i]
                    return twin && headingKey(twin) === headingKey(before[i]) ? twin.pos : null
                  }
                  // a fold's early end is a block boundary: the same block, counted from the top
                  const sameBlocks = tr.doc.childCount === oldState.doc.childCount
                  const blockStart = (end: number): number => {
                    const index = oldState.doc.resolve(end).index(0)
                    let at = 0
                    for (let i = 0; i < index; i++) at += tr.doc.child(i).nodeSize
                    return at
                  }
                  limits = new Map(
                    [...limits].flatMap(([pos, end]): Array<[number, number]> => {
                      const at = twinOf(pos)
                      return at === null || !sameBlocks ? [] : [[at, blockStart(end)]]
                    }),
                  )
                  wanted = wanted.flatMap((pos) => {
                    const at = twinOf(pos)
                    return at === null ? [] : [at]
                  })
                }
              } else {
                if (limits.size > 0)
                  limits = new Map(
                    [...limits].flatMap(([pos, end]): Array<[number, number]> => {
                      const at = follow(pos)
                      return at === null ? [] : [[at, tr.mapping.map(end, -1)]]
                    }),
                  )
                if (wanted.length > 0)
                  wanted = wanted.flatMap((pos) => {
                    const at = follow(pos)
                    return at === null ? [] : [at]
                  })
                if (tr.getMeta(PHASED_APPEND))
                  wanted = [...wanted, ...fresh(oldState.doc.content.size)]
              }
            }
            // undo or redo of "Collapsed by default" takes the fold it made with it
            if (tr.docChanged && tr.getMeta('history$') && !isFullReplace(tr, oldState.doc)) {
              const back = tr.mapping.invert()
              for (const ref of collapseHeadingsOf(tr.doc, styles())) {
                const now = tr.doc.nodeAt(ref.pos)
                const was = oldState.doc.nodeAt(back.map(ref.pos, 1))
                if (!now || !was || was.type !== now.type) continue
                const on = collapsedByDefault(now, styles())
                if (on === collapsedByDefault(was, styles())) continue
                wanted = on ? [...wanted, ref.pos] : wanted.filter((pos) => pos !== ref.pos)
              }
            }
            if (meta?.set) {
              wanted = meta.set
              limits = NO_LIMITS
            }
            if (meta?.collapse) wanted = [...wanted, ...meta.collapse]
            if (meta?.expand) {
              wanted = wanted.filter((pos) => !meta.expand!.includes(pos))
              if (limits.size > 0)
                limits = new Map([...limits].filter(([pos]) => !meta.expand!.includes(pos)))
            }
            if (meta?.limit) limits = new Map(limits).set(meta.limit.heading, meta.limit.end)
            if (wanted.length === 0) return value.collapsed.length === 0 ? value : EMPTY
            if (
              !tr.docChanged &&
              !meta?.refresh &&
              limits === value.limits &&
              wanted.length === value.collapsed.length &&
              wanted.every((pos, i) => pos === value.collapsed[i])
            )
              return value
            return build(tr.doc, wanted, limits)
          },
        },
        appendTransaction(trs, oldState, newState) {
          const plugin = headingCollapseKey.getState(newState)
          const before = headingCollapseKey.getState(oldState)
          if (!plugin || !before) return null
          if (before.ranges.length > 0 && trs.some((tr) => tr.docChanged)) {
            const edit = bodyEdit(trs, oldState, newState, before.ranges, plugin.ranges)
            if (edit) return edit
          }
          if (plugin.ranges.length === 0) return null
          const points: number[] = []
          // a document just loaded parks the selection at its end, which may be a hidden body
          const loaded =
            trs.some((tr) => tr.getMeta(PHASED_APPEND)) || isFullReplace(trs[0], oldState.doc)
          // Go To, the navigation pane or an edit put the caret in a hidden body: open it
          if (!loaded && !newState.selection.eq(oldState.selection))
            points.push(newState.selection.head, newState.selection.anchor)
          // so does Find stepping onto a match there
          for (const tr of trs) {
            const search = tr.getMeta(searchPluginKey) as
              { ranges: Array<{ from: number }>; activeIndex: number } | undefined
            const active = search?.ranges[search.activeIndex]
            if (active) points.push(active.from)
          }
          const inside = (r: { from: number; to: number }) =>
            points.some((pos) => pos >= r.from && pos < r.to)
          if (!plugin.ranges.some(inside)) return null
          // every fold on the way down to it, not only the outer one
          const headings = collapseHeadingsOf(newState.doc, styles())
          const opened = plugin.collapsed.filter((pos) => {
            const own = foldRange(newState.doc, headings, pos, isSectionEnd, plugin.limits)
            return !!own && inside(own)
          })
          if (opened.length === 0) return null
          return newState.tr
            .setMeta(headingCollapseKey, { expand: opened } satisfies CollapseMeta)
            .setMeta('addToHistory', false)
        },
        view(view) {
          return hoverToggle(view, {
            headingAt: (pos) => {
              // most blocks under the pointer are body text: no scan of the document for those
              const node = view.state.doc.nodeAt(pos)
              const style = typeof node?.attrs.styleId === 'string' ? node.attrs.styleId : ''
              if (node?.type.name !== 'docHeading' && !styles()?.get(style)?.headingLevel)
                return null
              const headings = collapseHeadingsOf(view.state.doc, styles())
              const index = headings.findIndex((h) => h.pos === pos)
              if (index < 0 || !headingBodyRange(view.state.doc, headings, index, isSectionEnd))
                return null
              const collapsed = !!headingCollapseKey.getState(view.state)?.collapsed.includes(pos)
              return { collapsed }
            },
            label: () => storage.labels.collapse,
            toggle: (pos) => toggleAt(view, pos, true),
          })
        },
        props: {
          decorations: (state) => headingCollapseKey.getState(state)?.decorations,
          handleDOMEvents: {
            mousedown(view, event) {
              const target = event.target
              if (!(target instanceof Element) || !target.classList.contains('heading-toggle'))
                return false
              event.preventDefault()
              const pos = view.posAtDOM(target, 0)
              const $pos = view.state.doc.resolve(pos)
              toggleAt(view, $pos.depth ? $pos.before(1) : pos, false)
              return true
            },
          },
          // copying or cutting a collapsed heading takes the body it hides
          transformCopied(slice, view) {
            return selectionWithHiddenBody(view.state)?.slice ?? slice
          },
        },
      }),
    ]
  },
})

/** a selected collapsed heading together with the body it hides, as a cut or copy takes it */
export function selectionWithHiddenBody(state: EditorState): { slice: Slice; to: number } | null {
  const { from, to } = state.selection
  const range = wholeHeadingRange(state, from, to)
  return range ? { slice: state.doc.slice(from, range.to), to: range.to } : null
}

/** the hidden range of the collapsed heading whose whole text [from, to] selects, short of its body */
function wholeHeadingRange(
  state: EditorState,
  from: number,
  to: number,
): HeadingCollapseState['ranges'][number] | undefined {
  if (from === to) return undefined
  return collapsedHeadingRanges(state).find((r) => {
    const node = state.doc.nodeAt(r.heading)
    if (!node || node.content.size === 0) return false
    return from <= r.heading + 1 && to >= r.from - 1 && to < r.to
  })
}

/**
 * What an edit of a collapsed heading means for its hidden body:
 * deleting the selected heading deletes the body (Word), and a paragraph opened
 * by Enter at the heading's end goes below the body, not into it.
 */
function bodyEdit(
  trs: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState,
  before: HeadingCollapseState['ranges'],
  after: HeadingCollapseState['ranges'],
): Transaction | null {
  if (
    trs.some(
      (tr) =>
        tr.getMeta(headingCollapseKey) || tr.getMeta('history$') || tr.getMeta(HEADING_BODY_KEEP),
    )
  )
    return null
  const map = (pos: number, assoc: number) => {
    let result = { pos, deletedAfter: false, deletedBefore: false }
    for (const tr of trs) {
      const r = tr.mapping.mapResult(result.pos, assoc)
      result = {
        pos: r.pos,
        deletedAfter: result.deletedAfter || r.deletedAfter,
        deletedBefore: result.deletedBefore || r.deletedBefore,
      }
    }
    return result
  }
  const { from, to, empty } = oldState.selection
  const whole = wholeHeadingRange(oldState, from, to)
  if (whole) {
    // typing over the selected heading renames it; only a plain deletion takes the body
    const removedOnly = newState.doc.content.size === oldState.doc.content.size - (to - from)
    const gone = map(whole.heading + 1, 1).deletedAfter && map(whole.from - 1, -1).deletedBefore
    const start = map(whole.from, 1).pos
    const end = map(whole.to, -1).pos
    if (removedOnly && gone && end > start) return newState.tr.delete(start, end)
    return null
  }
  if (!empty) return null
  const $old = oldState.selection.$head
  const range = $old.depth === 1 ? before.find((r) => r.heading === $old.before(1)) : undefined
  if (!range || $old.parentOffset !== $old.parent.content.size) return null
  const now = after.find((r) => r.heading === range.heading)
  const $new = newState.selection.$head
  if (!now || !newState.selection.empty || $new.depth !== 1 || $new.before(1) !== now.from)
    return null
  const opened = $new.parent
  if (opened.content.size > 0 || newState.doc.childCount !== oldState.doc.childCount + 1)
    return null
  const tr = newState.tr.delete(now.from, now.from + opened.nodeSize)
  const at = now.to - opened.nodeSize
  tr.insert(at, opened)
  tr.setSelection(TextSelection.create(tr.doc, at + 1)).scrollIntoView()
  tr.setMeta(headingCollapseKey, {
    limit: { heading: now.heading, end: at },
  } satisfies CollapseMeta)
  return tr.setMeta(TRACK_IGNORE, true)
}

function foldGapDom(): HTMLElement {
  const el = document.createElement('div')
  el.className = 'doc-fold-gap'
  el.contentEditable = 'false'
  return el
}

const PARAGRAPH_TAGS = /^(P|H[1-6]|LI)$/

/**
 * Word keeps the hidden paragraphs in the spacing arithmetic: the block after a
 * collapsed body gives up as much of its space before as the last hidden
 * paragraph had space after, while the heading keeps its own space after
 * (probes p13, p15a-g). The zero-height spacer between the two carries that as
 * a negative margin; with a page gap between them nothing is taken.
 */
function syncFoldGaps(view: EditorView): void {
  for (const range of collapsedHeadingRanges(view.state)) {
    const next = view.nodeDOM(range.to)
    const gap = next instanceof HTMLElement ? next.previousElementSibling : null
    if (!(gap instanceof HTMLElement) || !gap.classList.contains('doc-fold-gap')) continue
    const last = gap.previousElementSibling
    let take = 0
    if (
      last instanceof HTMLElement &&
      last.classList.contains('doc-collapsed-body') &&
      (PARAGRAPH_TAGS.test(last.tagName) || last.classList.contains('doc-li'))
    ) {
      const after = parseFloat(getComputedStyle(last).marginBottom) || 0
      const before = parseFloat(getComputedStyle(next as HTMLElement).marginTop) || 0
      take = Math.max(0, Math.min(after, before))
    }
    const want = take > 0 ? `-${take}px` : ''
    if (gap.style.marginBottom !== want) gap.style.marginBottom = want
  }
}

interface HoverHost {
  headingAt(pos: number): { collapsed: boolean } | null
  label(): string
  toggle(pos: number): void
}

/**
 * Word's triangle beside the heading under the pointer. It floats over the page
 * instead of sitting in the document DOM, so a document with nothing collapsed
 * renders the same nodes as before.
 */
function hoverToggle(view: EditorView, host: HoverHost) {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = 'heading-toggle heading-toggle-hover'
  button.tabIndex = -1
  button.hidden = true
  document.body.appendChild(button)
  let shown: { pos: number; el: HTMLElement } | null = null
  /** the block last found to have no triangle, so moving inside it costs nothing */
  let plain: HTMLElement | null = null
  const hide = () => {
    shown = null
    button.hidden = true
  }
  const place = (el: HTMLElement) => {
    const range = document.createRange()
    range.selectNodeContents(el)
    const line = range.getClientRects()[0] ?? el.getBoundingClientRect()
    const size = Math.max(10, Math.min(18, line.height * 0.6))
    button.style.width = `${size}px`
    button.style.height = `${size}px`
    button.style.left = `${el.getBoundingClientRect().left - size - 4}px`
    button.style.top = `${line.top + (line.height - size) / 2}px`
  }
  const onMove = (event: MouseEvent) => {
    const target = event.target
    if (target === button) return
    // Outline view folds its own rows
    if (view.dom.closest('.view-outline')) return hide()
    const block =
      target instanceof Element && view.dom.contains(target)
        ? (target.closest('.ProseMirror > *') as HTMLElement | null)
        : null
    if (!block || block.parentElement !== view.dom) {
      // the margin beside the heading leads to the triangle: keep it while the pointer is level with it
      if (shown) {
        const r = shown.el.getBoundingClientRect()
        if (event.clientY >= r.top && event.clientY <= r.bottom && event.clientX < r.right) return
      }
      return hide()
    }
    if (shown?.el === block || plain === block) return
    plain = block
    let pos: number
    try {
      pos = view.posAtDOM(block, 0)
    } catch {
      return hide()
    }
    const $pos = view.state.doc.resolve(pos)
    const top = $pos.depth ? $pos.before(1) : pos
    const info = host.headingAt(top)
    // a collapsed heading carries its own triangle
    if (!info || info.collapsed) return hide()
    plain = null
    shown = { pos: top, el: block }
    const label = host.label()
    button.setAttribute('aria-label', label)
    button.title = label
    button.setAttribute('aria-expanded', 'true')
    place(block)
    button.hidden = false
  }
  const onDown = (event: MouseEvent) => {
    event.preventDefault()
    const pos = shown?.pos
    hide()
    if (pos !== undefined) host.toggle(pos)
  }
  const scroller = view.dom.ownerDocument
  scroller.addEventListener('mousemove', onMove)
  scroller.addEventListener('scroll', hide, true)
  button.addEventListener('mousedown', onDown)
  syncFoldGaps(view)
  return {
    update: (v: EditorView, prev: EditorState) => {
      if (v.state !== prev) plain = null
      if (shown && v.state.doc !== prev.doc) hide()
      syncFoldGaps(v)
    },
    destroy: () => {
      scroller.removeEventListener('mousemove', onMove)
      scroller.removeEventListener('scroll', hide, true)
      button.remove()
    },
  }
}

/** fold facts of the top-level heading `pos` is in; null when it is not in one */
export function headingFoldAt(
  editor: Editor,
  pos: number,
): { pos: number; collapsed: boolean; foldable: boolean } | null {
  const { doc } = editor.state
  if (!headingCollapseKey.getState(editor.state)) return null
  const $pos = doc.resolve(Math.min(Math.max(pos, 0), doc.content.size))
  if ($pos.depth !== 1) return null
  const top = $pos.before(1)
  const headings = collapseHeadingsOf(doc, editor.storage.listNumbering?.styles)
  const index = headings.findIndex((h) => h.pos === top)
  if (index < 0) return null
  const storage = editor.storage.headingCollapse
  const foldable = !!headingBodyRange(doc, headings, index, sectionEndTest(storage?.sectionEnds))
  const collapsed = collapsedHeadingRanges(editor.state).some((r) => r.heading === top)
  return { pos: top, collapsed, foldable }
}

/** top-level headings the selection touches */
function selectedHeadings(editor: Editor): HeadingRef[] {
  const { doc, selection } = editor.state
  return collapseHeadingsOf(doc, editor.storage.listNumbering?.styles).filter((h) => {
    const end = h.pos + (doc.nodeAt(h.pos)?.nodeSize ?? 0)
    return selection.from < end && selection.to > h.pos
  })
}

/** Paragraph dialog: whether "Collapsed by default" applies at the selection and what it shows */
export function collapsedByDefaultAt(editor: Editor): { enabled: boolean; checked: boolean } {
  const first = selectedHeadings(editor)[0]
  const node = first ? editor.state.doc.nodeAt(first.pos) : null
  if (!node) return { enabled: false, checked: false }
  return { enabled: true, checked: collapsedByDefault(node, editor.storage.listNumbering?.styles) }
}

/** set "Collapsed by default" on the selected headings; the view folds or opens with it */
export function setCollapsedByDefault(editor: Editor, on: boolean): boolean {
  const styles: CollapseStyles | undefined = editor.storage.listNumbering?.styles
  const targets = selectedHeadings(editor)
  if (targets.length === 0) return false
  const { tr, doc } = editor.state
  for (const ref of targets) {
    const node = doc.nodeAt(ref.pos)!
    const styleId = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : ''
    const fromStyle = (styleId && styles?.get(styleId)?.collapsed) === true
    // only a value differing from the style's is written on the paragraph
    tr.setNodeMarkup(ref.pos, undefined, { ...node.attrs, collapsed: on === fromStyle ? null : on })
  }
  const at = targets.map((h) => h.pos)
  tr.setMeta(headingCollapseKey, (on ? { collapse: at } : { expand: at }) satisfies CollapseMeta)
  if (on) {
    const storage = editor.storage.headingCollapse
    const headings = collapseHeadingsOf(doc, styles)
    const hidden = hiddenRanges(doc, headings, at, sectionEndTest(storage?.sectionEnds))
    const caret = editor.state.selection.head
    const over = hidden.find((r) => caret >= r.from && caret < r.to)
    if (over) tr.setSelection(TextSelection.near(tr.doc.resolve(over.from - 1), -1))
  }
  editor.view.dispatch(tr)
  return true
}
