import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { collectHeadings, type HeadingRef } from './headings'
import { headingAtPos, headingKey } from './nav-outline'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    docsOutlineView: {
      /** draw the outline-view collapse handles (true) or clear them and refold */
      setOutlineViewActive: (active: boolean) => ReturnType
    }
  }
}

export interface OutlineViewState {
  active: boolean
  /** headingKey set of the folded headings (fold = hide the descendant heading rows) */
  collapsed: ReadonlySet<string>
  decorations: DecorationSet
}

export const outlineViewKey = new PluginKey<OutlineViewState>('docsOutlineView')

/**
 * View ▸ Outline View collapse handles.
 *
 * Word's outline view marks every heading with a ⊖/⊕ handle that folds the
 * heading's subtree. The outline view (styles.css `.view-outline`) already
 * hides every non-heading block, so the visible subtree of a heading is its
 * descendant heading rows: folding hides those rows down to the next heading
 * of the same or a higher level — the same visibility semantics as the
 * navigation pane's `visibleRows`. The old CSS painted the glyph with a
 * `::before` pseudo-element, which has no click target; this extension draws
 * a real DOM handle (widget decoration) per heading and keeps the collapsed
 * set as plugin state.
 *
 * The collapsed set is keyed by level + text (`headingKey`, like the
 * navigation pane), so folds survive edits that move headings; renaming a
 * heading re-expands it. Leaving outline view clears the set: the folds are
 * view state, not document state, and re-expanding everything on the next
 * visit keeps that honest without persisting anything.
 */

interface OutlineViewMeta {
  /** switch the view on/off; switching off also clears the collapsed set */
  active?: boolean
  /** heading key to flip in the collapsed set */
  toggle?: string
}

/** heading indices whose ancestor rows are folded away by a collapsed heading */
export function hiddenOutlineHeadings(
  headings: ReadonlyArray<Pick<HeadingRef, 'level' | 'text'>>,
  collapsed: ReadonlySet<string>,
): Set<number> {
  const hidden = new Set<number>()
  const stack: Array<Pick<HeadingRef, 'level' | 'text'>> = []
  headings.forEach((ref, index) => {
    while (stack.length && stack[stack.length - 1].level >= ref.level) stack.pop()
    for (const ancestor of stack) {
      if (collapsed.has(headingKey(ancestor))) {
        hidden.add(index)
        break
      }
    }
    stack.push(ref)
  })
  return hidden
}

/** whether a deeper heading sits under headings[index] (the handle has something to fold) */
export function outlineHasChildren(
  headings: ReadonlyArray<Pick<HeadingRef, 'level'>>,
  index: number,
): boolean {
  // the next heading is either inside the subtree (deeper level: a descendant
  // row) or it is the boundary that ends it (same or higher level: no rows)
  const next = headings[index + 1]
  return next != null && next.level > headings[index].level
}

function handleDom(leaf: boolean, collapsed: boolean): HTMLElement {
  const span = document.createElement('span')
  span.className = `outline-handle${leaf ? ' is-leaf' : ''}`
  // Word: ⊖ marks an expanded heading (click folds its subtree), ⊕ a collapsed one
  span.textContent = collapsed ? '⊕' : '⊖'
  span.contentEditable = 'false'
  return span
}

function buildOutlineDecorations(
  doc: ProseMirrorNode,
  collapsed: ReadonlySet<string>,
): DecorationSet {
  // includeEmpty: a blank heading shows its handle right away, like the old
  // always-on ::before glyph did
  const headings = collectHeadings(doc, undefined, true)
  if (headings.length === 0) return DecorationSet.empty
  const hidden = hiddenOutlineHeadings(headings, collapsed)
  const decos: Decoration[] = []
  headings.forEach((ref, index) => {
    const node = doc.nodeAt(ref.pos)
    if (!node) return
    if (hidden.has(index)) {
      decos.push(Decoration.node(ref.pos, ref.pos + node.nodeSize, { class: 'outline-collapsed' }))
      return
    }
    const leaf = !outlineHasChildren(headings, index)
    const isCollapsed = collapsed.has(headingKey(ref))
    // the key folds the glyph state in: widgets with the same pos + key keep
    // their DOM, so a plain key would never swap ⊖ for ⊕
    decos.push(
      Decoration.widget(ref.pos + 1, () => handleDom(leaf, isCollapsed), {
        key: `outline-handle:${leaf ? 'leaf' : isCollapsed ? 'folded' : 'open'}`,
      }),
    )
  })
  return DecorationSet.create(doc, decos)
}

export const OutlineViewExtension = Extension.create({
  name: 'docsOutlineView',

  addCommands() {
    return {
      setOutlineViewActive:
        (active: boolean) =>
        ({ state, dispatch }) => {
          if (outlineViewKey.getState(state)?.active === active) return true
          if (dispatch)
            dispatch(state.tr.setMeta(outlineViewKey, { active } satisfies OutlineViewMeta))
          return true
        },
    }
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: outlineViewKey,
        state: {
          init: () => ({
            active: false,
            collapsed: new Set<string>(),
            decorations: DecorationSet.empty,
          }),
          apply(tr, value) {
            const meta = tr.getMeta(outlineViewKey) as OutlineViewMeta | undefined
            let { active, collapsed } = value
            if (meta?.toggle) {
              collapsed = new Set(collapsed)
              if (collapsed.has(meta.toggle)) collapsed.delete(meta.toggle)
              else collapsed.add(meta.toggle)
            }
            if (meta?.active != null && meta.active !== active) {
              active = meta.active
              if (!active) collapsed = new Set()
            }
            // folds follow their key across edits, so only doc changes need a rebuild
            if (active === value.active && collapsed === value.collapsed && !tr.docChanged)
              return value
            return {
              active,
              collapsed,
              decorations: active
                ? buildOutlineDecorations(tr.doc, collapsed)
                : DecorationSet.empty,
            }
          },
        },
        props: {
          decorations: (state: EditorState) =>
            outlineViewKey.getState(state)?.decorations ?? DecorationSet.empty,
          handleDOMEvents: {
            mousedown(view, event) {
              const target = event.target
              if (!(target instanceof Element)) return false
              if (!target.classList.contains('outline-handle')) return false
              // a handle click must not move the caret into the heading
              event.preventDefault()
              if (target.classList.contains('is-leaf')) return true
              const plugin = outlineViewKey.getState(view.state)
              if (!plugin?.active) return true
              const pos = view.posAtDOM(target, 0)
              const headings = collectHeadings(view.state.doc, undefined, true)
              const index = headingAtPos(headings, pos)
              if (index < 0) return true
              view.dispatch(
                view.state.tr.setMeta(outlineViewKey, {
                  toggle: headingKey(headings[index]),
                } satisfies OutlineViewMeta),
              )
              return true
            },
          },
        },
      }),
    ]
  },
})
