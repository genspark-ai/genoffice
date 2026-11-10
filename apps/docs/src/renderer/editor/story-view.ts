import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { ReplaceStep } from '@tiptap/pm/transform'
import type { Node as PmNode } from '@tiptap/pm/model'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    docsStoryView: {
      /** switch the story-editor view (#2277) on or off */
      setStoryViewActive: (active: boolean) => ReturnType
    }
  }

  interface Storage {
    docsStoryView: StoryViewStorage
  }
}

export interface StoryViewState {
  active: boolean
  decorations: DecorationSet
}

export const storyViewKey = new PluginKey<StoryViewState>('docsStoryView')

/**
 * View ▸ Story Editor (#2277): a plain writing column beside Print/Web/Outline.
 *
 * The view half is CSS (`.view-story` in styles.css): the document's own
 * fonts, sizes, colors, highlights, indents and spacing are hidden with
 * `!important` overrides of the `docTextStyle` span's inline styles, while the
 * separate emphasis marks (bold/italic/underline/strike/link) and the list
 * markers keep rendering. Nothing in the document model changes on a switch.
 *
 * This extension draws the faded structure markers (`#`/`-`/`>`), the
 * manual-format dot, and the paragraph gutter (block index + style name, the
 * same indexes the CLI and MCP tools use), and carries the editing semantic:
 * text typed (or pasted) while the view is on takes the paragraph style's
 * formatting — the hidden run formatting of the neighbors is stripped from the
 * inserted text, so it cannot be picked up silently.
 */

/** docTextStyle attrs the story view hides (everything but emphasis lives here) */
const HIDDEN_STYLE_ATTRS = [
  'color',
  'sizeHalfPoints',
  'font',
  'fontAscii',
  'csFont',
  'charSpacingTwips',
  'charScaleEm',
  'charScaleX',
  'kern',
  'highlight',
  'shading',
  'textOutline',
  'textEffect',
  'dstrike',
  'glow',
  'positionHalfPoints',
  'bdr',
  'vertAlign',
  'em',
  'boldOff',
  'italicOff',
  'caps',
] as const

/** whether a docTextStyle mark carries formatting the story view hides */
export function hasHiddenOverride(mark: { attrs: Record<string, unknown> }): boolean {
  return HIDDEN_STYLE_ATTRS.some((attr) => mark.attrs[attr] != null)
}

/** whether a block has any run whose manual formatting the view hides (the margin dot) */
export function blockHasManualOverride(node: PmNode): boolean {
  let found = false
  node.descendants((child) => {
    if (found) return false
    if (child.isText && child.marks.some((mark) => hasHiddenOverride(mark))) {
      found = true
      return false
    }
    return true
  })
  return found
}

/// Word's built-in quote styles; the id is locale-independent even when the
/// display name is translated
const QUOTE_STYLE = /(^|[,\s])(intense)?quote([,\s]|$)/i

/** the faded margin marker for a top-level block, or null when it has none */
export function storyMarkerFor(node: PmNode): string | null {
  if (node.type.name === 'docHeading') {
    const level = Math.min(Math.max(Number(node.attrs.level) || 1, 1), 6)
    return '#'.repeat(level)
  }
  if (node.type.name === 'docListItem') return '-'
  if (node.type.name !== 'docParagraph') return null
  const styleId = node.attrs.styleId as string | null
  if (styleId && QUOTE_STYLE.test(styleId)) return '>'
  return null
}

export interface StoryViewStorage {
  /** paragraph styles of the open document (styleId → StyleInfo), for the gutter names */
  styles?: Map<string, { name: string; type: string }>
}

function markerDom(marker: string): HTMLElement {
  const span = document.createElement('span')
  span.className = 'story-marker'
  span.textContent = marker
  span.contentEditable = 'false'
  span.setAttribute('aria-hidden', 'true')
  return span
}

function gutterDom(index: number, styleName: string | null): HTMLElement {
  const span = document.createElement('span')
  span.className = 'story-gutter'
  span.contentEditable = 'false'
  span.setAttribute('aria-hidden', 'true')
  const num = document.createElement('span')
  num.className = 'story-gutter-num'
  num.textContent = String(index)
  span.appendChild(num)
  if (styleName) {
    const name = document.createElement('span')
    name.className = 'story-gutter-name'
    name.textContent = styleName
    span.appendChild(name)
  }
  return span
}

function buildStoryDecorations(doc: PmNode, storage: StoryViewStorage | undefined): DecorationSet {
  const decos: Decoration[] = []
  doc.forEach((node, offset, index) => {
    const marker = storyMarkerFor(node)
    if (marker) {
      decos.push(
        Decoration.widget(offset + 1, () => markerDom(marker), {
          key: `story-marker:${index}:${marker}`,
        }),
      )
    }
    const isPara =
      node.type.name === 'docParagraph' ||
      node.type.name === 'docHeading' ||
      node.type.name === 'docListItem'
    if (isPara && blockHasManualOverride(node)) {
      decos.push(Decoration.node(offset, offset + node.nodeSize, { class: 'story-override' }))
    }
    decos.push(
      Decoration.widget(
        offset + 1,
        () =>
          gutterDom(index, storage?.styles?.get(String(node.attrs.styleId ?? ''))?.name ?? null),
        { key: `story-gutter:${index}` },
      ),
    )
  })
  return DecorationSet.create(doc, decos)
}

export const StoryViewExtension = Extension.create({
  name: 'docsStoryView',

  addStorage() {
    // App fills this after mount with the open document's paragraph styles
    // (read lazily: the value captured here at creation would go stale)
    return { styles: undefined as StoryViewStorage['styles'] }
  },

  addCommands() {
    return {
      setStoryViewActive:
        (active: boolean) =>
        ({ state, dispatch }) => {
          if (storyViewKey.getState(state)?.active === active) return true
          if (dispatch)
            dispatch(state.tr.setMeta(storyViewKey, { active } satisfies Partial<StoryViewState>))
          return true
        },
    }
  },

  addProseMirrorPlugins() {
    const editorStorage = this.editor.storage as unknown as Record<
      string,
      StoryViewStorage | undefined
    >
    return [
      new Plugin({
        key: storyViewKey,
        state: {
          init: () => ({ active: false, decorations: DecorationSet.empty }),
          apply(tr, value) {
            const meta = tr.getMeta(storyViewKey) as Partial<StoryViewState> | undefined
            const active = meta?.active ?? value.active
            if (active === value.active && !tr.docChanged) return value
            return {
              active,
              decorations: active
                ? buildStoryDecorations(tr.doc, editorStorage.storyView)
                : DecorationSet.empty,
            }
          },
        },
        props: {
          decorations: (state: EditorState) =>
            storyViewKey.getState(state)?.decorations ?? DecorationSet.empty,
        },
        // The editing semantic: in story view the typed/pasted text takes the
        // paragraph style's formatting — the neighbors' hidden run formatting
        // (the docTextStyle mark) is stripped from what was inserted, so it
        // cannot be picked up silently. The emphasis marks (bold/italic/
        // underline/strike/link) and the revision/comment marks are separate
        // mark types and stay.
        appendTransaction(trs, _oldState, newState) {
          if (!trs.some((tr) => tr.docChanged)) return null
          const state = storyViewKey.getState(newState)
          if (!state?.active) return null
          const styleType = newState.schema.marks.docTextStyle
          if (!styleType) return null
          const tr = newState.tr
          let changed = false
          for (const source of trs) {
            source.steps.forEach((step, i) => {
              if (!(step instanceof ReplaceStep) || step.slice.size === 0) return
              const from = source.mapping.slice(i).map(step.from, -1)
              const to = source.mapping.slice(i).map(step.to, 1)
              newState.doc.nodesBetween(
                Math.max(0, from),
                Math.min(newState.doc.content.size, to),
                (node, pos) => {
                  if (node.isText && node.marks.some((mark) => mark.type === styleType)) {
                    tr.removeMark(pos, pos + node.nodeSize, styleType)
                    changed = true
                  }
                  return true
                },
              )
            })
          }
          return changed ? tr : null
        },
      }),
    ]
  },
})
