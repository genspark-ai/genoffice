/**
 * Story view: one column of text with the document's own formatting hidden.
 *
 * The way InDesign's story editor (⌘Y) and iA Writer show it: the words in a
 * single column, the document's look gone. It is a pure view — the document
 * keeps every byte it had, and the save path rewrites only the blocks that
 * were edited.
 *
 * The hidden formatting rides two carriers: the docTextStyle mark (a run's own
 * font, size, colour, highlight, spacing) and the block's own indent and
 * spacing. Both are CSS in the view (styles.css `.view-story`), not the model,
 * so hiding them needs no document change at all.
 *
 * Two things the CSS cannot do live here:
 *
 * - the margin markers. A heading gets one '#' per level, a quote-styled
 *   paragraph gets '>': faded Markdown-style marks that hang in the gutter so
 *   the structure the hidden formatting used to show stays readable. A heading
 *   also gets a class, so a heading that is a styled paragraph (Word's
 *   numbered headings are list items) takes the same look as a real h1..h6.
 *   When the document has no heading styles at all, a short bold or all-caps
 *   line (or a section number) is read as an inferred heading: same marker and
 *   class, faded so the guess stays distinguishable from a real style.
 * - the typing rule. Text typed in the view takes the paragraph style's
 *   formatting, so typing next to a run with a manual font or colour does not
 *   silently pick it up. Only the typed text loses the hidden mark; the text
 *   around it keeps its formatting, and the bold and italic marks stay.
 */
import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { ReplaceStep } from '@tiptap/pm/transform'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { TRACK_IGNORE } from './revisions'
import { collectStructuralHeadings, structuralMarker } from './structural-headings'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    docsStoryView: {
      /** draw the story-view margin markers (true) or clear them */
      setStoryViewActive: (active: boolean) => ReturnType
    }
  }
}

/** what the markers read off a paragraph style: its heading level and its name */
export interface StoryStyle {
  headingLevel?: number
  headingOutlineOff?: true
  name?: string
}

export type StoryStyles = ReadonlyMap<string, StoryStyle>

export interface StoryViewState {
  active: boolean
  decorations: DecorationSet
}

export const storyViewKey = new PluginKey<StoryViewState>('docsStoryView')

/** the mark the view hides: the run's own font, size, colour, highlight, spacing */
export const STORY_HIDDEN_MARK = 'docTextStyle'

/**
 * Word's built-in quote styles. The engine reads the heading levels the same
 * way (the canonical English name, with the styleId as the fallback for a
 * localized Word), so the names here are the canonical ones.
 */
const QUOTE_NAMES: ReadonlySet<string> = new Set(['quote', 'intense quote'])
const QUOTE_IDS: ReadonlySet<string> = new Set(['quote', 'intensequote'])

/** the node types that can be a heading without being a docHeading */
const HEADING_CAPABLE: Record<string, true> = { docParagraph: true, docListItem: true }

/**
 * The heading level a block counts as: its own h1..h6, or the level its
 * paragraph style confers (Word's numbered headings keep their numbering in
 * the style, so the parser classifies those paragraphs as list items).
 */
export function storyHeadingLevel(node: ProseMirrorNode, styles?: StoryStyles): number | null {
  if (node.type.name === 'docHeading')
    return Math.min(Math.max(Number(node.attrs.level) || 1, 1), 6)
  if (!styles || HEADING_CAPABLE[node.type.name] !== true) return null
  const styleId = typeof node.attrs?.styleId === 'string' ? node.attrs.styleId : ''
  if (!styleId) return null
  const style = styles.get(styleId)
  // w:outlineLvl 9 on the style means body text even when a basedOn ancestor is a heading
  if (!style?.headingLevel || style.headingOutlineOff) return null
  return style.headingLevel
}

/** whether a block sits in a quote style (Word's Quote / Intense Quote) */
export function storyIsQuote(node: ProseMirrorNode, styles?: StoryStyles): boolean {
  const styleId = typeof node.attrs?.styleId === 'string' ? node.attrs.styleId : ''
  if (!styleId || !styles) return false
  const style = styles.get(styleId)
  const name = (style?.name ?? '').trim().toLowerCase()
  if (QUOTE_NAMES.has(name)) return true
  return QUOTE_IDS.has(styleId.toLowerCase())
}

/**
 * The marker a block gets in the margin: one '#' per heading level, '>' for a
 * quote. A list item gets none — the document's own marker moves into the
 * gutter (styles.css), so a numbered list keeps its numbers.
 */
export function storyMarkerFor(node: ProseMirrorNode, styles?: StoryStyles): string | null {
  const level = storyHeadingLevel(node, styles)
  if (level != null) return '#'.repeat(Math.min(level, 6))
  if (storyIsQuote(node, styles)) return '>'
  return null
}

/** the class a heading block carries, so it takes the view's one heading look */
export function storyHeadingClass(level: number): string {
  return `story-heading story-h${Math.min(Math.max(level, 1), 6)}`
}

function markerDom(marker: string, inferred = false): HTMLElement {
  const span = document.createElement('span')
  span.className = inferred ? 'story-marker story-marker-inferred' : 'story-marker'
  span.textContent = marker
  span.contentEditable = 'false'
  return span
}

function buildStoryDecorations(doc: ProseMirrorNode, styles?: StoryStyles): DecorationSet {
  const decos: Decoration[] = []
  const structural = new Map(
    collectStructuralHeadings(doc, styles).map((ref) => [ref.pos, ref] as const),
  )
  doc.forEach((node, offset) => {
    const inferred = structural.get(offset)
    if (inferred) {
      decos.push(
        Decoration.node(offset, offset + node.nodeSize, {
          class: `${storyHeadingClass(inferred.level)} story-heading-inferred`,
        }),
      )
      decos.push(
        Decoration.widget(offset + 1, () => markerDom(structuralMarker(inferred.level), true), {
          key: `story-marker:${inferred.level}:inferred`,
        }),
      )
      return
    }
    const level = storyHeadingLevel(node, styles)
    if (level != null) {
      decos.push(
        Decoration.node(offset, offset + node.nodeSize, { class: storyHeadingClass(level) }),
      )
    }
    const marker = storyMarkerFor(node, styles)
    if (!marker) return
    // the key carries the marker's text: widgets with the same pos + key keep
    // their DOM, so a plain key would never swap '#' for '>'
    decos.push(
      Decoration.widget(offset + 1, () => markerDom(marker), { key: `story-marker:${marker}` }),
    )
  })
  return DecorationSet.create(doc, decos)
}

/**
 * Text typed in the view takes the paragraph style's formatting: the hidden run
 * mark is stripped from the typed text alone.
 *
 * Only what the user typed is touched — a paste keeps the formatting it
 * carries, and the text around the insertion keeps its own.
 */
export function stripHiddenMarksFromTypedText(
  transactions: readonly Transaction[],
  newState: EditorState,
): Transaction | null {
  const markType = newState.schema.marks[STORY_HIDDEN_MARK]
  if (!markType) return null
  // The mark removal on freshly typed text belongs to the insertion, not to a
  // formatting change: with Track Changes on, an untagged transaction would
  // have the plugin record a formatting edit on top of the insertion, and every
  // typed character would show a "Formatted" bubble in the review pane.
  const tr = newState.tr.setMeta(TRACK_IGNORE, true)
  let changed = false
  for (const t of transactions) {
    if (!t.docChanged) continue
    for (const [i, step] of t.steps.entries()) {
      if (!(step instanceof ReplaceStep)) continue
      // a pure insertion is typing; a replacement carries the text it replaced
      if (step.from !== step.to) continue
      const size = step.slice.size
      if (size <= 0) continue
      // the position in the new doc: the step's own, mapped over the steps
      // that follow it
      const from = t.mapping.slice(i).map(step.from, -1)
      tr.removeMark(from, from + size, markType)
      changed = true
    }
  }
  return changed ? tr : null
}

interface StoryViewMeta {
  active?: boolean
}

export const StoryViewExtension = Extension.create({
  name: 'docsStoryView',

  addCommands() {
    return {
      setStoryViewActive:
        (active: boolean) =>
        ({ state, dispatch }) => {
          if (storyViewKey.getState(state)?.active === active) return true
          if (dispatch) dispatch(state.tr.setMeta(storyViewKey, { active } satisfies StoryViewMeta))
          return true
        },
    }
  },

  addProseMirrorPlugins() {
    const editor = this.editor
    const styles = (): StoryStyles | undefined => editor.storage.listNumbering?.styles
    // The next dispatched transaction is the typed insertion. ProseMirror calls
    // handleTextInput right before it, and only real typed text reaches it — a
    // paste, a drop and an AI edit do not (the app's own convention, see
    // autocorrect.ts). The flag is consumed by the very next transaction.
    let typed = false
    return [
      new Plugin({
        key: storyViewKey,
        state: {
          init: () => ({ active: false, decorations: DecorationSet.empty }),
          apply(tr, value) {
            const meta = tr.getMeta(storyViewKey) as StoryViewMeta | undefined
            const active = meta?.active ?? value.active
            if (active === value.active && !tr.docChanged) return value
            return {
              active,
              decorations: active ? buildStoryDecorations(tr.doc, styles()) : DecorationSet.empty,
            }
          },
        },
        props: {
          decorations: (state: EditorState) =>
            storyViewKey.getState(state)?.decorations ?? DecorationSet.empty,
          handleTextInput(view, _from, _to, _text) {
            typed = storyViewKey.getState(view.state)?.active === true
            // the default insertion still runs; the mark is stripped after it
            return false
          },
        },
        appendTransaction: (transactions, _oldState, newState) => {
          if (!typed) return null
          typed = false
          if (!storyViewKey.getState(newState)?.active) return null
          return stripHiddenMarksFromTypedText(transactions, newState)
        },
      }),
    ]
  },
})
