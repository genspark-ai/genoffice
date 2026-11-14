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

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    docsStoryView: {
      /** draw the story-view margin markers (true) or clear them */
      setStoryViewActive: (active: boolean) => ReturnType
      /** Focus mode: dim everything but the sentence the caret sits in */
      setStoryViewFocus: (on: boolean) => ReturnType
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
  focus: boolean
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

function markerDom(marker: string): HTMLElement {
  const span = document.createElement('span')
  span.className = 'story-marker'
  span.textContent = marker
  span.contentEditable = 'false'
  return span
}

function buildStoryDecorations(doc: ProseMirrorNode, styles?: StoryStyles): Decoration[] {
  const decos: Decoration[] = []
  doc.forEach((node, offset) => {
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
  return decos
}

/** the characters that end a sentence for Focus mode; the CJK terminators read
    the same way their Latin counterparts do */
const SENTENCE_END = /[.!?。！？…]/u
/** the Latin terminators, which also occur inside words and numbers */
const SENTENCE_LATIN_END = /[.!?]/u
/** a closing quote or bracket belongs to the sentence it closes */
const SENTENCE_CLOSER = /["'\u201d\u2019\u300d\u300f\uff09)}\]]/u

/**
 * Whether the terminator at `at` really ends a sentence. A Latin one ends it
 * only where the text actually stops — whitespace, a closing quote/bracket, or
 * the end of the block — so "3.14", "example.com" and "v2.0.1" keep their
 * dots; the CJK terminators never sit inside a Latin run, so they always end
 * one.
 */
function sentenceStopsAt(text: string, at: number): boolean {
  if (!SENTENCE_LATIN_END.test(text[at]!)) return true
  if (at + 1 >= text.length) return true
  if (/\s/u.test(text[at + 1]!)) return true
  if (SENTENCE_CLOSER.test(text[at + 1]!)) return at + 2 >= text.length || /\s/u.test(text[at + 2]!)
  return false
}

/** One past the sentence a terminator at `at` ends: past a closing quote or
    bracket that carries the stop ("Stop." — the quote stays with the
    sentence), then past the trailing whitespace that forms the gap between
    the sentences. */
function sentenceEndAfter(text: string, at: number): number {
  let j = at + 1
  if (j < text.length && SENTENCE_CLOSER.test(text[j]!)) j++
  while (j < text.length && /\s/u.test(text[j]!)) j++
  return j
}

/**
 * The sentence the caret sits in, as ranges into the block's text: from after
 * the previous terminator's trailing whitespace to the next terminator's own
 * trailing whitespace. A block without a terminator reads as one sentence, and
 * the ranges always cover the whole block.
 */
export function storyFocusSentence(text: string, cursor: number): [number, number] {
  let start = 0
  for (let i = 0; i < cursor;) {
    if (!SENTENCE_END.test(text[i]!) || !sentenceStopsAt(text, i)) {
      i++
      continue
    }
    // the terminator's trailing whitespace belongs to the gap between the
    // sentences; the next sentence starts after it. A caret inside that gap is
    // about to write the next one — that one stays visible.
    i = start = sentenceEndAfter(text, i)
  }
  let end = text.length
  for (let k = cursor; k < text.length; k++) {
    if (!SENTENCE_END.test(text[k]!) || !sentenceStopsAt(text, k)) continue
    end = sentenceEndAfter(text, k)
    break
  }
  return [start, end]
}

/**
 * Whether an Escape reaches the story view's own "leave" action. An Escape
 * meant for something else never claims the view: a composing IME, an event a
 * transient overlay already consumed (the paste chip marks its Escape), one
 * aimed inside the find panel (it closes itself), or one landing while an
 * overlay the app knows about is open — the context menu, the comment popup
 * and the dialogs each close on their own Escape.
 */
export function storyEscLeavesView(e: KeyboardEvent, overlayOpen: boolean): boolean {
  if (e.key !== 'Escape') return false
  if (e.defaultPrevented || e.isComposing) return false
  if (overlayOpen) return false
  if (e.target instanceof Element && e.target.closest('.find-panel')) return false
  return true
}

/**
 * The dim ranges Focus mode paints: every text run except the sentence the
 * caret sits in. Inline decorations, so the dim rides the text itself — a
 * paragraph the caret is not in dims whole. Empty blocks dim nothing (an empty
 * paragraph is always the writing spot), and the active block keeps itself
 * visible entirely when it holds an inline object: a text index is not a
 * position there, so the sentence ranges could not be drawn safely.
 */
function buildFocusDim(doc: ProseMirrorNode, head: number): Decoration[] {
  const decos: Decoration[] = []
  doc.forEach((node, offset) => {
    if (!node.isTextblock || node.content.size === 0) return
    const start = offset + 1
    const end = offset + node.nodeSize - 1
    // a boundary position (caret between blocks) belongs to both blocks it
    // sits against: pos == this block's start, pos == this block's end + 1 —
    // each keeps its sentence visible around the caret
    if (head < offset || head > offset + node.nodeSize) {
      decos.push(Decoration.inline(start, end, { class: 'story-dim' }))
      return
    }
    if (node.content.size !== node.textContent.length) return
    const cursor = Math.min(Math.max(head - start, 0), node.textContent.length)
    const [sFrom, sTo] = storyFocusSentence(node.textContent, cursor)
    if (sFrom > 0) decos.push(Decoration.inline(start, start + sFrom, { class: 'story-dim' }))
    if (sTo < node.textContent.length)
      decos.push(Decoration.inline(start + sTo, end, { class: 'story-dim' }))
  })
  return decos
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
  focus?: boolean
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
      setStoryViewFocus:
        (on: boolean) =>
        ({ state, dispatch }) => {
          if (storyViewKey.getState(state)?.focus === on) return true
          if (dispatch)
            dispatch(state.tr.setMeta(storyViewKey, { focus: on } satisfies StoryViewMeta))
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
          init: () => ({ active: false, focus: false, decorations: DecorationSet.empty }),
          apply(tr, value) {
            const meta = tr.getMeta(storyViewKey) as StoryViewMeta | undefined
            const active = meta?.active ?? value.active
            const focus = meta?.focus ?? value.focus
            // focus rides the caret: a selection-only transaction repaints the
            // dim ranges too
            if (
              active === value.active &&
              focus === value.focus &&
              !tr.docChanged &&
              !(active && focus && tr.selectionSet)
            )
              return value
            return {
              active,
              focus,
              decorations: active
                ? DecorationSet.create(tr.doc, [
                    ...buildStoryDecorations(tr.doc, styles()),
                    ...(focus ? buildFocusDim(tr.doc, tr.selection.head) : []),
                  ])
                : DecorationSet.empty,
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
