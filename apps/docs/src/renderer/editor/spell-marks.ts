/**
 * Misspelled words shown on the page as Word shows them: a red wavy line. The
 * browser's own markers only appear beside typing, so the words a file came
 * with had none; these are drawn by us with the CSS Custom Highlight API,
 * which leaves the page's DOM, the saved file and the layout alone.
 */
import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { pageHighlight } from './page-highlight'
import { spellingDismissed, spellWordAt } from './spell-menu'
import { wordsOfBlock, type SpellWord } from './spell-walk'

export interface SpellMarksState {
  active: boolean
  /** the word being typed: it is checked once the caret leaves it */
  typing: { from: number; to: number } | null
  /** bumped when the dictionary changed and the page is to be read again */
  epoch: number
}

type SpellMarksMeta = { active: boolean } | { refresh: true }

export const spellMarksKey = new PluginKey<SpellMarksState>('spellMarks')
export const SPELL_HIGHLIGHT = 'docs-spelling'
/** how far past the window the marks are kept up, in document positions */
const AROUND = 6000
const MOST = 2000

const verdicts = new Map<string, boolean>()
const lookedUp = (word: string): boolean => {
  let wrong = verdicts.get(word)
  if (wrong === undefined) {
    wrong = window.desktop?.spellIsMisspelled?.(word) ?? false
    verdicts.set(word, wrong)
  }
  return wrong
}

type BlockWord = { at: number; length: number; word: string }
const wordsByBlock = new WeakMap<PmNode, BlockWord[]>()

function blockWords(block: PmNode): BlockWord[] {
  let words = wordsByBlock.get(block)
  if (!words) {
    const found: SpellWord[] = []
    wordsOfBlock(block, 0, found)
    words = found.map((w) => ({ at: w.from, length: w.to - w.from, word: w.word }))
    wordsByBlock.set(block, words)
  }
  return words
}

/** the words between two positions that get a line: misspelled, not set aside, not being typed */
export function spellMarkWords(
  state: EditorState,
  from: number,
  to: number,
  misspelled: (word: string) => boolean = lookedUp,
): SpellWord[] {
  const marks = spellMarksKey.getState(state)
  if (!marks?.active) return []
  const out: SpellWord[] = []
  const end = Math.min(to, state.doc.content.size)
  state.doc.nodesBetween(Math.max(from, 0), end, (node, pos) => {
    if (!node.isTextblock) return out.length < MOST
    for (const w of blockWords(node)) {
      const word = { from: pos + w.at, to: pos + w.at + w.length, word: w.word }
      if (word.to < from || word.from > end || out.length >= MOST) continue
      const typing = marks.typing
      if (typing && word.from <= typing.to && word.to >= typing.from) continue
      if (misspelled(word.word) && !spellingDismissed(state, word)) out.push(word)
    }
    return false
  })
  return out
}

/** whether everything the change put in lies inside that word, as typing and deleting in it do; a paste of several words, a rewrite of the document, formatting and edits elsewhere do not */
function editedInside(tr: Transaction, word: { from: number; to: number }): boolean {
  let touched = false
  let inside = true
  tr.mapping.maps.forEach((map, i) => {
    const rest = tr.mapping.slice(i + 1)
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      touched = true
      if (rest.map(newStart, -1) < word.from || rest.map(newEnd, 1) > word.to) inside = false
    })
  })
  return touched && inside
}

function apply(tr: Transaction, prev: SpellMarksState): SpellMarksState {
  const meta = tr.getMeta(spellMarksKey) as SpellMarksMeta | undefined
  let next = prev
  if (meta && 'active' in meta) next = { ...next, active: meta.active }
  if (meta && 'refresh' in meta) next = { ...next, epoch: next.epoch + 1 }
  const { empty, head } = tr.selection
  let typing = next.typing
  const word = tr.docChanged && empty ? spellWordAt(tr.doc, head) : null
  if (word && editedInside(tr, word)) {
    typing = { from: word.from, to: word.to }
  } else if (typing) {
    // what is put in beside the word, such as the space that ends it, is not part of it
    const from = tr.mapping.map(typing.from, 1)
    const to = tr.mapping.map(typing.to, -1)
    const left = !empty || head < from || head > to
    // a break put inside the word leaves two words: neither is the one that was being typed
    const held = left || !tr.docChanged ? null : spellWordAt(tr.doc, from)
    const whole = !tr.docChanged || (held?.from === from && held.to === to)
    typing =
      left || !whole ? null : from === typing.from && to === typing.to ? typing : { from, to }
  }
  return typing === next.typing ? next : { ...next, typing }
}

/** the part of the page the scrolling pane shows, in window coordinates */
function shownBox(view: EditorView): { top: number; bottom: number; left: number; right: number } {
  const page = view.dom.getBoundingClientRect()
  let top = Math.max(page.top, 0)
  let bottom = Math.min(page.bottom, window.innerHeight)
  for (let el = view.dom.parentElement; el; el = el.parentElement) {
    const flow = getComputedStyle(el).overflowY
    if (flow !== 'auto' && flow !== 'scroll') continue
    const pane = el.getBoundingClientRect()
    top = Math.max(top, pane.top)
    bottom = Math.min(bottom, pane.bottom)
  }
  return { top, bottom, left: page.left, right: page.right }
}

/** the first position found walking from `y` in steps of `step`: a gap between pages holds none */
function posNear(view: EditorView, box: ReturnType<typeof shownBox>, y: number, step: number) {
  const width = box.right - box.left
  for (let at = y; at > box.top && at < box.bottom; at += step)
    for (const part of [0.5, 0.25, 0.75]) {
      const found = view.posAtCoords({ left: box.left + width * part, top: at })
      if (found) return found.pos
    }
  return undefined
}

const lastSpan = new WeakMap<EditorView, { from: number; to: number }>()

/** the positions the window shows, with room around them; where that cannot be told, what it showed last */
function shownSpan(view: EditorView): { from: number; to: number } {
  const fallback = lastSpan.get(view) ?? { from: 0, to: view.state.doc.content.size }
  try {
    const box = shownBox(view)
    const first = posNear(view, box, box.top + 2, 24)
    const last = posNear(view, box, box.bottom - 2, -24)
    if (first === undefined || last === undefined) return fallback
    const span = { from: Math.min(first, last) - AROUND, to: Math.max(first, last) + AROUND }
    lastSpan.set(view, span)
    return span
  } catch {
    return fallback
  }
}

function drawMarks(view: EditorView, drawn: AbstractRange[]): AbstractRange[] {
  const highlight = pageHighlight(SPELL_HIGHLIGHT)
  if (!highlight) return []
  for (const range of drawn) highlight.delete(range)
  if (!spellMarksKey.getState(view.state)?.active) return []
  const span = shownSpan(view)
  const next: AbstractRange[] = []
  for (const word of spellMarkWords(view.state, span.from, span.to)) {
    try {
      const start = view.domAtPos(word.from, 1)
      const end = view.domAtPos(word.to, -1)
      const range = document.createRange()
      range.setStart(start.node, start.offset)
      range.setEnd(end.node, end.offset)
      highlight.add(range)
      next.push(range)
    } catch {
      // a position the view has no DOM for yet is drawn on the next update
    }
  }
  return next
}

export const SpellMarksExtension = Extension.create({
  name: 'spellMarks',

  addProseMirrorPlugins() {
    return [
      new Plugin<SpellMarksState>({
        key: spellMarksKey,
        state: { init: () => ({ active: false, typing: null, epoch: 0 }), apply },
        view(view) {
          let drawn: AbstractRange[] = []
          let frame = 0
          let live = view
          const redraw = () => {
            frame = 0
            drawn = drawMarks(live, drawn)
          }
          // scrolling brings other words into the window
          const later = () => {
            if (!frame && (drawn.length > 0 || spellMarksKey.getState(live.state)?.active))
              frame = requestAnimationFrame(redraw)
          }
          document.addEventListener('scroll', later, true)
          window.addEventListener('resize', later)
          return {
            update(updated) {
              live = updated
              // the view re-renders text nodes without a state change of ours, so the ranges are rebuilt each time
              if (drawn.length > 0 || spellMarksKey.getState(updated.state)?.active)
                drawn = drawMarks(updated, drawn)
            },
            destroy() {
              document.removeEventListener('scroll', later, true)
              window.removeEventListener('resize', later)
              if (frame) cancelAnimationFrame(frame)
              const highlight = pageHighlight(SPELL_HIGHLIGHT)
              for (const range of drawn) highlight?.delete(range)
            },
          }
        },
      }),
    ]
  },
})

const send = (editor: Editor, meta: SpellMarksMeta) =>
  editor.view.dispatch(editor.state.tr.setMeta(spellMarksKey, meta).setMeta('addToHistory', false))

/** Review > Spelling: the lines show or do not */
export function setSpellMarksActive(editor: Editor, active: boolean): void {
  if (spellMarksKey.getState(editor.state)?.active !== active) send(editor, { active })
}

/** the dictionary changed (a word learned, ignored or reset): every word is looked up again */
export function refreshSpellMarks(editor: Editor): void {
  verdicts.clear()
  send(editor, { refresh: true })
}
