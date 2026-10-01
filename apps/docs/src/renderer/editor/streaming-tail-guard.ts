import { Extension } from '@tiptap/core'
import { Plugin, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { PHASED_CONTENT_SETTLED_EVENT, isPhasedContentPending } from '../phased-content'

/** set on the transactions that append a streamed chunk */
export const PHASED_APPEND = 'phasedAppend'

/**
 * A phased open streams the document tail in behind the first screens. The
 * editor stays editable meanwhile; only the append boundary is off limits:
 * an edit inside or after the last mounted block would land in front of the
 * blocks still to come. Saves already wait for the full content.
 *
 * Refusing that position must not swallow the keystroke with it. A filtered
 * transaction leaves nothing behind — no state change, no undo step, nothing
 * on screen — so a character typed at the boundary simply disappeared and
 * Ctrl+Z had nothing to bring back. Typed text is therefore held here until
 * the tail has landed and then inserted as an ordinary, undoable edit at the
 * position it was typed. Everything else still goes through the guard as
 * before.
 */
export function touchesStreamingTail(tr: Transaction): boolean {
  if (!tr.docChanged || tr.getMeta(PHASED_APPEND)) return false
  const last = tr.before.lastChild
  const tailStart = tr.before.content.size - (last?.nodeSize ?? 0)
  let touches = false
  // only the first step's ranges are in `before` coordinates; a multi-step
  // transaction is judged on its earliest step (later steps build on it)
  tr.mapping.maps[0]?.forEach((_oldStart, oldEnd) => {
    if (oldEnd >= tailStart) touches = true
  })
  if (tr.mapping.maps.length === 0 && tr.steps.length > 0) touches = true
  return touches
}

/** text typed at the boundary while the tail was still streaming */
interface HeldText {
  /** insertion point, remapped as later edits shift the document */
  at: number
  text: string
}

/**
 * True when a step starts at 0, which is how a whole-document replace reaches
 * the model — a remount after a refused chunk, or a different file. Typed text
 * cannot survive one of those, and replaying it would paste it into whatever
 * document took its place.
 */
function replacesWholeDocument(tr: Transaction): boolean {
  return tr.steps.some((step) => {
    const s = step as { from?: unknown }
    return typeof s.from === 'number' && s.from === 0
  })
}

export const StreamingTailGuardExtension = Extension.create({
  name: 'streamingTailGuard',
  addProseMirrorPlugins() {
    // per editor: held text belongs to one document, never to two
    let held: HeldText | null = null
    const insertHeld = (view: EditorView): void => {
      const pending = held
      held = null
      if (!pending || view.isDestroyed) return
      // the stream only ever added blocks after this point, so the recorded
      // position still addresses the same text; clamp for a short document
      const at = Math.min(pending.at, view.state.doc.content.size)
      view.dispatch(view.state.tr.insertText(pending.text, at))
    }
    return [
      new Plugin({
        filterTransaction: (tr) => !isPhasedContentPending() || !touchesStreamingTail(tr),
        appendTransaction: (trs) => {
          if (!held) return null
          for (const tr of trs) {
            if (!tr.docChanged) continue
            // the document this text was typed into is gone: drop it rather
            // than paste it into whatever replaced it
            if (replacesWholeDocument(tr)) {
              held = null
              return null
            }
            // otherwise follow the text the caret was on as an edit shifts it
            held.at = tr.mapping.map(held.at)
          }
          return null
        },
        props: {
          handleTextInput: (view, from, to, text) => {
            if (!isPhasedContentPending()) {
              // a cancelled stream never announces itself, so release what it
              // held the next time the user types in this editor
              insertHeld(view)
              return false
            }
            // a keystroke the guard would allow takes the normal path
            if (!touchesStreamingTail(view.state.tr.insertText(text, from, to))) return false
            // a selection replace is a deletion: replaying it once the tail has
            // landed could cut streamed blocks out, so it stays refused
            if (from !== to) return false
            // consecutive characters share a position, so they read as one word
            held = held?.at === from ? { at: from, text: held.text + text } : { at: from, text }
            return true
          },
        },
        view: (view) => {
          const onSettled = (): void => insertHeld(view)
          document.addEventListener(PHASED_CONTENT_SETTLED_EVENT, onSettled)
          return {
            destroy: () => document.removeEventListener(PHASED_CONTENT_SETTLED_EVENT, onSettled),
          }
        },
      }),
    ]
  },
})
