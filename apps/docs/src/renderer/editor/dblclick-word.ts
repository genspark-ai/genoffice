import { Extension } from '@tiptap/core'
import { Plugin, TextSelection } from '@tiptap/pm/state'

/**
 * Word's double-click unit is the word plus every space after it: double-clicking
 * inside a run of spaces selects the preceding word (or punctuation mark) together
 * with the whole run, so the spaces are never selected on their own. Chromium's
 * native double-click selects only the spaces, which in a fill-in-the-blank line
 * of spaces leaves an invisible selection that the next keystroke replaces.
 */

const LEAF = '\ufffc'
let segmenter: Intl.Segmenter | null = null
function segments(text: string): Array<{ index: number; segment: string }> {
  segmenter ??= new Intl.Segmenter(undefined, { granularity: 'word' })
  return [...segmenter.segment(text)]
}

const isSpaceRun = (s: string) => /^ +$/.test(s)
const isWordHead = (s: string) => !isSpaceRun(s) && s !== LEAF && !/^[\n\t]+$/.test(s)

/** [from, to) of the Word word unit containing text[index], or null for a leaf stand-in */
export function wordUnitAt(text: string, index: number): [number, number] | null {
  if (index < 0 || index >= text.length || text[index] === LEAF) return null
  const segs = segments(text)
  let i = segs.findIndex((s, k) => index < s.index + s.segment.length || k === segs.length - 1)
  if (i < 0) return null
  if (isSpaceRun(segs[i].segment) && i > 0 && isWordHead(segs[i - 1].segment)) i -= 1
  const head = segs[i]
  if (!isWordHead(head.segment) && !isSpaceRun(head.segment)) return null
  let end = head.index + head.segment.length
  while (i + 1 < segs.length && isSpaceRun(segs[i + 1].segment)) {
    i += 1
    end = segs[i].index + segs[i].segment.length
  }
  return [head.index, end]
}

export const DoubleClickWordExtension = Extension.create({
  name: 'doubleClickWord',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          handleDoubleClick(view, pos, event) {
            // ProseMirror also reports two quick nearby clicks as a double click;
            // only the browser's own count is a real one
            if (event.detail < 2 || event.button !== 0) return false
            if (event.shiftKey || event.altKey || event.metaKey) return false
            const $pos = view.state.doc.resolve(pos)
            if (!$pos.parent.isTextblock) return false
            const start = $pos.start()
            const text = $pos.parent.textBetween(0, $pos.parent.content.size, LEAF, LEAF)
            let index = $pos.parentOffset
            try {
              if (event.clientX < view.coordsAtPos(pos).left) index -= 1
            } catch {
              /* no layout yet: take the character after the slot */
            }
            const unit = wordUnitAt(text, Math.min(index, text.length - 1))
            if (!unit) return false
            view.focus()
            view.dispatch(
              view.state.tr.setSelection(
                TextSelection.create(view.state.doc, start + unit[0], start + unit[1]),
              ),
            )
            return true
          },
        },
      }),
    ]
  },
})
