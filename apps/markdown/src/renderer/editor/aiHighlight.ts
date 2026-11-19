import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { AddMarkStep, RemoveMarkStep } from '@tiptap/pm/transform'
import { OP_META, blockIndexRange, blockRange, type OpMeta } from './ops'

const key = new PluginKey<DecorationSet>('aiHighlight')

type Range = { from: number; to: number }

/**
 * Ranges (in tr.doc coordinates) an AI op transaction rewrote, read off its
 * step maps. A structural step (heading conversion, wrap) only touches node
 * boundary tokens where an inline decoration paints nothing, so text-less
 * ranges widen to the top-level block(s) they sit in.
 */
function changedRanges(tr: Transaction): Range[] {
  const ranges: Range[] = []
  const push = (from: number, to: number) => {
    if (to <= from) return
    let inline = false
    tr.doc.nodesBetween(from, to, (node) => {
      if (node.isInline) inline = true
      return !inline
    })
    if (inline) return ranges.push({ from, to })
    const { startIndex, endIndex } = blockIndexRange(tr.doc, from, to)
    ranges.push(blockRange(tr.doc, startIndex, endIndex))
  }
  tr.steps.forEach((step, i) => {
    const rest = tr.mapping.slice(i + 1)
    // mark steps do not move positions, so their step map is empty — use the step's own range
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      push(rest.map(step.from, 1), rest.map(step.to, -1))
      return
    }
    step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      push(rest.map(newStart, 1), rest.map(newEnd, -1))
    })
  })
  return ranges
}

/** Remove every AI-change highlight (called when a run finishes) */
export function clearAiHighlights(editor: Editor): void {
  editor.view.dispatch(editor.state.tr.setMeta('aiHighlightClear', true))
}

/**
 * Transient yellow wash over content the AI just wrote (every transaction an
 * AI-sourced runOps dispatches), so the user can see what changed during a
 * run. Display-only: decorations never touch the document or the serialized
 * markdown.
 */
export const AiHighlight = Extension.create({
  name: 'aiHighlight',

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            if (tr.getMeta('aiHighlightClear')) return DecorationSet.empty
            let next = set.map(tr.mapping, tr.doc)
            const meta = tr.getMeta(OP_META) as OpMeta | undefined
            const add =
              meta?.source === 'ai' && !meta.rollback && tr.docChanged ? changedRanges(tr) : []
            for (const range of add) {
              const from = Math.max(0, Math.min(range.from, tr.doc.content.size))
              const to = Math.max(from, Math.min(range.to, tr.doc.content.size))
              if (to > from) {
                next = next.add(tr.doc, [Decoration.inline(from, to, { class: 'ai-changed' })])
              }
            }
            return next
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
