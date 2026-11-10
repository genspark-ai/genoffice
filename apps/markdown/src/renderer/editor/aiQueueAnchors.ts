import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

/**
 * Anchors for queued AI edits, kept as inline decorations rather than marks:
 * DecorationSet.map() migrates them through every edit (character-precise,
 * undo-safe), they never enter the document model, so saving mid-queue cannot
 * dirty untouched blocks or leak anything into the serialized markdown. An
 * anchor whose text is deleted collapses and drops out of the set — the queue
 * reads that back as "target gone". (Same design as the docs app.)
 */

interface AnchorMeta {
  add?: { qid: string; from: number; to: number }
  remove?: string[]
  clear?: boolean
}

export const aiQueueAnchorsKey = new PluginKey<DecorationSet>('aiQueueAnchors')

/** fired on clicking an anchored range in the document; detail = { qid } */
export const AI_QUEUE_ANCHOR_CLICK = 'markdown-ai-queue-anchor-click'

export const AiQueueAnchors = Extension.create({
  name: 'aiQueueAnchors',

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: aiQueueAnchorsKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            let set = old.map(tr.mapping, tr.doc)
            const meta = tr.getMeta(aiQueueAnchorsKey) as AnchorMeta | undefined
            if (!meta) return set
            if (meta.clear) return DecorationSet.empty
            if (meta.remove) {
              set = set.remove(
                set.find(undefined, undefined, (spec) =>
                  meta.remove!.includes((spec as { qid?: string }).qid ?? ''),
                ),
              )
            }
            if (meta.add && meta.add.from < meta.add.to) {
              set = set.add(tr.doc, [
                Decoration.inline(
                  meta.add.from,
                  meta.add.to,
                  { class: 'ai-queue-anchor' },
                  { qid: meta.add.qid },
                ),
              ])
            }
            return set
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
          handleClick(view, pos) {
            const hit = aiQueueAnchorsKey.getState(view.state)?.find(pos, pos) ?? []
            const qid = (hit[0]?.spec as { qid?: string } | undefined)?.qid
            if (qid) {
              window.dispatchEvent(new CustomEvent(AI_QUEUE_ANCHOR_CLICK, { detail: { qid } }))
            }
            return false // never consume: the caret still lands where the user clicked
          },
        },
      }),
    ]
  },
})

const anchorMeta = (editor: Editor, meta: AnchorMeta): void => {
  editor.view.dispatch(
    editor.state.tr.setMeta(aiQueueAnchorsKey, meta).setMeta('addToHistory', false),
  )
}

export function addQueueAnchor(editor: Editor, qid: string, from: number, to: number): void {
  anchorMeta(editor, { add: { qid, from, to } })
}

export function removeQueueAnchors(editor: Editor, qids: string[]): void {
  if (qids.length > 0) anchorMeta(editor, { remove: qids })
}

export function clearQueueAnchors(editor: Editor): void {
  anchorMeta(editor, { clear: true })
}

/** current positions of an anchor; null once its text was deleted */
export function queueAnchorRange(
  state: EditorState,
  qid: string,
): { from: number; to: number } | null {
  const hits =
    aiQueueAnchorsKey
      .getState(state)
      ?.find(undefined, undefined, (spec) => (spec as { qid?: string }).qid === qid) ?? []
  const deco = hits[0]
  return deco && deco.from < deco.to ? { from: deco.from, to: deco.to } : null
}
