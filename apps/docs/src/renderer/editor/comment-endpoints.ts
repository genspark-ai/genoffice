import { Extension } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import type { AnchorBlock } from './margin-annotations'

interface Endpoints {
  source?: AnchorBlock[]
  positions: Map<string, number>
}

const key = new PluginKey<Endpoints>('commentEndpoints')

function resolveEndpoints(doc: PmNode, positions: Map<string, number>): void {
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return
    const offsets = node.attrs.commentEndOffsets as Record<string, number> | null
    for (const [id, offset] of Object.entries(offsets ?? {})) {
      if (!positions.has(id)) positions.set(id, pos + 1 + offset)
    }
    // An empty paragraph has no run to carry boundary metadata. Its only
    // possible text endpoint is its content start, exactly offset zero.
    if (node.content.size === 0) {
      for (const id of (node.attrs.commentEnds as string[] | null) ?? []) {
        if (!positions.has(id)) positions.set(id, pos + 1)
      }
    }
  })
}

/** End offsets come from the actual Run-to-inline conversion. Transactions
 * then move screen-only positions without altering saved document markers. */
export const CommentEndpointsExtension = Extension.create({
  name: 'commentEndpoints',
  addProseMirrorPlugins() {
    return [
      new Plugin<Endpoints>({
        key,
        state: {
          init: (_, state) => {
            const positions = new Map<string, number>()
            resolveEndpoints(state.doc, positions)
            return { positions }
          },
          apply(tr, previous) {
            const seeded = tr.getMeta(key) as Endpoints | undefined
            if (seeded) return seeded
            if (!tr.docChanged) return previous
            const positions = new Map(
              [...previous.positions].map(([id, pos]) => [id, tr.mapping.map(pos, -1)]),
            )
            // Later paragraphs of phased imports may only have arrived now.
            resolveEndpoints(tr.doc, positions)
            return { ...previous, positions }
          },
        },
      }),
    ]
  },
})

/** Reset on document replacement; layout refreshes retain mapped positions. */
export function commentEndpoints(
  view: EditorView,
  source: AnchorBlock[] | undefined,
): Map<string, number> {
  const previous = key.getState(view.state)
  if (!previous || previous.source === source) return previous?.positions ?? new Map()
  const positions = new Map<string, number>()
  resolveEndpoints(view.state.doc, positions)
  view.dispatch(view.state.tr.setMeta(key, { source, positions }).setMeta('addToHistory', false))
  return positions
}
