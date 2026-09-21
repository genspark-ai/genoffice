import { Extension } from '@tiptap/core'
import { Plugin, type Transaction } from '@tiptap/pm/state'
import { isPhasedContentPending } from '../phased-content'

/** set on the transactions that append a streamed chunk */
export const PHASED_APPEND = 'phasedAppend'

/**
 * A phased open streams the document tail in behind the first screens. The
 * editor stays editable meanwhile; only the append boundary is off limits:
 * an edit inside or after the last mounted block would land in front of the
 * blocks still to come. Saves already wait for the full content.
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

export const StreamingTailGuardExtension = Extension.create({
  name: 'streamingTailGuard',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        filterTransaction: (tr) => !isPhasedContentPending() || !touchesStreamingTail(tr),
      }),
    ]
  },
})
