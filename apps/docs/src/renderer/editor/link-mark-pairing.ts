import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'

/** the marker where a link field (INCLUDETEXT, INCLUDEPICTURE, LINK) opens */
export const isLinkBegin = (node: PmNode): boolean =>
  node.type.name === 'docXeMark' && typeof node.attrs.linkBegin === 'string'
/** the marker where the link field opened before it closes */
export const isLinkEnd = (node: PmNode): boolean =>
  node.type.name === 'docXeMark' && node.attrs.linkEnd === true
const isBegin = isLinkBegin
const isEnd = isLinkEnd

const linkMarksKey = new PluginKey<number>('linkMarks')

const countMarks = (doc: PmNode): number => {
  let count = 0
  doc.descendants((node) => {
    if (isBegin(node) || isEnd(node)) count++
    return true
  })
  return count
}

/**
 * A link field needs both of its markers: when an edit takes one away (the
 * paragraph that held it was deleted, half of a field was pasted) the other
 * goes with it, and the content stays as plain content.
 */
export function linkMarkPairing(): Plugin<number> {
  return new Plugin<number>({
    key: linkMarksKey,
    state: {
      init: (_, state) => countMarks(state.doc),
      apply(tr, count) {
        if (!tr.docChanged) return count
        // a document that had none can only gain one through content that carries it
        if (count === 0) {
          const brings = tr.steps.some((step) => {
            const slice = (step as { slice?: { content: PmNode['content'] } }).slice
            let found = false
            slice?.content.descendants((node) => {
              if (isBegin(node) || isEnd(node)) found = true
              return !found
            })
            return found
          })
          if (!brings) return 0
        }
        return countMarks(tr.doc)
      },
    },
    appendTransaction(trs, oldState, newState) {
      if (!trs.some((tr) => tr.docChanged)) return null
      if (linkMarksKey.getState(newState) === 0) return null
      if (linkMarksKey.getState(oldState) === linkMarksKey.getState(newState)) {
        // the same markers as before unless a step put content in
        const same = trs.every((tr) =>
          tr.steps.every((step) => {
            const slice = (step as { slice?: { size: number } }).slice
            return !slice || slice.size === 0
          }),
        )
        if (same) return null
      }
      // a field that lost one marker to the edit: the other is the loose one, whatever
      // fields stand around it or inside it
      const loose = new Set<number>()
      const survivor = (pos: number, is: (node: PmNode) => boolean): number | null => {
        let gone = false
        for (const tr of trs) {
          const mapped = tr.mapping.mapResult(pos, 1)
          gone ||= mapped.deletedAfter
          pos = mapped.pos
        }
        // a marker whose lock or instruction was set is replaced where it stands
        const now = pos < newState.doc.content.size ? newState.doc.nodeAt(pos) : null
        return gone && !(now && is(now)) ? null : pos
      }
      const before: number[] = []
      oldState.doc.descendants((node, pos) => {
        if (isBegin(node)) before.push(pos)
        else if (isEnd(node)) {
          const begin = before.pop()
          if (begin === undefined) return true
          const ends = [survivor(begin, isBegin), survivor(pos, isEnd)]
          if ((ends[0] === null) !== (ends[1] === null)) loose.add(ends[0] ?? ends[1]!)
        }
        return true
      })
      // and what came in without its other half
      const open: number[] = []
      newState.doc.descendants((node, pos) => {
        if (loose.has(pos)) return true
        if (isBegin(node)) open.push(pos)
        else if (isEnd(node) && open.pop() === undefined) loose.add(pos)
        return true
      })
      for (const pos of open) loose.add(pos)
      if (loose.size === 0) return null
      const tr = newState.tr
      for (const pos of [...loose].sort((a, b) => b - a)) tr.delete(pos, pos + 1)
      return tr
    },
  })
}
