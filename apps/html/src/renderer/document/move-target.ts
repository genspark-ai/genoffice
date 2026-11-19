import { TABLE_LIST_PARTS } from './insert-presets'
import { childrenOf, type ParseMap } from './parse-map'

const STRUCTURAL = new Set(['html', 'head', 'body'])

/**
 * Where "move up / down" puts the element: past its neighbouring sibling, or — at
 * the first / last position — out of its parent, so repeated presses walk the
 * element through the whole page instead of stopping at a container edge.
 * List and table parts stay inside the parent they require.
 */
export function moveTarget(
  map: ParseMap,
  sid: number,
  dir: -1 | 1,
): { position: 'before' | 'after'; ref_sid: number } | null {
  const e = map.bySid.get(sid)
  if (!e || e.parentSid === null) return null
  const siblings = childrenOf(map, e.parentSid)
  const i = siblings.findIndex((s) => s.sid === sid)
  const neighbour = siblings[i + dir]
  const position = dir < 0 ? 'before' : 'after'
  if (neighbour) return { position, ref_sid: neighbour.sid }
  if (TABLE_LIST_PARTS.has(e.tag)) return null
  const parent = map.bySid.get(e.parentSid)
  if (!parent || STRUCTURAL.has(parent.tag) || parent.parentSid === null) return null
  const grandparent = map.bySid.get(parent.parentSid)
  if (!grandparent || grandparent.tag === 'html' || grandparent.tag === 'head') return null
  return { position, ref_sid: parent.sid }
}
