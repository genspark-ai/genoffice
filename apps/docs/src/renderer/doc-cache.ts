import type { Node as PmNode } from '@tiptap/pm/model'

/** Memoize an O(doc) derivation by PM doc reference (docs are immutable, so same ref ⇒ same content). */
export function cachedByDoc<T>(compute: (doc: PmNode) => T): (doc: PmNode) => T {
  const cache = new WeakMap<PmNode, { value: T }>()
  return (doc) => {
    const hit = cache.get(doc)
    if (hit) return hit.value
    const value = compute(doc)
    cache.set(doc, { value })
    return value
  }
}
