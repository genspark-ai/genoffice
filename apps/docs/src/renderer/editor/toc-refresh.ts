import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'

import type { HeadingRef } from './headings'

/**
 * Backfill tocLine fieldDisplay pages from freshly measured heading pages.
 *
 * `displays` is aligned with `headings` (document order): one formatted page
 * string per heading, or undefined when the heading was not measured
 * (unmeasured entries keep their tocLine untouched instead of shifting later
 * pages up). The caller formats via the owning section's pgNumType, so
 * Roman/letter/dashed numbers survive the refresh. Duplicate titles keep
 * their own entries — the Nth tocLine with a title takes the Nth same-titled
 * heading's page (extras reuse the last one, matching a stale TOC with more
 * lines than headings).
 *
 * Returns true when `tr` gained at least one node update.
 */
export function applyTocPageDisplays(
  doc: PmNode,
  tr: Transaction,
  headings: HeadingRef[],
  displays: Array<string | undefined>,
): boolean {
  const pagesOfTitle = new Map<string, Array<string | undefined>>()
  headings.forEach((h, i) => {
    const key = h.text.trim()
    const list = pagesOfTitle.get(key)
    if (list) list.push(displays[i])
    else pagesOfTitle.set(key, [displays[i]])
  })
  if (pagesOfTitle.size === 0) return false
  let changed = false
  const seen = new Map<string, number>()
  doc.forEach((node, offset) => {
    if (node.type.name !== 'docProtected') return
    const field = node.attrs.fieldDisplay as {
      kind?: string
      left?: string
      right?: string
    } | null
    if (field?.kind !== 'tocLine') return
    const key = (field.left ?? '').trim()
    const list = pagesOfTitle.get(key)
    if (!list) return
    const nth = seen.get(key) ?? 0
    seen.set(key, nth + 1)
    const right = list[Math.min(nth, list.length - 1)]
    if (right === undefined || (field.right ?? '') === right) return
    tr.setNodeMarkup(offset, undefined, {
      ...node.attrs,
      fieldDisplay: { ...field, right },
    })
    changed = true
  })
  return changed
}
