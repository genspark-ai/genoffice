import { Extension } from '@tiptap/core'
import { Plugin, type Transaction } from '@tiptap/pm/state'
import type { SectionInfo } from '@genoffice/docx-engine'
import { topBlockSections } from './section-endnotes'

/**
 * Forms protection section by section: a document protected for forms keeps
 * the sections whose sectPr switches w:formProt off open for editing, and in
 * the others only the form fields change. The editor stays editable and this
 * guard refuses the edits that reach into a protected section.
 */
export interface SectionEditGuardStorage {
  /** per section: protected; null while no forms protection is in force */
  locked: boolean[] | null
  sections: readonly Pick<SectionInfo, 'lastBlockIndex' | 'sectPrXml' | 'pendingBreak'>[]
  onBlocked: (() => void) | null
}

declare module '@tiptap/core' {
  interface Storage {
    sectionEditGuard: SectionEditGuardStorage
  }
}

/** set on a transaction that is not the user editing (a form field filled, the document shown anew): it goes through in a protected section */
export const SECTION_GUARD_PASS = 'sectionGuardPass'

/** the top-level blocks of `tr.before` a step reaches, as [first, last] */
function touchedBlocks(tr: Transaction, ends: number[]): Array<[number, number]> {
  const out: Array<[number, number]> = []
  tr.steps.forEach((step, i) => {
    const at = step as unknown as { from?: number; to?: number; pos?: number }
    const from = at.from ?? at.pos
    if (typeof from !== 'number') return
    const to = at.to ?? from + 1
    // a later step counts in the document the earlier ones left: back to tr.before
    const back = tr.mapping.slice(0, i).invert()
    const a = back.map(from, -1)
    const b = Math.max(a, back.map(to, 1))
    // a position between two blocks belongs to the one in front, whose section a new block joins
    let first = ends.findIndex((end) => a < end || (a === b && a === end))
    if (first < 0) first = ends.length - 1
    let last = first
    while (last + 1 < ends.length && b > ends[last]) last++
    out.push([first, last])
  })
  return out
}

export function touchesLockedSection(tr: Transaction, storage: SectionEditGuardStorage): boolean {
  const { locked } = storage
  if (!locked?.some(Boolean) || !tr.docChanged) return false
  // layout bookkeeping and what a plugin adds to an edit that was let through
  if (tr.getMeta('addToHistory') === false || tr.getMeta('appendedTransaction')) return false
  if (tr.getMeta(SECTION_GUARD_PASS)) return false
  const { sectionOf } = topBlockSections(tr.before, storage.sections)
  const ends: number[] = []
  let pos = 0
  tr.before.forEach((node) => ends.push((pos += node.nodeSize)))
  if (ends.length === 0) return false
  return touchedBlocks(tr, ends).some(([first, last]) => {
    for (let k = first; k <= last; k++) if (locked[sectionOf[k]] ?? true) return true
    return false
  })
}

export const SectionEditGuardExtension = Extension.create<object, SectionEditGuardStorage>({
  name: 'sectionEditGuard',

  addStorage() {
    return { locked: null, sections: [], onBlocked: null }
  },

  addProseMirrorPlugins() {
    const storage = this.storage
    return [
      new Plugin({
        filterTransaction(tr) {
          if (!touchesLockedSection(tr, storage)) return true
          storage.onBlocked?.()
          return false
        },
      }),
    ]
  },
})
