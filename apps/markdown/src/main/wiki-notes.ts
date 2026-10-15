/**
 * Listing the vault's notes for the `[[` suggester. Resolution
 * (`wiki-resolve.ts`) answers for a name that is already written; this answers
 * for the one being written — the same root, the same skips, walked in the
 * same breadth-first order, so what the suggester offers first is what the
 * resolver would have matched first.
 */
import { basename, extname, join, relative } from 'node:path'

import { MAX_DEPTH, SKIP_DIRS, type WikiFs } from './wiki-resolve'
import type { WikiNoteEntry } from '../shared/ipc'

export type { WikiNoteEntry }

/** entries one listing may return, so a huge library cannot stall the popup */
export const MAX_NOTES = 2000

/** file types the suggester offers; resolution's NOTE_EXTS plus embeddable media */
const SUGGEST_EXTS = new Set([
  '.md',
  '.markdown',
  '.txt',
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
  '.webp',
  '.svg',
  '.bmp',
  '.avif',
  '.ico',
])

export async function listWikiNotes(
  root: string,
  fs: WikiFs,
  cap = MAX_NOTES,
): Promise<WikiNoteEntry[]> {
  const out: WikiNoteEntry[] = []
  if (!root) return out
  let queue: { dir: string; depth: number }[] = [{ dir: root, depth: 0 }]
  let budget = cap * 4
  while (queue.length > 0 && out.length < cap && budget > 0) {
    const next: typeof queue = []
    for (const { dir, depth } of queue) {
      let listing: { files: string[]; dirs: string[] }
      try {
        listing = await fs.list(dir)
      } catch {
        continue // unreadable directory, same as an unplugged root
      }
      budget -= listing.files.length + listing.dirs.length
      for (const file of listing.files) {
        if (!SUGGEST_EXTS.has(extname(file).toLowerCase())) continue
        const path = join(dir, file)
        out.push({
          path,
          file,
          name: basename(file, extname(file)),
          relPath: relative(root, path),
        })
        if (out.length >= cap) break
      }
      if (out.length >= cap || depth + 1 > MAX_DEPTH) continue
      for (const sub of listing.dirs) {
        if (SKIP_DIRS.has(sub)) continue
        next.push({ dir: join(dir, sub), depth: depth + 1 })
      }
    }
    queue = next
  }
  return out
}
