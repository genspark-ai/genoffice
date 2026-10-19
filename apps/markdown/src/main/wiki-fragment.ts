/**
 * Where a wiki link pointed, carried from the tab that followed it to the tab
 * that shows the note.
 *
 * The shell activates an already-open tab without telling its renderer
 * anything, and the path it does queue carries no fragment, so `Note#Heading`
 * could not say where to land by riding along with the open request. The
 * fragment waits here instead, keyed by path, until a renderer loads that path
 * and takes it.
 */

/** paths whose fragment is still waiting; a click whose note is never opened must not pile up */
const MAX_PENDING = 32

const pending = new Map<string, string>()

export function rememberWikiFragment(path: string, fragment: string): void {
  if (!path) return
  // a link with no fragment cancels an earlier one aimed at the same note
  if (!fragment) {
    pending.delete(path)
    return
  }
  while (pending.size >= MAX_PENDING) {
    const oldest = pending.keys().next().value
    if (oldest === undefined) break
    pending.delete(oldest)
  }
  pending.set(path, fragment)
}

export function takeWikiFragment(path: string): string {
  if (!path) return ''
  const fragment = pending.get(path) ?? ''
  pending.delete(path)
  return fragment
}
