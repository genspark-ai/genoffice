/**
 * Cascading-offset bookkeeping for element paste. Clipboard items keep their
 * source coordinates, so a paste lands exactly where the element was copied
 * from — the PowerPoint-style 16px cascade applies only when that spot on the
 * target page is already taken: by the original (pasting back onto the copy
 * page) or by an earlier paste of the same clipboard onto that page. A paste
 * onto any other page therefore reproduces the source position 1:1, and a cut
 * (no original remains) pastes back in place even on its own page.
 */
export interface PasteCascade {
  /** Page the elements were copied from; null after a cut (original removed). */
  sourceKey: string | null
  /** How many pastes of this clipboard each page has received. */
  counts: Map<string, number>
}

export function newPasteCascade(sourceKey: string | null): PasteCascade {
  return { sourceKey, counts: new Map() }
}

/** Pages are identified per window and slide (the clipboard is app-wide across open decks). */
export function pageKey(webContentsId: number, slideIndex: number): string {
  return `${webContentsId}:${slideIndex}`
}

/** 16px per landing already on the page; the original on the copy page counts as one. */
export function pasteShiftPx(clip: PasteCascade, targetKey: string): number {
  const landed = (clip.counts.get(targetKey) ?? 0) + (targetKey === clip.sourceKey ? 1 : 0)
  return 16 * landed
}

export function recordPaste(clip: PasteCascade, targetKey: string): void {
  clip.counts.set(targetKey, (clip.counts.get(targetKey) ?? 0) + 1)
}
