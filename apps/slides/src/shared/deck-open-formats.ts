/**
 * The deck formats the Slides app opens. `.odp`/`.otp` join through the import
 * filter, which converts a copy — the same route the shell takes — so the
 * standalone open guard and the open dialog's picker filter read this one list
 * instead of drifting apart.
 */

export const OPENABLE_DECK_EXTENSIONS = ['pptx', 'ppt', 'odp', 'otp'] as const

const OPENABLE_DECK_PATTERN = new RegExp(`\\.(${OPENABLE_DECK_EXTENSIONS.join('|')})$`, 'i')

export function isOpenableDeckPath(path: string): boolean {
  return OPENABLE_DECK_PATTERN.test(path.trim())
}

/** lowercase extension without the dot, as the import route switches on it */
export function deckExtension(path: string): string {
  return path.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1] ?? ''
}
