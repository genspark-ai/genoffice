/**
 * Command-key matching that survives a non-Latin layout: `key` is the layout's
 * own character ('с' on a Cyrillic layout), `code` ('KeyS') the physical key.
 */

/** A keyboard event, narrowed to the fields the match needs. */
export type CommandLetterEvent = {
  key: string
  code?: string
  altKey?: boolean
  isComposing?: boolean
}

/**
 * True when the event is `letter` as the user's layout types it, or — when the
 * layout types no Latin letter for the keystroke — the physical key position.
 */
export function isCommandLetter(event: CommandLetterEvent, letter: string): boolean {
  const lower = letter.toLowerCase()
  const typed = event.key.toLowerCase()
  // A Latin letter is the layout's own answer: on Dvorak the key that types 'o'
  // sits at KeyS, and Open must not turn into Save because of the position.
  if (/^[a-z]$/.test(typed)) return typed === lower
  // A non-Latin layout names its own character, so the physical position is the
  // only stable signal. AltGr (Ctrl+Alt on Windows) types an accented letter at
  // that same position, and an IME has not decided on a character yet: both
  // belong to the document, not to a command.
  if (event.altKey || event.isComposing) return false
  return event.code === `Key${lower.toUpperCase()}`
}
