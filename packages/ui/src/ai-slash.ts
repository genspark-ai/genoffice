import type { AiSkillsMenuEntry } from './AiSkillsMenu'

/** Where a live `/token` sits in the composer text: `start` is the slash
 * itself, `end` the caret, so removing `[start, end)` is exactly what a pick
 * owes the draft. */
export interface AiSlashToken {
  start: number
  end: number
  query: string
}

/**
 * Locate the `/skill` token the caret sits in, if any. The nearest `/` before
 * the caret is the only candidate — with several slashes in the text, the one
 * closest to the caret wins — and it must stand at index 0 or right after
 * whitespace (a slash mid-word never triggers), with no whitespace between it
 * and the caret: a space ends the token, and whatever follows is the message
 * the picked skill runs on.
 */
export function detectSlashQuery(text: string, caret: number): AiSlashToken | null {
  const end = Math.max(0, Math.min(caret, text.length))
  let start = -1
  for (let i = end - 1; i >= 0; i--) {
    if (text[i] === '/') {
      start = i
      break
    }
  }
  if (start < 0) return null
  if (start > 0 && !isSlashSpace(text[start - 1] ?? '')) return null
  for (let i = start + 1; i < end; i++) {
    if (isSlashSpace(text[i] ?? '')) return null
  }
  return { start, end, query: text.slice(start + 1, end) }
}

/**
 * Skills matching a slash query: the same case-insensitive substring test over
 * name and description the "+" menu's search box applies. An empty query
 * matches every skill.
 */
export function filterSlashSkills(
  skills: readonly AiSkillsMenuEntry[],
  query: string,
): readonly AiSkillsMenuEntry[] {
  const q = query.trim().toLowerCase()
  if (!q) return skills
  return skills.filter(
    (s) => s.name.toLowerCase().includes(q) || s.description.toLowerCase().includes(q),
  )
}

function isSlashSpace(ch: string): boolean {
  return /\s/.test(ch)
}
