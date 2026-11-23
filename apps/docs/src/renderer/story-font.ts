/**
 * The writing view's face, picked by the user (#2654 next cut). One face set
 * by the view, never by the document — the bundled writing faces (Monaspace
 * Neon, IBM Plex Mono, SIL OFL) come first, and any installed font can take
 * over. The choice persists in localStorage and lands as one CSS rule that
 * overrides --story-face (App injects it like the authorship inks).
 */

const KEY = 'docs.story.face'

export interface StoryFaceChoice {
  /** the CSS font-family value */
  family: string
  bundled?: boolean
}

/** the faces the picker offers by name; everything else goes through the
 *  free-text field. The bundled faces come with the font bundling, the rest
 *  are the writing faces a machine may have installed. */
export const STORY_FACE_CHOICES: StoryFaceChoice[] = [
  { family: 'Monaspace Neon', bundled: true },
  { family: 'IBM Plex Mono', bundled: true },
  { family: 'iA Writer Mono S' },
  { family: 'iA Writer Quattro S' },
  { family: 'Iosevka Etoile' },
  { family: 'Recursive' },
]

/** the stack styles.css sets for --story-face; a picked face rides in front */
const FACE_FALLBACK_TAIL =
  "'Monaspace Neon', 'iA Writer Mono S', 'iA Writer Quattro S', 'IBM Plex Mono', " +
  "'Iosevka Etoile', 'Recursive', system-ui, -apple-system, 'Segoe UI', 'PingFang SC', " +
  "'Microsoft YaHei', sans-serif"

/** the stored choice; '' = the view's own default stack */
export function loadStoryFace(): string {
  try {
    return localStorage.getItem(KEY) ?? ''
  } catch {
    return ''
  }
}

export function saveStoryFace(family: string): void {
  try {
    localStorage.setItem(KEY, family)
  } catch {
    // storage disabled: the choice lasts for the session
  }
}

/** a font name is a CSS string value: quotes, backslashes and rule characters
 *  never belong in one. Returns '' when nothing usable remains. */
export function sanitizeFaceName(raw: string): string {
  const cleaned = raw
    .trim()
    .replace(/['"\\;{}]/g, '')
    .trim()
  return cleaned.length > 0 && cleaned.length <= 64 ? cleaned : ''
}

/** whether the machine actually has the face (bundled or installed) —
 *  document.fonts.check answers for a face the document can use right now.
 *  Unavailable in non-browser test environments: treat as unknown (false). */
export function faceInstalled(family: string): boolean {
  try {
    return document.fonts.check(`16px "${family}"`)
  } catch {
    return false
  }
}

/** the rule a picked face injects: the face first — a missing one just falls
 *  through the bundled stack below, exactly like the view's default does.
 *  '' when the family is the empty default. */
export function storyFaceRule(family: string): string {
  const name = sanitizeFaceName(family)
  if (!name) return ''
  const face = `'${name}'`
  const stack = face === "'Monaspace Neon'" ? FACE_FALLBACK_TAIL : `${face}, ${FACE_FALLBACK_TAIL}`
  return `.view-story{--story-face:${stack};}`
}
