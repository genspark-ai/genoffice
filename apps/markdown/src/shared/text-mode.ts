/**
 * Which editing surface a file gets inside the text app.
 *
 * The markdown surface is a WYSIWYG block editor whose whole model is "a
 * document is a tree of blocks, and its file form is markdown". Plain text and
 * JSON have no such structure: feeding either to that editor would parse
 * `---` as a frontmatter fence, read `*x*` as emphasis and `1.` as a list,
 * normalise CRLF, and force a trailing newline — so saving a file the user
 * never meaningfully edited would rewrite it. Both therefore open as source
 * text that round-trips byte for byte.
 */
export type TextMode = 'markdown' | 'plain' | 'json'

const JSON_RE = /\.json$/i
const MARKDOWN_RE = /\.(md|markdown)$/i

/**
 * The mode for a path, or 'markdown' for a null/untitled path.
 *
 * An untitled document is a new markdown document: that is what the New
 * button, the empty tab and the AI auto-naming path all mean by it.
 */
export function textModeForPath(path: string | null | undefined): TextMode {
  if (!path) return 'markdown'
  if (JSON_RE.test(path)) return 'json'
  if (MARKDOWN_RE.test(path)) return 'markdown'
  return 'plain'
}

/** True when the mode edits source text instead of a block document. */
export function isSourceMode(mode: TextMode): boolean {
  return mode !== 'markdown'
}
