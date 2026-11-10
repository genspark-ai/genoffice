// plain text and JSON open as source: the block editor would reinterpret
// `---`, `*x*`, `1.` and normalise line endings, rewriting an unedited file
export type TextMode = 'markdown' | 'plain' | 'json'

const JSON_RE = /\.json$/i
const MARKDOWN_RE = /\.(md|markdown)$/i

/** 'markdown' for an untitled path: that is what New and AI auto-naming mean by it. */
export function textModeForPath(path: string | null | undefined): TextMode {
  if (!path) return 'markdown'
  if (JSON_RE.test(path)) return 'json'
  if (MARKDOWN_RE.test(path)) return 'markdown'
  return 'plain'
}

export function isSourceMode(mode: TextMode): boolean {
  return mode !== 'markdown'
}
