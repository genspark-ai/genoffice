/**
 * Plain-text Markdown paste. Clipboards that carry only text/plain (fenced
 * code blocks, terminals, .md files in a text editor, LLM output) lose all
 * structure when pasted literally: "## Heading" keeps its hashes, "**bold**"
 * its asterisks. When such text looks like Markdown, it is converted to HTML
 * and inserted through the same pipeline as an HTML paste. Clipboards that
 * already carry text/html never reach this path.
 *
 * Detection errs toward literal paste: prose that merely mentions a "#" or a
 * lone "1." must not be reformatted, so conversion requires an unambiguous
 * Markdown construct (heading, fence, link, emphasis, table) or a repeated
 * one (several list/quote lines).
 */
import { marked } from 'marked'

/// Constructs that essentially never appear in prose accidentally.
const STRONG_SIGNALS: readonly RegExp[] = [
  /^#{1,6}\s+\S/m, // ATX heading
  /^\s{0,3}```/m, // fenced code block
  /\[[^\]\n]+\]\([^\s)]+\)/, // [text](url) link
  // **bold** only — the __bold__ form is indistinguishable from Python
  // dunder identifiers (__init__, __name__) and would convert pasted code.
  /(?:^|\W)\*\*[^*\n]+\*\*(?:\W|$)/,
  /^\s{0,3}\|.+\|\s*$\n^\s{0,3}\|[\s:|-]+\|\s*$/m, // table header + separator row
]

/// Constructs that appear in prose too; require several matching lines.
const LIST_LINE = /^\s{0,3}(?:[-*+]|\d{1,3}[.)])\s+\S/
const QUOTE_LINE = /^\s{0,3}>\s?\S/

export function looksLikeMarkdown(text: string): boolean {
  if (STRONG_SIGNALS.some((signal) => signal.test(text))) return true
  const lines = text.split('\n')
  const listLines = lines.filter((line) => LIST_LINE.test(line)).length
  const quoteLines = lines.filter((line) => QUOTE_LINE.test(line)).length
  return listLines >= 2 || quoteLines >= 2
}

/**
 * HTML for a plain-text paste that looks like Markdown, or null to leave the
 * default literal paste in place. The output feeds the same DOM → ProseMirror
 * parse as a native HTML paste, so unknown elements degrade the same way.
 */
export function markdownPasteHtml(text: string): string | null {
  const normalized = text.replace(/\r\n?/g, '\n')
  if (!looksLikeMarkdown(normalized)) return null
  try {
    return marked.parse(normalized, { gfm: true, breaks: false, async: false })
  } catch {
    return null
  }
}
