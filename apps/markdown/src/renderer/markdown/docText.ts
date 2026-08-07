/**
 * Round-trip envelope for a .md file: the TipTap editor only ever sees the
 * markdown body; the YAML frontmatter block, EOL style and EOF-newline state
 * are captured on load and re-applied verbatim on save, so a file written by
 * another tool survives an open→save cycle without envelope churn.
 */
export interface DocEnvelope {
  /** raw frontmatter block including both fences and any blank lines after it; '' when absent */
  frontmatter: string
  /** markdown body, \n line endings */
  body: string
  eol: '\n' | '\r\n'
  /** whether the original file ended with a newline (new documents: true) */
  trailingNewline: boolean
  /** the original file started with a UTF-8 BOM (Windows Notepad) — re-emitted on save */
  bom: boolean
}

const FENCE = '---'

export function parseDocText(raw: string): DocEnvelope {
  const bom = raw.startsWith('\uFEFF')
  if (bom) raw = raw.slice(1)
  const eol: DocEnvelope['eol'] = raw.includes('\r\n') ? '\r\n' : '\n'
  const text = raw.replace(/\r\n/g, '\n')
  const trailingNewline = text === '' || text.endsWith('\n')

  let frontmatter = ''
  let body = text
  if (text.startsWith(`${FENCE}\n`)) {
    let searchFrom = FENCE.length
    for (;;) {
      const close = text.indexOf(`\n${FENCE}`, searchFrom)
      if (close < 0) break
      const closeEnd = close + 1 + FENCE.length
      // the closing fence must be a whole line: EOF or followed by a newline
      if (closeEnd === text.length || text[closeEnd] === '\n') {
        let end = closeEnd
        while (text[end] === '\n') end++
        frontmatter = text.slice(0, end)
        body = text.slice(end)
        break
      }
      searchFrom = close + 1
    }
  }
  return { frontmatter, body, eol, trailingNewline, bom }
}

/** Inner YAML text of a raw frontmatter block (fences and trailing blank lines stripped) */
export function frontmatterInner(raw: string): string {
  if (!raw.startsWith(`${FENCE}\n`)) return ''
  const inner = raw.slice(FENCE.length + 1).replace(/\n+$/, '')
  return inner.endsWith(`\n${FENCE}`) ? inner.slice(0, -(FENCE.length + 1)) : ''
}

/** Rebuild a raw frontmatter block from edited inner YAML; blank input removes the block */
export function buildFrontmatterRaw(inner: string): string {
  const trimmed = inner.replace(/^\n+|\n+$/g, '')
  return trimmed === '' ? '' : `${FENCE}\n${trimmed}\n${FENCE}\n\n`
}

/** Reassemble the full file text from the envelope and the (re)serialized body */
export function serializeDocText(envelope: DocEnvelope, body: string): string {
  let text = envelope.frontmatter + body
  if (envelope.trailingNewline) {
    if (text !== '' && !text.endsWith('\n')) text += '\n'
  } else {
    text = text.replace(/\n+$/, '')
  }
  if (envelope.eol === '\r\n') text = text.replace(/\n/g, '\r\n')
  return envelope.bom ? `\uFEFF${text}` : text
}
