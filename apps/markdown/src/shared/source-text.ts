/** BOM, line endings and trailing newline must survive a source-mode edit byte for byte. */
export interface SourceTextFormat {
  bom: boolean
  eol: '\n' | '\r\n'
  trailingNewline: boolean
}

const BOM = '\uFEFF'

export function readSourceText(raw: string): { text: string; format: SourceTextFormat } {
  const bom = raw.startsWith(BOM)
  const body = bom ? raw.slice(1) : raw
  const eol: '\n' | '\r\n' = body.includes('\r\n') ? '\r\n' : '\n'
  // a lone \r (classic Mac) is a line terminator too
  const text = body.replace(/\r\n?/g, '\n')
  return {
    text,
    format: {
      bom,
      eol,
      trailingNewline: body === '' || body.endsWith('\n') || body.endsWith('\r'),
    },
  }
}

export function writeSourceText(text: string, format: SourceTextFormat): string {
  const normalized = text.replace(/\r\n/g, '\n')
  const body = format.eol === '\r\n' ? normalized.replace(/\n/g, '\r\n') : normalized
  const withEol =
    body === '' || body.endsWith('\n')
      ? body
      : format.trailingNewline
        ? `${body}${format.eol}`
        : body
  return format.bom ? BOM + withEol : withEol
}
