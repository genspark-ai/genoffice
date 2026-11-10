/** File-level facts that must survive an open → edit → save round trip untouched. */
export interface Envelope {
  bom: boolean
  eol: '\n' | '\r\n'
  trailingNewline: boolean
}

export interface DocText {
  /** body with EOL normalized to \n and no BOM */
  text: string
  envelope: Envelope
}

const BOM = '﻿'

export function parseDocText(raw: string): DocText {
  const bom = raw.startsWith(BOM)
  let body = bom ? raw.slice(1) : raw
  const crlf = (body.match(/\r\n/g) ?? []).length
  const lf = (body.match(/(?<!\r)\n/g) ?? []).length
  const eol: Envelope['eol'] = crlf > lf ? '\r\n' : '\n'
  body = body.replace(/\r\n?/g, '\n')
  const trailingNewline = body.endsWith('\n')
  if (trailingNewline) body = body.slice(0, -1)
  return { text: body, envelope: { bom, eol, trailingNewline } }
}

export function serializeDocText(doc: DocText): string {
  const { bom, eol, trailingNewline } = doc.envelope
  let out = eol === '\n' ? doc.text : doc.text.replace(/\n/g, eol)
  if (trailingNewline) out += eol
  return bom ? BOM + out : out
}
