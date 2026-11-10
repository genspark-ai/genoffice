/// ECMA-376 Part 1 §22.4.2.4: characters illegal in XML (and CR, which XML
/// parsers would fold into LF) are stored in cell text as `_xHHHH_`; a
/// literal `_x` that would read as an escape is itself written `_x005F_x…`.
/// Formula text never carries these escapes.

export function decodeXlsxEscapes(text: string): string {
  if (!text.includes('_x')) return text
  return text.replace(/_x([0-9A-Fa-f]{4})_/g, (match, hex: string) => {
    const code = Number.parseInt(hex, 16)
    return code >= 0xd800 && code <= 0xdfff ? match : String.fromCharCode(code)
  })
}

export function encodeXlsxEscapes(text: string): string {
  return text.replace(/_(?=x[0-9A-Fa-f]{4}_)/g, '_x005F_').replace(
    // eslint-disable-next-line no-control-regex -- the control range is the thing being escaped
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF\r]/g,
    (character) => `_x${character.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`,
  )
}

const XML_NAMED_ENTITIES: Record<string, string> = {
  quot: '"',
  apos: "'",
  lt: '<',
  gt: '>',
  amp: '&',
}

/// Single-pass decode of the XML named entities plus numeric character
/// references (&#dd; / &#xhh;), which some producers use for non-ASCII text.
export function decodeXmlText(input: string): string {
  return input.replace(
    /&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|(quot|apos|lt|gt|amp));/g,
    (match, hex: string | undefined, dec: string | undefined, named: string | undefined) => {
      if (named !== undefined) return XML_NAMED_ENTITIES[named] ?? match
      const code = hex !== undefined ? Number.parseInt(hex, 16) : Number(dec)
      return code <= 0x10ffff ? String.fromCodePoint(code) : match
    },
  )
}
