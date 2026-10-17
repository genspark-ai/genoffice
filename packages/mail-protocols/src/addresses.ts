/**
 * Address handling: parse and format RFC 5322 mailbox lists for the
 * envelope view-model. Display names may carry RFC 2047 encoded words;
 * parsing never throws on hostile input, it degrades to the raw atom.
 */

export interface MailAddress {
  name: string
  address: string
}

/** Split a header value on top-level commas (angle-bracket aware). */
function splitList(value: string): string[] {
  const parts: string[] = []
  let depth = 0
  let inQuote = false
  let cur = ''
  for (let i = 0; i < value.length; i++) {
    const ch = value[i]
    if (ch === '"' && value[i - 1] !== '\\') {
      inQuote = !inQuote
      cur += ch
    } else if (!inQuote && ch === '<') {
      depth++
      cur += ch
    } else if (!inQuote && ch === '>') {
      depth = Math.max(0, depth - 1)
      cur += ch
    } else if (!inQuote && depth === 0 && ch === ',') {
      parts.push(cur)
      cur = ''
    } else {
      cur += ch
    }
  }
  if (cur.trim() !== '') parts.push(cur)
  return parts
}

/** Decode one RFC 2047 encoded word; unknown charsets pass through raw. */
function decodeWord(word: string): string {
  const m = word.match(/^=\?([^?\s]+)\?([bBqQ])\?([^?]*)\?=$/)
  if (!m) return word
  const [, charset, enc, data] = m
  try {
    if (enc.toUpperCase() === 'B') {
      return Buffer.from(data, 'base64').toString(charset as BufferEncoding)
    }
    const qp = data
      .replace(/_/g, ' ')
      .replace(/=([0-9A-Fa-f]{2})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    return Buffer.from(qp, 'latin1').toString(charset as BufferEncoding)
  } catch {
    return word
  }
}

/** Decode all RFC 2047 encoded words embedded in a display name. */
export function decodeDisplayName(name: string): string {
  return name.replace(/=\?[^?\s]+\?[bBqQ]\?[^?]*\?=/g, (w) => decodeWord(w)).trim()
}

/** Parse one mailbox token into name plus address; never throws. */
export function parseAddress(token: string): MailAddress {
  const t = token.trim()
  const angle = t.match(/^(.*)<([^<>@\s]+@[^<>@\s]+)>\s*$/)
  if (angle) {
    const name = decodeDisplayName(angle[1].trim().replace(/^"|"$/g, '').trim())
    return { name, address: angle[2].trim() }
  }
  const bare = t.match(/^[^\s@<>]+@[^\s@<>]+$/)
  if (bare) return { name: '', address: bare[0] }
  return { name: '', address: t }
}

/** Parse a full address-list header into mailboxes; never throws. */
export function parseAddressList(header: string): MailAddress[] {
  if (!header || header.trim() === '') return []
  // Drop group syntax (Friends: a@x, b@y;) for envelope display.
  const deGroup =
    header.includes(':') && header.includes(';')
      ? header.slice(header.indexOf(':') + 1, header.lastIndexOf(';'))
      : header
  return splitList(deGroup)
    .map((p) => p.trim())
    .filter((p) => p !== '')
    .map(parseAddress)
    .filter((a) => a.address !== '')
}

/** Format one mailbox for display; quotes the name only when needed. */
export function formatAddress(addr: MailAddress): string {
  if (!addr.name) return addr.address
  const needsQuote = /[",;<>@]/.test(addr.name)
  return `${needsQuote ? `"${addr.name}"` : addr.name} <${addr.address}>`
}

/** Format a mailbox list for a compose header. */
export function formatAddressList(addrs: MailAddress[]): string {
  return addrs.map(formatAddress).join(', ')
}

/** Envelope-ready address strings (name dropped, address kept). */
export function toEnvelopeStrings(addrs: MailAddress[]): string[] {
  return addrs.map((a) => a.address)
}
