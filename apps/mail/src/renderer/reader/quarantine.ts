/**
 * Quarantined HTML-mail rendering (Lane R owns the reader layer).
 *
 * Message bodies carry AUTHORED colors and markup: the sanitizer below
 * keeps those while stripping everything that can reach outside the
 * quarantine — scripts, remote stylesheets, event handlers, non-http(s)
 * links, and (by default) remote images. Sanitized output must never
 * contain a chrome token reference (`var(--…)`); the companion
 * `mail-quarantine.css` styles bodies with app-scoped `--mq-*` custom
 * properties only. Remote-image gating, sender allowlist UX, link
 * confirm, and phishing flags live in `MessageReader`; this module is
 * pure string/DOM analysis with no storage and no network.
 */

export interface SanitizeOptions {
  /** When false (default), remote http(s) media becomes a placeholder. */
  allowRemoteImages?: boolean
  /** Accessible label for blocked-media placeholders. */
  blockedLabel?: string
}

export interface SanitizeResult {
  html: string
  /** Remote media nodes replaced by placeholders. */
  blockedCount: number
}

/** Elements removed with their entire subtree. */
const DROP_ENTIRELY = new Set([
  'script',
  'style',
  'iframe',
  'frame',
  'frameset',
  'object',
  'embed',
  'form',
  'button',
  'input',
  'select',
  'textarea',
  'meta',
  'base',
  'link',
  'title',
  'noscript',
])

/** Elements kept (attributes still scrubbed); anything else is unwrapped. */
const KEEP_TAGS = new Set([
  'a',
  'p',
  'div',
  'span',
  'br',
  'hr',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'ul',
  'ol',
  'li',
  'dl',
  'dt',
  'dd',
  'blockquote',
  'pre',
  'code',
  'table',
  'thead',
  'tbody',
  'tfoot',
  'tr',
  'td',
  'th',
  'caption',
  'colgroup',
  'col',
  'img',
  'picture',
  'figure',
  'figcaption',
  'strong',
  'em',
  'b',
  'i',
  'u',
  's',
  'strike',
  'del',
  'ins',
  'font',
  'center',
  'sub',
  'sup',
  'mark',
  'small',
  'big',
  'q',
  'cite',
  'dfn',
  'abbr',
  'address',
  'details',
  'summary',
  'section',
  'article',
  'header',
  'footer',
  'main',
  'aside',
  'video',
  'audio',
  'source',
])

const GLOBAL_ATTRS = new Set(['title', 'lang', 'dir', 'align'])

const TAG_ATTRS: Record<string, Set<string>> = {
  a: new Set(['href', 'title']),
  img: new Set(['src', 'alt', 'title', 'width', 'height']),
  video: new Set(['src', 'poster', 'width', 'height', 'controls']),
  audio: new Set(['src', 'controls']),
  source: new Set(['src', 'type']),
  table: new Set(['border', 'cellpadding', 'cellspacing', 'width']),
  td: new Set(['colspan', 'rowspan', 'width', 'align', 'valign']),
  th: new Set(['colspan', 'rowspan', 'width', 'align', 'valign']),
  ol: new Set(['type', 'start']),
  li: new Set(['value']),
}

/** Inline-style properties kept from authored bodies (colors included). */
const STYLE_PROPS = new Set([
  'color',
  'background-color',
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'text-decoration',
  'text-align',
  'line-height',
  'letter-spacing',
  'margin',
  'margin-top',
  'margin-right',
  'margin-bottom',
  'margin-left',
  'padding',
  'padding-top',
  'padding-right',
  'padding-bottom',
  'padding-left',
  'border',
  'border-color',
  'border-width',
  'border-style',
  'border-collapse',
  'width',
  'height',
  'max-width',
  'vertical-align',
  'list-style-type',
  'white-space',
  'word-break',
  'direction',
])

const FORBIDDEN_STYLE_VALUE =
  /url\s*\(|expression\s*\(|var\s*\(|behavior\s*:|@import|javascript\s*:/i

function isRemoteUrl(value: string): boolean {
  return /^https?:\/\//i.test(value.trim())
}

function isDataImage(value: string): boolean {
  return /^data:image\//i.test(value.trim())
}

function sanitizeStyle(raw: string): string {
  const kept: string[] = []
  for (const declaration of raw.split(';')) {
    const colon = declaration.indexOf(':')
    if (colon < 0) continue
    const prop = declaration.slice(0, colon).trim().toLowerCase()
    const value = declaration.slice(colon + 1).trim()
    if (!STYLE_PROPS.has(prop)) continue
    if (!value || FORBIDDEN_STYLE_VALUE.test(value)) continue
    kept.push(`${prop}: ${value}`)
  }
  return kept.join('; ')
}

function sanitizeHref(raw: string): string | null {
  const value = raw.trim()
  if (/^https?:\/\//i.test(value)) return value
  if (/^mailto:/i.test(value)) return value
  if (value.startsWith('#')) return value
  return null
}

/**
 * Sanitize authored message HTML into quarantine-safe markup. Pure:
 * no storage, no network. Remote http(s) media is replaced by a
 * labelled placeholder unless `allowRemoteImages` is set (per-message
 * "Show images" or an allowlisted sender, decided by the reader).
 */
export function sanitizeMailHtml(source: string, opts: SanitizeOptions = {}): SanitizeResult {
  const allowRemoteImages = opts.allowRemoteImages ?? false
  const blockedLabel = opts.blockedLabel ?? 'Blocked remote image'
  const doc = new DOMParser().parseFromString(
    '<!doctype html><html><body></body></html>',
    'text/html',
  )
  doc.body.innerHTML = source
  let blockedCount = 0

  const blockMedia = (node: Element): void => {
    blockedCount += 1
    const placeholder = doc.createElement('span')
    placeholder.setAttribute('class', 'mq-blocked')
    placeholder.setAttribute('data-mq-blocked', 'media')
    placeholder.setAttribute('aria-label', blockedLabel)
    const alt = node.getAttribute('alt')
    if (alt) placeholder.setAttribute('data-mq-alt', alt)
    node.replaceWith(placeholder)
  }

  const scrub = (root: Element): void => {
    for (const child of [...root.children]) scrub(child)
    const tag = root.tagName.toLowerCase()
    if (tag === 'html' || tag === 'head' || tag === 'body') return
    if (DROP_ENTIRELY.has(tag)) {
      root.remove()
      return
    }
    if (!KEEP_TAGS.has(tag)) {
      root.replaceWith(...[...root.childNodes])
      return
    }
    const allowed = new Set<string>(GLOBAL_ATTRS)
    for (const attr of TAG_ATTRS[tag] ?? []) allowed.add(attr)
    for (const attr of [...root.attributes]) {
      const name = attr.name.toLowerCase()
      if (name.startsWith('on')) {
        root.removeAttribute(attr.name)
        continue
      }
      if (name === 'style') {
        const clean = sanitizeStyle(attr.value)
        if (clean) root.setAttribute('style', clean)
        else root.removeAttribute(attr.name)
        continue
      }
      if (name === 'id') {
        // App ids must not leak into (or out of) the quarantine.
        root.removeAttribute(attr.name)
        continue
      }
      if (name === 'class') {
        // Only quarantine-namespaced classes survive; app classes stay out.
        const kept = attr.value.split(/\s+/).filter((token) => token.startsWith('mq-'))
        if (kept.length > 0) root.setAttribute('class', kept.join(' '))
        else root.removeAttribute(attr.name)
        continue
      }
      if (!allowed.has(name)) {
        root.removeAttribute(attr.name)
        continue
      }
      if (
        (tag === 'img' || tag === 'video' || tag === 'audio' || tag === 'source') &&
        name === 'src'
      ) {
        const value = attr.value.trim()
        if (isRemoteUrl(value) && !allowRemoteImages) {
          blockMedia(root)
          return
        }
        if (!isRemoteUrl(value) && !isDataImage(value) && !value.startsWith('cid:')) {
          root.removeAttribute(attr.name)
        }
        continue
      }
      if (tag === 'video' && name === 'poster') {
        if (isRemoteUrl(attr.value) && !allowRemoteImages) {
          root.removeAttribute(attr.name)
        } else if (!isRemoteUrl(attr.value) && !isDataImage(attr.value)) {
          root.removeAttribute(attr.name)
        }
        continue
      }
      if (tag === 'a' && name === 'href') {
        const clean = sanitizeHref(attr.value)
        if (clean === null) root.removeAttribute(attr.name)
        else {
          root.setAttribute('href', clean)
          root.setAttribute('target', '_blank')
          root.setAttribute('rel', 'noopener noreferrer')
        }
      }
    }
  }

  scrub(doc.body)
  return { html: doc.body.innerHTML, blockedCount }
}

/** Split `Name <addr>` down to the bare addr-spec, lowercased. */
export function extractAddress(raw: string): string {
  const match = raw.match(/<([^<>]+)>/)
  const addr = (match ? match[1] : raw).trim().toLowerCase()
  return addr
}

/** Domain part of an address, lowercased; empty when unparseable. */
export function extractDomain(raw: string): string {
  const addr = extractAddress(raw)
  const at = addr.lastIndexOf('@')
  if (at < 0) return ''
  return addr.slice(at + 1)
}

/** True when the Reply-To domain differs from the sender domain. */
export function replyToMismatch(from: string[], replyTo?: string): boolean {
  if (!replyTo || replyTo.trim() === '') return false
  if (from.length === 0) return false
  const senderDomain = extractDomain(from[0])
  const replyDomain = extractDomain(replyTo)
  if (!senderDomain || !replyDomain) return false
  return senderDomain !== replyDomain
}

function levenshtein(a: string, b: string): number {
  const prev = Array.from({ length: b.length + 1 }, (_, i) => i)
  for (let i = 1; i <= a.length; i += 1) {
    let diag = prev[0]
    prev[0] = i
    for (let j = 1; j <= b.length; j += 1) {
      const next = prev[j]
      const cost = a[i - 1] === b[j - 1] ? 0 : 1
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, diag + cost)
      diag = next
    }
  }
  return prev[b.length]
}

/**
 * Heuristic lookalike-domain flag: internationalised (xn--) domains are
 * always flagged; other domains within edit distance 2 of a trusted
 * domain are flagged as probable spoofs.
 */
export function lookalikeDomain(domain: string, trusted: string[]): boolean {
  const candidate = domain.trim().toLowerCase()
  if (!candidate) return false
  if (
    candidate.startsWith('xn--') ||
    candidate.split('.').some((label) => label.startsWith('xn--'))
  )
    return true
  return trusted.some((name) => {
    const known = name.trim().toLowerCase()
    if (!known || known === candidate) return false
    return levenshtein(candidate, known) <= 2
  })
}
