import type { JSONContent } from '@tiptap/core'

/**
 * `@tiptap/markdown` merges everything between an inline opening tag and its
 * closing tag into one HTML string, so markdown inside `<span>`...`</span>`
 * is never parsed, and it decodes only four entities. Tags with no schema
 * counterpart are dropped before that pass (GitHub strips them too), leaving
 * their content as ordinary inline tokens; text tokens are decoded up front.
 */

export interface InlineToken {
  type: string
  raw?: string
  text?: string
  tokens?: InlineToken[]
  /** set on the synthetic redaction token the inline patch produces */
  attrs?: { label?: unknown }
}

type InlineParser = { parseInlineTokens(tokens: InlineToken[]): JSONContent[] }

/** inline tags the schema maps to marks or nodes; everything else is inert */
/** the span a withheld span is persisted as; see editor/Redaction.ts */
const REDACTION_OPEN_RE = /<span\b[^>]*\bdata-redaction\b[^>]*>/i
const REDACTION_LABEL_RE = /\bdata-label\s*=\s*(?:"([^"]*)"|'([^']*)')/i
const SPAN_CLOSE_RE = /^<\/span\b/i

const FORMATTING_TAGS = new Set([
  'a',
  'b',
  'br',
  'code',
  'del',
  'em',
  'i',
  'img',
  's',
  'strike',
  'strong',
])
/** CommonMark tag names; marked also accepts `_`, which GitHub shows as literal text */
const TAG_NAME_RE = /^<\/?([a-zA-Z][a-zA-Z0-9-]*)(?=[\s/>])/

function isInertTag(raw: string): boolean {
  const match = TAG_NAME_RE.exec(raw)
  return match !== null && !FORMATTING_TAGS.has(match[1]!.toLowerCase())
}

const ENTITY_RE = /&(?:#(\d{1,8})|#[xX]([0-9a-fA-F]{1,8})|([a-zA-Z][a-zA-Z0-9]{1,31}));/g
const namedCache = new Map<string, string>()

function fromCodePoint(code: number): string {
  if (code === 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return '�'
  return String.fromCodePoint(code)
}

/** the browser knows the full HTML5 entity table; unknown names stay as typed */
function decodeNamed(entity: string): string {
  let text = namedCache.get(entity)
  if (text === undefined) {
    text = entity
    if (typeof DOMParser !== 'undefined') {
      const body = new DOMParser().parseFromString(entity, 'text/html').body
      text = body?.textContent ?? entity
    }
    namedCache.set(entity, text)
  }
  return text
}

export function decodeEntities(text: string): string {
  if (!text.includes('&')) return text
  return text.replace(ENTITY_RE, (entity, dec: string | undefined, hex: string | undefined) => {
    if (dec !== undefined) return fromCodePoint(parseInt(dec, 10))
    if (hex !== undefined) return fromCodePoint(parseInt(hex, 16))
    return decodeNamed(entity)
  })
}

/** the upstream pass still decodes these four; keep a decoded `&` in front of them literal */
const UPSTREAM_DECODED_RE = /&(?=(?:lt|gt|quot|amp);)/g

/**
 * Re-tag the span a withheld span is persisted as, so the parser hands it to
 * the redaction mark's own parseMarkdown instead of treating a <span> as an
 * inert tag and keeping only the words.
 *
 * It has to happen on the token stream: the manager only registers
 * markdownTokenizers from `baseExtensions`, which excludes marks by
 * construction, so a mark cannot hook the lexer from here. Rewriting the
 * tokens is the seam the existing inline patch already uses, and the opening
 * and closing tags are dropped so nothing downstream sees an unbalanced span.
 */
function redactSpans(tokens: InlineToken[]): InlineToken[] {
  const out: InlineToken[] = []
  let label: string | null = null
  for (const token of tokens) {
    const raw = token.raw ?? token.text ?? ''
    const open = REDACTION_OPEN_RE.exec(raw)
    if (open) {
      const found = REDACTION_LABEL_RE.exec(open[0])
      label = (found?.[1] ?? found?.[2] ?? 'private').trim() || 'private'
      out.push({ type: 'redaction', text: '', attrs: { label }, tokens: [] })
      continue
    }
    if (label !== null && SPAN_CLOSE_RE.test(raw)) {
      out.push({ type: 'redaction-close', text: '' })
      label = null
      continue
    }
    out.push(token)
  }
  return out
}

function prepareTokens(tokens: InlineToken[]): InlineToken[] {
  const out: InlineToken[] = []
  for (const token of tokens) {
    if (token.type === 'html' && isInertTag(token.raw ?? token.text ?? '')) continue
    if (token.type === 'text' && !token.tokens && token.text?.includes('&')) {
      const text = decodeEntities(token.text).replace(UPSTREAM_DECODED_RE, '&amp;')
      out.push({ ...token, text })
      continue
    }
    out.push(token)
  }
  return out
}

/** the newline after `<br>` is a soft break the browser collapses; the editor would show it */
function dropNewlineAfterHardBreak(nodes: JSONContent[]): JSONContent[] {
  for (let i = 0; i + 1 < nodes.length; i++) {
    const next = nodes[i + 1]!
    if (nodes[i]!.type !== 'hardBreak' || next.type !== 'text' || !next.text?.startsWith('\n')) {
      continue
    }
    const text = next.text.slice(1)
    if (text) nodes[i + 1] = { ...next, text }
    else nodes.splice(i + 1, 1)
  }
  return nodes
}

export function patchInlineParsing(manager: unknown): void {
  const parser = manager as InlineParser
  const upstream = parser.parseInlineTokens.bind(parser)
  // the span's opening tag becomes a redaction token, its content is gathered
  // into that token, and the closing tag ends it — after which the manager
  // routes the token to the mark's parseMarkdown and applies the mark
  parser.parseInlineTokens = (tokens) => {
    // both passes: the span rewrite below, and the entity/inert-tag cleanup
    // that was already here — the order matters, since redactSpans turns an
    // inert <span> into a token prepareTokens would otherwise have dropped
    const spans = markSpanTokens(redactSpans(tokens))
    return dropNewlineAfterHardBreak(upstream(prepareTokens(spans)))
  }
}

/** wrap the text between a redaction token and its close into that token */
function markSpanTokens(tokens: InlineToken[]): InlineToken[] {
  const out: InlineToken[] = []
  let open: InlineToken | null = null
  for (const token of tokens) {
    if (token.type === 'redaction') {
      open = token
      out.push(token)
      continue
    }
    if (token.type === 'redaction-close') {
      open = null
      continue
    }
    if (open) {
      const inner = open.tokens ?? []
      open.tokens = [
        ...inner,
        { type: 'text', raw: token.raw ?? token.text ?? '', text: token.text ?? token.raw ?? '' },
      ]
      continue
    }
    out.push(token)
  }
  return out
}
