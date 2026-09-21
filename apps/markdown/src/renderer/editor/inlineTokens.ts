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
}

type InlineParser = { parseInlineTokens(tokens: InlineToken[]): JSONContent[] }

/** inline tags the schema maps to marks or nodes; everything else is inert */
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
  parser.parseInlineTokens = (tokens) => dropNewlineAfterHardBreak(upstream(prepareTokens(tokens)))
}
