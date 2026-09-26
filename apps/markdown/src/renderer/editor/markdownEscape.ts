import { Markdown } from '@tiptap/markdown'
import { ImageParagraph } from './localImage'
import { Paragraph } from '@tiptap/extension-paragraph'
import { patchInlineParsing } from './inlineTokens'
import { containsMathSyntax } from './mathSyntax'

/**
 * `@tiptap/markdown` backslash-escapes every `[` and `]` in text, so an
 * authored `[[Foo]]`, `[1]` or `[TODO]` comes back from disk as `\[\[Foo\]\]`.
 * Plain brackets are inert in CommonMark; only escape them when the text would
 * otherwise turn into a link, a task marker or a reference definition.
 */
const LINK_LIKE_RE = /\]\(/
const TASK_MARKER_RE = /^\s*\[[ xX]\]/
const REF_DEFINITION_RE = /^\s*\[[^\]]+\]:/

export function escapeBrackets(text: string): string {
  if (!text.includes('[') && !text.includes(']')) return text
  const risky = LINK_LIKE_RE.test(text) || TASK_MARKER_RE.test(text) || REF_DEFINITION_RE.test(text)
  return risky ? text.replace(/([[\]])/g, '\\$1') : text
}

/** the table renderer joins cell text with bare pipes, so a `|` in a cell must be escaped */
let inTable = false

export function renderingTable(): boolean {
  return inTable
}

export function withTableRendering<T>(fn: () => T): T {
  const previous = inTable
  inTable = true
  try {
    return fn()
  } finally {
    inTable = previous
  }
}

export function escapeMarkdownText(text: string): string {
  // a literal no-break space is trimmed at cell and line edges when read back; the entity is not
  let escaped = escapeBrackets(text.replace(/([\\`*_~])/g, '\\$1')).replace(/\u00a0/g, '&nbsp;')
  // a `\$` the author wrote is plain text after parsing; without re-escaping, `a \$b\$ c` is a formula on reopen
  if (containsMathSyntax(text)) escaped = escaped.replace(/\$/g, '\\$')
  return inTable ? escaped.replace(/\|/g, '\\|') : escaped
}

/** block starts CommonMark would read from a paragraph line (`>` is already `&gt;` here) */
const BLOCK_START_RE =
  /^( {0,3})(?:(\d{1,9})(?=[.)](?:\s|$))|(#{1,6}(?=\s|$)|&gt;|[-+*](?=\s|$)|-{3,}\s*$|=+\s*$))/
const TABLE_DELIMITER_RE = /^\s*\|?\s*:?-+:?\s*(?:\|\s*:?-+:?\s*)*\|?\s*$/

export function escapeBlockStarts(text: string): string {
  return text
    .split('\n')
    .map((line, i) => {
      if (i > 0 && line.includes('|') && TABLE_DELIMITER_RE.test(line)) {
        return line.replace(/^(\s*)/, '$1\\')
      }
      // any run of dashes under a paragraph line is a setext underline
      if (i > 0 && /^ {0,3}-+\s*$/.test(line)) return line.replace(/^(\s*)/, '$1\\')
      // a lone tag followed by whitespace opens an HTML block that swallows the next lines:
      // hard-break the line with a backslash instead of trailing spaces
      if (i === 0 && /^\s*<[a-zA-Z](?:[^>"']|"[^"]*"|'[^']*')*>\s{2,}$/.test(line)) {
        return line.replace(/\s+$/, ' \\')
      }
      const m = BLOCK_START_RE.exec(line)
      if (!m) return line
      if (m[2]) return `${m[1]}${m[2]}\\${line.slice(m[0].length)}`
      if (m[3] === '&gt;') return `${m[1]}\\>${line.slice(m[0].length)}`
      if (m[3]!.startsWith('=') && i === 0) return line
      return `${m[1]}\\${line.slice(m[1]!.length)}`
    })
    .join('\n')
}

export const BlockStartEscapedParagraph = ImageParagraph.extend({
  renderMarkdown: (node, h, ctx) => {
    const text = Paragraph.config.renderMarkdown!(node, h, ctx)
    // table cells render their paragraph with the table as parent; `|` is escaped there
    return ctx?.parentType === 'table' ? text : escapeBlockStarts(text)
  },
})

/** CommonMark code span; travels with the text because marks render through a placeholder */
export function fenceCodeSpan(text: string): string {
  const runs = text.match(/`+/g) ?? []
  const fence = '`'.repeat(Math.max(1, ...runs.map((run) => run.length + 1)))
  const pad = /^`|`$/.test(text) || (/^ .* $/s.test(text) && text.trim() !== '') ? ' ' : ''
  return `${fence}${pad}${text}${pad}${fence}`
}

type JsonNode = { type?: string; marks?: Array<string | { type: string }> }
type Serializer = {
  escapeMarkdownSyntax: (text: string) => string
  encodeTextForMarkdown: (text: string, node: JsonNode, parent?: JsonNode) => string
}

const hasCodeMark = (node: JsonNode): boolean =>
  (node.marks ?? []).some((m) => (typeof m === 'string' ? m : m.type) === 'code')

export const SelectiveEscapeMarkdown = Markdown.extend({
  onBeforeCreate(props) {
    this.parent?.(props)
    const manager = this.editor.markdown as unknown as Serializer | undefined
    if (!manager) return
    manager.escapeMarkdownSyntax = escapeMarkdownText
    const encode = manager.encodeTextForMarkdown.bind(manager)
    manager.encodeTextForMarkdown = (text, node, parent) =>
      hasCodeMark(node) ? fenceCodeSpan(text) : encode(text, node, parent)
    patchInlineParsing(manager)
  },
})
