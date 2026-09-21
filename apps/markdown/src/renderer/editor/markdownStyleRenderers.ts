import type { JSONContent, MarkdownRendererHelpers, RenderContext } from '@tiptap/core'
import { Bold } from '@tiptap/extension-bold'
import { Code } from '@tiptap/extension-code'
import { HardBreak } from '@tiptap/extension-hard-break'
import { Heading } from '@tiptap/extension-heading'
import { HorizontalRule } from '@tiptap/extension-horizontal-rule'
import { Italic } from '@tiptap/extension-italic'
import { getListMarker, ListItem, TaskItem } from '@tiptap/extension-list'
import { Table } from '@tiptap/extension-table'
import { activeMarkdownStyle } from '../markdown/markdownStyle'
import { renderingTable } from './markdownEscape'

/**
 * The stock renderers hardcode `-`, `1.`, `*`, `**`, `#`, ```` ``` ````, `---`
 * and two-space breaks. These read the style in force instead (see
 * markdownStyle.ts); outside a style scope they produce the same output.
 */

export const StyledBold = Bold.extend({
  renderMarkdown: (node, h) => {
    const strong = activeMarkdownStyle().strong
    return `${strong}${h.renderChildren(node)}${strong}`
  },
})

export const StyledItalic = Italic.extend({
  renderMarkdown: (node, h) => {
    const em = activeMarkdownStyle().em
    return `${em}${h.renderChildren(node)}${em}`
  },
})

/** fence comes with the text (fenceCodeSpan); registered last among marks so it renders innermost */
export const CodeSpan = Code.extend({
  // other marks may wrap a code span (**`x`**); the default excludes '_' rejects them
  excludes: '',
  renderMarkdown: (node, h) => (node.content ? h.renderChildren(node.content) : ''),
})

/**
 * A setext underline only makes a heading of plain paragraph lines: a line
 * CommonMark reads as a list item, quote, heading, indented code or thematic
 * break first would come back as something else, and a fence opener or a
 * bare `===` line would end the heading early. Line breaks are fine — a
 * setext heading may span lines, which ATX cannot.
 */
const NOT_SETEXT_LINE =
  /^ {0,3}(?:[-+*]|\d{1,9}[.)])(?:\s|$)|^ {0,3}[>#]|^ {4}|^ {0,3}(?:[-*_]\s*){3,}$|^ {0,3}(?:`{3,}|~{3,})|^ {0,3}=+\s*$/

function setextSafe(text: string): boolean {
  if (text.trim() === '') return false
  return text.split('\n').every((line) => line.trim() !== '' && !NOT_SETEXT_LINE.test(line))
}

export const StyledHeading = Heading.extend({
  renderMarkdown: (node, h) => {
    if (!node.content) return ''
    const level = node.attrs?.level ? parseInt(String(node.attrs.level), 10) : 1
    const text = h.renderChildren(node.content)
    if (activeMarkdownStyle().setext && level <= 2 && setextSafe(text)) {
      const width = Math.max(3, ...text.split('\n').map((line) => line.length))
      return `${text}\n${(level === 1 ? '=' : '-').repeat(width)}`
    }
    return `${'#'.repeat(level)} ${text}`
  },
})

export const StyledHorizontalRule = HorizontalRule.extend({
  renderMarkdown: () => activeMarkdownStyle().rule,
})

const HARD_BREAKS = { spaces: '  \n', backslash: '\\\n', html: '<br>\n' }

/** the table renderer turns every newline in a cell into `<br>`, so a break there must not add one */
export const StyledHardBreak = HardBreak.extend({
  renderMarkdown: () => (renderingTable() ? '<br>' : HARD_BREAKS[activeMarkdownStyle().hardBreak]),
})

function parentLoose(ctx: RenderContext | undefined): boolean {
  return (ctx?.meta?.parentAttrs as Record<string, unknown> | undefined)?.loose === true
}

/** blocks that may follow another block without a blank line (they interrupt a paragraph) */
const INTERRUPTS = new Set([
  'codeBlock',
  'bulletList',
  'orderedList',
  'taskList',
  'heading',
  'horizontalRule',
  'blockquote',
])

/**
 * Child blocks sit at the item's content column (the marker width), and a
 * loose list keeps a blank line between them; the stock helper indents by the
 * fixed 4 spaces and only separates paragraphs, which glues a fence to the
 * image after it and turns a loose list tight.
 */
function renderItem(
  node: JSONContent,
  h: MarkdownRendererHelpers,
  prefix: string,
  loose: boolean,
  contentColumn = prefix.length,
): string {
  const children = node.content ?? []
  const indent = ' '.repeat(contentColumn)
  const reindent = (text: string, fromLine: number): string =>
    text
      .split('\n')
      .map((line, i) => (i < fromLine || line === '' ? line : indent + line))
      .join('\n')
  let out = prefix
  children.forEach((child, i) => {
    const rendered = h.renderChild?.(child, i) ?? h.renderChildren([child])
    if (i === 0) {
      out += reindent(rendered, 1)
      return
    }
    const blank = loose || !INTERRUPTS.has(String(child.type))
    out += (blank ? '\n\n' : '\n') + reindent(rendered, 0)
  })
  return out
}

export const StyledListItem = ListItem.extend({
  renderMarkdown: (node, h, ctx) => {
    const style = activeMarkdownStyle()
    let prefix = `${style.bullet} `
    if (ctx?.parentType === 'orderedList') {
      const attrs = ctx.meta?.parentAttrs as { start?: number; type?: string } | undefined
      const start = attrs?.start || 1
      const index = style.orderedRepeat ? 0 : ctx.index || 0
      prefix = getListMarker(attrs?.type, start - 1 + index, `${style.orderedDelimiter} `)
    }
    return renderItem(node, h, prefix, parentLoose(ctx))
  },
})

export const StyledTaskItem = TaskItem.extend({
  renderMarkdown: (node, h, ctx) => {
    // the checkbox is paragraph text: the content column is right after the bullet
    const bullet = `${activeMarkdownStyle().bullet} `
    const prefix = `${bullet}[${node.attrs?.checked ? 'x' : ' '}] `
    return renderItem(node, h, prefix, parentLoose(ctx), bullet.length)
  },
})

/** `loose` is not in the HTML; it is read off the markdown at parse time (see looseLists.ts) */
export function renderList(node: JSONContent, h: MarkdownRendererHelpers): string {
  if (!node.content) return ''
  return h.renderChildren(node.content, node.attrs?.loose === true ? '\n\n' : '\n')
}

/** the stock renderer pads every column to its widest cell; a compact source stays compact */
export function renderTable(
  node: JSONContent,
  h: MarkdownRendererHelpers,
  ctx: RenderContext,
): string {
  const out = Table.config.renderMarkdown!(node, h, ctx).replace(/^\n+|\n+$/g, '')
  if (activeMarkdownStyle().tableAligned) return out
  return out
    .split('\n')
    .map((line, i) => (i === 1 ? line.replace(/-{3,}/g, '---') : line).replace(/ {2,}\|/g, ' |'))
    .join('\n')
}

/** fence long enough that the code's own backtick or tilde runs cannot close it */
export function renderFencedCode(language: string, body: string | null): string {
  const char = activeMarkdownStyle().fence
  const runs = body?.match(char === '`' ? /`{3,}/g : /~{3,}/g) ?? []
  const fence = char.repeat(Math.max(3, ...runs.map((run) => run.length + 1)))
  return body === null
    ? `${fence}${language}\n\n${fence}`
    : `${fence}${language}\n${body}\n${fence}`
}
