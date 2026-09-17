import { renderNestedMarkdownContent } from '@tiptap/core'
import { Bold } from '@tiptap/extension-bold'
import { HardBreak } from '@tiptap/extension-hard-break'
import { Heading } from '@tiptap/extension-heading'
import { HorizontalRule } from '@tiptap/extension-horizontal-rule'
import { Italic } from '@tiptap/extension-italic'
import { getListMarker, ListItem, TaskItem } from '@tiptap/extension-list'
import { activeMarkdownStyle } from '../markdown/markdownStyle'

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

export const StyledHardBreak = HardBreak.extend({
  renderMarkdown: () => (activeMarkdownStyle().hardBreak === 'backslash' ? '\\\n' : '  \n'),
})

export const StyledListItem = ListItem.extend({
  renderMarkdown: (node, h, ctx) =>
    renderNestedMarkdownContent(
      node,
      h,
      (context) => {
        const style = activeMarkdownStyle()
        if (context.parentType !== 'orderedList') return `${style.bullet} `
        const attrs = context.meta?.parentAttrs as { start?: number; type?: string } | undefined
        const start = attrs?.start || 1
        const index = style.orderedRepeat ? 0 : context.index || 0
        return getListMarker(attrs?.type, start - 1 + index, `${style.orderedDelimiter} `)
      },
      ctx,
      { alignNestedToPrefix: ctx?.parentType === 'orderedList' },
    ),
})

export const StyledTaskItem = TaskItem.extend({
  renderMarkdown: (node, h) => {
    const prefix = `${activeMarkdownStyle().bullet} [${node.attrs?.checked ? 'x' : ' '}] `
    return renderNestedMarkdownContent(node, h, prefix)
  },
})

/** fence long enough that the code's own backtick or tilde runs cannot close it */
export function renderFencedCode(language: string, body: string | null): string {
  const char = activeMarkdownStyle().fence
  const runs = body?.match(char === '`' ? /`{3,}/g : /~{3,}/g) ?? []
  const fence = char.repeat(Math.max(3, ...runs.map((run) => run.length + 1)))
  return body === null
    ? `${fence}${language}\n\n${fence}`
    : `${fence}${language}\n${body}\n${fence}`
}
