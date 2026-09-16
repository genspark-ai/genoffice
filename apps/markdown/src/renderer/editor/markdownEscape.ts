import { Markdown } from '@tiptap/markdown'

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

export function escapeMarkdownText(text: string): string {
  return escapeBrackets(text.replace(/([\\`*_~])/g, '\\$1'))
}

type Escaper = { escapeMarkdownSyntax: (text: string) => string }

export const SelectiveEscapeMarkdown = Markdown.extend({
  onBeforeCreate(props) {
    this.parent?.(props)
    const manager = this.editor.markdown as unknown as Escaper | undefined
    if (manager) manager.escapeMarkdownSyntax = escapeMarkdownText
  },
})
