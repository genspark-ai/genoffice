import { Mark, mergeAttributes } from '@tiptap/core'

/**
 * `==highlight==` — Obsidian's text highlight, applied to the real text so it
 * round-trips as the same syntax. `inclusive: false` and `exitable: true` stop
 * the mark from growing or bleeding as the caret moves around it (the same
 * choice the redaction mark makes).
 *
 * **In (an open).** marked has no `==` tokenizer; the inline patch in
 * editor/inlineTokens.ts rewrites `==text==` runs into a synthetic `highlight`
 * token, and the manager routes that token to `parseMarkdown` below — the same
 * path the redaction mark rides on.
 *
 * **Out (a save).** `renderMarkdown` writes the delimiters back. The
 * tokenizer's run cannot contain `=` (and cannot span a newline), so an inner
 * `==` can never close the span early.
 */
export const Highlight = Mark.create({
  name: 'highlight',
  inclusive: false,
  exitable: true,

  parseHTML() {
    return [{ tag: 'mark' }]
  },

  renderHTML({ HTMLAttributes }) {
    return ['mark', mergeAttributes(HTMLAttributes), 0]
  },

  markdownTokenName: 'highlight',

  parseMarkdown: (token, helpers) => ({
    mark: 'highlight',
    content: helpers.parseInline(
      (token.tokens as Array<{ type: string; text?: string }> | undefined) ?? [],
    ),
  }),

  renderMarkdown(node, helpers) {
    return `==${helpers.renderChildren(node)}==`
  },
})
