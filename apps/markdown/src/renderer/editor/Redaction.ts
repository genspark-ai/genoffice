import { Mark, mergeAttributes, type Editor } from '@tiptap/core'
import { REDACT_MARK } from './redact'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    redaction: {
      /** withhold the current selection from the model */
      setRedaction: (label: string) => ReturnType
      /** stop withholding, keeping the text */
      unsetRedaction: () => ReturnType
    }
  }
}

/**
 * A mark over the real text, so the reader keeps their data and only the
 * model's view changes.
 *
 * `inclusive: false` and `exitable: true` stop the mark from growing or
 * bleeding as the caret moves around it; without them, typing at the edge of
 * a withheld span would silently pull neighbouring words in with it.
 *
 * ## The two directions
 *
 * **Out (a save).** `renderMarkdown` is how the serializer writes a mark —
 * Bold emits `**` through the same hook. It writes a `<span data-redaction>`
 * with the reader's own words inside, so the file shows the text and keeps
 * the marker. The model-facing replacement is *not* here: that happens on the
 * JSON, on the request path, in editor/redact.ts.
 *
 * **In (an open).** The markdown parser treats a `<span>` as an inert tag: it
 * keeps the words and drops the tag, so the mark would be lost on every
 * reopen. The inline patch in editor/inlineTokens.ts rewrites the persisted
 * span into a `redaction` token, and the manager routes that token to
 * `parseMarkdown` below. It returns a mark result rather than nodes because
 * that is the shape the manager's own applyMarkToContent expects.
 */
export const Redaction = Mark.create({
  name: REDACT_MARK,
  inclusive: false,
  exitable: true,
  spanning: false,

  addAttributes() {
    return {
      label: {
        default: 'private',
        parseHTML: (element) => element.getAttribute('data-label') ?? 'private',
        renderHTML: (attributes) => ({ 'data-label': String(attributes.label ?? 'private') }),
      },
    }
  },

  parseHTML() {
    // recovers the mark when the document is opened as HTML; the markdown path
    // goes through parseMarkdown below
    return [{ tag: 'span[data-redaction]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, { class: 'redact-span', 'data-redaction': '' }),
      0,
    ]
  },

  addCommands() {
    return {
      setRedaction:
        (label: string) =>
        ({ editor, tr, dispatch }) => {
          if (label.trim() === '') return false
          if (dispatch) {
            tr.addMark(
              editor.state.selection.from,
              editor.state.selection.to,
              this.type.create({ label }),
            )
            dispatch(tr)
          }
          return true
        },
      unsetRedaction:
        () =>
        ({ editor, tr, dispatch }) => {
          const { from, to, empty, $from } = editor.state.selection
          // no selection: clear the mark across the block the caret is in
          const range = empty
            ? { from: $from.pos, to: $from.pos + $from.parent.content.size }
            : { from, to }
          if (dispatch) {
            tr.removeMark(range.from, range.to, this.type)
            dispatch(tr)
          }
          return true
        },
    }
  },

  markdownTokenName: REDACT_MARK,

  /** recovers a persisted span; see the class comment */
  parseMarkdown: (token, helpers) => {
    const attrs = token.attrs as { label?: unknown } | undefined
    const label =
      typeof attrs?.label === 'string' && attrs.label.trim() !== '' ? attrs.label : 'private'
    return {
      mark: REDACT_MARK,
      attrs: { label },
      content: helpers.parseInline(
        (token.tokens as Array<{ type: string; text?: string }> | undefined) ?? [],
      ),
    }
  },

  /** writes the persisted span; see the class comment */
  renderMarkdown(node, helpers) {
    const label = String(node.attrs?.label ?? 'private')
    const text = helpers.renderChildren(node)
    return `<span class="redact-span" data-redaction="" data-label="${label.replace(/"/g, '&quot;')}">${text}</span>`
  },
})

/**
 * Whether any part of a range carries the redaction mark.
 *
 * The un-hide entry needs this before it can be offered: a reader who hid a
 * span weeks ago and no longer remembers where would otherwise have no way back
 * short of deleting the words, which is the state this branch exists to undo.
 *
 * Overlap is the test, not equality, because a selection that clips the edge of
 * a hidden span is still part of one — and `unsetRedaction` clears the whole
 * selection, so offering the entry on a partial selection does what it says.
 *
 * A bare caret answers false, and that needs no guard here: a range with
 * `from === to` holds no marked text, so ProseMirror reports false whatever the
 * document carries. The reader selects the words, as they did to hide them.
 */
export function hasRedactionIn(editor: Editor, from: number, to: number): boolean {
  const type = editor.state.schema.marks[REDACT_MARK]
  if (!type) return false
  return editor.state.doc.rangeHasMark(from, to, type)
}
