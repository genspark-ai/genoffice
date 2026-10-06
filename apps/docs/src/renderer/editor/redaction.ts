import { Mark } from '@tiptap/core'
import { REDACT_EL } from '@genoffice/docx-engine'
import { REDACT_MARK } from '../ai/redact'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    docsRedaction: {
      /** withhold the current selection from the model */
      setRedaction: (label: string) => ReturnType
      /** stop withholding, keeping the text */
      unsetRedaction: () => ReturnType
    }
  }
}

/** 0.75pt solid line, the same weight the docs border UI offers */
const BDR = { val: 'single', sz: 6, color: '6B4FC0' }

/**
 * A mark over the real words, so the reader keeps their data and only the
 * model's view changes.
 *
 * ## How it is stored
 *
 * The visual is a plain character border (`w:bdr`) on the run — real OOXML, so
 * Word renders the words with a line under them and round-trips the mark
 * without this app. The *label* rides alongside in a custom element inside the
 * run's `rPr`, which the docx engine keeps verbatim (`rawRPr`): the save path
 * compares each modelled group against the original bytes and writes back what
 * it does not manage, so an element the model never mentions survives
 * untouched. A comment in marks.ts already relies on the same property for
 * `w:em` and `w:textOutline`.
 *
 * `inclusive: false` stops the mark from swallowing what the reader types at
 * its edge, which would silently withhold their next word.
 *
 * ## What the model sees
 *
 * Nothing here: the placeholder is substituted in redact-view.ts, on the
 * request path only. A save writes the words and the border, never `{{…}}`.
 */
export const DocRedaction = Mark.create({
  name: REDACT_MARK,
  inclusive: false,
  exitable: true,
  spanning: false,

  addAttributes() {
    return {
      label: { default: 'private' },
      /**
       * The complete run properties: the w:bdr Word draws plus the custom
       * element carrying the label. The save path writes this back verbatim,
       * so the mark needs no separate modelling.
       */
      rawRPr: { default: null as string | null, rendered: false },
    }
  },

  parseHTML() {
    return [{ tag: 'span[data-redaction]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      {
        'data-redaction': '',
        'data-label': String(HTMLAttributes.label ?? 'private'),
        class: 'redact-span',
      },
      0,
    ]
  },

  addCommands() {
    return {
      setRedaction:
        (label: string) =>
        ({ editor, tr, dispatch }) => {
          const clean = label.trim()
          if (!clean) return false
          if (dispatch) {
            tr.addMark(
              editor.state.selection.from,
              editor.state.selection.to,
              this.type.create({
                label: clean,
                // the label has to live in the file, not only in the editor
                rawRPr: `<w:rPr><w:bdr w:val="${BDR.val}" w:sz="${BDR.sz}" w:color="${BDR.color}"/><${REDACT_EL} w:label="${escapeXmlAttr(clean)}"/></w:rPr>`,
              }),
            )
            dispatch(tr)
          }
          return true
        },
      unsetRedaction:
        () =>
        ({ editor, tr, dispatch }) => {
          const { from, to, empty, $from } = editor.state.selection
          // no selection: clear across the block the caret is in
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
})

function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
