import { Extension } from '@tiptap/core'
import type { Mark as PmMark } from '@tiptap/pm/model'
import { Plugin, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { sdtCheckboxGlyphs } from '@genoffice/docx-engine'
import { TRACK_IGNORE } from './revisions'

/**
 * Click-to-toggle for the three kinds of checkbox a docx can hold: legacy
 * FORMCHECKBOX form fields, w14:checkbox content controls, and ballot-box
 * characters typed as plain text. The first two are one glyph run in the
 * editor whose write-back reads the state off the glyph, so toggling is
 * swapping the character; the marks stay so the field or control survives.
 * A plain ballot-box character (☐ ☑ ☒) is just text — no control survives a
 * round trip through editors that drop them — but it is still a box a reader
 * expects to click: the swap is an ordinary text edit (undoable, saved as the
 * new character). Other text never toggles.
 */

export interface CheckboxToggleStorage {
  /** forms protection: the body is read-only, filling the form fields is not */
  formsFill: boolean
}

declare module '@tiptap/core' {
  interface Storage {
    checkboxToggle: CheckboxToggleStorage
  }
}

const FORM_FIELD_GLYPHS = { checked: '\u2612', unchecked: '\u2610' }

/** plain ballot-box characters: box↔check and box/X↔box, as Word's glyphs read */
const PLAIN_GLYPHS: Record<string, string> = {
  '\u2610': '\u2611', // ☐ → ☑
  '\u2611': '\u2610', // ☑ → ☐
  '\u2612': '\u2610', // ☒ → ☐
}

export interface CheckboxToggleOptions {
  /** forms protection: plain text is locked, only the field and control glyphs toggle */
  allowPlainGlyphs?: boolean
}

function toggledGlyph(
  glyph: string,
  marks: readonly PmMark[],
  options: CheckboxToggleOptions = {},
): string | null {
  const field = marks.find((m) => m.type.name === 'instrField')
  const control = marks.find((m) => m.type.name === 'ctrlCheckbox')
  const glyphs = control
    ? sdtCheckboxGlyphs(String(control.attrs.sdtPr ?? ''))
    : field?.attrs.instr === 'FORMCHECKBOX'
      ? FORM_FIELD_GLYPHS
      : null
  if (!glyphs) {
    // a plain ballot-box character toggles as ordinary text; other plain
    // characters are prose, not boxes
    if (options.allowPlainGlyphs !== false) return PLAIN_GLYPHS[glyph] ?? null
    return null
  }
  if (glyph === glyphs.checked) return glyphs.unchecked
  if (glyph === glyphs.unchecked) return glyphs.checked
  return null
}

/**
 * The transaction toggling the checkbox whose glyph sits right after (`side`
 * 'after') or right before `pos`, or null when that character is not one. The
 * toggle is not a tracked change: Word records neither kind of checkbox flip.
 */
export function toggleCheckboxAt(
  state: EditorState,
  pos: number,
  side: 'before' | 'after',
  options: CheckboxToggleOptions = {},
): Transaction | null {
  const from = side === 'after' ? pos : pos - 1
  if (from < 0 || from >= state.doc.content.size) return null
  const node = state.doc.resolve(from).nodeAfter
  if (!node?.isText || !node.text) return null
  const glyph = node.text[0]
  const next = toggledGlyph(glyph, node.marks, options)
  if (next === null) return null
  const tr = state.tr.replaceWith(from, from + glyph.length, state.schema.text(next, node.marks))
  tr.setSelection(TextSelection.create(tr.doc, from + next.length))
  return tr.setMeta(TRACK_IGNORE, true).setMeta('addToHistory', true)
}

export const CheckboxToggleExtension = Extension.create<object, CheckboxToggleStorage>({
  name: 'checkboxToggle',

  addStorage() {
    return { formsFill: false }
  },

  addProseMirrorPlugins() {
    const storage = this.storage
    return [
      new Plugin({
        props: {
          handleClick(view, pos, event) {
            if (event.button !== 0 || event.detail > 1) return false
            if (event.shiftKey || event.altKey || event.ctrlKey || event.metaKey) return false
            if (!view.editable && !storage.formsFill) return false
            // pos is the caret slot nearest the click; the glyph under the
            // pointer is the character on the side the pointer is on
            let side: 'before' | 'after' = 'after'
            try {
              side = event.clientX >= view.coordsAtPos(pos).left ? 'after' : 'before'
            } catch {
              // a position without layout (mid-rerender): fall through
            }
            // forms protection: plain text is locked, only the field and
            // control glyphs fill
            const tr = toggleCheckboxAt(view.state, pos, side, {
              allowPlainGlyphs: view.editable,
            })
            if (!tr) return false
            view.dispatch(tr)
            return true
          },
        },
      }),
    ]
  },
})
