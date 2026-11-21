import type { Editor } from '@tiptap/core'
import type { SectionSettings } from '@genoffice/docx-engine'
import { getActiveSubEditor } from '../editor/active-editor'
import { enterSelectedTable } from '../editor/table-ops'
import { HORIZONTAL, selectedBoxDirection } from '../editor/textbox-direction'
import type { StringKey } from '../i18n/locale'
import { sectionVertical } from '../pagination-sections'

/** what Text Direction turns: the selected text box, the cells at the caret, else the caret's section */
export type TextDirectionTarget = 'textBox' | 'cells' | 'section'

export interface TextDirectionChoice {
  /** the payload of format-text-direction */
  value: string
  label: StringKey
}

export const SECTION_VERTICAL = 'tbRl'

/** per target, only the directions the page lays out */
const CHOICES: Record<TextDirectionTarget, readonly TextDirectionChoice[]> = {
  textBox: [
    { value: HORIZONTAL, label: 'ribbonTextDirectionHorizontal' },
    { value: 'eaVert', label: 'ribbonDocVertical' },
    { value: 'vert', label: 'ribbonTextDirectionRotate90' },
    { value: 'vert270', label: 'ribbonTextDirectionRotate270' },
    { value: 'wordArtVert', label: 'ribbonTextDirectionStacked' },
  ],
  cells: [
    { value: 'lrTb', label: 'ribbonTextDirectionHorizontal' },
    { value: 'tbRl', label: 'ribbonTextDirectionRotate90' },
    { value: 'btLr', label: 'ribbonTextDirectionRotate270' },
  ],
  section: [
    { value: 'lrTb', label: 'ribbonTextDirectionHorizontal' },
    { value: SECTION_VERTICAL, label: 'ribbonDocVertical' },
  ],
}

export const TEXT_DIRECTION_APPLY_TO: Record<TextDirectionTarget, StringKey> = {
  textBox: 'ribbonTextBox',
  cells: 'ribbonTextDirectionApplyCells',
  section: 'ribbonPaperApplySection',
}

export function textDirectionTarget(editor: Editor): TextDirectionTarget {
  if (selectedBoxDirection(editor) !== undefined) return 'textBox'
  return !getActiveSubEditor() && enterSelectedTable(editor.state) ? 'cells' : 'section'
}

export interface TextDirectionContext {
  target: TextDirectionTarget
  choices: readonly TextDirectionChoice[]
  /** the choice the target has now; null for a direction that is not offered */
  current: string | null
}

/** what the dialog and the ribbon show for a target: `box` is the text box's vert, `cell` the cell's direction */
export function textDirectionOf(
  target: TextDirectionTarget,
  own: { box?: string | null; cell?: string | null; section?: SectionSettings | null },
): TextDirectionContext {
  const choices = CHOICES[target]
  const value =
    target === 'textBox'
      ? (own.box ?? HORIZONTAL)
      : target === 'cells'
        ? (own.cell ?? 'lrTb')
        : sectionVertical(own.section ?? undefined)
          ? SECTION_VERTICAL
          : 'lrTb'
  return { target, choices, current: choices.some((c) => c.value === value) ? value : null }
}

export function textDirectionContext(
  editor: Editor,
  cell: string | null,
  section: SectionSettings | null,
): TextDirectionContext {
  return textDirectionOf(textDirectionTarget(editor), {
    box: selectedBoxDirection(editor),
    cell,
    section,
  })
}
