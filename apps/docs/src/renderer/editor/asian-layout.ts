import type { Editor } from '@tiptap/core'
import type { EditorState } from '@tiptap/pm/state'

/** the text mark's `eaLayout` value: the run's w:eastAsianLayout, and how wide two lines in one sit */
export interface EaLayoutAttr {
  combine?: true
  vert?: true
  /** groups the runs of one two-lines-in-one region */
  id?: number
  /** half the length of the combined text, in em of its half-size font: the width of its two lines */
  w?: number
}

export function parseEaLayoutAttr(value: unknown): EaLayoutAttr | null {
  if (typeof value !== 'string' || !value) return null
  try {
    const v = JSON.parse(value) as EaLayoutAttr
    return v && typeof v === 'object' ? v : null
  } catch {
    return null
  }
}

/** how many em wide the two lines of a combined text are: half its characters, rounded up */
export const twoLinesWidthEm = (text: string): number =>
  Math.max(1, Math.ceil([...text].length / 2))

/** CSS of fitted text: its characters spread across the given width */
export function fitTextDecls(twips: number): string[] {
  return [
    'display:inline-block',
    `width:${twips / 20}pt`,
    'white-space:nowrap',
    'text-indent:0',
    'text-align-last:justify',
    'text-justify:inter-character',
  ]
}

/** CSS of an East Asian layout: two half-size lines in the space of one, or upright in vertical text */
export function eaLayoutDecls(ea: EaLayoutAttr, sizeHalfPoints: number | null): string[] {
  if (ea.combine) {
    return [
      'display:inline-block',
      `font-size:${sizeHalfPoints ? `${sizeHalfPoints / 4}pt` : '50%'}`,
      `width:${ea.w ?? 1}em`,
      'line-height:1.1',
      'vertical-align:middle',
      'white-space:normal',
      'word-break:break-all',
      'text-indent:0',
      'text-align:start',
    ]
  }
  return ea.vert ? ['text-combine-upright:all'] : []
}

/** the selection is one stretch of text inside one paragraph: what a layout is put on */
export function textStretch(state: EditorState): { from: number; to: number; text: string } | null {
  const { from, to, $from, $to } = state.selection
  if (from === to || !$from.sameParent($to) || !$from.parent.isTextblock) return null
  const text = state.doc.textBetween(from, to)
  return text ? { from, to, text } : null
}

/**
 * Turns two lines in one, or horizontal in vertical, on for the selected text,
 * or off when the selection already has it. The two exclude each other. False
 * when there is no stretch of text selected.
 */
export function toggleEaLayout(editor: Editor, kind: 'combine' | 'vert'): boolean {
  const stretch = textStretch(editor.state)
  if (!stretch) return false
  const current = parseEaLayoutAttr(editor.getAttributes('docTextStyle').eaLayout)
  const next: EaLayoutAttr = current?.[kind]
    ? {}
    : kind === 'combine'
      ? { combine: true, id: newRegionId(), w: twoLinesWidthEm(stretch.text) }
      : { vert: true, id: newRegionId() }
  return editor
    .chain()
    .focus()
    .setMark('docTextStyle', { eaLayout: JSON.stringify(next) })
    .run()
}

const newRegionId = (): number => Math.floor(Math.random() * 0x7fffffff) + 1

/** the width the selected text is fitted into, in twips; 0 when it is not fitted */
export const fitTextOf = (editor: Editor): number =>
  Math.max(0, Number(editor.getAttributes('docTextStyle').fitText) || 0)

/** fits the selected text into the width, or takes the fitting off for 0 */
export function setFitText(editor: Editor, twips: number): boolean {
  if (!textStretch(editor.state)) return false
  return editor
    .chain()
    .focus()
    .setMark('docTextStyle', { fitText: Math.max(0, Math.round(twips)) })
    .run()
}
