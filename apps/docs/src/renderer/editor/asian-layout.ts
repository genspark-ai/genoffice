import type { Editor } from '@tiptap/core'
import type { EditorState } from '@tiptap/pm/state'
import type { CombineBrackets } from '@genoffice/docx-engine'

/** the text mark's `eaLayout` value: the run's w:eastAsianLayout, and how wide two lines in one sit */
export interface EaLayoutAttr {
  combine?: true
  /** what encloses the two lines */
  brackets?: CombineBrackets
  vert?: true
  /** Word's Fit in line: kept for the file, Word draws the run the same with and without it */
  vertCompress?: true
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

/** CSS of two lines in one: two half-size lines in the space of one */
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
  // horizontal in vertical is drawn by the stylesheet ([data-ea-vert]): it depends on the line's direction
  return []
}

export const COMBINE_BRACKET_PAIRS: Record<CombineBrackets, [string, string]> = {
  round: ['(', ')'],
  square: ['[', ']'],
  angle: ['<', '>'],
  curly: ['{', '}'],
}

/** the selection is one stretch of text inside one paragraph: what a layout is put on */
export function textStretch(state: EditorState): { from: number; to: number; text: string } | null {
  const { from, to, $from, $to } = state.selection
  if (from === to || !$from.sameParent($to) || !$from.parent.isTextblock) return null
  const text = state.doc.textBetween(from, to)
  return text ? { from, to, text } : null
}

export type EaLayoutChoice =
  { kind: 'combine'; brackets: CombineBrackets | null } | { kind: 'vert'; fit: boolean }

/** the layout the selected text has, as its dialog shows it */
export const eaLayoutOf = (editor: Editor): EaLayoutAttr | null =>
  parseEaLayoutAttr(editor.getAttributes('docTextStyle').eaLayout)

/**
 * Puts two lines in one, or horizontal in vertical, on the selected text; null
 * takes the layout off. The two exclude each other. False when there is no
 * stretch of text selected.
 */
export function setEaLayout(editor: Editor, choice: EaLayoutChoice | null): boolean {
  const stretch = textStretch(editor.state)
  if (!stretch) return false
  const id = eaLayoutOf(editor)?.id ?? newRegionId()
  const next: EaLayoutAttr = !choice
    ? {}
    : choice.kind === 'combine'
      ? {
          combine: true,
          ...(choice.brackets ? { brackets: choice.brackets } : {}),
          id,
          w: twoLinesWidthEm(stretch.text),
        }
      : { vert: true, ...(choice.fit ? { vertCompress: true as const } : {}), id }
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
