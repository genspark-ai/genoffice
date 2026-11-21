import type { StyleParaProps } from '@genoffice/docx-engine'

const STYLE_FLAGS = [
  'widowControl',
  'keepNext',
  'keepLines',
  'suppressLineNumbers',
  'pageBreakBefore',
  'contextualSpacing',
  'suppressAutoHyphens',
] as const

export interface ParagraphDefaultInput {
  /** the paragraph's own attributes when the dialog opened */
  attrs: Record<string, unknown>
  align: 'left' | 'center' | 'right' | 'justify'
  line: {
    lineSpacing: number | null
    lineRule: 'exact' | 'atLeast' | null
    lineRawTwips: number | null
  }
  indentLeft: number
  indentRight: number
  firstLine: number | null
  spaceBefore: number
  spaceAfter: number
  flags: Record<string, boolean>
  initialFlags: Record<string, boolean>
}

/**
 * Set As Default: the default paragraph style takes what the paragraph declares
 * itself and what the dialog changed. A field showing an inherited value is
 * left out, so the style keeps its own.
 */
export function paragraphDefaultProps(d: ParagraphDefaultInput): StyleParaProps {
  const { attrs } = d
  const own = (key: string) => attrs[key] != null
  const twips = (key: string) => Number(attrs[key]) || 0
  const out: StyleParaProps = {}
  const startAlign = attrs.bidi === true ? 'right' : 'left'
  if (own('align') || d.align !== startAlign) out.align = d.align
  if (own('indentLeft') || d.indentLeft !== 0) out.indentLeftTwips = d.indentLeft
  if (own('indentRight') || d.indentRight !== 0) out.indentRightTwips = d.indentRight
  if (own('indentFirstLine') || d.firstLine != null) out.firstLineTwips = d.firstLine ?? 0
  if (own('spaceBefore') || d.spaceBefore !== 0) out.spaceBeforeTwips = d.spaceBefore
  if (own('spaceAfter') || d.spaceAfter !== 0) out.spaceAfterTwips = d.spaceAfter
  const { lineSpacing, lineRule, lineRawTwips } = d.line
  if (lineRule && lineRawTwips) out.lineFixed = { rule: lineRule, twips: lineRawTwips }
  else if (lineSpacing != null) out.lineSpacing = lineSpacing
  // single spacing is the absence of a value: written only when it replaces a declared one
  else if (own('lineSpacing') || own('lineRule') || twips('lineRawTwips')) out.lineSpacing = 1
  for (const key of STYLE_FLAGS) {
    const declared =
      key === 'pageBreakBefore' ? attrs[key] === true : typeof attrs[key] === 'boolean'
    if (declared || d.flags[key] !== d.initialFlags[key]) out[key] = d.flags[key]
  }
  return out
}
