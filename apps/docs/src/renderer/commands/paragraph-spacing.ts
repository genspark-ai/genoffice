/** Word's Design > Paragraph Spacing built-ins: space after in twips, line spacing as a multiple */
export const PARAGRAPH_SPACINGS = [
  { id: 'none', afterTwips: 0, line: 1 },
  { id: 'compact', afterTwips: 80, line: 1 },
  { id: 'tight', afterTwips: 120, line: 1.15 },
  { id: 'open', afterTwips: 200, line: 1.15 },
  { id: 'relaxed', afterTwips: 120, line: 1.5 },
  { id: 'double', afterTwips: 160, line: 2 },
] as const

export type ParagraphSpacing = (typeof PARAGRAPH_SPACINGS)[number]
