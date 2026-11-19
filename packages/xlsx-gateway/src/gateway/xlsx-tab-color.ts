/// Sheet tab color: rewrites `<sheetPr><tabColor rgb="FFRRGGBB"/></sheetPr>`
/// on a worksheet part, keeping the other sheetPr children verbatim.

const TAB_COLOR_PATTERN = /<tabColor\b[^>]*\/>|<tabColor\b[^>]*>[\s\S]*?<\/tabColor>/
const SHEET_PR_PATTERN = /<sheetPr\b[^>]*\/>|<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/

export function applyTabColor(worksheetXml: string, color: string | null): string {
  const rgb = color === null ? null : toArgb(color)
  const existing = SHEET_PR_PATTERN.exec(worksheetXml)
  if (rgb === null) {
    if (!existing) return worksheetXml
    const stripped = existing[0].replace(TAB_COLOR_PATTERN, '')
    const emptied = /^<sheetPr\b([^>]*)>\s*<\/sheetPr>$/.exec(stripped)
    const next = emptied?.[1]?.trim() === '' ? '' : stripped
    return worksheetXml.replace(SHEET_PR_PATTERN, () => next)
  }
  const element = `<tabColor rgb="${rgb}"/>`
  if (!existing) {
    return worksheetXml.replace(
      /(<worksheet\b[^>]*>)/,
      (_full, open: string) => `${open}<sheetPr>${element}</sheetPr>`,
    )
  }
  const sheetPr = existing[0]
  if (TAB_COLOR_PATTERN.test(sheetPr)) {
    return worksheetXml.replace(SHEET_PR_PATTERN, () =>
      sheetPr.replace(TAB_COLOR_PATTERN, () => element),
    )
  }
  // Schema order: tabColor comes first among sheetPr children.
  const opened = sheetPr.endsWith('/>')
    ? `${sheetPr.slice(0, -2)}>${element}</sheetPr>`
    : sheetPr.replace(/(<sheetPr\b[^>]*>)/, (_full, open: string) => `${open}${element}`)
  return worksheetXml.replace(SHEET_PR_PATTERN, () => opened)
}

function toArgb(color: string): string {
  const hex = /^#?([0-9a-f]{6})$/i.exec(color.trim())
  if (!hex) throw new Error(`Unsupported tab color "${color}".`)
  return `FF${hex[1]!.toUpperCase()}`
}
