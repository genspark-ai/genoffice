/// Worksheet protection toggle: adds or removes `<sheetProtection>` (no
/// password support — unprotecting a password-protected sheet fails closed).

export class SheetProtectionError extends Error {}

const ELEMENT_PATTERN = /<sheetProtection\b[^>]*\/>|<sheetProtection\b[^>]*>\s*<\/sheetProtection>/

export function applySheetProtection(worksheetXml: string, protect: boolean): string {
  const existing = ELEMENT_PATTERN.exec(worksheetXml)
  if (!protect) {
    if (!existing) return worksheetXml
    if (/\b(?:password|hashValue)="/.test(existing[0])) {
      throw new SheetProtectionError(
        'This sheet is protected with a password — removing its protection is not ' + 'supported.',
      )
    }
    return worksheetXml.replace(existing[0], '')
  }
  if (existing) {
    if (/\bsheet="(?:1|true)"/.test(existing[0])) return worksheetXml
    const updated = existing[0].includes(' sheet="')
      ? existing[0].replace(/ sheet="[^"]*"/, ' sheet="1"')
      : existing[0].replace(/<sheetProtection\b/, '<sheetProtection sheet="1"')
    return worksheetXml.replace(existing[0], updated)
  }
  // Excel's defaults when protecting without a password. Schema order: the
  // element follows sheetData (and sheetCalcPr when present).
  const element = '<sheetProtection sheet="1" objects="1" scenarios="1"/>'
  const anchor =
    /<sheetCalcPr\b[^>]*\/?>/.exec(worksheetXml) ??
    /<\/sheetData>|<sheetData\b[^>]*\/>/.exec(worksheetXml)
  if (!anchor) throw new SheetProtectionError('Worksheet has no sheetData element.')
  const at = anchor.index + anchor[0].length
  return worksheetXml.slice(0, at) + element + worksheetXml.slice(at)
}
