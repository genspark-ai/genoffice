/// Worksheet protection toggle: adds or removes `<sheetProtection>` (no
/// password support — unprotecting a password-protected sheet fails closed).
/// Also: workbook structure protection (`<workbookProtection>`) and
/// allow-edit ranges (`<protectedRanges>`), same no-password rules.

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

const WORKBOOK_PROTECTION_PATTERN =
  /<workbookProtection\b[^>]*\/>|<workbookProtection\b[^>]*>\s*<\/workbookProtection>/

/// Workbook structure lock in workbook.xml. Unlocking a password-protected
/// structure fails closed; other workbookProtection attributes stay verbatim.
export function applyWorkbookProtection(workbookXml: string, lockStructure: boolean): string {
  const existing = WORKBOOK_PROTECTION_PATTERN.exec(workbookXml)
  if (!lockStructure) {
    if (!existing) return workbookXml
    if (/\bworkbook(?:Password|HashValue)="/.test(existing[0])) {
      throw new SheetProtectionError(
        'The workbook structure is protected with a password — removing its protection ' +
          'is not supported.',
      )
    }
    const stripped = existing[0].replace(/\s+lockStructure="[^"]*"/, '')
    // Drop the element entirely once no protection attribute remains.
    const empty = /^<workbookProtection\s*(?:\/>|>\s*<\/workbookProtection>)$/.test(stripped)
    return (
      workbookXml.slice(0, existing.index) +
      (empty ? '' : stripped) +
      workbookXml.slice(existing.index + existing[0].length)
    )
  }
  if (existing) {
    if (/\blockStructure="(?:1|true)"/.test(existing[0])) return workbookXml
    const updated = /\slockStructure="/.test(existing[0])
      ? existing[0].replace(/(\s+lockStructure=)"[^"]*"/, '$1"1"')
      : existing[0].replace(/<workbookProtection\b/, '<workbookProtection lockStructure="1"')
    return (
      workbookXml.slice(0, existing.index) +
      updated +
      workbookXml.slice(existing.index + existing[0].length)
    )
  }
  // Schema order: workbookProtection follows fileVersion/fileSharing/
  // workbookPr/alternateContent and precedes bookViews/sheets.
  const element = '<workbookProtection lockStructure="1"/>'
  const anchor = /<bookViews\b|<sheets\b/.exec(workbookXml)
  if (!anchor) throw new SheetProtectionError('Workbook has no sheets element.')
  return workbookXml.slice(0, anchor.index) + element + workbookXml.slice(anchor.index)
}

export interface ProtectedRangeState {
  readonly name: string
  readonly sqref: string
}

const PROTECTED_RANGES_PATTERN =
  /<protectedRanges\b[^>]*\/>|<protectedRanges\b[^>]*>[\s\S]*?<\/protectedRanges>/

function escapeAttr(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/// Replaces the sheet's allow-edit ranges with the session's snapshot; an
/// empty set removes the element. Replacing password-protected ranges fails
/// closed (their hashes cannot be preserved through the rewrite).
export function applyProtectedRanges(
  worksheetXml: string,
  ranges: readonly ProtectedRangeState[],
): string {
  const existing = PROTECTED_RANGES_PATTERN.exec(worksheetXml)
  // securityDescriptor carries per-user permissions (attribute or child
  // element form); rewriting name+sqref only would silently fail open.
  if (existing && /\b(?:password|hashValue)="|securityDescriptor/.test(existing[0])) {
    throw new SheetProtectionError(
      'This sheet has password- or permission-protected edit ranges — editing them is not ' +
        'supported.',
    )
  }
  const stripped = existing
    ? worksheetXml.slice(0, existing.index) +
      worksheetXml.slice(existing.index + existing[0].length)
    : worksheetXml
  if (ranges.length === 0) return stripped
  const body = ranges
    .map(
      (range) =>
        `<protectedRange sqref="${escapeAttr(range.sqref)}" name="${escapeAttr(range.name)}"/>`,
    )
    .join('')
  const element = `<protectedRanges>${body}</protectedRanges>`
  // Schema order: protectedRanges follows sheetProtection (or sheetCalcPr/
  // sheetData when absent) and precedes scenarios/autoFilter.
  const anchor =
    /<sheetProtection\b[^>]*\/?>/.exec(stripped) ??
    /<sheetCalcPr\b[^>]*\/?>/.exec(stripped) ??
    /<\/sheetData>|<sheetData\b[^>]*\/>/.exec(stripped)
  if (!anchor) throw new SheetProtectionError('Worksheet has no sheetData element.')
  const at = anchor.index + anchor[0].length
  return stripped.slice(0, at) + element + stripped.slice(at)
}
