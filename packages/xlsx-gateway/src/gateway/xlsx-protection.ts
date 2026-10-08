/// Worksheet protection: `<sheetProtection>` with Excel's permission flags
/// and password hash. Also: workbook structure protection
/// (`<workbookProtection>`) and allow-edit ranges (`<protectedRanges>`).
import type { SheetPasswordHash } from './xlsx-protection-hash'

export class SheetProtectionError extends Error {}

export const SHEET_PROTECTION_PERMISSIONS = [
  'selectLockedCells',
  'selectUnlockedCells',
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'sort',
  'autoFilter',
  'pivotTables',
  'objects',
  'scenarios',
] as const

export type SheetProtectionPermission = (typeof SHEET_PROTECTION_PERMISSIONS)[number]
export type SheetProtectionAllow = Record<SheetProtectionPermission, boolean>

/// Attributes whose absence means "allowed" (schema default 0 = unlocked);
/// every other permission attribute defaults to 1 = locked.
const DEFAULT_ALLOWED = new Set<SheetProtectionPermission>([
  'selectLockedCells',
  'selectUnlockedCells',
  'objects',
  'scenarios',
])

/// What Excel's Protect Sheet dialog pre-checks.
export const DEFAULT_SHEET_PROTECTION_ALLOW: SheetProtectionAllow = {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  objects: false,
  scenarios: false,
}

export interface SheetProtectionWrite {
  readonly protected: boolean
  /// Absent together with `password`: keep the file's own flags (legacy toggle).
  readonly allow?: SheetProtectionAllow | undefined
  readonly password?: SheetPasswordHash | null | undefined
  /// The caller checked the file's password; unprotecting may drop the hash.
  readonly verified?: boolean | undefined
}

export interface ParsedSheetProtection {
  readonly protected: boolean
  readonly allow: SheetProtectionAllow
  readonly password: SheetPasswordHash | null
}

const ELEMENT_PATTERN = /<sheetProtection\b[^>]*\/>|<sheetProtection\b[^>]*>\s*<\/sheetProtection>/

function readAttribute(element: string, name: string): string | undefined {
  return new RegExp(`\\s${name}="([^"]*)"`).exec(element)?.[1]
}

function isTrue(value: string | undefined): boolean {
  return value === '1' || value === 'true'
}

function unescapeAttr(value: string): string {
  return value
    .replaceAll('&quot;', '"')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
}

export function parseSheetProtectionElement(element: string): ParsedSheetProtection {
  const allow = { ...DEFAULT_SHEET_PROTECTION_ALLOW }
  for (const permission of SHEET_PROTECTION_PERMISSIONS) {
    const value = readAttribute(element, permission)
    allow[permission] = DEFAULT_ALLOWED.has(permission)
      ? !isTrue(value)
      : value === '0' || value === 'false'
  }
  const hashValue = readAttribute(element, 'hashValue')
  const legacy = readAttribute(element, 'password')
  const password: SheetPasswordHash | null = hashValue
    ? {
        algorithmName: readAttribute(element, 'algorithmName') ?? 'SHA-512',
        hashValue: unescapeAttr(hashValue),
        saltValue: unescapeAttr(readAttribute(element, 'saltValue') ?? ''),
        spinCount: Number(readAttribute(element, 'spinCount') ?? '0') || 0,
      }
    : legacy
      ? { legacy }
      : null
  return { protected: isTrue(readAttribute(element, 'sheet')), allow, password }
}

/// Worksheet-level parse: null when the sheet has no `<sheetProtection>`.
export function readSheetProtection(worksheetXml: string): ParsedSheetProtection | null {
  const existing = ELEMENT_PATTERN.exec(worksheetXml)
  return existing ? parseSheetProtectionElement(existing[0]) : null
}

function escapeAttrValue(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/// Schema attribute order; lock-style flags are written only when they
/// differ from their default, matching Excel's own output.
export function sheetProtectionElement(
  allow: SheetProtectionAllow,
  password: SheetPasswordHash | null,
): string {
  const attributes: string[] = []
  if (password && 'legacy' in password) attributes.push(`password="${password.legacy}"`)
  else if (password) {
    attributes.push(
      `algorithmName="${escapeAttrValue(password.algorithmName)}"`,
      `hashValue="${escapeAttrValue(password.hashValue)}"`,
      `saltValue="${escapeAttrValue(password.saltValue)}"`,
      `spinCount="${password.spinCount}"`,
    )
  }
  attributes.push('sheet="1"')
  const order: readonly SheetProtectionPermission[] = [
    'objects',
    'scenarios',
    'formatCells',
    'formatColumns',
    'formatRows',
    'insertColumns',
    'insertRows',
    'insertHyperlinks',
    'deleteColumns',
    'deleteRows',
    'selectLockedCells',
    'sort',
    'autoFilter',
    'pivotTables',
    'selectUnlockedCells',
  ]
  for (const permission of order) {
    if (DEFAULT_ALLOWED.has(permission)) {
      if (!allow[permission]) attributes.push(`${permission}="1"`)
    } else if (allow[permission]) {
      attributes.push(`${permission}="0"`)
    }
  }
  return `<sheetProtection ${attributes.join(' ')}/>`
}

export function applySheetProtection(
  worksheetXml: string,
  state: boolean | SheetProtectionWrite,
): string {
  const write: SheetProtectionWrite = typeof state === 'boolean' ? { protected: state } : state
  const existing = ELEMENT_PATTERN.exec(worksheetXml)
  if (!write.protected) {
    if (!existing) return worksheetXml
    if (!write.verified && /\b(?:password|hashValue)="/.test(existing[0])) {
      throw new SheetProtectionError(
        'This sheet is protected with a password — removing its protection is not ' + 'supported.',
      )
    }
    return worksheetXml.replace(existing[0], '')
  }
  const element =
    write.allow === undefined && write.password === undefined
      ? null
      : sheetProtectionElement(
          write.allow ?? DEFAULT_SHEET_PROTECTION_ALLOW,
          write.password ?? null,
        )
  if (existing) {
    if (element) return worksheetXml.replace(existing[0], element)
    if (/\bsheet="(?:1|true)"/.test(existing[0])) return worksheetXml
    const updated = existing[0].includes(' sheet="')
      ? existing[0].replace(/ sheet="[^"]*"/, ' sheet="1"')
      : existing[0].replace(/<sheetProtection\b/, '<sheetProtection sheet="1"')
    return worksheetXml.replace(existing[0], updated)
  }
  // Schema order: the element follows sheetData (and sheetCalcPr when present).
  const anchor =
    /<sheetCalcPr\b[^>]*\/?>/.exec(worksheetXml) ??
    /<\/sheetData>|<sheetData\b[^>]*\/>/.exec(worksheetXml)
  if (!anchor) throw new SheetProtectionError('Worksheet has no sheetData element.')
  const at = anchor.index + anchor[0].length
  return (
    worksheetXml.slice(0, at) +
    (element ?? '<sheetProtection sheet="1" objects="1" scenarios="1"/>') +
    worksheetXml.slice(at)
  )
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
