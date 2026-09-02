/**
 * Effective print settings for the PDF export: the session's Page Layout
 * journal overlaid on the workbook's saved page setup (worksheet pageSetup /
 * pageMargins / printOptions / headerFooter plus the workbook-level
 * _xlnm.Print_Area / _xlnm.Print_Titles defined names).
 */
import { columnLabel, parseRange } from '../domain/cell-address'
import type { WorkbookPagePrintSettings } from '../shared/desktop-api'
import type { HeaderFooterParts, PageSetupJournalState, StructuralJournalOp } from './edit-journal'
import { fileRangeToScreenRange, fileToScreen } from './view-transform'

export interface PrintMargins {
  readonly left: number
  readonly right: number
  readonly top: number
  readonly bottom: number
  readonly header: number
  readonly footer: number
}

/// A page's printed header and footer (either half may be blank).
export interface HeaderFooterPair {
  readonly header: HeaderFooterParts | null
  readonly footer: HeaderFooterParts | null
}

/// One `&G` picture slot the file declares (legacyDrawingHF); the bytes are
/// fetched through readWorkbookMedia at export time.
export type HeaderFooterPictureSlot = NonNullable<
  WorkbookPagePrintSettings['headerFooterPictures']
>[number]

/// Concrete values the print layout consumes; every field resolved.
export interface EffectivePageSetup {
  readonly orientation: 'portrait' | 'landscape'
  /// OOXML paper-size code.
  readonly paperSize: number
  /// Percent; applies only when fitToPage is off.
  readonly scale: number
  /// Pages across / tall; 0 = automatic on that axis. Apply only when
  /// fitToPage is on.
  readonly fitToWidth: number
  readonly fitToHeight: number
  readonly fitToPage: boolean
  /// Inches.
  readonly margins: PrintMargins
  readonly printGridlines: boolean
  readonly printHeadings: boolean
  /// Plain A1 ranges to print ([] = the used range).
  readonly printAreas: readonly string[]
  /// Rows repeated at the top of every page ("1:2"), or null.
  readonly printTitles: string | null
  /// The odd-page (default) header/footer.
  readonly header: HeaderFooterParts | null
  readonly footer: HeaderFooterParts | null
  /// differentFirst: page 1's header/footer (blank halves print nothing),
  /// or null when page 1 uses the default.
  readonly firstPage: HeaderFooterPair | null
  /// differentOddEven: even pages' header/footer, or null when off.
  readonly evenPages: HeaderFooterPair | null
  /// scaleWithDoc (Excel's default): header/footer text and pictures follow
  /// the print scale like the sheet does.
  readonly headerFooterScaleWithDoc: boolean
  readonly headerFooterPictures: readonly HeaderFooterPictureSlot[]
}

/// Inches, mirroring the gateway's margin presets.
const MARGIN_PRESETS: Record<'normal' | 'wide' | 'narrow', PrintMargins> = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
}

/// The export wire caps margins at 3in per side.
const MAX_MARGIN_INCHES = 3

export interface FilePrintNames {
  readonly printArea?: string | undefined
  readonly printTitles?: string | undefined
}

/// File-space A1 areas → this session's screen space (envelope semantics,
/// like streamed merges). Fully deleted areas drop out; when nothing
/// survives the caller falls back to the used range.
function mapAreasToScreen(areas: readonly string[], ops: readonly StructuralJournalOp[]): string[] {
  if (ops.length === 0) return [...areas]
  const mapped: string[] = []
  for (const area of areas) {
    let bounds
    try {
      bounds = parseRange(area)
    } catch {
      continue
    }
    const screen = fileRangeToScreenRange(ops, bounds)
    if (screen === null) continue
    mapped.push(
      `${columnLabel(screen.startColumn)}${screen.startRow + 1}` +
        `:${columnLabel(screen.endColumn)}${screen.endRow + 1}`,
    )
  }
  return mapped
}

/// File-space title rows → screen space (row axis only — column edits must
/// not drop them). Deleted rows shrink the span; all deleted → null.
function mapTitleRowsToScreen(
  titles: string | null,
  ops: readonly StructuralJournalOp[],
): string | null {
  if (titles === null || ops.length === 0) return titles
  const match = /^([0-9]{1,7}):([0-9]{1,7})$/.exec(titles)
  if (!match) return null
  const rows: number[] = []
  for (let row = Number(match[1]) - 1; row <= Number(match[2]) - 1; row += 1) {
    const screen = fileToScreen(ops, 'row', row)
    if (screen !== null) rows.push(screen)
  }
  if (rows.length === 0) return null
  const start = Math.min(...rows)
  const end = Math.max(...rows)
  // Inserts/moves between title rows can stretch the surviving span past
  // the layout's 21-row cap (parseTitleRows would throw): drop the titles
  // instead, matching the cap semantics at parse time.
  if (end - start > 20) return null
  return `${start + 1}:${end + 1}`
}

export function resolveEffectivePageSetup(
  journal: PageSetupJournalState,
  file: WorkbookPagePrintSettings | null,
  names: FilePrintNames | null,
  /// This session's structural edits: file-sourced print names live in
  /// original workbook coordinates and must shift into screen space
  /// (journal-sourced values are already screen space).
  ops: readonly StructuralJournalOp[] = [],
): EffectivePageSetup {
  const clampMargin = (value: number): number => Math.min(Math.max(value, 0), MAX_MARGIN_INCHES)
  const fileMargins = file?.margins
  const margins =
    journal.margins !== undefined
      ? MARGIN_PRESETS[journal.margins]
      : fileMargins !== undefined
        ? {
            left: clampMargin(fileMargins.left),
            right: clampMargin(fileMargins.right),
            top: clampMargin(fileMargins.top),
            bottom: clampMargin(fileMargins.bottom),
            header: clampMargin(fileMargins.header),
            footer: clampMargin(fileMargins.footer),
          }
        : MARGIN_PRESETS.normal
  const fitToPage = journal.fitToPage ?? file?.fitToPage ?? false
  // OOXML default when fitToPage is on and fitToWidth/fitToHeight are
  // absent is 1 page on that axis.
  const fitToWidth = journal.fitToWidth ?? file?.fitToWidth ?? (file?.fitToPage === true ? 1 : 0)
  const fitToHeight = journal.fitToHeight ?? file?.fitToHeight ?? (file?.fitToPage === true ? 1 : 0)
  const printArea =
    journal.printArea !== undefined
      ? journal.printArea === null
        ? []
        : [journal.printArea]
      : mapAreasToScreen(printAreasFromFormula(names?.printArea), ops)
  const printTitles =
    journal.printTitles !== undefined
      ? journal.printTitles
      : mapTitleRowsToScreen(printTitleRowsFromFormula(names?.printTitles), ops)
  const header =
    journal.header !== undefined
      ? journal.header
      : file?.oddHeader !== undefined
        ? decodeHeaderFooter(file.oddHeader)
        : null
  const footer =
    journal.footer !== undefined
      ? journal.footer
      : file?.oddFooter !== undefined
        ? decodeHeaderFooter(file.oddFooter)
        : null
  // The session edits the odd header/footer only (like Excel's Page Setup
  // dialog); the file's first/even variants keep applying to their pages.
  const decodeOptional = (encoded: string | undefined): HeaderFooterParts | null =>
    encoded === undefined ? null : decodeHeaderFooter(encoded)
  const firstPage =
    file?.differentFirst === true
      ? { header: decodeOptional(file.firstHeader), footer: decodeOptional(file.firstFooter) }
      : null
  const evenPages =
    file?.differentOddEven === true
      ? { header: decodeOptional(file.evenHeader), footer: decodeOptional(file.evenFooter) }
      : null
  return {
    orientation: journal.orientation ?? file?.orientation ?? 'portrait',
    paperSize: journal.paperSize ?? file?.paperSize ?? 9,
    scale: journal.scale ?? file?.scale ?? 100,
    fitToWidth,
    fitToHeight,
    fitToPage,
    margins,
    printGridlines: journal.printGridlines ?? file?.printGridlines ?? false,
    printHeadings: journal.printHeadings ?? file?.printHeadings ?? false,
    printAreas: printArea,
    printTitles,
    header,
    footer,
    firstPage,
    evenPages,
    headerFooterScaleWithDoc: file?.headerFooterFixedSize !== true,
    headerFooterPictures: file?.headerFooterPictures ?? [],
  }
}

/// Splits a defined-name formula on commas outside quoted sheet names.
function splitAreas(formula: string): string[] {
  const parts: string[] = []
  let current = ''
  let quoted = false
  for (const character of formula) {
    if (character === "'") quoted = !quoted
    if (character === ',' && !quoted) {
      parts.push(current)
      current = ''
    } else {
      current += character
    }
  }
  parts.push(current)
  return parts
}

/// Strips the (optional) sheet qualifier and `$` anchors of one reference.
function plainReference(part: string): string {
  return part
    .slice(part.lastIndexOf('!') + 1)
    .replace(/\$/g, '')
    .trim()
}

/// `'S 1'!$A$1:$K$84,'S 1'!$M$1:$N$9` → ['A1:K84', 'M1:N9']. Anything the
/// print layout cannot crop to (full-column spans, 3-D refs, #REF!) yields
/// [] so the export falls back to the used range instead of dropping content.
export function printAreasFromFormula(formula: string | undefined): string[] {
  if (formula === undefined || formula === '') return []
  const areas: string[] = []
  for (const part of splitAreas(formula)) {
    const reference = plainReference(part).toUpperCase()
    if (/^[A-Z]{1,3}[0-9]{1,7}$/.test(reference)) {
      areas.push(`${reference}:${reference}`)
      continue
    }
    if (/^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$/.test(reference)) {
      areas.push(reference)
      continue
    }
    return []
  }
  return areas
}

/// `'S'!$1:$2` → '1:2'; column-repeat parts are skipped (unsupported), spans
/// beyond the layout's 21-row title cap are dropped.
export function printTitleRowsFromFormula(formula: string | undefined): string | null {
  if (formula === undefined || formula === '') return null
  for (const part of splitAreas(formula)) {
    const match = /^([0-9]{1,7}):([0-9]{1,7})$/.exec(plainReference(part))
    if (!match) continue
    const start = Number(match[1])
    const end = Number(match[2])
    if (start <= end && end - start <= 20) return `${start}:${end}`
  }
  return null
}

/// Excel's encoded header/footer → left/center/right parts. Field codes the
/// layout resolves (&P &N &D &T &F &A &G picture, && literal) stay verbatim;
/// formatting codes (font/size/color/bold/…) and unsupported codes (&Z
/// path) are stripped. Text before the first section marker is centered.
export function decodeHeaderFooter(encoded: string): HeaderFooterParts | null {
  const sections = { L: '', C: '', R: '' }
  let current: 'L' | 'C' | 'R' = 'C'
  let index = 0
  while (index < encoded.length) {
    const character = encoded[index] ?? ''
    if (character !== '&') {
      sections[current] += character
      index += 1
      continue
    }
    const next = encoded[index + 1]
    if (next === undefined) break
    if (next === '&') {
      sections[current] += '&&'
      index += 2
      continue
    }
    // &"Font Name,Style" — skip the whole quoted spec.
    if (next === '"') {
      const close = encoded.indexOf('"', index + 2)
      index = close === -1 ? encoded.length : close + 1
      continue
    }
    // &nn font size.
    if (/[0-9]/.test(next)) {
      index += 2
      while (index < encoded.length && /[0-9]/.test(encoded[index] ?? '')) index += 1
      continue
    }
    const code = next.toUpperCase()
    // &K followed by an RRGGBB or theme color spec.
    if (code === 'K') {
      index += 8
      continue
    }
    if (code === 'L' || code === 'C' || code === 'R') {
      current = code
      index += 2
      continue
    }
    if (
      code === 'P' ||
      code === 'N' ||
      code === 'D' ||
      code === 'T' ||
      code === 'F' ||
      code === 'A' ||
      code === 'G'
    ) {
      sections[current] += `&${code}`
    }
    index += 2
  }
  if (sections.L === '' && sections.C === '' && sections.R === '') return null
  return {
    ...(sections.L === '' ? {} : { left: sections.L }),
    ...(sections.C === '' ? {} : { center: sections.C }),
    ...(sections.R === '' ? {} : { right: sections.R }),
  }
}
