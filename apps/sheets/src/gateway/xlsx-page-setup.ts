/// Page Layout writer: merges the user's session changes into the worksheet's
/// `<printOptions>`, `<pageMargins>`, `<pageSetup>` (and `sheetView` display
/// attributes), and maintains the sheet-scoped `_xlnm.Print_Area` defined
/// name in workbook.xml. Untouched attributes and elements stay verbatim.

import { parseSheetElements } from './xlsx-sheets'

export class PageSetupError extends Error {}

/// One printed header or footer: Excel's left/center/right sections. Text
/// carries Excel field codes verbatim (&P page, &N pages, &D date, &T time,
/// &F file name, &A sheet name, && literal ampersand).
export interface HeaderFooterParts {
  readonly left?: string | undefined
  readonly center?: string | undefined
  readonly right?: string | undefined
}

export interface SheetPageSetupState {
  readonly sheetName: string
  readonly orientation?: 'portrait' | 'landscape' | undefined
  readonly paperSize?: number | undefined
  readonly scale?: number | undefined
  readonly fitToWidth?: number | undefined
  readonly fitToHeight?: number | undefined
  readonly fitToPage?: boolean | undefined
  readonly margins?: 'normal' | 'wide' | 'narrow' | undefined
  readonly printGridlines?: boolean | undefined
  readonly printHeadings?: boolean | undefined
  readonly showGridlines?: boolean | undefined
  readonly showFormulas?: boolean | undefined
  readonly printArea?: string | null | undefined
  readonly printTitles?: string | null | undefined
  readonly frozenRows?: number | undefined
  readonly frozenColumns?: number | undefined
  /// Printed header/footer; null clears that half, undefined keeps it.
  readonly header?: HeaderFooterParts | null | undefined
  readonly footer?: HeaderFooterParts | null | undefined
}

/// Inches, using the standard margin presets.
const MARGIN_PRESETS = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
} as const

/// CT_Worksheet order: elements that may follow printOptions/pageMargins/
/// pageSetup — inserting before the first of these keeps the schema valid.
const AFTER_PAGE_SETUP =
  /<headerFooter\b|<rowBreaks\b|<colBreaks\b|<customProperties\b|<cellWatches\b|<ignoredErrors\b|<smartTags\b|<drawing\b|<legacyDrawing\b|<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b/

/// CT_Worksheet order: elements that may follow headerFooter.
const AFTER_HEADER_FOOTER =
  /<rowBreaks\b|<colBreaks\b|<customProperties\b|<cellWatches\b|<ignoredErrors\b|<smartTags\b|<drawing\b|<legacyDrawing\b|<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b/

function insertWorksheetElement(xml: string, element: string, anchor: RegExp): string {
  const found = anchor.exec(xml)
  if (found) return xml.slice(0, found.index) + element + xml.slice(found.index)
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new PageSetupError('Worksheet has no closing element.')
  return xml.slice(0, end) + element + xml.slice(end)
}

/// Sets (or removes, on null) attributes on the first match of `tag`,
/// creating the element when absent and any attribute is set.
function mergeElementAttrs(
  xml: string,
  tag: string,
  attrs: Record<string, string | null>,
  insertAnchor: RegExp,
): string {
  const entries = Object.entries(attrs)
  const pattern = new RegExp(`<${tag}\\b[^>]*?(/?)>`)
  const existing = pattern.exec(xml)
  if (existing) {
    let element = existing[0]
    for (const [name, value] of entries) {
      const attrPattern = new RegExp(` ${name}="[^"]*"`)
      if (value === null) {
        element = element.replace(attrPattern, '')
      } else if (attrPattern.test(element)) {
        element = element.replace(attrPattern, ` ${name}="${value}"`)
      } else {
        element = element.replace(new RegExp(`<${tag}\\b`), `<${tag} ${name}="${value}"`)
      }
    }
    return xml.slice(0, existing.index) + element + xml.slice(existing.index + existing[0].length)
  }
  const kept = entries.filter(([, value]) => value !== null)
  if (kept.length === 0) return xml
  const body = kept.map(([name, value]) => ` ${name}="${value}"`).join('')
  return insertWorksheetElement(xml, `<${tag}${body}/>`, insertAnchor)
}

/// Fit-to-page lives in `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`,
/// the first element of CT_Worksheet.
function setFitToPage(xml: string, enabled: boolean): string {
  const pageSetUpPr = /<pageSetUpPr\b[^>]*\/?>/.exec(xml)
  if (pageSetUpPr) {
    const updated = pageSetUpPr[0].includes(' fitToPage="')
      ? pageSetUpPr[0].replace(/ fitToPage="[^"]*"/, enabled ? ' fitToPage="1"' : '')
      : enabled
        ? pageSetUpPr[0].replace(/<pageSetUpPr\b/, '<pageSetUpPr fitToPage="1"')
        : pageSetUpPr[0]
    return xml.replace(pageSetUpPr[0], updated)
  }
  if (!enabled) return xml
  const sheetPr = /<sheetPr\b[^>]*(\/?)>/.exec(xml)
  if (sheetPr) {
    if (sheetPr[1] === '/') {
      const expanded = `${sheetPr[0].slice(0, -2)}><pageSetUpPr fitToPage="1"/></sheetPr>`
      return xml.replace(sheetPr[0], expanded)
    }
    const at = sheetPr.index + sheetPr[0].length
    return `${xml.slice(0, at)}<pageSetUpPr fitToPage="1"/>${xml.slice(at)}`
  }
  const worksheetOpen = /<worksheet\b[^>]*>/.exec(xml)
  if (!worksheetOpen) throw new PageSetupError('Worksheet has no root element.')
  const at = worksheetOpen.index + worksheetOpen[0].length
  return `${xml.slice(0, at)}<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>${xml.slice(at)}`
}

function setSheetViewAttr(xml: string, name: string, value: string | null): string {
  const sheetView = /<sheetView\b[^>]*?\/?>/.exec(xml)
  if (!sheetView) {
    // No sheetViews section: only a non-default value needs one.
    if (value === null) return xml
    const element = `<sheetViews><sheetView workbookViewId="0" ${name}="${value}"/></sheetViews>`
    const anchor = /<sheetFormatPr\b|<cols\b|<sheetData\b/.exec(xml)
    if (!anchor) throw new PageSetupError('Worksheet has no sheetData element.')
    return xml.slice(0, anchor.index) + element + xml.slice(anchor.index)
  }
  let element = sheetView[0]
  const attrPattern = new RegExp(` ${name}="[^"]*"`)
  if (value === null) {
    element = element.replace(attrPattern, '')
  } else if (attrPattern.test(element)) {
    element = element.replace(attrPattern, ` ${name}="${value}"`)
  } else {
    element = element.replace(/<sheetView\b/, `<sheetView ${name}="${value}"`)
  }
  return xml.slice(0, sheetView.index) + element + xml.slice(sheetView.index + sheetView[0].length)
}

/// Column index (0-based) → letters, for the pane's topLeftCell.
function paneColumnLabel(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

/// Replaces the first sheetView's frozen `<pane>` (both counts 0 removes it).
/// Pane-scoped `<selection>` elements are dropped too: they reference panes
/// that may no longer exist, and Excel restores a default selection anyway.
function setFrozenPane(xml: string, rows: number, columns: number): string {
  let result = xml
  const sheetView = /<sheetView\b[^>]*?(\/?)>/.exec(result)
  if (!sheetView) {
    if (rows === 0 && columns === 0) return result
    const element = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    const anchor = /<sheetFormatPr\b|<cols\b|<sheetData\b/.exec(result)
    if (!anchor) throw new PageSetupError('Worksheet has no sheetData element.')
    result = result.slice(0, anchor.index) + element + result.slice(anchor.index)
    return setFrozenPane(result, rows, columns)
  }
  if (sheetView[1] === '/') {
    // Self-closing sheetView becomes a container so it can hold the pane.
    const expanded = `${sheetView[0].slice(0, -2)}></sheetView>`
    result =
      result.slice(0, sheetView.index) +
      expanded +
      result.slice(sheetView.index + sheetView[0].length)
  }
  result = result.replace(/<pane\b[^>]*\/?>(?:[\s\S]*?<\/pane>)?/, '')
  result = result.replace(/<selection\b[^>]*\bpane="[^"]*"[^>]*\/?>/g, '')
  if (rows === 0 && columns === 0) return result
  const activePane = rows > 0 && columns > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight'
  const pane =
    '<pane' +
    (columns > 0 ? ` xSplit="${columns}"` : '') +
    (rows > 0 ? ` ySplit="${rows}"` : '') +
    ` topLeftCell="${paneColumnLabel(columns)}${rows + 1}"` +
    ` activePane="${activePane}" state="frozen"/>`
  const opened = /<sheetView\b[^>]*?>/.exec(result)
  if (!opened) throw new PageSetupError('Worksheet lost its sheetView element.')
  const at = opened.index + opened[0].length
  return result.slice(0, at) + pane + result.slice(at)
}

/// Left/center/right → Excel's `&L..&C..&R..` code string, empty sections
/// omitted; the caller XML-escapes the result.
function encodeHeaderFooterSections(parts: HeaderFooterParts): string {
  const sections: readonly (readonly [string, string | undefined])[] = [
    ['&L', parts.left],
    ['&C', parts.center],
    ['&R', parts.right],
  ]
  return sections
    .filter(([, text]) => text !== undefined && text !== '')
    .map(([marker, text]) => marker + (text ?? ''))
    .join('')
}

/// Assembles `<headerFooter>` from already-XML-escaped odd header/footer
/// texts; both empty means no element at all ('').
function assembleHeaderFooterXml(oddHeader: string, oddFooter: string): string {
  if (oddHeader === '' && oddFooter === '') return ''
  return (
    '<headerFooter>' +
    (oddHeader === '' ? '' : `<oddHeader>${oddHeader}</oddHeader>`) +
    (oddFooter === '' ? '' : `<oddFooter>${oddFooter}</oddFooter>`) +
    '</headerFooter>'
  )
}

/// Header/footer parts → the worksheet `<headerFooter>` fragment ('' when
/// both are empty). Field codes and literal ampersands alike XML-escape to
/// `&amp;` — Excel reads the parsed `&` back as header/footer code.
export function buildHeaderFooterXml(
  header: HeaderFooterParts | null | undefined,
  footer: HeaderFooterParts | null | undefined,
): string {
  return assembleHeaderFooterXml(
    header ? escapeXml(encodeHeaderFooterSections(header)) : '',
    footer ? escapeXml(encodeHeaderFooterSections(footer)) : '',
  )
}

/// Replaces the first `<headerFooter>` (or inserts one after pageSetup).
/// An undefined half keeps its existing odd text verbatim, null (or
/// all-empty parts) clears it; both empty removes the element
/// entirely. Even/first-page variants and the element's
/// attributes are dropped: the session edits the odd header/footer, which
/// applies to every page once differentOddEven/differentFirst are gone.
function setHeaderFooter(
  xml: string,
  header: HeaderFooterParts | null | undefined,
  footer: HeaderFooterParts | null | undefined,
): string {
  const existing = /<headerFooter\b[^>]*?(?:\/>|>[\s\S]*?<\/headerFooter>)/.exec(xml)
  const keep = (tag: 'oddHeader' | 'oddFooter'): string => {
    if (!existing) return ''
    const section = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`).exec(existing[0])
    return section?.[1] ?? ''
  }
  const oddHeader =
    header === undefined
      ? keep('oddHeader')
      : header === null
        ? ''
        : escapeXml(encodeHeaderFooterSections(header))
  const oddFooter =
    footer === undefined
      ? keep('oddFooter')
      : footer === null
        ? ''
        : escapeXml(encodeHeaderFooterSections(footer))
  const element = assembleHeaderFooterXml(oddHeader, oddFooter)
  if (existing) {
    return xml.slice(0, existing.index) + element + xml.slice(existing.index + existing[0].length)
  }
  if (element === '') return xml
  return insertWorksheetElement(xml, element, AFTER_HEADER_FOOTER)
}

export function applyPageSetupState(worksheetXml: string, state: SheetPageSetupState): string {
  let xml = worksheetXml

  if (state.frozenRows !== undefined || state.frozenColumns !== undefined) {
    xml = setFrozenPane(xml, state.frozenRows ?? 0, state.frozenColumns ?? 0)
  }

  if (state.showGridlines !== undefined) {
    // showGridLines defaults to true; write "0" to hide, drop the attribute
    // to restore the default.
    xml = setSheetViewAttr(xml, 'showGridLines', state.showGridlines ? null : '0')
  }

  if (state.showFormulas !== undefined) {
    // showFormulas defaults to false; drop the attribute to restore it.
    xml = setSheetViewAttr(xml, 'showFormulas', state.showFormulas ? '1' : null)
  }

  const printOptions: Record<string, string | null> = {}
  if (state.printGridlines !== undefined) {
    printOptions.gridLines = state.printGridlines ? '1' : null
  }
  if (state.printHeadings !== undefined) {
    printOptions.headings = state.printHeadings ? '1' : null
  }
  if (Object.keys(printOptions).length > 0) {
    xml = mergeElementAttrs(
      xml,
      'printOptions',
      printOptions,
      /<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/,
    )
  }

  if (state.margins !== undefined) {
    const preset = MARGIN_PRESETS[state.margins]
    xml = mergeElementAttrs(
      xml,
      'pageMargins',
      {
        left: String(preset.left),
        right: String(preset.right),
        top: String(preset.top),
        bottom: String(preset.bottom),
        header: String(preset.header),
        footer: String(preset.footer),
      },
      /<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/,
    )
  }

  const pageSetup: Record<string, string | null> = {}
  if (state.orientation !== undefined) pageSetup.orientation = state.orientation
  if (state.paperSize !== undefined) pageSetup.paperSize = String(state.paperSize)
  if (state.scale !== undefined) {
    pageSetup.scale = state.scale === 100 ? null : String(state.scale)
  }
  if (state.fitToWidth !== undefined) {
    pageSetup.fitToWidth = state.fitToWidth === 1 ? null : String(state.fitToWidth)
  }
  if (state.fitToHeight !== undefined) {
    pageSetup.fitToHeight = state.fitToHeight === 1 ? null : String(state.fitToHeight)
  }
  if (Object.keys(pageSetup).length > 0) {
    xml = mergeElementAttrs(xml, 'pageSetup', pageSetup, AFTER_PAGE_SETUP)
  }
  if (state.fitToPage !== undefined) {
    xml = setFitToPage(xml, state.fitToPage)
  }
  if (state.header !== undefined || state.footer !== undefined) {
    xml = setHeaderFooter(xml, state.header, state.footer)
  }
  return xml
}

/// Maintains the sheet-scoped `_xlnm.Print_Area` / `_xlnm.Print_Titles`
/// names. localSheetId is positional, so entries follow workbook.xml's final
/// `<sheet>` order.
export function applyPrintAreas(
  workbookXml: string,
  areas: readonly {
    sheetName: string
    printArea?: string | null | undefined
    printTitles?: string | null | undefined
  }[],
): string {
  let xml = workbookXml
  // Parse elements instead of pattern-matching serialized XML: attribute
  // order and entity encoding vary by producer.
  const sheetOrder = parseSheetElements(xml).map((element) => element.name)
  for (const { sheetName, printArea, printTitles } of areas) {
    const sheetIndex = sheetOrder.indexOf(sheetName)
    if (sheetIndex === -1) {
      throw new PageSetupError(`Sheet "${sheetName}" was not found in the workbook.`)
    }
    const quoted = `'${sheetName.replace(/'/g, "''")}'`
    if (printArea !== undefined) {
      xml = setSheetScopedName(
        xml,
        '_xlnm.Print_Area',
        sheetIndex,
        printArea === null ? null : `${quoted}!${toAbsoluteRange(printArea)}`,
      )
    }
    if (printTitles !== undefined) {
      xml = setSheetScopedName(
        xml,
        '_xlnm.Print_Titles',
        sheetIndex,
        printTitles === null ? null : `${quoted}!${toAbsoluteRowSpan(printTitles)}`,
      )
    }
  }
  // An emptied definedNames section is dropped entirely rather than written empty.
  return xml.replace(/<definedNames>\s*<\/definedNames>/, '')
}

function setSheetScopedName(
  workbookXml: string,
  name: string,
  sheetIndex: number,
  reference: string | null,
): string {
  const escapedName = name.replace(/\./g, '\\.')
  const namePattern = new RegExp(
    `<definedName[^>]*name="${escapedName}"[^>]*localSheetId="${sheetIndex}"[^>]*>[\\s\\S]*?</definedName>` +
      `|<definedName[^>]*localSheetId="${sheetIndex}"[^>]*name="${escapedName}"[^>]*>[\\s\\S]*?</definedName>`,
  )
  let xml = workbookXml.replace(namePattern, '')
  if (reference === null) return xml
  const element =
    `<definedName name="${name}" localSheetId="${sheetIndex}">` +
    `${escapeXml(reference)}</definedName>`
  const section = /<definedNames\b[^>]*>/.exec(xml)
  if (section) {
    const at = section.index + section[0].length
    xml = xml.slice(0, at) + element + xml.slice(at)
  } else {
    const sheetsEnd = /<\/sheets>/.exec(xml)
    if (!sheetsEnd) throw new PageSetupError('workbook.xml has no sheets section.')
    const at = sheetsEnd.index + sheetsEnd[0].length
    xml = `${xml.slice(0, at)}<definedNames>${element}</definedNames>${xml.slice(at)}`
  }
  return xml
}

/// "1:3" → "$1:$3" (title rows repeated at the top of each page).
function toAbsoluteRowSpan(rows: string): string {
  const match = /^\$?(\d{1,7}):\$?(\d{1,7})$/.exec(rows)
  if (!match || Number(match[1]) > Number(match[2])) {
    throw new PageSetupError(`Invalid print titles "${rows}".`)
  }
  return `$${match[1]}:$${match[2]}`
}

/// "A1:C10" → "$A$1:$C$10" (already-absolute refs pass through).
function toAbsoluteRange(range: string): string {
  if (!/^[$A-Za-z0-9:]+$/.test(range)) {
    throw new PageSetupError(`Invalid print area "${range}".`)
  }
  return range
    .split(':')
    .map((part) => part.replace(/^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/, '$$$1$$$2'))
    .join(':')
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
