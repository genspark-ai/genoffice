import JSZip from 'jszip'
import { parseChartPartXml } from './chart'
import { findInkRuns, stripInkRuns } from './ink'
import { computeListMarkers, type ListItemRef } from './list-markers'
import { ommlFragmentsOf, ommlToLatex, ommlToMathML } from './math'
import { splitXmlChildren } from './generate'
import { NOTE_PART_PATH, parseNotesXml } from './notes'
import { scanBody, type BodyElement } from './scan'
import { findSourcesPart, parseSourcesXml } from './sources'
import { decodeSymbolChar, decodeSymbolText } from './symbol-fonts'
import { THEME_PART_PATH, readThemeColors, readThemeFonts, resolveThemeColor } from './theme'
import { TOTAL_PAGES_MARK } from './types'
import type {
  Block,
  ChartDisplay,
  CommentInfo,
  DocDefaults,
  DocProtection,
  FieldDisplay,
  InkInfo,
  NoteInfo,
  HfImage,
  HfParagraph,
  HfPartInfo,
  NumberingDef,
  NumberingLevel,
  ParaFormat,
  ParsedDoc,
  CellBorders,
  CellMargins,
  RevisionInfo,
  Run,
  SdtShell,
  SourceInfo,
  StyleDisplay,
  StyleInfo,
  TableBorders,
  TableStyleDisplay,
  TableCell,
  TableModel,
  TextboxDisplay,
  TextboxParaDisplay,
  ThemeColors,
  ThemeFonts,
} from './types'
import { readWatermarkText } from './watermark'
import {
  attrsOf,
  boolProp,
  childrenOf,
  childrenThroughSdt,
  findChild,
  findChildren,
  nameOf,
  serializeXNode,
  textOf,
  underlineProp,
  xmlParser,
  type XNode,
} from './xml-utils'

const MAX_ZIP_PARTS = 10000
const MAX_PART_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
const MAX_TOTAL_UNCOMPRESSED_BYTES = 1.5 * 1024 * 1024 * 1024

/**
 * Reject zip bombs before any part is inflated, using the declared
 * uncompressed sizes from the central directory (JSZip keeps them in
 * the lazy `_data` compressed object).
 */
export function assertZipWithinLimits(zip: JSZip): void {
  const files = Object.values(zip.files).filter((f) => !f.dir)
  if (files.length > MAX_ZIP_PARTS) {
    throw new Error(`docx rejected: ${files.length} parts exceeds the ${MAX_ZIP_PARTS} limit`)
  }
  let total = 0
  for (const file of files) {
    const size =
      (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0
    if (size > MAX_PART_UNCOMPRESSED_BYTES) {
      throw new Error(
        `docx rejected: part ${file.name} declares ${size} uncompressed bytes ` +
          `(limit ${MAX_PART_UNCOMPRESSED_BYTES})`,
      )
    }
    if (size > 0) total += size
  }
  if (total > MAX_TOTAL_UNCOMPRESSED_BYTES) {
    throw new Error(
      `docx rejected: total uncompressed size ${total} exceeds the ` +
        `${MAX_TOTAL_UNCOMPRESSED_BYTES} limit`,
    )
  }
}

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  emf: 'image/emf',
  wmf: 'image/wmf',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}

interface RelInfo {
  target: string
  type: string
  targetMode?: string
}

export interface ParseExtras {
  /** ranges of top-level body elements, aligned with docxIndex */
  elements: BodyElement[]
  /** original XML of chart parts referenced by chart blocks (partPath -> xml) */
  chartParts: Record<string, string>
}

export async function parseDocx(bytes: Uint8Array): Promise<ParsedDoc & { extras: ParseExtras }> {
  const zip = await JSZip.loadAsync(bytes)
  assertZipWithinLimits(zip)
  const docFile = zip.file('word/document.xml')
  if (!docFile) throw new Error('not a docx: missing word/document.xml')
  const documentXml = await docFile.async('string')

  const theme = await parseTheme(zip)
  const { styles, docDefaults } = await parseStyles(zip, theme.colors, theme.fonts)
  const headingStyleIds = new Map<number, string>()
  let listParagraphStyleId: string | undefined
  for (const info of styles.values()) {
    if (info.headingLevel && !headingStyleIds.has(info.headingLevel)) {
      headingStyleIds.set(info.headingLevel, info.styleId)
    }
    if (!listParagraphStyleId && /^listparagraph$/i.test(info.styleId)) {
      listParagraphStyleId = info.styleId
    }
  }

  const rels = await parseRels(zip, 'word/_rels/document.xml.rels')
  const { formats: numFormats, defs: numbering } = await parseNumbering(zip)
  const comments = await parseComments(zip)
  const protection = await parseProtection(zip)
  const footnotes = await parseNotesPart(zip, 'footnote')
  const endnotes = await parseNotesPart(zip, 'endnote')
  const sources = await parseSources(zip)

  // display numbers for note reference markers, by part order
  const noteNumbers = new Map<string, number>()
  footnotes.forEach((n, i) => noteNumbers.set(`footnote:${n.id}`, i + 1))
  endnotes.forEach((n, i) => noteNumbers.set(`endnote:${n.id}`, i + 1))

  const scan = scanBody(documentXml)
  const elements: BodyElement[] = []
  const blocks: Block[] = []
  const chartParts: Record<string, string> = {}
  const buildCtx: BuildContext = {
    zip,
    styles,
    rels,
    numFormats,
    numbering,
    chartParts,
    noteNumbers,
    themeColors: theme.colors,
    themeFonts: theme.fonts,
  }
  let sdtGroupSeq = 0
  for (const el of scan.elements) {
    const xml = documentXml.slice(el.start, el.end)
    const sdtParts = el.name === 'w:sdt' ? splitSdtParts(xml) : null
    if (sdtParts) {
      const meta = sdtMeta(xml)
      const group = sdtGroupSeq++
      for (const part of sdtParts) {
        const i = elements.length
        elements.push({ name: part.name, start: el.start + part.start, end: el.start + part.end })
        const childXml = xml.slice(part.childStart, part.childEnd)
        const block = await buildBlock(
          { name: part.name, start: 0, end: childXml.length },
          i,
          childXml,
          buildCtx,
        )
        block.originalXml = xml.slice(part.start, part.end)
        block.sdtShell = {
          ...meta,
          openXml: xml.slice(part.start, part.childStart),
          closeXml: xml.slice(part.childEnd, part.end),
          group,
        }
        if (!block.label) block.label = meta.alias || meta.tag || 'Content control'
        blocks.push(block)
      }
      continue
    }
    const i = elements.length
    elements.push(el)
    blocks.push(await buildBlock(el, i, xml, buildCtx))
  }
  applyTocEntryNumbers(blocks, numbering)

  const header = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'header',
    'default',
    theme.colors,
  )
  const footer = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'footer',
    'default',
    theme.colors,
  )
  const headerFirst = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'header',
    'first',
    theme.colors,
  )
  const footerFirst = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'footer',
    'first',
    theme.colors,
  )
  const headerEven = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'header',
    'even',
    theme.colors,
  )
  const footerEven = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'footer',
    'even',
    theme.colors,
  )
  const titlePg = /<w:titlePg\s*\/>/.test(documentXml)
  const evenAndOddHeaders = await parseEvenAndOddHeaders(zip)
  const hfParts = await parseAllHfParts(zip, rels, theme.colors)

  // ink annotations (freehand strokes): our own anchored floating pictures, restored
  // into an editable overlay layer instead of being shown as image blocks
  const inks: InkInfo[] = []
  for (const block of blocks) {
    if (block.docxIndex === null || !block.originalXml) continue
    for (const run of findInkRuns(block.originalXml)) {
      inks.push({
        anchorIndex: block.docxIndex,
        offsetXPx: run.offsetXPx,
        offsetYPx: run.offsetYPx,
        widthPx: run.widthPx,
        heightPx: run.heightPx,
        dataUrl: run.embedRId ? await mediaDataUrl(zip, rels, run.embedRId) : null,
        payload: run.payload,
      })
    }
  }

  return {
    blocks,
    comments,
    protection,
    footnotes,
    endnotes,
    sources,
    inks,
    themeFonts: theme.fonts,
    themeColors: theme.colors,
    watermarkText: header?.watermark ?? null,
    headerText: header?.text ?? null,
    headerParas: header?.paras ?? null,
    footerParas: footer?.paras ?? null,
    headerImages: header?.images ?? null,
    footerImages: footer?.images ?? null,
    footerText: footer?.text ?? null,
    footerHasPageNumber: footer?.hasPageNumber ?? false,
    headerHasPageNumber: header?.hasPageNumber ?? false,
    titlePg,
    evenAndOddHeaders,
    headerFirst: hfPartInfo(headerFirst),
    footerFirst: hfPartInfo(footerFirst),
    headerEven: hfPartInfo(headerEven),
    footerEven: hfPartInfo(footerEven),
    hfParts,
    styles,
    docDefaults,
    headingStyleIds,
    listParagraphStyleId,
    numbering,
    internal: {
      originalBytes: bytes,
      documentXml,
      bodyInnerStart: scan.innerStart,
      bodyInnerEnd: scan.innerEnd,
    },
    extras: { elements, chartParts },
  }
}

interface BuildContext {
  zip: JSZip
  styles: Map<string, StyleInfo>
  rels: Map<string, RelInfo>
  numFormats: Map<string, 'bullet' | 'ordered'>
  /** full per-level definitions, for level-aware kind classification */
  numbering: Map<string, NumberingDef>
  /** collector: chart part XML seen while building blocks (partPath -> xml) */
  chartParts: Record<string, string>
  /** "footnote:<id>" / "endnote:<id>" -> display number */
  noteNumbers: Map<string, number>
  /** live palette for w:themeColor resolution, null when the doc has no theme */
  themeColors?: ThemeColors | null
  /** theme font scheme for w:asciiTheme/... resolution */
  themeFonts?: ThemeFonts | null
}

/** numbering reference of a paragraph: direct w:numPr, falling back to the pStyle's
 * numPr (ListBullet/ListNumber style-driven lists carry no numPr on the paragraph).
 * numId="0" is Word's explicit "no numbering" and yields undefined. */
function listRefOf(
  ctx: BuildContext,
  pPr: XNode | undefined,
  styleId: string | undefined,
): { numId: string; ilvl: number } | undefined {
  const numPr = pPr ? findChild(pPr, 'w:numPr') : undefined
  const directNumId = numPr ? attrsOf(findChild(numPr, 'w:numId') ?? {})['w:val'] : undefined
  const directIlvl = numPr ? attrsOf(findChild(numPr, 'w:ilvl') ?? {})['w:val'] : undefined
  if (directNumId === '0') return undefined
  const styleNumPr = styleId ? ctx.styles.get(styleId)?.numPr : undefined
  const numId = directNumId ?? styleNumPr?.numId
  if (!numId) return undefined
  const ilvl = directIlvl !== undefined ? parseInt(directIlvl, 10) || 0 : (styleNumPr?.ilvl ?? 0)
  return { numId, ilvl }
}

/** bullet/ordered classification of one list level (mixed lists differ per ilvl) */
function listKindOf(ctx: BuildContext, numId: string, ilvl: number): 'bullet' | 'ordered' {
  const fmt = ctx.numbering.get(numId)?.levels[ilvl]?.numFmt
  if (fmt !== undefined) return fmt === 'bullet' ? 'bullet' : 'ordered'
  return ctx.numFormats.get(numId) ?? 'bullet'
}

/** w:color -> display hex; w:themeColor resolves against the live palette (beats stale w:val) */
function colorFrom(container: XNode | undefined, theme?: ThemeColors | null): string | undefined {
  if (!container) return undefined
  const a = attrsOf(findChild(container, 'w:color') ?? {})
  if (a['w:themeColor'] && theme) {
    const resolved = resolveThemeColor(
      a['w:themeColor'],
      theme,
      a['w:themeTint'],
      a['w:themeShade'],
    )
    if (resolved) return resolved
  }
  const val = a['w:val']
  return val && val !== 'auto' ? val : undefined
}

/**
 * Extract a SdtShell from a raw <w:sdt>…</w:sdt> XML string.
 * Returns the shell metadata + the first <w:p> found in <w:sdtContent>,
 * or null if no usable paragraph is found.
 */
function parseSdtBlock(sdtXml: string): { shell: SdtShell; pXml: string } | null {
  // --- find the sdtContent region ---
  // sdtContent is a direct child of w:sdt; its content may be a w:p or w:tbl
  const contentOpen = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  if (!contentOpen) return null

  const contentTagEnd = contentOpen.index + contentOpen[0].length
  // find the matching </w:sdtContent>
  const contentClose = sdtXml.indexOf('</w:sdtContent>', contentTagEnd)
  if (contentClose === -1) return null

  // extract the first w:p inside sdtContent
  const innerContent = sdtXml.slice(contentTagEnd, contentClose)
  const pStart = innerContent.search(/<w:p[\s/>]/)
  if (pStart === -1) return null

  // find matching closing </w:p>
  let depth = 0
  let pEnd = -1
  const tagRe = /<\/?w:p(?=[\s/>])/g
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(innerContent)) !== null) {
    if (m[0].startsWith('</')) {
      if (depth === 1) {
        pEnd = m.index + m[0].length + 1 // include '>'
        break
      }
      depth--
    } else {
      depth++
    }
  }
  if (pEnd === -1) {
    // self-closing <w:p/> or no </w:p> found — unlikely but safe
    const selfClose = /<w:p\/>/.exec(innerContent)
    pEnd = selfClose ? selfClose.index + selfClose[0].length : innerContent.length
  }
  const pXml = innerContent.slice(pStart, pEnd)

  // openXml = everything from start of sdt to end of <w:sdtContent> open tag
  const openXml = sdtXml.slice(0, contentTagEnd)
  const closeXml = sdtXml.slice(contentClose) // "</w:sdtContent></w:sdt>"

  return {
    shell: { ...sdtMeta(sdtXml), openXml, closeXml },
    pXml,
  }
}

function sdtMeta(sdtXml: string): Pick<SdtShell, 'alias' | 'tag' | 'controlType'> {
  const sdtPrXml = /<w:sdtPr>([\s\S]*?)<\/w:sdtPr>/.exec(sdtXml)?.[1] ?? ''
  const alias = /w:val="([^"]*)"/.exec(/<w:alias[^>]*>/.exec(sdtPrXml)?.[0] ?? '')?.[1] ?? ''
  const tag = /w:val="([^"]*)"/.exec(/<w:tag[^>]*>/.exec(sdtPrXml)?.[0] ?? '')?.[1] ?? ''
  let controlType: SdtShell['controlType'] = 'text'
  if (/<w:date[\s/>]/.test(sdtPrXml)) controlType = 'date'
  else if (/<w:dropDownList[\s/>]|<w:comboBox[\s/>]/.test(sdtPrXml)) controlType = 'dropdown'
  else if (/<w:checkbox[\s/>]/.test(sdtPrXml)) controlType = 'checkbox'
  else if (/<w:text[\s/>]|<w:richText[\s/>]/.test(sdtPrXml)) controlType = 'text'
  return { alias, tag, controlType }
}

interface SdtPart {
  name: string
  /** slice of the whole <w:sdt> xml owned by this part (parts partition the sdt exactly) */
  start: number
  end: number
  /** the w:p / w:tbl child inside [start, end) */
  childStart: number
  childEnd: number
}

/**
 * When <w:sdtContent> holds several top-level w:p / w:tbl children (Word TOC,
 * multi-paragraph rich-text controls), each child becomes its own block.
 * Returns null for 0-1 children (single-block path keeps its behavior).
 */
function splitSdtParts(sdtXml: string): SdtPart[] | null {
  const contentOpen = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  if (!contentOpen) return null
  const innerStart = contentOpen.index + contentOpen[0].length
  const innerEnd = sdtXml.lastIndexOf('</w:sdtContent>')
  if (innerEnd <= innerStart) return null

  const children: Array<{ name: string; start: number; end: number }> = []
  const tagRe = /<(\/?)([A-Za-z0-9:._-]+)((?:"[^"]*"|'[^']*'|[^"'>])*)>/g
  tagRe.lastIndex = innerStart
  let depth = 0
  let start = -1
  let name = ''
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(sdtXml)) !== null && m.index < innerEnd) {
    if (m[1] === '/') {
      depth--
      if (depth === 0) children.push({ name, start, end: m.index + m[0].length })
    } else if (m[3].endsWith('/')) {
      if (depth === 0) children.push({ name: m[2], start: m.index, end: m.index + m[0].length })
    } else {
      if (depth === 0) {
        start = m.index
        name = m[2]
      }
      depth++
    }
  }

  const parts = children.filter((c) => c.name === 'w:p' || c.name === 'w:tbl')
  if (parts.length < 2) return null
  return parts.map((c, k) => ({
    name: c.name,
    start: k === 0 ? 0 : c.start,
    end: k === parts.length - 1 ? sdtXml.length : parts[k + 1].start,
    childStart: c.start,
    childEnd: c.end,
  }))
}

/**
 * When a top-level <w:sdt>'s content begins with a table (no paragraph before
 * it), return the balanced <w:tbl>…</w:tbl> slice, else null.
 */
function sdtTableXml(sdtXml: string): string | null {
  const contentOpen = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  if (!contentOpen) return null
  const contentTagEnd = contentOpen.index + contentOpen[0].length
  const contentClose = sdtXml.lastIndexOf('</w:sdtContent>')
  if (contentClose <= contentTagEnd) return null
  const inner = sdtXml.slice(contentTagEnd, contentClose)
  const tblStart = inner.search(/<w:tbl[\s>]/)
  if (tblStart === -1) return null
  const pStart = inner.search(/<w:p[\s/>]/)
  if (pStart !== -1 && pStart < tblStart) return null
  // balanced </w:tbl> for the opening tag (tables can nest)
  let depth = 0
  const tagRe = /<\/?w:tbl(?=[\s/>])/g
  tagRe.lastIndex = tblStart
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(inner)) !== null) {
    if (m[0].startsWith('</')) {
      depth--
      if (depth === 0) return inner.slice(tblStart, m.index + '</w:tbl>'.length)
    } else depth++
  }
  return null
}

async function buildBlock(
  el: BodyElement,
  index: number,
  xml: string,
  ctx: BuildContext,
): Promise<Block> {
  const base = { id: `b${index}`, docxIndex: index, originalXml: xml }

  if (el.name === 'w:ins' || el.name === 'w:del') {
    const openEnd = xml.indexOf('>') + 1
    const closeStart = xml.lastIndexOf(`</${el.name}>`)
    const child = splitXmlChildren(xml.slice(openEnd, closeStart)).find(
      (entry) => entry.name === 'w:p' || entry.name === 'w:tbl',
    )
    if (child) {
      const inner = await buildBlock(
        { name: child.name, start: 0, end: child.xml.length },
        index,
        child.xml,
        ctx,
      )
      let revisionAttrs: Record<string, string> = {}
      try {
        const parsed = xmlParser.parse(xml) as XNode[]
        const revisionNode = parsed.find((node) => nameOf(node) === el.name)
        if (revisionNode) revisionAttrs = attrsOf(revisionNode)
      } catch {
        /* malformed wrapper remains a protected passthrough */
      }
      inner.originalXml = xml
      inner.blockRevision = {
        kind: el.name === 'w:ins' ? 'ins' : 'del',
        author: revisionAttrs['w:author'] ?? '',
        ...(revisionAttrs['w:date'] ? { date: revisionAttrs['w:date'] } : {}),
        ...(revisionAttrs['w:id'] ? { id: revisionAttrs['w:id'] } : {}),
      }
      return inner
    }
  }

  if (el.name === 'w:sectPr') {
    return { ...base, type: 'passthrough', label: 'Section properties', hidden: true }
  }
  if (el.name === 'w:tbl') {
    return { ...base, type: 'table', ...tableSummary(xml), table: extractTable(xml, ctx) }
  }

  // --- SDT (structured document tag): extract sdtContent paragraph as editable ---
  if (el.name === 'w:sdt') {
    // sdtContent that starts with a table (research-report templates wrap whole
    // tables in content controls): display as a real table. Untouched
    // it saves byte-identical; cell-text edits patch inside the sdt shell.
    const tblXml = sdtTableXml(xml)
    if (tblXml) {
      return { ...base, type: 'table', ...tableSummary(tblXml), table: extractTable(tblXml, ctx) }
    }
    const sdtResult = parseSdtBlock(xml)
    if (sdtResult) {
      const { shell, pXml } = sdtResult
      // Build a synthetic BodyElement for the inner w:p
      const syntheticEl: BodyElement = { name: 'w:p', start: 0, end: pXml.length }
      // Build the block from the inner paragraph XML (keeps the sdt originalXml for passthrough)
      const innerBlock = await buildBlock(syntheticEl, index, pXml, ctx)
      // Attach the sdt shell and preserve the full sdt XML as the original
      innerBlock.originalXml = xml
      innerBlock.sdtShell = shell
      // Label the block with the alias for UI affordance
      if (!innerBlock.label) {
        const aliasLabel = shell.alias || shell.tag || 'Content control'
        innerBlock.label = aliasLabel
      }
      return innerBlock
    }
    // No usable paragraph found → passthrough
    return { ...base, type: 'passthrough', label: 'Content control', previewText: '' }
  }
  if (el.name !== 'w:p') {
    return { ...base, type: 'passthrough', label: el.name, previewText: '' }
  }

  // Feature-detect on XML with mc:Fallback stripped: Word pairs every modern
  // DrawingML shape (mc:Choice) with a legacy VML twin (<mc:Fallback><w:pict>),
  // so matching the raw bytes would misclassify every decorated paragraph as an
  // embedded object. Only detection uses this; saving still passes through the
  // original bytes untouched.
  // aidocs-ink runs are parsed separately into ParsedDoc.inks and re-emitted
  // from the save options; hide them from detection so an annotated text
  // paragraph stays an editable paragraph instead of a protected drawing.
  const detect = stripInkRuns(
    xml.includes('<mc:Fallback') ? xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '') : xml,
  )

  // Paragraph: certain constructs are protected as whole passthrough blocks.
  // Regenerating them would silently drop structure (section breaks, fields,
  // footnote anchors...), which is exactly the kind of damage patch-save exists
  // to prevent.
  if (detect.includes('<w:sectPr')) {
    return {
      ...base,
      type: 'passthrough',
      label: 'Section break paragraph',
      previewText: plainText(detect),
    }
  }
  // Field paragraphs whose visible result is a picture (e.g. INCLUDEPICTURE)
  // should still display the image; the block stays protected either way.
  if (
    detect.includes('<w:drawing') &&
    !detect.includes('<c:chart') &&
    !detect.includes('r:dm=') &&
    !detect.includes('<dgm:')
  ) {
    const image = await extractImage(detect, ctx)
    if (image) {
      return { ...base, type: 'image', label: 'Image', imageDataUrl: image, ...imageMeta(detect) }
    }
  }
  // Field chars inside textbox content don't make the paragraph itself a field
  // paragraph: research-report sidebars are VML textboxes embedding PAGE/date
  // fields, and those paragraphs must reach the textbox display path below
  //. Textbox blocks are protected passthrough anyway.
  const fieldDetect = detect.includes('<w:txbxContent') ? stripTextboxes(detect) : detect
  if (
    fieldDetect.includes('<w:fldChar') ||
    fieldDetect.includes('<w:fldSimple') ||
    fieldDetect.includes('<w:instrText')
  ) {
    // XE (index entry) fields are invisible markers; a paragraph whose only
    // fields are XE stays editable (extractRuns round-trips the markers).
    if (!onlyXeFields(detect)) {
      return {
        ...base,
        type: 'passthrough',
        label: fieldLabel(xml),
        previewText: plainText(xml),
        fieldDisplay: fieldDisplayOf(xml),
      }
    }
  }
  // TOC-styled paragraphs are part of a TOC field result even when they carry
  // no field chars themselves (entries with literal page numbers). Editing
  // them individually would corrupt the field, so they stay protected.
  // Word writes styleIds "TOC1".."TOC9"; Pages exports "TOC 1"/"TOC 2" (with space).
  if (/<w:pStyle w:val="TOC ?[1-9]"/.test(xml)) {
    return {
      ...base,
      type: 'passthrough',
      label: 'TOC entry',
      previewText: plainText(xml),
      fieldDisplay: fieldDisplayOf(xml),
    }
  }
  // footnote / endnote references are editable: extractRuns turns them into
  // Run.noteRef markers that regenerate as w:footnoteReference / w:endnoteReference
  // Run-level w:ins / w:del are parsed into Run.ins / Run.del (editable).
  // Paragraph-property revisions (pPrChange / numberingChange / paragraph-mark
  // ins/del) live in pPr and survive editing via Block.rawPPr passthrough.
  // moveFrom / moveTo are now parsed as editable with move-revision markers.
  // Only run-level revision constructs the run model cannot round-trip stay
  // protected: run-property changes, deleted field instructions, table-cell ins/del.
  if (/<w:(delInstrText|cellIns|cellDel)[ />]/.test(detect)) {
    return {
      ...base,
      type: 'passthrough',
      label: 'Revised paragraph',
      previewText: plainText(detect),
    }
  }
  // display equations (oMathPara / math-only paragraphs) stay protected as a
  // whole; paragraphs mixing math with plain text fall through to
  // buildTextParagraph, where each m:oMath becomes an atomic inline math run
  if (
    detect.includes('<m:oMath') &&
    (detect.includes('<m:oMathPara') || plainText(detect).trim() === '')
  ) {
    const tokens = mathTokens(detect)
    const omml = ommlFragmentsOf(detect).join('')
    // 2D MathML only for pure equations; an oMathPara paragraph that also has
    // plain runs keeps the flat token strip so the surrounding text stays visible
    const mathml = plainText(detect).trim() === '' ? ommlToMathML(omml) : ''
    const latex = omml ? ommlToLatex(omml) : null
    return {
      ...base,
      type: 'passthrough',
      label: 'Equation',
      previewText: tokens.join(''),
      formulaDisplay: {
        tokens,
        ...(mathml ? { mathml } : {}),
        ...(omml ? { omml } : {}),
        ...(latex ? { latex } : {}),
      },
    }
  }
  if (detect.includes('<w:object') || detect.includes('<w:pict')) {
    // Legacy VML textboxes (v:shape/v:textbox in w:pict): extract the
    // structured display model instead of flattening every nested paragraph and
    // table into one unreadable run. w:object (OLE) keeps the plain preview.
    if (!detect.includes('<w:object') && detect.includes('<w:txbxContent')) {
      const textboxes = extractTextboxes(detect, ctx)
      if (textboxes.length > 0 && plainText(stripTextboxes(detect)).trim() === '') {
        return {
          ...base,
          type: 'passthrough',
          label: 'Text box',
          previewText: textboxes
            .flatMap((t) => t.paras.map((p) => p.runs.map((r) => r.text).join('')))
            .join('\n'),
          textboxes,
        }
      }
    }
    return {
      ...base,
      type: 'passthrough',
      label: 'Embedded object',
      previewText: plainText(detect),
      ...(await oleDisplay(detect, ctx)),
    }
  }
  if (detect.includes('<w:drawing')) {
    if (detect.includes('<c:chart') || detect.includes('r:dm=') || detect.includes('<dgm:')) {
      const isChart = detect.includes('<c:chart')
      const chartDisplay = isChart ? await extractChart(detect, ctx) : null
      const diagramText = isChart ? null : await extractDiagramText(detect, ctx)
      return {
        ...base,
        type: 'passthrough',
        label: isChart ? 'Chart' : 'SmartArt',
        ...(chartDisplay ? { chartDisplay, previewText: chartDisplay.title ?? '' } : {}),
        ...(diagramText ? { previewText: diagramText } : {}),
      }
    }
    const image = await extractImage(detect, ctx)
    if (image) {
      return { ...base, type: 'image', label: 'Image', imageDataUrl: image, ...imageMeta(detect) }
    }
    // Anchored decorative shape (underline rule, background box...) in a
    // paragraph that also carries real text: parse it as a normal paragraph so
    // the text stays readable/editable. The shape lives in runs we do not
    // regenerate — the block still saves byte-identical while untouched, and
    // only loses the decoration if the user actually edits this paragraph.
    if (plainText(stripTextboxes(detect)).trim() !== '') {
      return buildTextParagraph(base, xml, ctx)
    }
    // Anchored textboxes (code boxes, callout cards): all visible text lives in
    // w:txbxContent. Extract a display-only model so content and box styling
    // render; the block stays protected and saves byte-identical.
    const textboxes = extractTextboxes(detect, ctx)
    if (textboxes.length > 0) {
      return {
        ...base,
        type: 'passthrough',
        label: 'Text box',
        previewText: textboxes
          .flatMap((t) => t.paras.map((p) => p.runs.map((r) => r.text).join('')))
          .join('\n'),
        textboxes,
        ...imageMeta(detect),
      }
    }
    return { ...base, type: 'passthrough', label: 'Drawing object', decorative: isThinRule(detect) }
  }

  return buildTextParagraph(base, xml, ctx)
}

/**
 * Effective heading level 1-9 (Word TOC/outline semantics): direct pPr
 * w:outlineLvl wins (9 = body text), then the style's level, then a built-in
 * HeadingN styleId the document never defined.
 */
function headingLevelOf(
  pPr: XNode | undefined,
  styleId: string | undefined,
  ctx: BuildContext,
): number | undefined {
  const direct = pPr ? attrsOf(findChild(pPr, 'w:outlineLvl') ?? {})['w:val'] : undefined
  if (direct !== undefined) {
    const lvl = parseInt(direct, 10)
    return lvl >= 0 && lvl <= 8 ? lvl + 1 : undefined
  }
  if (!styleId) return undefined
  const info = ctx.styles.get(styleId)
  if (info) return info.headingLevel
  const m = /^Heading([1-9])$/i.exec(styleId)
  return m ? parseInt(m[1], 10) : undefined
}

/** parse a w:p as editable text content (paragraph / heading / listItem) */
function buildTextParagraph(
  base: Pick<Block, 'id' | 'docxIndex' | 'originalXml'>,
  xml: string,
  ctx: BuildContext,
): Block {
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    // unparseable paragraph (e.g. pathological nesting): keep the original bytes
    return { ...base, type: 'passthrough', label: 'Paragraph', previewText: plainText(xml) }
  }
  const pNode = parsed.find((n) => nameOf(n) === 'w:p')
  if (!pNode) {
    return { ...base, type: 'passthrough', label: 'Unknown paragraph', previewText: plainText(xml) }
  }

  const pPr = findChild(pNode, 'w:pPr')
  const styleId = pPr ? attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] : undefined
  const format = pPr ? extractParaFormat(pPr) : undefined
  const rawPPr = rawPPrOf(xml)
  // inline math: raw <m:oMath> fragments in document order, aligned with the
  // walk below (strip fallback/textbox copies so indexes match visited nodes)
  const mathXml = stripTextboxes(
    xml.includes('<mc:Fallback') ? xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '') : xml,
  )
  const runs = extractRuns(pNode, ctx, ommlFragmentsOf(mathXml), rubyFragmentsOf(mathXml))
  const { bookmarks, hiddenBookmarks } = bookmarkNamesOf(stripTextboxes(xml))
  const { commentStarts, commentEnds } = crossParaCommentMarkers(stripTextboxes(xml))

  // --- move revision detection ---
  // A paragraph with moveFrom/moveTo at run level gets a block-level marker
  // for visual styling (strikethrough+red bg for moveFrom, green bg for moveTo).
  let moveRevision: 'from' | 'to' | undefined
  if (/<w:moveFrom[\s/>]/.test(xml)) moveRevision = 'from'
  else if (/<w:moveTo[\s/>]/.test(xml)) moveRevision = 'to'

  // --- pPrChange info extraction ---
  // When the paragraph has a pPrChange (tracked format change), extract the
  // revision author/date/id for the review badge and navigation.
  let pPrChangeInfo: Block['pPrChangeInfo']
  if (pPr) {
    const pPrChangeEl = findChild(pPr, 'w:pPrChange')
    if (pPrChangeEl) {
      const attrs = attrsOf(pPrChangeEl)
      pPrChangeInfo = { author: attrs['w:author'] ?? '' }
      if (attrs['w:date']) pPrChangeInfo.date = attrs['w:date']
      if (attrs['w:id']) pPrChangeInfo.id = attrs['w:id']
      const oldPPr = findChild(pPrChangeEl, 'w:pPr')
      if (oldPPr) {
        const old: NonNullable<NonNullable<Block['pPrChangeInfo']>['old']> = {
          ...(extractParaFormat(oldPPr) ?? {}),
        }
        const oldStyleId = attrsOf(findChild(oldPPr, 'w:pStyle') ?? {})['w:val']
        if (oldStyleId) old.styleId = oldStyleId
        const oldNumPr = findChild(oldPPr, 'w:numPr')
        const oldNumId = oldNumPr
          ? attrsOf(findChild(oldNumPr, 'w:numId') ?? {})['w:val']
          : undefined
        if (oldNumId) {
          old.type = 'docListItem'
          old.numId = oldNumId
          old.ilvl =
            parseInt(attrsOf(findChild(oldNumPr!, 'w:ilvl') ?? {})['w:val'] ?? '0', 10) || 0
          old.kind = listKindOf(ctx, oldNumId, old.ilvl)
        } else {
          const oldLevel = headingLevelOf(oldPPr, oldStyleId, ctx)
          if (oldLevel) {
            old.type = 'docHeading'
            old.level = oldLevel
          } else if (oldStyleId) {
            old.type = 'docParagraph'
          }
        }
        if (old && Object.keys(old).length > 0) pPrChangeInfo.old = old
      }
    }
  }

  // list item?
  const listRef = listRefOf(ctx, pPr, styleId)
  /** extra revision fields shared across all return paths */
  const revExtras = {
    ...(moveRevision ? { moveRevision } : {}),
    ...(pPrChangeInfo ? { pPrChangeInfo } : {}),
  }
  if (listRef) {
    const kind = listKindOf(ctx, listRef.numId, listRef.ilvl)
    return {
      ...base,
      type: 'listItem',
      styleId,
      list: { kind, numId: listRef.numId, ilvl: listRef.ilvl },
      format,
      rawPPr,
      bookmarks,
      hiddenBookmarks,
      commentStarts,
      commentEnds,
      runs,
      ...revExtras,
    }
  }

  // heading?
  const headingLevel = headingLevelOf(pPr, styleId, ctx)
  if (headingLevel) {
    return {
      ...base,
      type: 'heading',
      level: headingLevel,
      styleId,
      format,
      rawPPr,
      bookmarks,
      hiddenBookmarks,
      commentStarts,
      commentEnds,
      runs,
      ...revExtras,
    }
  }

  return {
    ...base,
    type: 'paragraph',
    styleId,
    format,
    rawPPr,
    bookmarks,
    hiddenBookmarks,
    commentStarts,
    commentEnds,
    runs,
    ...revExtras,
  }
}

/**
 * Cross-paragraph comment range endpoints: comment ids where only one end falls in this
 * paragraph (the other end is in a different paragraph). Ranges fully within one
 * paragraph are handled by run.commentIds; this only catches cross-paragraph ones so a
 * paragraph rebuild does not leave orphaned commentRangeEnd/Start markers.
 */
function crossParaCommentMarkers(xml: string): {
  commentStarts: string[] | undefined
  commentEnds: string[] | undefined
} {
  const ids = (re: RegExp) => [...xml.matchAll(re)].map((m) => m[1])
  const starts = ids(/<w:commentRangeStart [^>]*w:id="([^"]+)"/g)
  const ends = ids(/<w:commentRangeEnd [^>]*w:id="([^"]+)"/g)
  const onlyStarts = starts.filter((id) => !ends.includes(id))
  const onlyEnds = ends.filter((id) => !starts.includes(id))
  return {
    commentStarts: onlyStarts.length ? onlyStarts : undefined,
    commentEnds: onlyEnds.length ? onlyEnds : undefined,
  }
}

/** user bookmark names starting in this paragraph; Word internals (_Toc/_Ref/_GoBack…) split out as hiddenBookmarks */
function bookmarkNamesOf(xml: string): {
  bookmarks: string[] | undefined
  hiddenBookmarks: string[] | undefined
} {
  const names: string[] = []
  const hidden: string[] = []
  for (const m of xml.matchAll(/<w:bookmarkStart [^>]*w:name="([^"]+)"/g)) {
    const name = decodeEntities(m[1])
    // A _ prefix marks Word internal bookmarks (_Ref/_Toc/_Hlk): hidden from the UI, but
    // they must be re-emitted when the paragraph rebuilds, otherwise REF cross-references
    // and TOC anchors pointing at them break
    const list = name.startsWith('_') ? hidden : names
    if (!list.includes(name)) list.push(name)
  }
  return {
    bookmarks: names.length > 0 ? names : undefined,
    hiddenBookmarks: hidden.length > 0 ? hidden : undefined,
  }
}

/**
 * Exact <w:pPr>...</w:pPr> slice of a paragraph (depth-aware: pPrChange nests
 * another w:pPr inside). undefined when the paragraph has no properties.
 */
function rawPPrOf(xml: string): string | undefined {
  // the paragraph's own pPr must be the first child of w:p; a later match
  // would belong to nested content (textbox paragraphs)
  const openEnd = xml.indexOf('>') + 1
  if (openEnd === 0 || !xml.startsWith('<w:pPr', openEnd)) return undefined
  const start = openEnd
  const re = /<w:pPr(?=[\s/>])|<\/w:pPr>/g
  re.lastIndex = start
  let depth = 0
  let match: RegExpExecArray | null
  while ((match = re.exec(xml)) !== null) {
    if (match[0] === '</w:pPr>') {
      depth--
      if (depth === 0) return xml.slice(start, match.index + match[0].length)
    } else {
      // self-closing <w:pPr/> never opens
      const gt = xml.indexOf('>', match.index)
      if (xml[gt - 1] === '/') {
        if (depth === 0) return xml.slice(start, gt + 1)
        continue
      }
      depth++
    }
  }
  return undefined
}

/**
 * A text-less anchored shape no taller than ~10px is almost always a
 * decorative horizontal rule (heading underlines, dividers); render those as a
 * line instead of a drawing-object chip.
 */
function isThinRule(xml: string): boolean {
  const cy = parseInt(/<wp:extent cx="\d+" cy="(\d+)"/.exec(xml)?.[1] ?? '', 10)
  return Number.isFinite(cy) && cy > 0 && cy <= 130000
}

/** drop textbox content so "does the paragraph itself have text" checks work */
function stripTextboxes(xml: string): string {
  return xml.includes('<w:txbxContent')
    ? xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '')
    : xml
}

/** paragraphs (and tables, one display line per row) of a w:txbxContent node */
function txbxContentParas(content: XNode, ctx: BuildContext): TextboxParaDisplay[] {
  const out: TextboxParaDisplay[] = []
  for (const child of childrenOf(content)) {
    const name = nameOf(child)
    if (name === 'w:p') {
      const para: TextboxParaDisplay = { runs: extractRuns(child, ctx) }
      const pPr = findChild(child, 'w:pPr')
      if (pPr) Object.assign(para, extractParaFormat(pPr))
      out.push(para)
    } else if (name === 'w:tbl') {
      out.push(...txbxTableParas(child, ctx))
    } else if (name === 'w:sdt') {
      const inner = findChild(child, 'w:sdtContent')
      if (inner) out.push(...txbxContentParas(inner, ctx))
    }
  }
  return out
}

/**
 * Table inside a textbox (sidebar key-data blocks): the display model has no
 * grid, so render one line per row with the cells spaced apart — readable rows
 * instead of a single concatenated blob. Tables nested inside a
 * cell (research templates wrap analyst info this way) recurse into their own
 * lines after the host row.
 */
function txbxTableParas(tbl: XNode, ctx: BuildContext): TextboxParaDisplay[] {
  const out: TextboxParaDisplay[] = []
  for (const tr of childrenThroughSdt(tbl, 'w:tr')) {
    const runs: Run[] = []
    const nestedLines: TextboxParaDisplay[] = []
    for (const tc of childrenThroughSdt(tr, 'w:tc')) {
      const cellRuns: Run[] = []
      for (const p of childrenThroughSdt(tc, 'w:p')) {
        const pRuns = extractRuns(p, ctx)
        if (pRuns.every((r) => r.text.trim() === '')) continue
        if (cellRuns.length > 0) cellRuns.push({ text: ' ' })
        cellRuns.push(...pRuns)
      }
      if (!cellRuns.every((r) => r.text.trim() === '')) {
        if (runs.length > 0) runs.push({ text: '\u2002\u2002' })
        runs.push(...cellRuns)
      }
      for (const nested of childrenThroughSdt(tc, 'w:tbl')) {
        nestedLines.push(...txbxTableParas(nested, ctx))
      }
    }
    if (runs.length > 0 || nestedLines.length === 0) out.push({ runs })
    out.push(...nestedLines)
  }
  return out
}

/**
 * w:tbl or w:sdt among the txbxContent children: the display lines no longer
 * map 1:1 onto the w:p segments patch-save rewrites, so the box must stay
 * read-only (editing would drop the table / sdt shells).
 */
function txbxHasStructuredContent(content: XNode): boolean {
  return childrenOf(content).some((c) => {
    const n = nameOf(c)
    return n === 'w:tbl' || n === 'w:sdt'
  })
}

/** VML style="width:189.9pt;height:626pt" dimension → CSS px */
function vmlStyleDimPx(style: string, key: 'width' | 'height'): number | undefined {
  const m = new RegExp(`(?:^|;)\\s*${key}:([0-9.]+)(pt|px|in|mm|cm)?`).exec(style)
  if (!m) return undefined
  const v = parseFloat(m[1]!)
  if (!Number.isFinite(v) || v <= 0) return undefined
  const unit = m[2] ?? 'pt'
  const px =
    unit === 'px'
      ? v
      : unit === 'in'
        ? v * 96
        : unit === 'mm'
          ? (v / 25.4) * 96
          : unit === 'cm'
            ? (v / 2.54) * 96
            : (v * 96) / 72
  return Math.round(px)
}

/** VML color attr ("#dbe5f1", "#dbe5f1 [3204]", "windowText") → hex without '#', or undefined */
function vmlColorHex(value: string | undefined): string | undefined {
  const m = value && /^#?([0-9a-fA-F]{6})/.exec(value.trim())
  return m ? m[1] : undefined
}

/**
 * Display-only extraction of anchored textboxes: DrawingML (wps:wsp, converter
 * output for code boxes / callout cards) and legacy VML (v:shape etc. inside
 * w:pict, common in broker research templates). Expects
 * fallback-stripped XML, otherwise the mc:Fallback VML twin would duplicate
 * every box.
 */
const LINE_PRSTS_RE =
  /<a:prstGeom[^>]*prst="(?:line|straightConnector1|bentConnector[234]|curvedConnector[234])"/

/** stroke-only line/connector prsts shown as display boxes despite no text body */
const LINE_PRSTS = new Set([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
])

/** wps line shape → display-only line box (synthetic prst carries the arrow ends) */
function lineBoxOf(shape: XNode): TextboxDisplay | null {
  const spPr = findChild(shape, 'wps:spPr')
  if (!spPr) return null
  const prst = attrsOf(findChild(spPr, 'a:prstGeom') ?? {})['prst']
  if (!prst || !LINE_PRSTS.has(prst)) return null
  const box: TextboxDisplay = { paras: [], readOnly: true }
  const ln = findChild(spPr, 'a:ln')
  const border = ln
    ? attrsOf(findChild(findChild(ln, 'a:solidFill') ?? {}, 'a:srgbClr') ?? {})['val']
    : undefined
  box.borderColor = border ?? '000000'
  const arrowEnd = (name: string): boolean => {
    const type = attrsOf(findChild(ln ?? {}, name) ?? {})['type']
    return !!type && type !== 'none'
  }
  const head = arrowEnd('a:headEnd')
  const tail = arrowEnd('a:tailEnd')
  box.prst = prst.startsWith('bentConnector')
    ? 'lineBent'
    : prst.startsWith('curvedConnector')
      ? 'lineCurved'
      : head && tail
        ? 'lineArrowDouble'
        : head || tail
          ? 'lineArrow'
          : 'line'
  const ext = findChild(findChild(spPr, 'a:xfrm') ?? {}, 'a:ext')
  const cx = ext ? parseInt(attrsOf(ext)['cx'] ?? '', 10) : NaN
  const cy = ext ? parseInt(attrsOf(ext)['cy'] ?? '', 10) : NaN
  if (Number.isFinite(cx) && cx > 0) box.widthPx = Math.round(cx / EMU_PER_PX)
  if (Number.isFinite(cy) && cy > 0) {
    box.heightPx = Math.round(cy / EMU_PER_PX)
    box.minHeightPx = box.heightPx
  } else {
    // zero-height extent = Word's horizontal line; keep a 12 px grab band
    box.heightPx = 12
  }
  box.insetTopPx = 0
  box.insetRightPx = 0
  box.insetBottomPx = 0
  box.insetLeftPx = 0
  return box
}

function extractTextboxes(xml: string, ctx: BuildContext): TextboxDisplay[] {
  // wrapSquare gate keeps converter-emitted decorative rules on the thin-rule path
  const hasLineShapes = xml.includes('<wp:wrapSquare') && LINE_PRSTS_RE.test(xml)
  if (!xml.includes('<w:txbxContent') && !hasLineShapes) return []
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    return []
  }
  const shapes: XNode[] = []
  collectNodes(parsed, 'wps:wsp', shapes)
  for (const vmlName of ['v:shape', 'v:rect', 'v:roundrect']) {
    collectNodes(parsed, vmlName, shapes)
  }
  const out: TextboxDisplay[] = []
  for (const shape of shapes) {
    const contents: XNode[] = []
    collectNodes(childrenOf(shape), 'w:txbxContent', contents)
    if (contents.length === 0) {
      if (hasLineShapes) {
        const lineBox = lineBoxOf(shape)
        if (lineBox) out.push(lineBox)
      }
      continue
    }
    const box: TextboxDisplay = { paras: [] }
    const shapeAttrs = attrsOf(shape)
    if (nameOf(shape) !== 'wps:wsp') {
      // VML shape: geometry from the style attribute, colors from fillcolor/strokecolor
      const style = shapeAttrs['style'] ?? ''
      const w = vmlStyleDimPx(style, 'width')
      if (w) box.widthPx = w
      const h = vmlStyleDimPx(style, 'height')
      if (h) {
        box.heightPx = h
        box.minHeightPx = h
      }
      const fill = vmlColorHex(shapeAttrs['fillcolor'])
      if (fill && shapeAttrs['filled'] !== 'f') box.fill = fill
      const stroke = vmlColorHex(shapeAttrs['strokecolor'])
      if (stroke && shapeAttrs['stroked'] !== 'f') box.borderColor = stroke
      for (const content of contents) box.paras.push(...txbxContentParas(content, ctx))
      if (contents.some(txbxHasStructuredContent)) box.readOnly = true
      if (box.paras.some((p) => p.runs.length > 0)) out.push(box)
      continue
    }
    const spPr = findChild(shape, 'wps:spPr')
    if (spPr) {
      const fill = attrsOf(findChild(findChild(spPr, 'a:solidFill') ?? {}, 'a:srgbClr') ?? {})[
        'val'
      ]
      if (fill) box.fill = fill
      const ln = findChild(spPr, 'a:ln')
      const border = ln
        ? attrsOf(findChild(findChild(ln, 'a:solidFill') ?? {}, 'a:srgbClr') ?? {})['val']
        : undefined
      if (border) box.borderColor = border
      const prstGeom = findChild(spPr, 'a:prstGeom')
      const prst = prstGeom ? attrsOf(prstGeom)['prst'] : undefined
      if (prst && prst !== 'rect') box.prst = prst
      const ext = findChild(findChild(spPr, 'a:xfrm') ?? {}, 'a:ext')
      const cx = ext ? parseInt(attrsOf(ext)['cx'] ?? '', 10) : NaN
      if (Number.isFinite(cx) && cx > 0) box.widthPx = Math.round(cx / EMU_PER_PX)
      // Word clips overflowing text unless the shape auto-fits its content;
      // carrying the fixed height keeps tall sparse boxes from exploding layout
      const bodyPr = findChild(shape, 'wps:bodyPr')
      const autoFit = bodyPr ? !!findChild(bodyPr, 'a:spAutoFit') : false
      const cy = ext ? parseInt(attrsOf(ext)['cy'] ?? '', 10) : NaN
      if (!autoFit && Number.isFinite(cy) && cy > 0) {
        box.heightPx = Math.round(cy / EMU_PER_PX)
        box.minHeightPx = box.heightPx
      }
    }
    const bodyPr = findChild(shape, 'wps:bodyPr')
    if (bodyPr) {
      const attrs = attrsOf(bodyPr)
      const inset = (name: string): number | undefined => {
        const emu = parseInt(attrs[name] ?? '', 10)
        return Number.isFinite(emu) && emu >= 0
          ? Math.round((emu / EMU_PER_PX) * 100) / 100
          : undefined
      }
      box.insetLeftPx = inset('lIns')
      box.insetTopPx = inset('tIns')
      box.insetRightPx = inset('rIns')
      box.insetBottomPx = inset('bIns')
    }
    for (const content of contents) box.paras.push(...txbxContentParas(content, ctx))
    if (contents.some(txbxHasStructuredContent)) box.readOnly = true
    if (box.paras.some((p) => p.runs.length > 0)) out.push(box)
  }
  return out
}

function collectNodes(nodes: XNode[], name: string, out: XNode[]): void {
  for (const node of nodes) {
    if (nameOf(node) === name) out.push(node)
    collectNodes(childrenOf(node), name, out)
  }
}

const JC_ALIGN: Record<string, ParaFormat['align']> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  both: 'justify',
  distribute: 'distribute',
}

function extractParaFormat(pPr: XNode): ParaFormat | undefined {
  const format: ParaFormat = {}
  if (boolProp(pPr, 'w:bidi')) format.bidi = true
  const jc = attrsOf(findChild(pPr, 'w:jc') ?? {})['w:val']
  if (jc && JC_ALIGN[jc]) format.align = JC_ALIGN[jc]
  // Word quirk: in bidi paragraphs w:jc left/right are logical values (start/end); convert to visual direction
  if (format.bidi && (format.align === 'left' || format.align === 'right')) {
    format.align = format.align === 'left' ? 'right' : 'left'
  }
  const spacing = findChild(pPr, 'w:spacing')
  if (spacing) {
    const attrs = attrsOf(spacing)
    const rule = (attrs['w:lineRule'] ?? 'auto') as 'auto' | 'atLeast' | 'exact'
    const line = parseInt(attrs['w:line'] ?? '', 10)
    if (line > 0) {
      format.lineRawTwips = line
      if (rule === 'auto') {
        format.lineSpacing = Math.round((line / 240) * 100) / 100
        format.lineRule = 'auto'
      } else {
        format.lineRule = rule
        // lineSpacing in 'auto' sense is not applicable for atLeast/exact, keep undefined
      }
    }
    // autospacing=1 means Word ignores the literal before/after and computes its own;
    // dropping the literal (→ style/docDefaults cascade) is closer than honoring it.
    // Explicit "0" must be kept — it overrides the style's spacing.
    const autoBefore =
      attrs['w:beforeAutospacing'] === '1' || attrs['w:beforeAutospacing'] === 'true'
    const autoAfter = attrs['w:afterAutospacing'] === '1' || attrs['w:afterAutospacing'] === 'true'
    const before = parseInt(attrs['w:before'] ?? '', 10)
    if (before >= 0 && attrs['w:before'] !== undefined && !autoBefore) format.spaceBefore = before
    const after = parseInt(attrs['w:after'] ?? '', 10)
    if (after >= 0 && attrs['w:after'] !== undefined && !autoAfter) format.spaceAfter = after
  }
  const ind = findChild(pPr, 'w:ind')
  if (ind) {
    const attrs = attrsOf(ind)
    const left = parseInt(attrs['w:left'] ?? attrs['w:start'] ?? '', 10)
    // Negative indent (hanging past the left margin) is legal, keep it; 0 stays out of the model
    if (Number.isFinite(left) && left !== 0) format.indentLeft = left
    const right = parseInt(attrs['w:right'] ?? attrs['w:end'] ?? '', 10)
    if (Number.isFinite(right) && right !== 0) format.indentRight = right
    const firstLine = parseInt(attrs['w:firstLine'] ?? '', 10)
    const hanging = parseInt(attrs['w:hanging'] ?? '', 10)
    if (hanging > 0) format.indentFirstLine = -hanging
    else if (firstLine > 0) format.indentFirstLine = firstLine
  }
  if (boolProp(pPr, 'w:pageBreakBefore')) format.pageBreakBefore = true
  if (boolProp(pPr, 'w:keepNext')) format.keepNext = true
  if (boolProp(pPr, 'w:keepLines')) format.keepLines = true
  // widowControl: Word default is ON; only store when explicitly set to OFF
  const wcEl = findChild(pPr, 'w:widowControl')
  if (wcEl) {
    const wcVal = attrsOf(wcEl)['w:val']
    // w:widowControl/ (no val) = on; w:widowControl w:val="0" or "false" = off
    if (wcVal === '0' || wcVal === 'false') format.widowControl = false
  }
  if (boolProp(pPr, 'w:contextualSpacing')) format.contextualSpacing = true
  const shd = findChild(pPr, 'w:shd')
  if (shd) {
    const fill = attrsOf(shd)['w:fill']
    if (fill && fill !== 'auto') format.shadingFill = fill
  }
  const pBdr = findChild(pPr, 'w:pBdr')
  if (pBdr) {
    let borders = ''
    for (const [side, ch] of [
      ['top', 't'],
      ['bottom', 'b'],
      ['left', 'l'],
      ['right', 'r'],
    ] as const) {
      const el = findChild(pBdr, `w:${side}`)
      if (el && attrsOf(el)['w:val'] !== 'none') borders += ch
    }
    if (borders) format.borders = borders
  }
  const tabsEl = findChild(pPr, 'w:tabs')
  if (tabsEl) {
    const stops: import('./types').TabStop[] = []
    for (const tab of findChildren(tabsEl, 'w:tab')) {
      const attrs = attrsOf(tab)
      const pos = parseInt(attrs['w:pos'] ?? '', 10)
      const val = attrs['w:val'] ?? 'left'
      if (!Number.isFinite(pos)) continue
      const validVals = ['left', 'center', 'right', 'decimal', 'bar', 'clear'] as const
      const safeVal = validVals.includes(val as (typeof validVals)[number])
        ? (val as (typeof validVals)[number])
        : 'left'
      const stop: import('./types').TabStop = { pos, val: safeVal }
      const leader = attrs['w:leader']
      if (leader && leader !== 'none') {
        const validLeaders = ['dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] as const
        if (validLeaders.includes(leader as (typeof validLeaders)[number])) {
          stop.leader = leader as (typeof validLeaders)[number]
        }
      }
      stops.push(stop)
    }
    if (stops.length > 0) format.tabStops = stops
  }
  // Drop cap: w:framePr w:dropCap="drop"|"margin"
  const framePr = findChild(pPr, 'w:framePr')
  if (framePr) {
    const dropCapVal = attrsOf(framePr)['w:dropCap']
    if (dropCapVal === 'drop' || dropCapVal === 'margin') {
      const lines = parseInt(attrsOf(framePr)['w:lines'] ?? '3', 10) || 3
      format.dropCap = { type: dropCapVal as 'drop' | 'margin', lines }
    }
  }
  return Object.keys(format).length > 0 ? format : undefined
}

/**
 * True when every field in the paragraph is an XE (index entry) marker.
 * Multi-fragment instructions or fldSimple fields fail the check, so anything
 * unusual falls back to the protected-passthrough path.
 */
/** paragraphs whose only fields are XE / REF stay editable (extractRuns round-trips them) */
/** Simple instructions foldable into an editable inline-field run (the cached result is the display text) */
const SIMPLE_INLINE_FIELD_RE = /^\s*(DATE|TIME|CREATEDATE|SAVEDATE|NUMPAGES|FILENAME|AUTHOR|PAGE)\b/

function onlyXeFields(xml: string): boolean {
  if (xml.includes('<w:fldSimple')) return false
  const instrs = xml.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? []
  if (instrs.length === 0) return false
  return instrs.every((fragment) => {
    const text = decodeEntities(fragment.replace(/<[^>]+>/g, ''))
    return /^\s*XE[\s"]/.test(text) || /^\s*REF\s/.test(text) || SIMPLE_INLINE_FIELD_RE.test(text)
  })
}

function extractRuns(
  pNode: XNode,
  ctx: BuildContext,
  mathFragments: string[] = [],
  rubyFragments: string[] = [],
): Run[] {
  const runs: Run[] = []
  let mathIndex = 0
  let rubyIndex = 0
  // Comments are only tracked when the whole range lives inside this paragraph:
  // a regenerated paragraph can then re-emit its own markers, while ranges that
  // span paragraphs are left untouched (their runs get no commentIds).
  const starts = new Set<string>()
  const ends = new Set<string>()
  const collectRangeIds = (nodes: XNode[]) => {
    for (const node of nodes) {
      const name = nameOf(node)
      if (name === 'w:commentRangeStart' || name === 'w:commentRangeEnd') {
        const id = attrsOf(node)['w:id']
        if (id) (name === 'w:commentRangeStart' ? starts : ends).add(id)
      }
      collectRangeIds(childrenOf(node))
    }
  }
  collectRangeIds(childrenOf(pNode))
  const complete = new Set([...starts].filter((id) => ends.has(id)))
  const activeComments = new Set<string>()
  type RevCtx = { ins?: RevisionInfo; del?: RevisionInfo }
  // inline field state: XE folds into Run.xeTerm, REF (cross-reference) into Run.refField
  let fieldDepth = 0
  let fieldInstr = ''
  let fieldSeparated = false
  let fieldCached = ''
  const pushRun = (run: Run, rev?: RevCtx) => {
    if (activeComments.size > 0) run.commentIds = [...activeComments].sort()
    if (rev?.ins) run.ins = rev.ins
    if (rev?.del) run.del = rev.del
    runs.push(run)
  }
  const handleRun = (node: XNode, link: Run['link'] | undefined, rev?: RevCtx) => {
    const fldChar = findChild(node, 'w:fldChar')
    if (fldChar) {
      const type = attrsOf(fldChar)['w:fldCharType']
      if (type === 'begin') {
        fieldDepth++
        if (fieldDepth === 1) {
          fieldInstr = ''
          fieldSeparated = false
          fieldCached = ''
        }
      } else if (type === 'separate') {
        if (fieldDepth === 1) fieldSeparated = true
      } else if (type === 'end') {
        fieldDepth = Math.max(0, fieldDepth - 1)
        if (fieldDepth === 0) {
          const xe = /^\s*XE\s+(?:"([^"]*)"|(\S+))/.exec(fieldInstr)
          const ref = /^\s*REF\s+(?:"([^"]+)"|([^\s\\]+))/.exec(fieldInstr)
          if (xe) pushRun({ text: '', xeTerm: xe[1] ?? xe[2] }, rev)
          else if (ref) {
            const name = ref[1] ?? ref[2]
            pushRun({ text: fieldCached || name, refField: name, refInstr: fieldInstr }, rev)
          } else if (SIMPLE_INLINE_FIELD_RE.test(fieldInstr)) {
            pushRun({ text: fieldCached || ' ', instrField: fieldInstr.trim() }, rev)
          }
          fieldInstr = ''
          fieldSeparated = false
          fieldCached = ''
        }
      }
      return
    }
    if (fieldDepth > 0) {
      // keep the fragment index aligned even for ruby inside a field cache
      if (findChild(node, 'w:ruby')) rubyIndex++
      const instr = findChild(node, 'w:instrText')
      if (instr) fieldInstr += textOf(instr)
      else if (fieldSeparated && fieldDepth === 1) {
        // REF cached result is the display text; other fields' caches are dropped
        const cached = buildRun(node, link, ctx.themeColors, ctx.themeFonts)
        if (cached) fieldCached += cached.text
      }
      return
    }
    const rubyNode = findChild(node, 'w:ruby')
    if (rubyNode) {
      const xml = rubyFragments[rubyIndex++]
      const base = rubyPartText(rubyNode, 'w:rubyBase')
      const rt = rubyPartText(rubyNode, 'w:rt')
      // no fragment (textbox paths): degrade to the base characters
      if (base) pushRun(xml ? { text: base, ruby: { rt, xml } } : { text: base }, rev)
      return
    }
    const noteRefNode =
      findChild(node, 'w:footnoteReference') ?? findChild(node, 'w:endnoteReference')
    if (noteRefNode) {
      const kind = nameOf(noteRefNode) === 'w:footnoteReference' ? 'footnote' : 'endnote'
      const id = attrsOf(noteRefNode)['w:id']
      if (id) {
        const num = ctx.noteNumbers.get(`${kind}:${id}`)
        pushRun({ text: String(num ?? '*'), noteRef: { kind, id } }, rev)
        return
      }
    }
    const run = buildRun(node, link, ctx.themeColors, ctx.themeFonts)
    if (run) pushRun(run, rev)
  }
  const walk = (nodes: XNode[], link?: Run['link'], rev?: RevCtx) => {
    for (const node of nodes) {
      const name = nameOf(node)
      if (name === 'w:commentRangeStart' || name === 'w:commentRangeEnd') {
        const id = attrsOf(node)['w:id']
        if (id && complete.has(id)) {
          if (name === 'w:commentRangeStart') activeComments.add(id)
          else activeComments.delete(id)
        }
      } else if (name === 'w:ins' || name === 'w:del') {
        const attrs = attrsOf(node)
        const info: RevisionInfo = { author: attrs['w:author'] ?? '' }
        if (attrs['w:date']) info.date = attrs['w:date']
        if (attrs['w:id']) info.id = attrs['w:id']
        const next: RevCtx = name === 'w:ins' ? { ...rev, ins: info } : { ...rev, del: info }
        walk(childrenOf(node), link, next)
      } else if (name === 'w:moveFrom' || name === 'w:moveTo') {
        // Treat moveFrom like del (content was moved away) and moveTo like ins (content arrived here).
        // This allows the existing accept/reject mechanism to handle moves via del/ins marks.
        const attrs = attrsOf(node)
        const info: RevisionInfo = { author: attrs['w:author'] ?? '' }
        if (attrs['w:date']) info.date = attrs['w:date']
        if (attrs['w:id']) info.id = attrs['w:id']
        const next: RevCtx = name === 'w:moveFrom' ? { ...rev, del: info } : { ...rev, ins: info }
        walk(childrenOf(node), link, next)
      } else if (name === 'w:r') {
        handleRun(node, link, rev)
      } else if (name === 'm:oMath') {
        // atomic inline formula; the raw fragment saves verbatim on regeneration
        const omml = mathFragments[mathIndex++]
        if (omml) pushRun({ text: mathTokens(omml).join(''), math: { omml } }, rev)
      } else if (name === 'w:hyperlink') {
        const attrs = attrsOf(node)
        const rId = attrs['r:id']
        const anchor = attrs['w:anchor']
        const tooltip = attrs['w:tooltip']
        const href = rId ? (ctx.rels.get(rId)?.target ?? '') : anchor ? `#${anchor}` : ''
        walk(childrenOf(node), { href, rId, ...(tooltip ? { tooltip } : {}) }, rev)
      } else if (name === 'w:smartTag' || name === 'w:sdt' || name === 'w:sdtContent') {
        walk(childrenOf(node), link, rev)
      }
    }
  }
  walk(childrenOf(pNode))
  return mergeRuns(runs)
}

/** exact <w:ruby> fragments in document order (w:ruby has no attributes and cannot nest) */
function rubyFragmentsOf(xml: string): string[] {
  return xml.match(/<w:ruby>[\s\S]*?<\/w:ruby>/g) ?? []
}

/** concatenated w:t text of the runs inside a ruby part (w:rubyBase / w:rt) */
function rubyPartText(rubyNode: XNode, part: 'w:rubyBase' | 'w:rt'): string {
  const partNode = findChild(rubyNode, part)
  if (!partNode) return ''
  let text = ''
  for (const r of childrenOf(partNode)) {
    if (nameOf(r) !== 'w:r') continue
    for (const c of childrenOf(r)) {
      if (nameOf(c) === 'w:t') text += decodeEntities(textOf(c))
    }
  }
  return text
}

/** OOXML on/off toggle, three-state: absent → undefined, explicit off → false */
function onOffOf(parent: XNode, name: string): boolean | undefined {
  const child = findChild(parent, name)
  if (!child) return undefined
  const val = attrsOf(child)['w:val']
  if (val === undefined) return true
  return !['0', 'false', 'none', 'off'].includes(val.toLowerCase())
}

/** w:rFonts with theme references resolved: theme attrs supersede same-slot literal
 * values (ECMA-376 §17.3.2.26). Unresolvable references fall back to the literal. */
function themedRFonts(
  attrs: Record<string, string | undefined>,
  fonts: ThemeFonts | null | undefined,
): { ascii?: string; hAnsi?: string; eastAsia?: string } {
  const themeVal = (ref: string | undefined): string | undefined => {
    if (!ref || !fonts) return undefined
    switch (ref) {
      case 'majorAscii':
      case 'majorHAnsi':
        return fonts.major || undefined
      case 'minorAscii':
      case 'minorHAnsi':
        return fonts.minor || undefined
      case 'majorEastAsia':
        return fonts.majorEastAsia || undefined
      case 'minorEastAsia':
        return fonts.eastAsia || undefined
      case 'majorBidi':
        return fonts.majorCs || undefined
      case 'minorBidi':
        return fonts.minorCs || undefined
      default:
        return undefined
    }
  }
  return {
    ascii: themeVal(attrs['w:asciiTheme']) ?? attrs['w:ascii'],
    hAnsi: themeVal(attrs['w:hAnsiTheme']) ?? attrs['w:hAnsi'],
    eastAsia: themeVal(attrs['w:eastAsiaTheme']) ?? attrs['w:eastAsia'],
  }
}

function buildRun(
  rNode: XNode,
  link?: Run['link'],
  theme?: ThemeColors | null,
  themeFonts?: ThemeFonts | null,
): Run | null {
  let text = ''
  for (const child of childrenOf(rNode)) {
    const name = nameOf(child)
    if (name === 'w:t' || name === 'w:delText') text += decodeEntities(textOf(child))
    else if (name === 'w:tab') text += '\t'
    // In-paragraph page breaks (w:br w:type="page") are encoded as \f and preserved; column/soft breaks become \n
    else if (name === 'w:br') text += attrsOf(child)['w:type'] === 'page' ? '\f' : '\n'
    else if (name === 'w:cr') text += '\n'
    else if (name === 'w:noBreakHyphen') text += '\u2011'
    else if (name === 'w:sym') {
      const a = attrsOf(child)
      const code = parseInt(a['w:char'] ?? '', 16)
      if (Number.isFinite(code))
        text +=
          decodeSymbolChar(a['w:font'] ?? '', code) ?? String.fromCodePoint((code & 0xff) + 0xf000)
    }
  }
  if (text === '') return null

  const rPr = findChild(rNode, 'w:rPr')
  const run: Run = { text }
  if (link) run.link = link
  if (rPr) {
    run.rawRPr = serializeXNode(rPr)
    const rStyle = attrsOf(findChild(rPr, 'w:rStyle') ?? {})['w:val']
    if (rStyle && rStyle !== 'Hyperlink') run.styleId = rStyle
    const bold = onOffOf(rPr, 'w:b')
    if (bold !== undefined) run.bold = bold
    const italic = onOffOf(rPr, 'w:i')
    if (italic !== undefined) run.italic = italic
    if (underlineProp(rPr)) run.underline = true
    else if (attrsOf(findChild(rPr, 'w:u') ?? {})['w:val'] === 'none') run.underline = false
    const strike = onOffOf(rPr, 'w:strike')
    if (strike !== undefined) run.strike = strike
    const color = colorFrom(rPr, theme)
    if (color) run.color = color
    const sz = attrsOf(findChild(rPr, 'w:sz') ?? {})['w:val']
    if (sz) run.sizeHalfPoints = parseInt(sz, 10) || undefined
    const rf = themedRFonts(attrsOf(findChild(rPr, 'w:rFonts') ?? {}), themeFonts)
    const font = rf.eastAsia ?? rf.ascii ?? rf.hAnsi
    if (font) run.font = font
    const fontAscii = rf.ascii ?? rf.hAnsi
    if (fontAscii) run.fontAscii = fontAscii
    const spc = parseInt(attrsOf(findChild(rPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
    if (spc) run.charSpacingTwips = spc
    const wScale = parseInt(attrsOf(findChild(rPr, 'w:w') ?? {})['w:val'] ?? '', 10)
    if (wScale > 0 && wScale !== 100) run.charScalePct = wScale
    const highlight = attrsOf(findChild(rPr, 'w:highlight') ?? {})['w:val']
    if (highlight && highlight !== 'none') run.highlight = highlight
    const vertAlign = attrsOf(findChild(rPr, 'w:vertAlign') ?? {})['w:val']
    if (vertAlign === 'superscript' || vertAlign === 'subscript') run.vertAlign = vertAlign
    const em = attrsOf(findChild(rPr, 'w:em') ?? {})['w:val']
    if (em && em !== 'none') run.em = em as NonNullable<Run['em']>
    const rPrChange = findChild(rPr, 'w:rPrChange')
    if (rPrChange) {
      const a = attrsOf(rPrChange)
      const oldRPr = findChild(rPrChange, 'w:rPr')
      const old: NonNullable<Run['rPrChange']>['old'] = {}
      if (oldRPr) {
        if (boolProp(oldRPr, 'w:b')) old.bold = true
        if (boolProp(oldRPr, 'w:i')) old.italic = true
        if (underlineProp(oldRPr)) old.underline = true
        if (boolProp(oldRPr, 'w:strike')) old.strike = true
        const oc = colorFrom(oldRPr, theme)
        if (oc) old.color = oc
        const osz = attrsOf(findChild(oldRPr, 'w:sz') ?? {})['w:val']
        if (osz) old.sizeHalfPoints = parseInt(osz, 10) || undefined
        const ofonts = attrsOf(findChild(oldRPr, 'w:rFonts') ?? {})
        const of = ofonts['w:eastAsia'] ?? ofonts['w:ascii'] ?? ofonts['w:hAnsi']
        if (of) old.font = of
        const ofa = ofonts['w:ascii'] ?? ofonts['w:hAnsi']
        if (ofa) old.fontAscii = ofa
        const ospc = parseInt(attrsOf(findChild(oldRPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
        if (ospc) old.charSpacingTwips = ospc
        const owScale = parseInt(attrsOf(findChild(oldRPr, 'w:w') ?? {})['w:val'] ?? '', 10)
        if (owScale > 0 && owScale !== 100) old.charScalePct = owScale
        const ohighlight = attrsOf(findChild(oldRPr, 'w:highlight') ?? {})['w:val']
        if (ohighlight && ohighlight !== 'none') old.highlight = ohighlight
        const overtAlign = attrsOf(findChild(oldRPr, 'w:vertAlign') ?? {})['w:val']
        if (overtAlign === 'superscript' || overtAlign === 'subscript') old.vertAlign = overtAlign
        const ostyle = attrsOf(findChild(oldRPr, 'w:rStyle') ?? {})['w:val']
        if (ostyle && ostyle !== 'Hyperlink') old.styleId = ostyle
      }
      run.rPrChange = {
        author: a['w:author'] ?? '',
        ...(a['w:date'] ? { date: a['w:date'] } : {}),
        ...(a['w:id'] ? { id: a['w:id'] } : {}),
        ...(Object.keys(old).length > 0 ? { old } : {}),
      }
    }
  }
  // Symbol-encoded fonts (Symbol/Wingdings…): swap the glyph codes for their Unicode
  // equivalents and drop the font, so the text survives systems without those fonts
  if (run.font) {
    const decoded = decodeSymbolText(run.font, run.text)
    if (decoded !== null) {
      run.text = decoded
      delete run.font
      delete run.fontAscii
      if (run.rawRPr) run.rawRPr = run.rawRPr.replace(/<w:rFonts[^>]*\/>/, '')
    }
  }
  return run
}

function mergeRuns(runs: Run[]): Run[] {
  const merged: Run[] = []
  for (const run of runs) {
    const prev = merged[merged.length - 1]
    if (prev && sameStyle(prev, run)) prev.text += run.text
    else merged.push({ ...run })
  }
  return merged
}

function sameStyle(a: Run, b: Run): boolean {
  // reference markers, index entries, cross-references and inline math are atomic; never merge
  if (a.noteRef || b.noteRef || a.xeTerm !== undefined || b.xeTerm !== undefined) return false
  if (a.refField !== undefined || b.refField !== undefined) return false
  if (a.instrField !== undefined || b.instrField !== undefined) return false
  if (a.math || b.math) return false
  if (a.ruby || b.ruby) return false
  return (
    (a.rawRPr ?? '') === (b.rawRPr ?? '') &&
    a.styleId === b.styleId &&
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    !!a.strike === !!b.strike &&
    a.color === b.color &&
    a.sizeHalfPoints === b.sizeHalfPoints &&
    a.font === b.font &&
    a.fontAscii === b.fontAscii &&
    a.highlight === b.highlight &&
    a.vertAlign === b.vertAlign &&
    (a.link?.href ?? '') === (b.link?.href ?? '') &&
    (a.link?.rId ?? '') === (b.link?.rId ?? '') &&
    (a.commentIds ?? []).join(',') === (b.commentIds ?? []).join(',') &&
    sameRevision(a.ins, b.ins) &&
    sameRevision(a.del, b.del)
  )
}

function sameRevision(a: RevisionInfo | undefined, b: RevisionInfo | undefined): boolean {
  if (!a || !b) return !a === !b
  return a.author === b.author && a.date === b.date && a.id === b.id
}

function tableSummary(xml: string): { label: string; previewText: string } {
  const rows = (xml.match(/<w:tr[\s>]/g) ?? []).length
  const firstRow = /<w:tr[\s>][\s\S]*?<\/w:tr>/.exec(xml)?.[0] ?? ''
  const cols = (firstRow.match(/<w:tc[\s>]/g) ?? []).length
  return { label: `Table ${rows}×${cols}`, previewText: plainText(xml).slice(0, 120) }
}

/**
 * Display-only table structure. Nested tables render as read-only sub-tables
 * inside their cell; the exact original bytes are what get saved, so lossiness
 * here only affects on-screen rendering.
 */
function extractTable(xml: string, ctx: BuildContext): TableModel | undefined {
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    return undefined
  }
  const tbl = parsed.find((n) => nameOf(n) === 'w:tbl')
  if (!tbl) return undefined
  const model = extractTableModel(tbl, ctx)
  if (!model) return undefined
  const rawTrPrs: Array<string | null> = model.rows.map(() => null)
  attachRawTablePr(xml, model.rows, rawTrPrs)
  if (rawTrPrs.some((r) => r !== null)) model.rawTrPrs = rawTrPrs
  return model
}

/** One w:tbl node → display model (shared by top-level tables and tables nested in cells) */
/**
 * Per-column widths reconstructed from cell w:tcW (dxa) across all rows: the first
 * un-spanned cell seen per grid slot wins. Undefined unless every column got a value —
 * partial data would skew the ratio worse than the tblGrid fallback.
 */
function tcwColumnWidths(tbl: XNode): number[] | undefined {
  const cols: number[] = []
  let colCount = 0
  for (const tr of childrenThroughSdt(tbl, 'w:tr')) {
    let idx = 0
    for (const tc of childrenThroughSdt(tr, 'w:tc')) {
      const tcPr = findChild(tc, 'w:tcPr')
      const span = Math.max(
        1,
        Number(attrsOf(findChild(tcPr ?? {}, 'w:gridSpan') ?? {})['w:val']) || 1,
      )
      const a = attrsOf(findChild(tcPr ?? {}, 'w:tcW') ?? {})
      const w = !a['w:type'] || a['w:type'] === 'dxa' ? Number(a['w:w']) || 0 : 0
      if (span === 1 && w > 0 && !(cols[idx] > 0)) cols[idx] = w
      idx += span
    }
    colCount = Math.max(colCount, idx)
  }
  if (colCount === 0) return undefined
  for (let i = 0; i < colCount; i++) if (!(cols[i] > 0)) return undefined
  return cols.slice(0, colCount)
}

function extractTableModel(tbl: XNode, ctx: BuildContext): TableModel | undefined {
  const grid = findChild(tbl, 'w:tblGrid')
  let colWidthsPct: number[] | undefined
  let colWidthsTwips: number[] | undefined
  if (grid) {
    const widths = findChildren(grid, 'w:gridCol').map((c) => Number(attrsOf(c)['w:w']) || 0)
    const total = widths.reduce((a, b) => a + b, 0)
    if (total > 0) {
      colWidthsPct = widths.map((w) => (w / total) * 100)
      if (widths.every((w) => w > 0)) colWidthsTwips = widths
    }
  }
  // Cell-level w:tcW is Word's actual layout input for auto tables; generators often
  // leave a stale evenly-split tblGrid behind. When the two disagree, tcW wins.
  const tcwWidths = tcwColumnWidths(tbl)
  if (tcwWidths) {
    const tcwTotal = tcwWidths.reduce((a, b) => a + b, 0)
    const tcwPct = tcwWidths.map((w) => (w / tcwTotal) * 100)
    const disagree =
      !colWidthsPct ||
      colWidthsPct.length !== tcwPct.length ||
      colWidthsPct.some((w, i) => Math.abs(w - tcwPct[i]) > 2)
    if (disagree) {
      colWidthsPct = tcwPct
      colWidthsTwips = tcwWidths
    }
  }
  const tblPrNode = findChild(tbl, 'w:tblPr')
  const tblW = attrsOf(findChild(tblPrNode ?? {}, 'w:tblW') ?? {})
  let widthPct: number | undefined
  if (tblW['w:type'] === 'pct') {
    const raw = String(tblW['w:w'] ?? '')
    // The pct unit is 1/50 of a percentage point; some generators write a literal "NN%"
    const pct = raw.endsWith('%') ? parseFloat(raw) : Number(raw) / 50
    if (Number.isFinite(pct) && pct > 0 && pct <= 100) widthPct = pct
  }
  const cellMar = cellMarginsOf(findChild(tblPrNode ?? {}, 'w:tblCellMar'))
  const tblBorders = borderLinesOf(findChild(tblPrNode ?? {}, 'w:tblBorders'), true)
  const tblJc = attrsOf(findChild(tblPrNode ?? {}, 'w:jc') ?? {})['w:val']
  const tblAlign =
    tblJc === 'center' ? 'center' : tblJc === 'right' || tblJc === 'end' ? 'right' : undefined
  const tblInd = attrsOf(findChild(tblPrNode ?? {}, 'w:tblInd') ?? {})
  const tblIndTwips = !tblInd['w:type'] || tblInd['w:type'] === 'dxa' ? Number(tblInd['w:w']) : NaN

  const rows: TableCell[][] = []
  const rowHeightsTwips: Array<number | null> = []
  const rowRevisions: NonNullable<TableModel['rowRevisions']> = []
  for (const tr of childrenThroughSdt(tbl, 'w:tr')) {
    const cells: TableCell[] = []
    for (const tc of childrenThroughSdt(tr, 'w:tc')) {
      const cell = extractCell(tc, ctx)
      const prev = cells[cells.length - 1]
      // Legacy horizontal merge: continue cells fold into the cell to their left (same effect
      // as gridSpan)
      if (cell.hMerge === 'continue' && prev) {
        prev.colSpan = (prev.colSpan ?? 1) + (cell.colSpan ?? 1)
        continue
      }
      cells.push(cell)
    }
    if (cells.length > 0) {
      rows.push(cells)
      const trPr = findChild(tr, 'w:trPr')
      const h = trPr ? Number(attrsOf(findChild(trPr, 'w:trHeight') ?? {})['w:val']) : NaN
      rowHeightsTwips.push(Number.isFinite(h) && h > 0 ? h : null)
      rowRevisions.push(trPr ? rowRevisionOf(trPr) : null)
    }
  }
  if (rows.length === 0) return undefined
  applyTableStyleDisplay(rows, findChild(tbl, 'w:tblPr'), ctx)
  // Borders/margins from the table style (styles.xml, basedOn chain included): fallback
  // when the document level declares none
  const styleIdEarly = attrsOf(findChild(tblPrNode ?? {}, 'w:tblStyle') ?? {})['w:val']
  const styleTable = styleIdEarly ? ctx.styles.get(styleIdEarly)?.tableDisplay : undefined
  const effBorders = tblBorders ?? styleTable?.borders
  const effCellMar = cellMar ?? styleTable?.cellMarTwips
  const model: TableModel = { rows, colWidthsPct }
  if (colWidthsTwips) model.colWidthsTwips = colWidthsTwips
  if (widthPct) model.widthPct = widthPct
  if (effCellMar) model.cellMarTwips = effCellMar
  if (effBorders) model.borders = effBorders
  if (tblAlign) model.align = tblAlign
  if (Number.isFinite(tblIndTwips) && tblIndTwips !== 0) model.indentTwips = tblIndTwips
  const tblStyle = attrsOf(findChild(findChild(tbl, 'w:tblPr') ?? {}, 'w:tblStyle') ?? {})['w:val']
  if (tblStyle) model.tblStyleId = tblStyle
  if (tblPrNode && boolProp(tblPrNode, 'w:bidiVisual')) model.bidiVisual = true
  if (rowHeightsTwips.some((h) => h !== null)) model.rowHeightsTwips = rowHeightsTwips
  if (rowRevisions.some((r) => r !== null)) model.rowRevisions = rowRevisions
  return model
}

/** trPr w:ins / w:del → row-level revision (inserted/deleted row) */
function rowRevisionOf(trPr: XNode): ({ kind: 'ins' | 'del' } & RevisionInfo) | null {
  for (const kind of ['ins', 'del'] as const) {
    const node = findChild(trPr, `w:${kind}`)
    if (!node) continue
    const a = attrsOf(node)
    return {
      kind,
      author: a['w:author'] ?? '',
      ...(a['w:date'] ? { date: a['w:date'] } : {}),
      ...(a['w:id'] ? { id: a['w:id'] } : {}),
    }
  }
  return null
}

/**
 * Attach each cell's rawTcPr and each row's rawTrPr from the original table XML (byte
 * fidelity: surgically patched on regeneration so unmodeled tcMar/textDirection/
 * tblHeader etc. are not lost). Uses depth-aware splitXmlChildren so nested tables do
 * not misalign; gives up when row/column counts disagree with the parse result
 * (conservative — never attach to the wrong cell).
 */
function attachRawTablePr(xml: string, rows: TableCell[][], rawTrPrs: Array<string | null>): void {
  const open = /<w:tbl[\s>]/.exec(xml)
  if (!open) return
  const innerStart = xml.indexOf('>', open.index) + 1
  const innerEnd = xml.lastIndexOf('</w:tbl>')
  if (innerStart <= 0 || innerEnd < 0) return
  const trs = splitXmlChildren(xml.slice(innerStart, innerEnd)).filter((c) => c.name === 'w:tr')
  if (trs.length !== rows.length) return
  trs.forEach((tr, ri) => {
    const trOpenEnd = tr.xml.indexOf('>') + 1
    const trInner = tr.xml.slice(trOpenEnd, tr.xml.lastIndexOf('</w:tr>'))
    const kids = splitXmlChildren(trInner)
    const trPr = kids.find((k) => k.name === 'w:trPr')
    if (trPr) rawTrPrs[ri] = trPr.xml
    const tcs = kids.filter((k) => k.name === 'w:tc')
    if (tcs.length !== rows[ri].length) return
    tcs.forEach((tc, ci) => {
      const tcOpenEnd = tc.xml.indexOf('>') + 1
      const tcInner = tc.xml.slice(tcOpenEnd, tc.xml.lastIndexOf('</w:tc>'))
      const tcPr = splitXmlChildren(tcInner).find((k) => k.name === 'w:tcPr')
      if (tcPr) rows[ri][ci].rawTcPr = tcPr.xml
    })
  })
}

/**
 * Layer the referenced table style's fills / first-row formatting under the
 * cells' explicit properties, honoring the w:tblLook flags. Display-only:
 * untouched tables still save byte-identically.
 */
function applyTableStyleDisplay(
  rows: TableCell[][],
  tblPr: XNode | undefined,
  ctx: BuildContext,
): void {
  if (!tblPr) return
  const styleId = attrsOf(findChild(tblPr, 'w:tblStyle') ?? {})['w:val']
  const ts = styleId ? ctx.styles.get(styleId)?.tableDisplay : undefined
  if (!ts) return

  const look = attrsOf(findChild(tblPr, 'w:tblLook') ?? {})
  const bits = parseInt(look['w:val'] ?? '', 16)
  const flag = (attr: string, bit: number, dflt: boolean): boolean =>
    look[attr] !== undefined
      ? look[attr] !== '0' && look[attr] !== 'false'
      : Number.isFinite(bits)
        ? (bits & bit) !== 0
        : dflt
  const firstRowOn = flag('w:firstRow', 0x20, true)
  const hBandOn = !flag('w:noHBand', 0x200, false)

  rows.forEach((row, r) => {
    const isFirst = firstRowOn && r === 0
    const bandRow = firstRowOn ? r - 1 : r
    const bandFill =
      hBandOn && bandRow >= 0 ? (bandRow % 2 === 0 ? ts.band1Fill : ts.band2Fill) : undefined
    for (const cell of row) {
      if (cell.fill === undefined) {
        cell.fill = (isFirst ? (ts.firstRow?.fill ?? ts.fill) : (bandFill ?? ts.fill)) ?? undefined
      }
      if (isFirst) {
        if (ts.firstRow?.bold && cell.bold === undefined) cell.bold = true
        if (ts.firstRow?.color && !cell.color) cell.color = ts.firstRow.color
      }
    }
  })
}

function borderLinesOf(node: XNode | undefined, withInside: true): TableBorders | undefined
function borderLinesOf(node: XNode | undefined, withInside: false): CellBorders | undefined
function borderLinesOf(node: XNode | undefined, withInside: boolean): TableBorders | undefined {
  if (!node) return undefined
  const ALIAS: Record<string, keyof TableBorders> = {
    'w:top': 'top',
    'w:left': 'left',
    'w:bottom': 'bottom',
    'w:right': 'right',
    'w:start': 'left',
    'w:end': 'right',
    ...(withInside ? { 'w:insideH': 'insideH', 'w:insideV': 'insideV' } : {}),
  }
  const borders: TableBorders = {}
  for (const [tag, side] of Object.entries(ALIAS)) {
    const child = findChild(node, tag)
    if (!child || borders[side]) continue
    const a = attrsOf(child)
    if (!a['w:val']) continue
    borders[side] = {
      style: a['w:val'],
      ...(a['w:sz'] ? { szEighths: Number(a['w:sz']) || undefined } : {}),
      ...(a['w:color'] ? { color: a['w:color'] } : {}),
    }
  }
  return Object.keys(borders).length > 0 ? borders : undefined
}

function cellMarginsOf(node: XNode | undefined): CellMargins | undefined {
  if (!node) return undefined
  const SIDES: Array<[string, keyof CellMargins]> = [
    ['w:top', 'top'],
    ['w:left', 'left'],
    ['w:bottom', 'bottom'],
    ['w:right', 'right'],
    ['w:start', 'left'],
    ['w:end', 'right'],
  ]
  const m: CellMargins = {}
  for (const [tag, side] of SIDES) {
    const a = attrsOf(findChild(node, tag) ?? {})
    if (a['w:type'] && a['w:type'] !== 'dxa') continue
    const v = Number(a['w:w'])
    if (Number.isFinite(v) && v >= 0 && m[side] === undefined) m[side] = v
  }
  return Object.keys(m).length > 0 ? m : undefined
}

function extractCell(tc: XNode, ctx: BuildContext): TableCell {
  const cell: TableCell = { paras: [] }
  const richParas: NonNullable<TableCell['richParas']> = []

  const tcPr = findChild(tc, 'w:tcPr')
  if (tcPr) {
    const span = Number(attrsOf(findChild(tcPr, 'w:gridSpan') ?? {})['w:val'])
    if (span > 1) cell.colSpan = span
    const vMerge = findChild(tcPr, 'w:vMerge')
    if (vMerge) {
      cell.vMerge = attrsOf(vMerge)['w:val'] === 'restart' ? 'restart' : 'continue'
    }
    const fill = attrsOf(findChild(tcPr, 'w:shd') ?? {})['w:fill']
    if (fill && fill !== 'auto') cell.fill = fill
    const vAlign = attrsOf(findChild(tcPr, 'w:vAlign') ?? {})['w:val']
    if (vAlign === 'center' || vAlign === 'bottom' || vAlign === 'top') cell.vAlign = vAlign
    const tcMar = cellMarginsOf(findChild(tcPr, 'w:tcMar'))
    if (tcMar) cell.cellMarTwips = tcMar
    const dir = attrsOf(findChild(tcPr, 'w:textDirection') ?? {})['w:val']
    if (dir === 'tbRl' || dir === 'tbRlV') cell.textDirection = 'tbRl'
    else if (dir === 'btLr' || dir === 'btLrV') cell.textDirection = 'btLr'
    const hMerge = findChild(tcPr, 'w:hMerge')
    if (hMerge) cell.hMerge = attrsOf(hMerge)['w:val'] === 'restart' ? 'restart' : 'continue'
    const borders = borderLinesOf(findChild(tcPr, 'w:tcBorders'), false)
    if (borders) cell.borders = borders
    for (const kind of ['ins', 'del'] as const) {
      const node = findChild(tcPr, kind === 'ins' ? 'w:cellIns' : 'w:cellDel')
      if (!node) continue
      const a = attrsOf(node)
      cell.cellRevision = {
        kind,
        author: a['w:author'] ?? '',
        ...(a['w:date'] ? { date: a['w:date'] } : {}),
        ...(a['w:id'] ? { id: a['w:id'] } : {}),
      }
      break
    }
  }

  // Tables nested in a cell: parsed as read-only sub-tables (byte fidelity is the
  // outer table's responsibility)
  const nested = childrenThroughSdt(tc, 'w:tbl')
    .map((n) => extractTableModel(n, ctx))
    .filter((m): m is TableModel => !!m)
  if (nested.length > 0) cell.nestedTables = nested

  let sawBold = false
  let sawNonBold = false
  for (const p of childrenThroughSdt(tc, 'w:p')) {
    cell.paras.push(textOf(p))
    const pPr = findChild(p, 'w:pPr')
    const format = extractParaFormat(pPr ?? {})
    const cellStyleId = pPr ? attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] : undefined
    const cellRef = listRefOf(ctx, pPr, cellStyleId)
    const list = cellRef
      ? {
          kind: listKindOf(ctx, cellRef.numId, cellRef.ilvl),
          numId: cellRef.numId,
          ilvl: cellRef.ilvl,
        }
      : undefined
    richParas.push({ ...format, ...(list ? { list } : {}), runs: extractRuns(p, ctx) })
    if (!cell.align) {
      const jc = attrsOf(findChild(findChild(p, 'w:pPr') ?? {}, 'w:jc') ?? {})['w:val']
      if (jc === 'center' || jc === 'right' || jc === 'left' || jc === 'justify') {
        cell.align = jc
      }
    }
    for (const r of findChildren(p, 'w:r')) {
      const rPr = findChild(r, 'w:rPr')
      if (rPr && boolProp(rPr, 'w:b')) sawBold = true
      else sawNonBold = true
      if (!cell.color && rPr) {
        const color = colorFrom(rPr, ctx.themeColors)
        if (color) cell.color = color
      }
    }
  }
  // drop trailing empty paragraphs so cells don't get artificially tall
  while (cell.paras.length > 1 && cell.paras[cell.paras.length - 1] === '') {
    cell.paras.pop()
    richParas.pop()
  }
  cell.richParas = richParas
  if (sawBold && !sawNonBold) cell.bold = true
  return cell
}

function hfPartInfo(
  part: { text: string; hasPageNumber: boolean; paras: HfParagraph[]; images?: HfImage[] } | null,
): HfPartInfo | null {
  if (!part) return null
  return {
    text: part.text,
    hasPageNumber: part.hasPageNumber,
    paras: part.paras,
    ...(part.images?.length ? { images: part.images } : {}),
  }
}

/**
 * display-only images of a header/footer part (logos etc.): resolves a:blip r:embed
 * and VML v:imagedata r:id from the part's own rels; watermarks (v:textpath) excluded.
 * The save path does not go through here -- image paragraphs keep their original bytes
 * when the part is regenerated.
 */
async function hfImages(zip: JSZip, partPath: string, partXml: string): Promise<HfImage[]> {
  if (!partXml.includes('<a:blip') && !partXml.includes('<v:imagedata')) return []
  const relsPath = partPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const rels = await parseRels(zip, relsPath)
  const out: HfImage[] = []
  for (const frag of partXml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g) ?? []) {
    const rId = /<a:blip[^>]*r:embed="([^"]+)"/.exec(frag)?.[1]
    if (!rId) continue
    const dataUrl = await mediaDataUrl(zip, rels, rId)
    if (!dataUrl) continue
    const image: HfImage = { dataUrl }
    const extent = /<wp:extent[^>]*\/?>/.exec(frag)?.[0] ?? ''
    const cx = parseInt(/cx="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    const cy = parseInt(/cy="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    if (Number.isFinite(cx) && cx > 0) image.widthPx = Math.round(cx / EMU_PER_PX)
    if (Number.isFinite(cy) && cy > 0) image.heightPx = Math.round(cy / EMU_PER_PX)
    if (/<wp:anchor[\s>]/.test(frag)) image.floating = true
    out.push(image)
  }
  for (const frag of partXml.match(/<w:pict>[\s\S]*?<\/w:pict>/g) ?? []) {
    if (frag.includes('<v:textpath')) continue
    const rId = /<v:imagedata[^>]*r:id="([^"]+)"/.exec(frag)?.[1]
    if (!rId) continue
    const dataUrl = await mediaDataUrl(zip, rels, rId)
    if (!dataUrl) continue
    // VML dimensions live in the v:shape style attribute (pt)
    const style = /<v:shape[^>]*style="([^"]*)"/.exec(frag)?.[1] ?? ''
    const w = parseFloat(/width:([\d.]+)pt/.exec(style)?.[1] ?? '')
    const h = parseFloat(/height:([\d.]+)pt/.exec(style)?.[1] ?? '')
    const image: HfImage = { dataUrl }
    if (Number.isFinite(w) && w > 0) image.widthPx = Math.round((w / 72) * 96)
    if (Number.isFinite(h) && h > 0) image.heightPx = Math.round((h / 72) * 96)
    out.push(image)
  }
  return out
}

/** settings.xml <w:evenAndOddHeaders/> (w:val="0|false" counts as off) */
async function parseEvenAndOddHeaders(zip: JSZip): Promise<boolean> {
  const file = zip.file('word/settings.xml')
  if (!file) return false
  const m = /<w:evenAndOddHeaders[^>]*\/>/.exec(await file.async('string'))
  return m !== null && !/w:val="(?:0|false)"/.test(m[0])
}

/** header/footer part XML -> display content (PAGE fields shown as '#') */
function hfContentFromXml(
  xml: string,
  kind: 'header' | 'footer',
  theme?: ThemeColors | null,
): { text: string; hasPageNumber: boolean; watermark: string | null; paras: HfParagraph[] } {
  // Rewrite each field span (begin..end) for display. PAGE becomes the '#'
  // marker and NUMPAGES the private-use marker (the renderer substitutes real
  // numbers), dropping their stale cached results; other fields (DATE, STYLEREF,
  // ...) keep their cached result runs (Word refreshes them on open). fldChar
  // attribute matching is tolerant (Pages writes w:fldLock="0" etc.).
  // hasPageNumber is set by the same match that emits '#', so the two can't drift
  // (Word may split "PAGE" across several instrText runs).
  let hasPageNumber = false
  let cleaned = xml.replace(
    /<w:fldChar[^>]*w:fldCharType="begin"[^>]*\/>[\s\S]*?<w:fldChar[^>]*w:fldCharType="end"[^>]*\/>/g,
    (span) => {
      const instr = (span.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? [])
        .map((m) => m.replace(/<[^>]+>/g, ''))
        .join('')
      const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(span)?.[0] ?? ''
      // the span starts inside the begin run and ends inside the end run, so
      // the replacement closes/reopens the enclosing w:r to stay balanced
      // (the leftover edge runs end up empty and are dropped later)
      const emit = (inner: string) => `</w:r>${inner}<w:r>`
      if (/\bNUMPAGES\b/.test(instr)) {
        return emit(`<w:r>${rPr}<w:t>${TOTAL_PAGES_MARK}</w:t></w:r>`)
      }
      if (/\bPAGE\b/.test(instr)) {
        hasPageNumber = true
        return emit(`<w:r>${rPr}<w:t>#</w:t></w:r>`)
      }
      const cached = /<w:fldChar[^>]*w:fldCharType="separate"[^>]*\/>([\s\S]*)$/.exec(span)?.[1]
      // complete result runs between separate and end (partial run fragments at the edges drop out)
      return emit(
        (cached?.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g) ?? [])
          .filter((run) => run.includes('<w:t'))
          .join(''),
      )
    },
  )
  // <w:fldSimple w:instr=" PAGE "> single-element field form
  cleaned = cleaned.replace(
    /<w:fldSimple[^>]*w:instr="([^"]*)"[^>]*(?:\/>|>([\s\S]*?)<\/w:fldSimple>)/g,
    (whole, instr: string, inner: string | undefined) => {
      const rPr = inner ? (/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(inner)?.[0] ?? '') : ''
      if (/\bNUMPAGES\b/.test(instr)) return `<w:r>${rPr}<w:t>${TOTAL_PAGES_MARK}</w:t></w:r>`
      if (/\bPAGE\b/.test(instr)) {
        hasPageNumber = true
        return `<w:r>${rPr}<w:t>#</w:t></w:r>`
      }
      return inner ?? whole
    },
  )
  return {
    text: plainText(cleaned),
    hasPageNumber,
    watermark: kind === 'header' ? readWatermarkText(xml) : null,
    // strip leftover field chars so the '#' page marker parses as plain text
    paras: hfParagraphs(cleaned.replace(/<w:fldChar[^>]*\/>/g, ''), theme),
  }
}

/** Plain-text content of a header/footer part referenced by any sectPr. */
async function readHeaderFooterPart(
  zip: JSZip,
  documentXml: string,
  rels: Map<string, RelInfo>,
  kind: 'header' | 'footer',
  hfType: 'default' | 'first' | 'even' = 'default',
  theme?: ThemeColors | null,
): Promise<{
  text: string
  hasPageNumber: boolean
  watermark: string | null
  paras: HfParagraph[]
  images?: HfImage[]
} | null> {
  const refs = documentXml.match(new RegExp(`<w:${kind}Reference[^>]*/>`, 'g')) ?? []
  const typed = refs.find((r) => r.includes(`w:type="${hfType}"`))
  // untyped references count as default (w:type is technically required but often omitted)
  const ref = hfType === 'default' ? (typed ?? refs.find((r) => !/w:type="/.test(r))) : typed
  if (!ref) return null
  const rId = /r:id="([^"]+)"/.exec(ref)?.[1]
  const target = rId ? rels.get(rId)?.target : undefined
  if (!target) return null
  const path = target.startsWith('/') ? target.slice(1) : `word/${target}`
  const file = zip.file(path)
  if (!file) return null
  const xml = await file.async('string')
  const content = hfContentFromXml(xml, kind, theme)
  const images = await hfImages(zip, path, xml)
  return images.length > 0 ? { ...content, images } : content
}

/** All header/footer parts by rId (multi-section docs look them up via each section's sectPr refs) */
async function parseAllHfParts(
  zip: JSZip,
  rels: Map<string, RelInfo>,
  theme?: ThemeColors | null,
): Promise<Record<string, HfPartInfo>> {
  const out: Record<string, HfPartInfo> = {}
  for (const [rId, rel] of rels) {
    const kind = rel.type.endsWith('/header')
      ? 'header'
      : rel.type.endsWith('/footer')
        ? 'footer'
        : null
    if (!kind) continue
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`
    const file = zip.file(path)
    if (!file) continue
    const xml = await file.async('string')
    const content = hfContentFromXml(xml, kind, theme)
    const images = await hfImages(zip, path, xml)
    out[rId] = {
      text: content.text,
      hasPageNumber: content.hasPageNumber,
      paras: content.paras,
      ...(images.length > 0 ? { images } : {}),
    }
  }
  return out
}

/** rich paragraphs of a header/footer part (watermark/drawing paragraphs skipped) */
function hfParagraphs(partXml: string, theme?: ThemeColors | null): HfParagraph[] {
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(partXml) as XNode[]
  } catch {
    return []
  }
  const root = parsed.find((n) => nameOf(n) === 'w:hdr' || nameOf(n) === 'w:ftr')
  if (!root) return []
  // header parts have their own rels; hyperlink targets are not resolved here
  const ctx = {
    rels: new Map(),
    noteNumbers: new Map(),
    themeColors: theme,
  } as unknown as BuildContext
  const out: HfParagraph[] = []
  for (const pNode of findChildren(root, 'w:p')) {
    const runs = extractRuns(pNode, ctx)
    if (runs.length === 0 && (findChild(pNode, 'w:r') || findChild(pNode, 'w:pict'))) {
      // Government-style footers keep their text (e.g. the "— PAGE —" page number)
      // inside a VML textbox shape; surface those inner paragraphs instead of dropping
      // the content. Watermark / decorative drawing paragraphs still skip.
      out.push(...textboxParagraphs(pNode, ctx))
      continue
    }
    const pPr = findChild(pNode, 'w:pPr')
    out.push({ ...(pPr ? extractParaFormat(pPr) : {}), runs })
  }
  // trailing all-empty paragraphs are layout noise
  while (out.length > 0 && out[out.length - 1].runs.length === 0) out.pop()
  return out
}

/** text paragraphs nested inside textbox shapes (VML v:textbox / DrawingML wps:txbx → w:txbxContent) */
function textboxParagraphs(pNode: XNode, ctx: BuildContext): HfParagraph[] {
  const out: HfParagraph[] = []
  const walk = (node: XNode) => {
    if (nameOf(node) === 'w:txbxContent') {
      for (const inner of findChildren(node, 'w:p')) {
        const runs = extractRuns(inner, ctx)
        if (runs.length === 0) continue
        const pPr = findChild(inner, 'w:pPr')
        out.push({ ...(pPr ? extractParaFormat(pPr) : {}), runs })
      }
      return
    }
    for (const child of childrenOf(node)) walk(child)
  }
  walk(pNode)
  return out
}

async function parseNotesPart(zip: JSZip, kind: 'footnote' | 'endnote'): Promise<NoteInfo[]> {
  const file = zip.file(NOTE_PART_PATH[kind])
  if (!file) return []
  return parseNotesXml(await file.async('string'), kind)
}

async function parseSources(zip: JSZip): Promise<SourceInfo[]> {
  const path = await findSourcesPart(zip)
  if (!path) return []
  return parseSourcesXml(await zip.file(path)!.async('string'))
}

async function parseTheme(
  zip: JSZip,
): Promise<{ fonts: ThemeFonts | null; colors: ThemeColors | null }> {
  const file = zip.file(THEME_PART_PATH)
  if (!file) return { fonts: null, colors: null }
  const xml = await file.async('string')
  return { fonts: readThemeFonts(xml), colors: readThemeColors(xml) }
}

function plainText(xml: string): string {
  const texts: string[] = []
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) texts.push(m[1])
  return decodeEntities(texts.join(''))
}

/** Visible OMML leaf tokens; editing these preserves the surrounding formula tree. */
function mathTokens(xml: string): string[] {
  const tokens: string[] = []
  const re = /<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) tokens.push(decodeEntities(m[1]))
  return tokens
}

function decodeEntities(text: string): string {
  return text
    .replace(
      /&#(?:x([0-9a-f]+)|([0-9]+));/gi,
      (entity, hex: string | undefined, decimal: string | undefined) => {
        const codePoint = parseInt(hex ?? decimal ?? '', hex ? 16 : 10)
        return Number.isFinite(codePoint) &&
          codePoint >= 0 &&
          codePoint <= 0x10ffff &&
          !(codePoint >= 0xd800 && codePoint <= 0xdfff)
          ? String.fromCodePoint(codePoint)
          : entity
      },
    )
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/**
 * Display-only rendering hint for protected field paragraphs. The visible
 * field *result* (w:t runs; instruction text lives in w:instrText and is
 * excluded) is shown instead of a generic chip. Original XML still saves
 * byte-identical.
 */
function fieldDisplayOf(xml: string): FieldDisplay | undefined {
  const styleId = /<w:pStyle w:val="([^"]+)"/.exec(xml)?.[1] ?? ''
  // "TOC1" (Word) or "TOC 1" (Pages export)
  const tocLevel = /^TOC ?([1-9])$/i.exec(styleId)
  if (tocLevel) {
    // TOC entry: title <tab with dot leader> page number
    let left = ''
    let right = ''
    let seenTab = false
    const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\/>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(xml)) !== null) {
      if (m[0] === '<w:tab/>') seenTab = true
      else if (seenTab) right += m[1]
      else left += m[1]
    }
    const anchor = /<w:hyperlink [^>]*w:anchor="([^"]+)"/.exec(xml)?.[1]
    return {
      kind: 'tocLine',
      left: decodeEntities(left).trim(),
      right: decodeEntities(right).trim(),
      level: parseInt(tocLevel[1], 10),
      ...(anchor ? { anchor } : {}),
    }
  }
  const visible = plainText(xml).trim()
  if (visible === '' && /<w:br\s[^>]*w:type="page"/.test(xml)) {
    return { kind: 'pageBreak' }
  }
  if (visible !== '') {
    return { kind: 'text', left: visible }
  }
  return undefined
}

/**
 * TOC entries carry their outline number ("1.", "1.1.") as w:numPr numbering
 * (Pages exports one numId per entry with startOverride restarts). The field
 * result is a display-only cache, so the marker is computed once at parse time
 * and stored on the tocLine FieldDisplay. Counters run document-wide in block
 * order, shared with editable list items (same abstractNum semantics).
 */
function applyTocEntryNumbers(blocks: Block[], numbering: Map<string, NumberingDef>): void {
  if (numbering.size === 0) return
  const items: ListItemRef[] = []
  const tocAt = new Map<number, FieldDisplay>()
  for (const block of blocks) {
    if (block.list?.numId) {
      items.push({ numId: block.list.numId, ilvl: block.list.ilvl })
      continue
    }
    const fd = block.fieldDisplay
    if (block.type !== 'passthrough' || fd?.kind !== 'tocLine' || !block.originalXml) continue
    const numPr = /<w:numPr>[\s\S]*?<\/w:numPr>/.exec(block.originalXml)?.[0]
    if (!numPr) continue
    const numId = /<w:numId w:val="([^"]+)"/.exec(numPr)?.[1]
    if (!numId) continue
    const ilvl = parseInt(/<w:ilvl w:val="(\d+)"/.exec(numPr)?.[1] ?? '0', 10)
    tocAt.set(items.length, fd)
    items.push({ numId, ilvl })
  }
  if (tocAt.size === 0) return
  const markers = computeListMarkers(items, numbering)
  for (const [i, fd] of tocAt) {
    const marker = markers[i]
    // bullets make no sense in front of a TOC entry; only ordered markers show
    if (marker && !/^[•◦▪➢❖✓]$/.test(marker)) fd.num = marker
  }
}

const FIELD_LABELS: Record<string, string> = {
  TOC: 'Auto TOC (updates when opened in Word)',
  PAGE: 'Page number field',
  NUMPAGES: 'Page count field',
  PAGEREF: 'Page reference field',
  REF: 'Cross-reference field',
  SEQ: 'Caption number field',
  HYPERLINK: 'Hyperlink field',
  DATE: 'Date field',
  TIME: 'Time field',
  INCLUDEPICTURE: 'Linked picture field',
  STYLEREF: 'Style reference field',
}

/** Human-readable label for a protected field paragraph, based on its field code. */
function fieldLabel(xml: string): string {
  const instr =
    /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/.exec(xml)?.[1] ??
    /<w:fldSimple[^>]*w:instr="([^"]*)"/.exec(xml)?.[1] ??
    ''
  const keyword = instr.trim().split(/\s+/)[0]?.toUpperCase() ?? ''
  if (keyword && FIELD_LABELS[keyword]) return FIELD_LABELS[keyword]
  if (keyword) return `Field (${keyword})`
  // No field code in this paragraph: it only closes a field started earlier
  // (e.g. the paragraph holding the TOC's fldChar end + page break).
  if (xml.includes('fldCharType="end"') && !xml.includes('fldCharType="begin"')) {
    return xml.includes('w:type="page"') ? 'Field end marker + page break' : 'Field end marker'
  }
  return 'Field (TOC/page number/etc.)'
}

const EMU_PER_PX = 9525

/** display size (wp:extent), paragraph alignment and wrap mode of an image paragraph */
type ImageMeta = Pick<
  Block,
  | 'imageWidthPx'
  | 'imageHeightPx'
  | 'imageAlign'
  | 'imageWrap'
  | 'imageOffsetXEmu'
  | 'imageOffsetYEmu'
  | 'imagePosH'
  | 'imagePosV'
>

function imageMeta(xml: string): ImageMeta {
  const meta: ImageMeta = {}
  const extent = /<wp:extent[^>]*\/?>/.exec(xml)?.[0]
  if (extent) {
    const cx = parseInt(/cx="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    const cy = parseInt(/cy="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    if (Number.isFinite(cx) && cx > 0) meta.imageWidthPx = Math.round(cx / EMU_PER_PX)
    if (Number.isFinite(cy) && cy > 0) meta.imageHeightPx = Math.round(cy / EMU_PER_PX)
  }
  const jc = /<w:jc w:val="([^"]+)"/.exec(xml)?.[1]
  if (jc === 'center') meta.imageAlign = 'center'
  else if (jc === 'right' || jc === 'end') meta.imageAlign = 'right'
  const anchor = /<wp:anchor[^>]*>/.exec(xml)?.[0]
  if (anchor) {
    if (/behindDoc="1"/.test(anchor)) meta.imageWrap = 'behind'
    else if (/<wp:wrapTopAndBottom/.test(xml)) meta.imageWrap = 'topBottom'
    else if (/<wp:wrap(Square|Tight|Through)/.test(xml)) {
      const kind = /<wp:wrap(Square|Tight|Through)/.exec(xml)![1]
      const alignRight =
        /<wp:positionH[^>]*>(?:(?!<\/wp:positionH>)[\s\S])*?<wp:align>right<\/wp:align>/.test(xml)
      const side = alignRight ? 'right' : 'left'
      meta.imageWrap =
        kind === 'Tight'
          ? `tight-${side}`
          : kind === 'Through'
            ? `through-${side}`
            : `square-${side}`
    } else meta.imageWrap = 'front'
    // Parse numeric posOffset for free-position floating images
    const posHBody = /<wp:positionH[^>]*>([\s\S]*?)<\/wp:positionH>/.exec(xml)?.[1] ?? ''
    const posVBody = /<wp:positionV[^>]*>([\s\S]*?)<\/wp:positionV>/.exec(xml)?.[1] ?? ''
    const offsetX = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(posHBody)?.[1]
    const offsetY = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(posVBody)?.[1]
    if (offsetX !== undefined) meta.imageOffsetXEmu = parseInt(offsetX, 10)
    if (offsetY !== undefined) meta.imageOffsetYEmu = parseInt(offsetY, 10)
    // margin-relative wp:align pair = Word position-gallery preset
    const posHFrom = /<wp:positionH[^>]*relativeFrom="([^"]+)"/.exec(xml)?.[1]
    const posVFrom = /<wp:positionV[^>]*relativeFrom="([^"]+)"/.exec(xml)?.[1]
    const alignH = /<wp:align>(left|center|right)<\/wp:align>/.exec(posHBody)?.[1]
    const alignV = /<wp:align>(top|center|bottom)<\/wp:align>/.exec(posVBody)?.[1]
    if (posHFrom === 'margin' && posVFrom === 'margin' && alignH && alignV) {
      meta.imagePosH = alignH as ImageMeta['imagePosH']
      meta.imagePosV = alignV as ImageMeta['imagePosV']
    }
  }
  return meta
}

/** resolve an image relationship id to a data URL (embedded parts only) */
async function mediaDataUrl(
  zip: JSZip,
  rels: Map<string, RelInfo>,
  rId: string,
): Promise<string | null> {
  const rel = rels.get(rId)
  if (!rel || rel.targetMode === 'External') return null
  const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`
  const file = zip.file(path.replace(/^word\/\.\.\//, ''))
  if (!file) return null
  const mime = IMAGE_MIME[path.split('.').pop()?.toLowerCase() ?? '']
  if (!mime) return null
  return `data:${mime};base64,${await file.async('base64')}`
}

async function extractImage(xml: string, ctx: BuildContext): Promise<string | null> {
  // embedded (r:embed -> word/media/...) or linked (r:link -> external URL)
  const rId =
    /<a:blip[^>]*r:embed="([^"]+)"/.exec(xml)?.[1] ?? /<a:blip[^>]*r:link="([^"]+)"/.exec(xml)?.[1]
  if (!rId) return null
  const rel = ctx.rels.get(rId)
  if (!rel) return null

  // linked pictures (Word downloads them on open; we let <img> do the same)
  if (rel.targetMode === 'External' || /^https?:\/\//i.test(rel.target)) {
    return rel.target
  }

  const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`
  const file = ctx.zip.file(path.replace(/^word\/\.\.\//, ''))
  if (!file) return null
  const ext = path.split('.').pop()?.toLowerCase() ?? ''
  const mime = IMAGE_MIME[ext]
  if (!mime || mime === 'image/emf' || mime === 'image/wmf') return null
  const base64 = await file.async('base64')
  return `data:${mime};base64,${base64}`
}

/** resolve a document rel to its zip path ("word/…"), or null for external targets */
function relPartPath(ctx: BuildContext, rId: string | undefined): string | null {
  const rel = rId ? ctx.rels.get(rId) : undefined
  if (!rel || rel.targetMode === 'External') return null
  return (rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`).replace(
    /^word\/\.\.\//,
    '',
  )
}

/**
 * SmartArt degrade: the node texts from the diagram data part (r:dm)
 * become the preview, so the reader still sees the labels the diagram holds.
 */
async function extractDiagramText(xml: string, ctx: BuildContext): Promise<string | null> {
  const path = relPartPath(ctx, /r:dm="([^"]+)"/.exec(xml)?.[1])
  const file = path ? ctx.zip.file(path) : null
  if (!file) return null
  const dataXml = await file.async('string')
  const texts: string[] = []
  for (const t of dataXml.match(/<a:t>[^<]*<\/a:t>/g) ?? []) {
    const s = decodeEntities(t.slice(5, -6)).trim()
    if (s) texts.push(s)
  }
  return texts.length > 0 ? texts.join('\n') : null
}

/**
 * OLE embed degrade: the original packages a preview picture
 * (v:imagedata) and declares its kind (o:OLEObject ProgID) — surface both
 * instead of a bare type label.
 */
async function oleDisplay(
  xml: string,
  ctx: BuildContext,
): Promise<Pick<Block, 'imageDataUrl' | 'oleProgId'>> {
  const out: Pick<Block, 'imageDataUrl' | 'oleProgId'> = {}
  const progId = /<o:OLEObject[^>]*ProgID="([^"]+)"/.exec(xml)?.[1]
  if (progId) out.oleProgId = progId
  const path = relPartPath(ctx, /<v:imagedata[^>]*r:id="([^"]+)"/.exec(xml)?.[1])
  const file = path ? ctx.zip.file(path) : null
  if (file && path) {
    const mime = IMAGE_MIME[path.split('.').pop()?.toLowerCase() ?? '']
    if (mime && mime !== 'image/emf' && mime !== 'image/wmf') {
      out.imageDataUrl = `data:${mime};base64,${await file.async('base64')}`
    }
  }
  return out
}

/**
 * Load and parse the chart part referenced by a chart drawing paragraph.
 * The part's original XML is kept in ctx.chartParts so data edits can be
 * patched into it at save time (the body paragraph itself never changes).
 */
async function extractChart(xml: string, ctx: BuildContext): Promise<ChartDisplay | null> {
  const rId = /<c:chart [^>]*r:id="([^"]+)"/.exec(xml)?.[1]
  const rel = rId ? ctx.rels.get(rId) : undefined
  if (!rel || rel.targetMode === 'External') return null
  const path = (rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`).replace(
    /^word\/\.\.\//,
    '',
  )
  const file = ctx.zip.file(path)
  if (!file) return null
  const partXml = await file.async('string')
  const display = parseChartPartXml(partXml, path)
  if (display) {
    ctx.chartParts[path] = partXml
    const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml)
    const cx = extent ? parseInt(extent[1]!, 10) : NaN
    const cy = extent ? parseInt(extent[2]!, 10) : NaN
    if (Number.isFinite(cx) && cx > 0) display.widthPx = Math.round(cx / EMU_PER_PX)
    if (Number.isFinite(cy) && cy > 0) display.heightPx = Math.round(cy / EMU_PER_PX)
  }
  return display
}

async function parseStyles(
  zip: JSZip,
  theme?: ThemeColors | null,
  themeFonts?: ThemeFonts | null,
): Promise<{ styles: Map<string, StyleInfo>; docDefaults?: DocDefaults }> {
  const styles = new Map<string, StyleInfo>()
  const file = zip.file('word/styles.xml')
  if (!file) return { styles }
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(await file.async('string')) as XNode[]
  } catch (err) {
    console.warn('styles.xml unparseable, styles degraded to empty:', err)
    return { styles }
  }
  const root = parsed.find((n) => nameOf(n) === 'w:styles')
  if (!root) return { styles }

  let docDefaults: DocDefaults | undefined
  const defaultsNode = findChild(root, 'w:docDefaults')
  if (defaultsNode) {
    const dd: DocDefaults = {}
    const rPr = findChild(findChild(defaultsNode, 'w:rPrDefault') ?? {}, 'w:rPr')
    const sz = rPr ? attrsOf(findChild(rPr, 'w:sz') ?? {})['w:val'] : undefined
    if (sz) dd.sizeHalfPoints = parseInt(sz, 10) || undefined
    const ddRf = themedRFonts(rPr ? attrsOf(findChild(rPr, 'w:rFonts') ?? {}) : {}, themeFonts)
    if (ddRf.ascii ?? ddRf.hAnsi) dd.asciiFont = ddRf.ascii ?? ddRf.hAnsi
    if (ddRf.eastAsia) dd.eastAsiaFont = ddRf.eastAsia
    if (rPr) {
      const onFlag = (tag: string) => {
        const node = findChild(rPr, tag)
        if (!node) return undefined
        const val = attrsOf(node)['w:val']
        return val === '0' || val === 'false' ? undefined : true
      }
      if (onFlag('w:b')) dd.bold = true
      if (onFlag('w:i')) dd.italic = true
      const color = colorFrom(rPr, theme)
      if (color) dd.color = color
    }
    const pPr = findChild(findChild(defaultsNode, 'w:pPrDefault') ?? {}, 'w:pPr')
    const spacingAttrs = pPr ? attrsOf(findChild(pPr, 'w:spacing') ?? {}) : {}
    if (spacingAttrs['w:line']) {
      const line = parseInt(spacingAttrs['w:line'], 10)
      const rule = (spacingAttrs['w:lineRule'] ?? 'auto') as 'auto' | 'atLeast' | 'exact'
      if (line > 0) {
        dd.lineRawTwips = line
        dd.lineRule = rule
        if (rule === 'auto') dd.lineSpacing = line / 240
      }
    }
    if (spacingAttrs['w:before'] !== undefined) {
      dd.spaceBeforeTwips = parseInt(spacingAttrs['w:before'], 10) || 0
    }
    if (spacingAttrs['w:after'] !== undefined) {
      dd.spaceAfterTwips = parseInt(spacingAttrs['w:after'], 10) || 0
    }
    if (Object.keys(dd).length > 0) docDefaults = dd
  }

  const basedOnIds = new Map<string, string>()
  const linkedIds = new Map<string, string>()
  // styles with an explicit w:outlineLvl 9 (body text, e.g. TOCHeading basedOn Heading1)
  const outlineOffIds = new Set<string>()
  for (const styleNode of findChildren(root, 'w:style')) {
    const attrs = attrsOf(styleNode)
    const type = attrs['w:type']
    if (type !== 'paragraph' && type !== 'character' && type !== 'table') continue
    const styleId = attrs['w:styleId']
    if (!styleId) continue
    const name = attrsOf(findChild(styleNode, 'w:name') ?? {})['w:val'] ?? styleId
    let headingLevel: number | undefined
    if (type === 'paragraph') {
      const nameMatch = /^heading\s*([1-9])$/i.exec(name) ?? /^Heading([1-9])$/.exec(styleId)
      if (nameMatch) headingLevel = parseInt(nameMatch[1], 10)
      else {
        const pPr = findChild(styleNode, 'w:pPr')
        const outline = pPr ? attrsOf(findChild(pPr, 'w:outlineLvl') ?? {})['w:val'] : undefined
        if (outline !== undefined) {
          const lvl = parseInt(outline, 10)
          if (lvl >= 0 && lvl <= 8) headingLevel = lvl + 1
          else outlineOffIds.add(styleId)
        }
      }
    }
    const basedOn = attrsOf(findChild(styleNode, 'w:basedOn') ?? {})['w:val']
    if (basedOn) basedOnIds.set(styleId, basedOn)
    const link = attrsOf(findChild(styleNode, 'w:link') ?? {})['w:val']
    if (link) linkedIds.set(styleId, link)
    const onFlag = (tag: string): boolean | undefined => {
      const node = findChild(styleNode, tag)
      if (!node) return undefined
      const val = attrsOf(node)['w:val']
      return val === '0' || val === 'false' ? undefined : true
    }
    let numPr: StyleInfo['numPr']
    if (type === 'paragraph') {
      const styleNumPr = findChild(findChild(styleNode, 'w:pPr') ?? {}, 'w:numPr')
      if (styleNumPr) {
        const numId = attrsOf(findChild(styleNumPr, 'w:numId') ?? {})['w:val']
        if (numId && numId !== '0') {
          const ilvl = parseInt(attrsOf(findChild(styleNumPr, 'w:ilvl') ?? {})['w:val'] ?? '0', 10)
          numPr = { numId, ilvl: ilvl || 0 }
        }
      }
    }
    styles.set(styleId, {
      styleId,
      name,
      type,
      headingLevel,
      semiHidden: onFlag('w:semiHidden'),
      qFormat: onFlag('w:qFormat'),
      display: type === 'table' ? undefined : styleDisplayOf(styleNode, theme, themeFonts),
      tableDisplay: type === 'table' ? tableStyleDisplayOf(styleNode, theme) : undefined,
      numPr,
      isDefault: attrs['w:default'] === '1' || attrs['w:default'] === 'true' ? true : undefined,
    })
  }

  // resolve basedOn chains: a style inherits every display prop it doesn't set itself
  const resolved = new Set<string>()
  const resolve = (styleId: string, seen: Set<string>): StyleInfo | undefined => {
    const info = styles.get(styleId)
    if (!info) return undefined
    const parentId = basedOnIds.get(styleId)
    if (resolved.has(styleId) || !parentId || seen.has(styleId)) return info
    seen.add(styleId)
    const parent = resolve(parentId, seen)
    resolved.add(styleId)
    if (parent?.display) {
      info.display = { ...parent.display, ...(info.display ?? {}) }
      if (Object.keys(info.display).length === 0) info.display = undefined
    }
    if (parent?.tableDisplay) {
      info.tableDisplay = mergeTableDisplay(parent.tableDisplay, info.tableDisplay)
    }
    if (
      info.type === 'paragraph' &&
      info.headingLevel === undefined &&
      !outlineOffIds.has(styleId) &&
      parent?.headingLevel
    ) {
      info.headingLevel = parent.headingLevel
    }
    if (info.type === 'paragraph' && !info.numPr && parent?.numPr) info.numPr = parent.numPr
    return info
  }
  for (const styleId of styles.keys()) resolve(styleId, new Set())

  // linkedStyle (w:link): a paragraph style and a character style form one unit (Word
  // "linked styles"). Fill in run-level display properties in both directions (never
  // overriding a style's own) — the common gap is a character-style shell with no rPr,
  // where all run properties live on the linked paragraph style.
  const RUN_KEYS = [
    'sizeHalfPoints',
    'color',
    'bold',
    'italic',
    'underline',
    'strike',
    'font',
    'fontAscii',
  ] as const
  for (const [fromId, toId] of linkedIds) {
    const a = styles.get(fromId)
    const b = styles.get(toId)
    if (!a || !b) continue
    for (const [self, other] of [
      [a, b],
      [b, a],
    ] as const) {
      if (self.type !== 'character' && self.type !== 'paragraph') continue
      const fill: Partial<StyleDisplay> = {}
      for (const key of RUN_KEYS) {
        if (self.display?.[key] === undefined && other.display?.[key] !== undefined) {
          ;(fill as Record<string, unknown>)[key] = other.display[key]
        }
      }
      if (Object.keys(fill).length > 0) self.display = { ...fill, ...(self.display ?? {}) }
    }
    if (a.type === 'character' && b.type === 'paragraph') a.linkedCharShell = true
    if (b.type === 'character' && a.type === 'paragraph') b.linkedCharShell = true
  }

  return { styles, docDefaults }
}

function mergeTableDisplay(
  parent: TableStyleDisplay,
  child: TableStyleDisplay | undefined,
): TableStyleDisplay | undefined {
  const merged: TableStyleDisplay = {
    ...parent,
    ...(child ?? {}),
    ...(parent.firstRow || child?.firstRow
      ? { firstRow: { ...(parent.firstRow ?? {}), ...(child?.firstRow ?? {}) } }
      : {}),
    ...(parent.paraSpacing || child?.paraSpacing
      ? { paraSpacing: { ...(parent.paraSpacing ?? {}), ...(child?.paraSpacing ?? {}) } }
      : {}),
  }
  return Object.keys(merged).length > 0 ? merged : undefined
}

/** fills / first-row formatting a table style contributes on screen */
function tableStyleDisplayOf(
  styleNode: XNode,
  theme?: ThemeColors | null,
): TableStyleDisplay | undefined {
  const display: TableStyleDisplay = {}
  const shdFill = (node: XNode | undefined): string | undefined => {
    const fill = node ? attrsOf(findChild(node, 'w:shd') ?? {})['w:fill'] : undefined
    return fill && fill !== 'auto' ? fill : undefined
  }
  const baseFill = shdFill(findChild(styleNode, 'w:tcPr'))
  if (baseFill) display.fill = baseFill
  for (const cond of findChildren(styleNode, 'w:tblStylePr')) {
    const type = attrsOf(cond)['w:type']
    const tcPr = findChild(cond, 'w:tcPr')
    const fill = shdFill(tcPr)
    if (type === 'firstRow') {
      const rPr = findChild(cond, 'w:rPr')
      const firstRow: NonNullable<TableStyleDisplay['firstRow']> = {}
      if (fill) firstRow.fill = fill
      if (rPr && boolProp(rPr, 'w:b')) firstRow.bold = true
      const color = colorFrom(rPr, theme)
      if (color) firstRow.color = color
      if (Object.keys(firstRow).length > 0) display.firstRow = firstRow
    } else if (type === 'band1Horz' && fill) {
      display.band1Fill = fill
    } else if (type === 'band2Horz' && fill) {
      display.band2Fill = fill
    }
  }
  const styleTblPr = findChild(styleNode, 'w:tblPr')
  const borders = borderLinesOf(findChild(styleTblPr ?? {}, 'w:tblBorders'), true)
  if (borders) display.borders = borders
  const cellMar = cellMarginsOf(findChild(styleTblPr ?? {}, 'w:tblCellMar'))
  if (cellMar) display.cellMarTwips = cellMar
  const stylePPrSpacing = findChild(findChild(styleNode, 'w:pPr') ?? {}, 'w:spacing')
  if (stylePPrSpacing) {
    const a = attrsOf(stylePPrSpacing)
    const ps: NonNullable<TableStyleDisplay['paraSpacing']> = {}
    const before = parseInt(a['w:before'] ?? '', 10)
    if (before >= 0 && a['w:before'] !== undefined) ps.beforeTwips = before
    const after = parseInt(a['w:after'] ?? '', 10)
    if (after >= 0 && a['w:after'] !== undefined) ps.afterTwips = after
    const line = parseInt(a['w:line'] ?? '', 10)
    if (line > 0) {
      ps.lineRawTwips = line
      const rule = (a['w:lineRule'] ?? 'auto') as 'auto' | 'atLeast' | 'exact'
      ps.lineRule = rule
      if (rule === 'auto') ps.lineSpacing = Math.round((line / 240) * 100) / 100
    }
    if (Object.keys(ps).length > 0) display.paraSpacing = ps
  }
  return Object.keys(display).length > 0 ? display : undefined
}

/** display-only formatting the style contributes on screen (Word renders these from styles.xml) */
function styleDisplayOf(
  styleNode: XNode,
  theme?: ThemeColors | null,
  themeFonts?: ThemeFonts | null,
): StyleDisplay | undefined {
  const display: StyleDisplay = {}
  const rPr = findChild(styleNode, 'w:rPr')
  if (rPr) {
    const sz = attrsOf(findChild(rPr, 'w:sz') ?? {})['w:val']
    if (sz) display.sizeHalfPoints = parseInt(sz, 10) || undefined
    const color = colorFrom(rPr, theme)
    if (color) display.color = color
    const bold = onOffOf(rPr, 'w:b')
    if (bold !== undefined) display.bold = bold
    const italic = onOffOf(rPr, 'w:i')
    if (italic !== undefined) display.italic = italic
    const u = attrsOf(findChild(rPr, 'w:u') ?? {})['w:val']
    if (u) display.underline = u !== 'none'
    const strike = onOffOf(rPr, 'w:strike')
    if (strike !== undefined) display.strike = strike
    const rf = themedRFonts(attrsOf(findChild(rPr, 'w:rFonts') ?? {}), themeFonts)
    const font = rf.eastAsia ?? rf.ascii ?? rf.hAnsi
    const fontAscii = rf.ascii ?? rf.hAnsi
    if (fontAscii) display.fontAscii = fontAscii
    if (font) display.font = font
    const spc = parseInt(attrsOf(findChild(rPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
    if (spc) display.charSpacingTwips = spc
  }
  const pPr = findChild(styleNode, 'w:pPr')
  if (pPr) {
    const spacing = attrsOf(findChild(pPr, 'w:spacing') ?? {})
    const line = parseInt(spacing['w:line'] ?? '', 10)
    if (line > 0) {
      const rule = (spacing['w:lineRule'] ?? 'auto') as 'auto' | 'atLeast' | 'exact'
      display.lineRule = rule
      display.lineRawTwips = line
      if (rule === 'auto') {
        display.lineSpacing = line / 240
      }
    }
    if (spacing['w:before'] !== undefined) {
      display.spaceBeforeTwips = parseInt(spacing['w:before'], 10) || 0
    }
    if (spacing['w:after'] !== undefined) {
      display.spaceAfterTwips = parseInt(spacing['w:after'], 10) || 0
    }
    if (boolProp(pPr, 'w:keepNext')) display.keepNext = true
    if (boolProp(pPr, 'w:keepLines')) display.keepLines = true
    if (boolProp(pPr, 'w:contextualSpacing')) display.contextualSpacing = true
    const jc = attrsOf(findChild(pPr, 'w:jc') ?? {})['w:val']
    if (jc === 'center' || jc === 'right' || jc === 'left' || jc === 'justify') display.align = jc
    else if (jc === 'both' || jc === 'distribute') display.align = 'justify'
    const ind = findChild(pPr, 'w:ind')
    if (ind) {
      const a = attrsOf(ind)
      const left = parseInt(a['w:left'] ?? a['w:start'] ?? '', 10)
      if (Number.isFinite(left) && left !== 0) display.indentLeftTwips = left
      const right = parseInt(a['w:right'] ?? a['w:end'] ?? '', 10)
      if (Number.isFinite(right) && right !== 0) display.indentRightTwips = right
      const firstLine = parseInt(a['w:firstLine'] ?? '', 10)
      const hanging = parseInt(a['w:hanging'] ?? '', 10)
      if (hanging > 0) display.indentFirstLineTwips = -hanging
      else if (firstLine > 0) display.indentFirstLineTwips = firstLine
    }
  }
  return Object.keys(display).length > 0 ? display : undefined
}

async function parseRels(zip: JSZip, path: string): Promise<Map<string, RelInfo>> {
  const rels = new Map<string, RelInfo>()
  const file = zip.file(path)
  if (!file) return rels
  const parsed = xmlParser.parse(await file.async('string')) as XNode[]
  const root = parsed.find((n) => nameOf(n) === 'Relationships')
  if (!root) return rels
  for (const relNode of findChildren(root, 'Relationship')) {
    const attrs = attrsOf(relNode)
    if (!attrs['Id']) continue
    rels.set(attrs['Id'], {
      target: attrs['Target'] ?? '',
      type: attrs['Type'] ?? '',
      targetMode: attrs['TargetMode'],
    })
  }
  return rels
}

/** word/comments.xml (+ reply/resolved relations from commentsExtended) -> display list, file order */
async function parseComments(zip: JSZip): Promise<CommentInfo[]> {
  const file = zip.file('word/comments.xml')
  if (!file) return []
  const parsed = xmlParser.parse(await file.async('string')) as XNode[]
  const root = parsed.find((n) => nameOf(n) === 'w:comments')
  if (!root) return []
  const out: CommentInfo[] = []
  for (const node of findChildren(root, 'w:comment')) {
    const attrs = attrsOf(node)
    if (!attrs['w:id']) continue
    const paras = findChildren(node, 'w:p')
    const paraId = paras.length > 0 ? attrsOf(paras[paras.length - 1])['w14:paraId'] : undefined
    out.push({
      id: attrs['w:id'],
      author: attrs['w:author'] ?? '',
      initials: attrs['w:initials'],
      date: attrs['w:date'],
      text: paras.map((p) => textOf(p)).join('\n'),
      ...(paraId ? { paraId } : {}),
    })
  }
  // commentsExtended.xml: paraId → parent paraId / done (Word 2013+ replies and resolution)
  const extFile = zip.file('word/commentsExtended.xml')
  if (extFile) {
    const extXml = await extFile.async('string')
    const byParaId = new Map(out.filter((c) => c.paraId).map((c) => [c.paraId!, c]))
    for (const m of extXml.match(/<w15:commentEx [^>]*\/>/g) ?? []) {
      const paraId = /w15:paraId="([^"]+)"/.exec(m)?.[1]
      const parentParaId = /w15:paraIdParent="([^"]+)"/.exec(m)?.[1]
      const done = /w15:done="(?:1|true)"/.test(m)
      const c = paraId ? byParaId.get(paraId) : undefined
      if (!c) continue
      if (done) c.done = true
      if (parentParaId) {
        const parent = byParaId.get(parentParaId)
        if (parent) c.parentId = parent.id
      }
    }
  }
  return out
}

/** w:documentProtection from word/settings.xml (editing restriction) */
async function parseProtection(zip: JSZip): Promise<DocProtection | null> {
  const file = zip.file('word/settings.xml')
  if (!file) return null
  const xml = await file.async('string')
  const tag = /<w:documentProtection[^>]*\/>/.exec(xml)?.[0]
  if (!tag) return null
  const edit = /w:edit="([^"]+)"/.exec(tag)?.[1]
  if (!edit || edit === 'none') return null
  const enforcement = /w:enforcement="([^"]+)"/.exec(tag)?.[1]
  const hash = /w:hash="([^"]+)"/.exec(tag)?.[1]
  const salt = /w:salt="([^"]+)"/.exec(tag)?.[1]
  const spin = /w:cryptSpinCount="(\d+)"/.exec(tag)?.[1]
  const sid = /w:cryptAlgorithmSid="(\d+)"/.exec(tag)?.[1]
  return {
    edit,
    enforced: enforcement === '1' || enforcement === 'true',
    ...(hash ? { hash } : {}),
    ...(salt ? { salt } : {}),
    ...(spin ? { spinCount: parseInt(spin, 10) } : {}),
    ...(sid ? { algorithmSid: parseInt(sid, 10) } : {}),
  }
}

function parseNumberingLevel(lvlNode: XNode): NumberingLevel {
  const start = parseInt(attrsOf(findChild(lvlNode, 'w:start') ?? {})['w:val'] ?? '1', 10)
  const level: NumberingLevel = {
    numFmt: attrsOf(findChild(lvlNode, 'w:numFmt') ?? {})['w:val'] ?? 'decimal',
    lvlText: decodeEntities(attrsOf(findChild(lvlNode, 'w:lvlText') ?? {})['w:val'] ?? ''),
    start: Number.isFinite(start) ? start : 1,
  }
  const lvlPPr = findChild(lvlNode, 'w:pPr')
  const ind = lvlPPr ? findChild(lvlPPr, 'w:ind') : undefined
  if (ind) {
    const attrs = attrsOf(ind)
    const left = parseInt(attrs['w:left'] ?? attrs['w:start'] ?? '', 10)
    if (left > 0) level.indentLeft = left
    const hanging = parseInt(attrs['w:hanging'] ?? '', 10)
    if (hanging > 0) level.hanging = hanging
  }
  const lvlRPr = findChild(lvlNode, 'w:rPr')
  const sz = lvlRPr ? parseInt(attrsOf(findChild(lvlRPr, 'w:sz') ?? {})['w:val'] ?? '', 10) : NaN
  if (sz > 0) level.szHalfPoints = sz
  const fonts = lvlRPr ? attrsOf(findChild(lvlRPr, 'w:rFonts') ?? {}) : {}
  const font = fonts['w:ascii'] ?? fonts['w:hAnsi'] ?? fonts['w:eastAsia']
  if (font) level.font = font
  return level
}

/** word/numbering.xml -> per-numId level definitions + the bullet/ordered classification */
async function parseNumbering(
  zip: JSZip,
): Promise<{ formats: Map<string, 'bullet' | 'ordered'>; defs: Map<string, NumberingDef> }> {
  const formats = new Map<string, 'bullet' | 'ordered'>()
  const defs = new Map<string, NumberingDef>()
  const file = zip.file('word/numbering.xml')
  if (!file) return { formats, defs }
  const parsed = xmlParser.parse(await file.async('string')) as XNode[]
  const root = parsed.find((n) => nameOf(n) === 'w:numbering')
  if (!root) return { formats, defs }

  const absLevels = new Map<string, Record<number, NumberingLevel>>()
  for (const abs of findChildren(root, 'w:abstractNum')) {
    const absId = attrsOf(abs)['w:abstractNumId']
    if (!absId) continue
    const levels: Record<number, NumberingLevel> = {}
    for (const lvl of findChildren(abs, 'w:lvl')) {
      const ilvl = parseInt(attrsOf(lvl)['w:ilvl'] ?? '', 10)
      if (Number.isFinite(ilvl)) levels[ilvl] = parseNumberingLevel(lvl)
    }
    absLevels.set(absId, levels)
  }
  for (const num of findChildren(root, 'w:num')) {
    const numId = attrsOf(num)['w:numId']
    const absId = attrsOf(findChild(num, 'w:abstractNumId') ?? {})['w:val']
    if (!numId || !absId) continue
    const levels: Record<number, NumberingLevel> = { ...(absLevels.get(absId) ?? {}) }
    const startOverrides: Record<number, number> = {}
    for (const over of findChildren(num, 'w:lvlOverride')) {
      const ilvl = parseInt(attrsOf(over)['w:ilvl'] ?? '', 10)
      if (!Number.isFinite(ilvl)) continue
      const startVal = attrsOf(findChild(over, 'w:startOverride') ?? {})['w:val']
      if (startVal !== undefined) {
        const n = parseInt(startVal, 10)
        if (Number.isFinite(n)) startOverrides[ilvl] = n
      }
      const lvl = findChild(over, 'w:lvl')
      if (lvl) levels[ilvl] = parseNumberingLevel(lvl)
    }
    defs.set(numId, { numId, abstractNumId: absId, levels, startOverrides })
    formats.set(numId, levels[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered')
  }
  return { formats, defs }
}
