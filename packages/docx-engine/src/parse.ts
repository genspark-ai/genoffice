import JSZip from 'jszip'
import { parseChartPartXml } from './chart'
import { findInkRuns, stripInkRuns } from './ink'
import { computeListMarkers, customEnumItems, type ListItemRef } from './list-markers'
import { isMetafileMime, metafileToDataUrl } from './metafile'
import { isTiffMime, tiffToDataUrl } from './tiff'
import { ommlFragmentsOf, ommlToLatex, ommlToMathML } from './math'
import { splitXmlChildren } from './generate'
import { NOTE_PART_PATH, parseNotesXml } from './notes'
import { scanBody, type BodyElement } from './scan'
import { sectionSettingsFromXml, xmlFlagOn } from './section'
import { findSourcesPart, parseSourcesXml } from './sources'
import { decodeSymbolChar, decodeSymbolText } from './symbol-fonts'
import { THEME_PART_PATH, readThemeColors, readThemeFonts, resolveThemeColor } from './theme'
import { PAGE_MARK, TOTAL_PAGES_MARK } from './types'
import { loadDocxZip } from './zip-load'

import type {
  Block,
  ChartDisplay,
  CommentInfo,
  DiagramDisplay,
  DiagramShape,
  DocDefaults,
  DocProtection,
  FieldDisplay,
  InkInfo,
  NoteInfo,
  HfImage,
  HfParagraph,
  HfTableCell,
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
  SectionSettings,
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
  deepXmlParser,
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
  emz: 'image/x-emz',
  wmz: 'image/x-wmz',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}

const contentTypesCache = new WeakMap<
  JSZip,
  Promise<{ defaults: Map<string, string>; overrides: Map<string, string> }>
>()

function contentTypesOf(zip: JSZip) {
  let cached = contentTypesCache.get(zip)
  if (!cached) {
    cached = (async () => {
      const defaults = new Map<string, string>()
      const overrides = new Map<string, string>()
      const file = zip.file('[Content_Types].xml')
      if (!file) return { defaults, overrides }
      const parsed = xmlParser.parse(await file.async('string')) as XNode[]
      const root = parsed.find((n) => nameOf(n) === 'Types')
      for (const node of root ? findChildren(root, 'Default') : []) {
        const attrs = attrsOf(node)
        const ext = attrs['Extension']?.toLowerCase()
        if (ext && attrs['ContentType']) defaults.set(ext, attrs['ContentType'])
      }
      for (const node of root ? findChildren(root, 'Override') : []) {
        const attrs = attrsOf(node)
        if (attrs['PartName'] && attrs['ContentType'])
          overrides.set(attrs['PartName'], attrs['ContentType'])
      }
      return { defaults, overrides }
    })()
    contentTypesCache.set(zip, cached)
  }
  return cached
}

/**
 * Mime for an image part: extension table first, then [Content_Types].xml
 * Override/Default — covers parts with opaque extensions (media/*.bin).
 */
async function imagePartMime(zip: JSZip, path: string): Promise<string | undefined> {
  const ext = path.split('.').pop()?.toLowerCase() ?? ''
  const fromExt = IMAGE_MIME[ext]
  if (fromExt) return fromExt
  const { defaults, overrides } = await contentTypesOf(zip)
  const fromPart = overrides.get(`/${path}`) ?? defaults.get(ext)
  return fromPart?.startsWith('image/') ? fromPart : undefined
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
  const zip = await loadDocxZip(bytes)
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

  const rangedCommentIds = new Set(
    [...documentXml.matchAll(/<w:commentRangeStart [^>]*w:id="([^"]+)"/g)].map((m) => m[1]),
  )
  const referenceOnlyComments = new Set(
    comments.map((c) => c.id).filter((id) => !rangedCommentIds.has(id)),
  )

  const scan = scanBody(documentXml)
  const mediaByRid = await tableBlipMedia(scan.elements, documentXml, zip, rels)
  // sections end at their w:sectPr: the one governing an offset is the first
  // sectPr at or after it (page/margin-anchored drawing placement)
  const sectSlices = [...documentXml.matchAll(/<w:sectPr[^>]*\/>|<w:sectPr[\s\S]*?<\/w:sectPr>/g)]
  const sectCache = new Map<number, SectionSettings>()
  const sectionAt = (docOffset: number): SectionSettings => {
    let i = sectSlices.findIndex((m) => (m.index ?? 0) + m[0].length > docOffset)
    if (i === -1) i = sectSlices.length - 1
    let settings = sectCache.get(i)
    if (!settings) {
      settings = sectionSettingsFromXml(i >= 0 ? sectSlices[i][0] : '')
      sectCache.set(i, settings)
    }
    return settings
  }
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
    mediaByRid,
    referenceOnlyComments,
    sectionAt,
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
        // real document offsets (not 0): sectionAt resolves page geometry by them
        const block = await buildBlock(
          { name: part.name, start: el.start + part.childStart, end: el.start + part.childEnd },
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
    styles,
  )
  const footer = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'footer',
    'default',
    theme.colors,
    styles,
  )
  const headerFirst = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'header',
    'first',
    theme.colors,
    styles,
  )
  const footerFirst = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'footer',
    'first',
    theme.colors,
    styles,
  )
  const headerEven = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'header',
    'even',
    theme.colors,
    styles,
  )
  const footerEven = await readHeaderFooterPart(
    zip,
    documentXml,
    rels,
    'footer',
    'even',
    theme.colors,
    styles,
  )
  const titlePg = xmlFlagOn(documentXml, 'w:titlePg')
  const evenAndOddHeaders = await parseEvenAndOddHeaders(zip)
  const compatibilityMode = await parseCompatibilityMode(zip)
  const hfParts = await parseAllHfParts(zip, rels, styles, theme.colors)

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
    compatibilityMode,
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
  /** pre-resolved a:blip rId -> data/external URL for pictures inside w:tbl (extractCell is sync, media reads are async) */
  mediaByRid?: Map<string, string>
  /** comment ids anchored only by a bare w:commentReference (no range markers anywhere in document.xml, LibreOffice style) */
  referenceOnlyComments?: Set<string>
  /** page geometry of the section governing a document.xml byte offset (page/margin-anchored drawings) */
  sectionAt?: (docOffset: number) => SectionSettings
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
  const styleNum = styleId ? ctx.styles.get(styleId)?.numPr : undefined
  const styleNumPr = styleNum === 'none' ? undefined : styleNum
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
  // find the matching </w:sdtContent> — last one, so nested content controls
  // don't truncate the region (same boundary rule as splitSdtParts/sdtTableXml)
  const contentClose = sdtXml.lastIndexOf('</w:sdtContent>')
  if (contentClose < contentTagEnd) return null

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
    // nested content controls are transparent containers (Word cover-page
    // building blocks wrap each field in its own sdt): descend instead of
    // treating the inner sdt as one opaque child
    if (m[2] === 'w:sdt' || m[2] === 'w:sdtContent') continue
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

/** Range/marker elements that are invisible in Word when they land at body top level */
const INVISIBLE_BODY_MARKERS = new Set([
  'w:bookmarkStart',
  'w:bookmarkEnd',
  'w:commentRangeStart',
  'w:commentRangeEnd',
  'w:proofErr',
  'w:permStart',
  'w:permEnd',
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
  'w:customXmlInsRangeStart',
  'w:customXmlInsRangeEnd',
  'w:customXmlDelRangeStart',
  'w:customXmlDelRangeEnd',
])

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
      // wrapper offsets keep sectionAt on the right section for the inner block
      const inner = await buildBlock(
        { name: child.name, start: el.start, end: el.end },
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
      // Build a synthetic BodyElement for the inner w:p; the sdt's own document
      // offsets keep sectionAt on the right section
      const syntheticEl: BodyElement = { name: 'w:p', start: el.start, end: el.end }
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
    // No usable paragraph found → passthrough; keep the text visible at least.
    // A content-less sdt (w:sdtPr only, or empty sdtContent) renders as nothing
    // in Word, so it must not produce a visible placeholder chip.
    const sdtPreview = plainText(xml)
    if (!sdtPreview.trim() && !xml.includes('<w:drawing') && !xml.includes('<w:pict')) {
      return { ...base, type: 'passthrough', label: 'Content control', invisibleMarker: true }
    }
    return { ...base, type: 'passthrough', label: 'Content control', previewText: sdtPreview }
  }
  if (INVISIBLE_BODY_MARKERS.has(el.name)) {
    return { ...base, type: 'passthrough', label: el.name, invisibleMarker: true }
  }
  if (el.name === 'w:br') {
    // Word honors a <w:br> sitting directly in the body: page-type turns the page
    if (/w:type="page"/.test(xml)) {
      return {
        ...base,
        type: 'passthrough',
        label: 'Page break',
        fieldDisplay: { kind: 'pageBreak' },
      }
    }
    return { ...base, type: 'passthrough', label: el.name, invisibleMarker: true }
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
  // Field chars inside textbox content don't make the paragraph itself a field
  // paragraph: research-report sidebars are VML textboxes embedding PAGE/date
  // fields, and those paragraphs must reach the textbox display path below
  //. Textbox blocks are protected passthrough anyway.
  const fieldDetect = detect.includes('<w:txbxContent') ? stripTextboxes(detect) : detect
  const hasFields =
    fieldDetect.includes('<w:fldChar') ||
    fieldDetect.includes('<w:fldSimple') ||
    fieldDetect.includes('<w:instrText')
  // Field paragraphs whose visible result is a picture (e.g. INCLUDEPICTURE)
  // should still display the image; the block stays protected either way.
  // Only when the picture is the whole visible content — a paragraph that also
  // carries text falls through to the field passthrough, whose preview keeps
  // the text instead of silently dropping it behind an image block.
  // Non-field drawing paragraphs take the drawing branch below, which keeps
  // mixed text + inline-image paragraphs editable.
  if (
    hasFields &&
    detect.includes('<w:drawing') &&
    !detect.includes('<c:chart') &&
    !detect.includes('r:dm=') &&
    !detect.includes('<dgm:') &&
    plainText(stripTextboxes(detect)).trim() === ''
  ) {
    const image = await extractImage(detect, ctx)
    if (image) {
      return { ...base, type: 'image', label: 'Image', imageDataUrl: image, ...imageMeta(detect) }
    }
  }
  if (hasFields) {
    // Legacy field-form OLE ({ EMBED ... } / { LINK ... } around a w:object):
    // take the OLE display path so the packaged preview picture and its
    // declared size survive instead of a bare "Field (EMBED)" chip.
    if (detect.includes('<w:object') && onlyOleFields(fieldDetect)) {
      return {
        ...base,
        type: 'passthrough',
        label: 'Embedded object',
        previewText: plainText(detect),
        ...(await oleDisplay(detect, ctx)),
      }
    }
    // XE (index entry) fields are invisible markers; a paragraph whose only
    // fields are XE stays editable (extractRuns round-trips the markers).
    if (!onlyXeFields(detect)) {
      const pStyle = /<w:pStyle w:val="([^"]+)"/.exec(xml)?.[1]
      return {
        ...base,
        type: 'passthrough',
        label: fieldLabel(xml),
        previewText: plainText(xml),
        fieldDisplay: fieldDisplayOf(xml),
        ...(pStyle ? { styleId: pStyle } : {}),
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
      styleId: /<w:pStyle w:val="([^"]+)"/.exec(xml)![1],
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
    // Legacy VML textboxes (v:shape/v:textbox) and WordArt (v:textpath) in
    // w:pict: extract the structured display model instead of flattening every
    // nested paragraph and table into one unreadable run — or degrading the
    // whole paragraph to an opaque chip. w:object (OLE) keeps the plain preview.
    if (
      !detect.includes('<w:object') &&
      (detect.includes('<w:txbxContent') || VML_WORDART_RE.test(detect))
    ) {
      const textboxes = extractTextboxes(detect, ctx)
      const strayText = plainText(stripTextboxes(detect)).trim()
      // paragraphs mixing an inline VML picture with real text stay on the
      // editable run-image path below; boxes would drop the picture
      const keepForImages = strayText !== '' && detect.includes('<v:imagedata')
      if (textboxes.length > 0 && !keepForImages) {
        // text the paragraph carries next to the shape (canvas + hyperlinks):
        // shown as a display-only line so it stays visible
        if (strayText !== '') {
          const strayBox = paragraphStrayBox(detect, ctx)
          if (strayBox) textboxes.push(strayBox)
        }
        // host paragraph jc only — a jc inside nested txbxContent must not
        // decide the outer block alignment
        const jc = /<w:jc w:val="([^"]+)"/.exec(stripTextboxes(detect))?.[1]
        return {
          ...base,
          type: 'passthrough',
          label: 'Text box',
          previewText: textboxes
            .flatMap((t) => t.paras.map((p) => p.runs.map((r) => r.text).join('')))
            .join('\n'),
          textboxes,
          ...(hostPageBreak(detect) ? { fieldDisplay: { kind: 'pageBreak' as const } } : {}),
          ...(jc === 'center'
            ? { imageAlign: 'center' as const }
            : jc === 'right' || jc === 'end'
              ? { imageAlign: 'right' as const }
              : {}),
        }
      }
    }
    // Plain VML picture (v:imagedata without OLE): stamps, watermark pictures,
    // Word-2003-era inline images. Render as a real image instead of an opaque
    // "Embedded object" chip (which loses both the picture and any run text).
    if (!detect.includes('<w:object') && detect.includes('<v:imagedata')) {
      if (plainText(stripTextboxes(detect)).trim() !== '') {
        // picture shares the paragraph with real text: keep the text editable
        // with run-level images (the pict fragment round-trips verbatim)
        await resolveBlipMedia(detect, ctx)
        return buildTextParagraph(base, xml, ctx, true)
      }
      const rId = /<v:imagedata[^>]*r:id="([^"]+)"/.exec(detect)?.[1]
      const image = rId ? await mediaDataUrl(ctx.zip, ctx.rels, rId) : null
      if (image) {
        return {
          ...base,
          type: 'image',
          label: 'Image',
          imageDataUrl: image,
          ...vmlImageMeta(detect),
        }
      }
    }
    if (
      !detect.includes('<w:object') &&
      plainText(detect).trim() === '' &&
      isInvisibleVmlPict(detect)
    ) {
      return { ...base, type: 'passthrough', label: 'Drawing object', invisibleMarker: true }
    }
    // VML horizontal rule (<hr> import: <v:rect o:hr="t">): Word draws a thin
    // line at the declared height; a full-width chip both misrenders it and
    // eats ~a line of page budget per rule
    const hrRect = !detect.includes('<w:object')
      ? /<v:rect\b[^>]*\bo:hr="t"[^>]*>/.exec(detect)?.[0]
      : undefined
    if (hrRect && plainText(detect).trim() === '') {
      const fill = /fillcolor="#?([0-9A-Fa-f]{6})"/.exec(hrRect)?.[1]
      const hPt = parseFloat(/height:([\d.]+)pt/.exec(hrRect)?.[1] ?? '')
      return {
        ...base,
        type: 'passthrough',
        label: 'Drawing object',
        decorative: true,
        ...(fill ? { ruleColorHex: fill.toUpperCase() } : {}),
        // VML HR width:0 means "fill the available width" — ruleWidthPx stays unset
        ...(hPt > 0 ? { ruleThicknessPx: Math.max(1, Math.round((hPt / 72) * 96)) } : {}),
      }
    }
    // OLE embed sharing the paragraph with real text: an "Embedded object"
    // block renders only the preview picture, silently dropping the text.
    // Keep the text editable with the object as a run-level image (the
    // w:object fragment round-trips verbatim). Only when every object's
    // preview resolved — otherwise the passthrough below at least keeps the
    // object bytes safe.
    if (detect.includes('<w:object') && plainText(stripTextboxes(detect)).trim() !== '') {
      await resolveBlipMedia(detect, ctx)
      const objects = detect.match(/<w:object[\s\S]*?<\/w:object>/g) ?? []
      const displayable = objects.every((o) => {
        const rId = /<v:imagedata[^>]*r:id="([^"]+)"/.exec(o)?.[1]
        return rId !== undefined && ctx.mediaByRid?.has(rId)
      })
      if (displayable && objects.length > 0) return buildTextParagraph(base, xml, ctx, true)
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
      // chartex (cx 2014 extension: sunburst/boxWhisker/waterfall…): Word pairs
      // the Choice with a pre-rendered picture fallback — exactly what other
      // renderers show. Prefer it; the data-model degrade only covers
      // fallback-less parts.
      if (isChart && detect.includes('/drawing/2014/chartex"')) {
        const fbImage = await extractImage(xml, ctx)
        if (fbImage) {
          return {
            ...base,
            type: 'image',
            label: 'Image',
            imageDataUrl: fbImage,
            ...imageMeta(xml),
          }
        }
      }
      const chartDisplay = isChart ? await extractChart(detect, ctx) : null
      const diagramText = isChart ? null : await extractDiagramText(detect, ctx)
      const diagramDisplay = isChart ? null : await extractDiagramDrawing(detect, ctx)
      // the diagram may share its paragraph with other anchored drawings
      // (photos, gallery shapes): extract those too, each at its own anchor,
      // instead of silently dropping everything but the diagram
      const frags = topLevelDrawings(detect)
      let siblingBoxes: TextboxDisplay[] = []
      if (!isChart && frags.length > 1) {
        if (diagramDisplay) {
          const dmFrag = frags.find((f) => f.includes('r:dm='))
          const meta = dmFrag ? drawingAnchorMeta(dmFrag) : {}
          if (meta.anchored) {
            diagramDisplay.offsetXEmu = meta.offsetXEmu
            diagramDisplay.offsetYEmu = meta.offsetYEmu
            diagramDisplay.floating = true
          }
        }
        await resolveBlipMedia(detect, ctx)
        siblingBoxes = extractTextboxes(detect, ctx, {
          shapes: true,
          pictures: true,
          section: ctx.sectionAt?.(el.start),
        })
      }
      return {
        ...base,
        type: 'passthrough',
        label: isChart ? 'Chart' : 'SmartArt',
        ...(chartDisplay ? { chartDisplay, previewText: chartDisplay.title ?? '' } : {}),
        ...(diagramText ? { previewText: diagramText } : {}),
        ...(diagramDisplay ? { diagramDisplay } : {}),
        ...(siblingBoxes.length > 0 ? { textboxes: siblingBoxes } : {}),
      }
    }
    // drawing canvas (lockedCanvas): scaled child geometry + raw-size text
    if (detect.includes('<lc:lockedCanvas')) {
      await resolveBlipMedia(detect, ctx)
      const canvas = extractLockedCanvas(detect, ctx)
      if (canvas) {
        const frags = topLevelDrawings(detect)
        const meta = frags.length > 0 ? drawingAnchorMeta(frags[0]) : {}
        if (meta.anchored) {
          // horizontal anchor offset only: LO renders these canvases from the
          // anchor paragraph's top, dropping the vertical offset
          canvas.offsetXEmu = meta.offsetXEmu
          if (meta.noWrap) canvas.floating = true
        }
        return {
          ...base,
          type: 'passthrough',
          label: 'Drawing object',
          diagramDisplay: canvas,
          previewText: canvas.shapes.flatMap((s) => s.texts ?? []).join('\n'),
        }
      }
    }
    const image = await extractImage(detect, ctx)
    if (image) {
      // shapes carrying textbox content alongside a blip (photo-framed quote
      // boxes, image-filled shapes with captions): an image block would drop
      // the text, so fall through to the textbox display below instead
      const boxed = detect.includes('<w:txbxContent')
        ? extractTextboxes(detect, ctx).some((t) =>
            t.paras.some((p) => p.runs.some((r) => r.text.trim() !== '')),
          )
        : false
      if (!boxed) {
        // picture(s) sharing the paragraph with real text — or several inline
        // pictures in one paragraph: an image block keeps only the first blip
        // and drops every text run, so stay an editable paragraph with
        // run-level images. Anchored (floating) pictures keep their wp:anchor
        // geometry on Run.image and float in the editor like Word.
        const multiPic = (detect.match(/<a:blip[ />]/g) ?? []).length > 1
        const hasText = plainText(stripTextboxes(detect)).trim() !== ''
        if (hasText || (multiPic && !detect.includes('<wp:anchor'))) {
          await resolveBlipMedia(detect, ctx)
          return buildTextParagraph(base, xml, ctx, true)
        }
        return { ...base, type: 'image', label: 'Image', imageDataUrl: image, ...imageMeta(detect) }
      }
    }
    // picture fills inside shapes (a:blipFill) resolve from pre-fetched media
    if (detect.includes('<a:blipFill')) await resolveBlipMedia(detect, ctx)
    // shapes on: textless preset shapes (stars, triangles, block arrows,
    // anchored connectors) render instead of degrading to a chip
    const textboxes = extractTextboxes(detect, ctx, {
      shapes: true,
      section: ctx.sectionAt?.(el.start),
    })
    const boxTexts = textboxes.flatMap((t) =>
      t.paras.map((p) => p.runs.map((r) => r.text).join('')),
    )
    const strayText = plainText(stripTextboxes(detect)).trim()
    // Anchored decorative shape (underline rule, background box...) in a
    // paragraph that also carries real text: parse it as a normal paragraph so
    // the text stays readable/editable. The shape lives in runs we do not
    // regenerate — the block still saves byte-identical while untouched, and
    // only loses the decoration if the user actually edits this paragraph.
    // Content-carrying textboxes are the exception: the plain-paragraph path
    // would silently drop their text, so the block stays a textbox passthrough
    // (the stray text joins the preview instead of vanishing).
    if (strayText !== '' && !boxTexts.some((t) => t.trim() !== '')) {
      return buildTextParagraph(base, xml, ctx)
    }
    // Anchored textboxes (code boxes, callout cards): all visible text lives in
    // w:txbxContent. Extract a display-only model so content and box styling
    // render; the block stays protected and saves byte-identical. Text the
    // paragraph itself carries next to the boxes is kept as display-only runs
    // (strayRuns) so it doesn't vanish from the page.
    if (textboxes.length > 0) {
      const stray = strayText !== '' ? strayParaRuns(detect, ctx) : null
      return {
        ...base,
        type: 'passthrough',
        label: 'Text box',
        previewText: (strayText !== '' ? [strayText, ...boxTexts] : boxTexts).join('\n'),
        textboxes,
        ...(hostPageBreak(detect) ? { fieldDisplay: { kind: 'pageBreak' as const } } : {}),
        ...(stray && stray.runs.length > 0
          ? { strayRuns: stray.runs, ...(stray.styleId ? { strayStyleId: stray.styleId } : {}) }
          : {}),
        ...imageMeta(detect),
      }
    }
    // Picture whose media cannot be shown (rel missing / pointing at a
    // non-media part, or metafile conversion failed): empty frame at the
    // declared extent
    // instead of a bare chip. Bytes still pass through untouched.
    if (detect.includes('<a:blip') || detect.includes('<pic:pic')) {
      const docPr = /<wp:docPr [^>]*\/?>/.exec(detect)?.[0] ?? ''
      const alt = /\bdescr="([^"]+)"/.exec(docPr)?.[1] ?? /\bname="([^"]+)"/.exec(docPr)?.[1]
      return {
        ...base,
        type: 'passthrough',
        label: 'Image',
        brokenImage: true,
        ...(alt ? { previewText: decodeEntities(alt) } : {}),
        ...imageMeta(detect),
      }
    }
    if (isInvisibleEmptyShape(detect)) {
      return { ...base, type: 'passthrough', label: 'Drawing object', invisibleMarker: true }
    }
    const decorative = isThinRule(detect)
    return {
      ...base,
      type: 'passthrough',
      label: 'Drawing object',
      decorative,
      ...(decorative ? ruleDisplayOf(detect) : {}),
    }
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

/** No run un-hides itself and nothing anchors here (bookmarks, comments, sectPr,
 *  drawings, numbering): safe to collapse a style-vanished paragraph entirely */
function staysVanished(xml: string): boolean {
  if (/<w:vanish\s[^>]*w:val="(?:0|false|off)"/.test(xml)) return false
  return !/<w:(?:drawing|pict|object|sectPr|bookmarkStart|commentRangeStart|commentRangeEnd|numPr)[\s/>]/.test(
    xml,
  )
}

/** parse a w:p as editable text content (paragraph / heading / listItem) */
function buildTextParagraph(
  base: Pick<Block, 'id' | 'docxIndex' | 'originalXml'>,
  xml: string,
  ctx: BuildContext,
  withImages = false,
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
  // whole paragraph hidden by style-level w:vanish (z-TopofForm/z-BottomofForm HTML
  // form markers): Word shows nothing; keep the original bytes at their body position
  if (styleId && ctx.styles.get(styleId)?.display?.vanish === true && staysVanished(xml)) {
    return { ...base, type: 'passthrough', label: 'Hidden paragraph', invisibleMarker: true }
  }
  let format = pPr ? extractParaFormat(pPr) : undefined
  // style-chain autoSpace off reaches the block: the renderer reads it per paragraph
  if (format?.autoSpace === undefined && styleId) {
    if (ctx.styles.get(styleId)?.display?.autoSpace === false) {
      format = { ...(format ?? {}), autoSpace: false }
    }
  }
  const rawPPr = rawPPrOf(xml)
  // inline math: raw <m:oMath> fragments in document order, aligned with the
  // walk below (strip fallback/textbox copies so indexes match visited nodes)
  const mathXml = stripTextboxes(
    xml.includes('<mc:Fallback') ? xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '') : xml,
  )
  const runs = extractRuns(
    pNode,
    ctx,
    ommlFragmentsOf(mathXml),
    rubyFragmentsOf(mathXml),
    withImages,
  )
  if (runs.length === 0) {
    const emptySz = emptyParaSizeHalfPoints(pNode, pPr)
    if (emptySz) format = { ...(format ?? {}), emptyRunSizeHalfPoints: emptySz }
  }
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

/** stroke color/thickness (a:ln) + extent width of a decorative rule drawing */
function ruleDisplayOf(
  xml: string,
): Pick<Block, 'ruleColorHex' | 'ruleThicknessPx' | 'ruleWidthPx'> {
  const out: Pick<Block, 'ruleColorHex' | 'ruleThicknessPx' | 'ruleWidthPx'> = {}
  const ln = /<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/.exec(xml)?.[0]
  if (ln) {
    const color = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(ln)?.[1]
    if (color) out.ruleColorHex = color.toUpperCase()
    const w = parseInt(/<a:ln\b[^>]*\bw="(\d+)"/.exec(ln)?.[1] ?? '', 10)
    if (Number.isFinite(w) && w > 0) out.ruleThicknessPx = Math.max(1, Math.round(w / EMU_PER_PX))
  }
  const cx = parseInt(/<wp:extent cx="(\d+)"/.exec(xml)?.[1] ?? '', 10)
  if (Number.isFinite(cx) && cx > 0) out.ruleWidthPx = Math.round(cx / EMU_PER_PX)
  return out
}

/**
 * Converter artifact: every shape explicitly declares noFill + noFill outline
 * and carries no picture, no text and no effects — Word renders nothing, so
 * the block renders as nothing too (still passthrough, bytes untouched).
 */
function isInvisibleEmptyShape(xml: string): boolean {
  if (!xml.includes('<wps:wsp') || xml.includes('<a:blip') || plainText(xml).trim() !== '') {
    return false
  }
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml) as XNode[]
  } catch {
    return false
  }
  const shapes: XNode[] = []
  collectNodes(parsed, 'wps:wsp', shapes)
  if (shapes.length === 0) return false
  return shapes.every((shape) => {
    const spPr = findChild(shape, 'wps:spPr')
    if (!spPr || !findChild(spPr, 'a:noFill')) return false
    const ln = findChild(spPr, 'a:ln')
    if (!ln || !findChild(ln, 'a:noFill')) return false
    const effects = findChild(spPr, 'a:effectLst')
    return !effects || childrenOf(effects).length === 0
  })
}

/**
 * Legacy VML pict that draws nothing: only v:shapetype definitions (geometry
 * templates, not drawn), shapes hidden via style visibility:hidden, or white
 * strokeless placeholder rectangles (LO fixtures / converter watermark
 * furniture). Word renders nothing there — an "Embedded object" chip would
 * paint a box on an otherwise blank page.
 */
function isInvisibleVmlPict(xml: string): boolean {
  if (!/<v:(?:shapetype|shape|rect|roundrect|oval|line|polyline)\b/.test(xml)) return false
  if (xml.includes('<v:imagedata') || xml.includes('<w:txbxContent')) return false
  const shapes = xml.match(/<v:(?:shape|rect|roundrect|oval|line|polyline)\b[^>]*>/g) ?? []
  // no drawable instances at all (shapetype-only pict) also renders nothing
  return shapes.every((tag) => {
    const style = /style="([^"]*)"/.exec(tag)?.[1] ?? ''
    if (/visibility:\s*hidden/.test(style)) return true
    const fill = /fillcolor="([^"]+)"/.exec(tag)?.[1]?.trim().toLowerCase()
    const unstroked = /\bstroked="(?:f|false|0)"/.test(tag)
    return unstroked && (fill === 'white' || fill === '#ffffff' || fill === '#fff')
  })
}

/** drop textbox content so "does the paragraph itself have text" checks work */
function stripTextboxes(xml: string): string {
  return xml.includes('<w:txbxContent')
    ? xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '')
    : xml
}

/** page-type w:br carried by the anchor paragraph itself (not inside box content) */
function hostPageBreak(xml: string): boolean {
  return /<w:br\s[^>]*w:type="page"/.test(stripTextboxes(xml))
}

/**
 * Text runs the anchor paragraph carries alongside content textboxes
 * (heading text sharing its paragraph with an anchored sidebar). Display-only:
 * the block still saves byte-identical.
 */
function strayParaRuns(
  paragraphXml: string,
  ctx: BuildContext,
): { runs: Run[]; styleId?: string } | null {
  try {
    const parsed = xmlParser.parse(stripTextboxes(paragraphXml)) as XNode[]
    const pNode = parsed.find((n) => nameOf(n) === 'w:p')
    if (!pNode) return null
    const runs = extractRuns(pNode, ctx).filter((r) => r.text !== '')
    if (runs.length === 0) return null
    const pPr = findChild(pNode, 'w:pPr')
    const styleId = pPr ? attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] : undefined
    return styleId ? { runs, styleId } : { runs }
  } catch {
    return null
  }
}

/** paragraphs (and tables, one display line per row) of a w:txbxContent node */
function txbxContentParas(content: XNode, ctx: BuildContext): TextboxParaDisplay[] {
  const out: TextboxParaDisplay[] = []
  for (const child of childrenOf(content)) {
    const name = nameOf(child)
    if (name === 'w:p') {
      const para: TextboxParaDisplay = { runs: extractRuns(child, ctx) }
      const pPr = findChild(child, 'w:pPr')
      if (pPr) {
        Object.assign(para, extractParaFormat(pPr))
        const styleId = attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val']
        if (styleId) para.styleId = styleId
      }
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
 * Display-only line for text a paragraph carries alongside its VML shapes
 * (canvas + trailing hyperlinks): the shape XML is stripped and the remaining
 * runs rendered as a boxless read-only paragraph, so the text stays visible.
 */
function paragraphStrayBox(pXml: string, ctx: BuildContext): TextboxDisplay | null {
  const stripped = stripTextboxes(pXml).replace(/<w:pict>[\s\S]*?<\/w:pict>/g, '')
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(stripped) as XNode[]
  } catch {
    return null
  }
  const pNodes: XNode[] = []
  collectNodes(parsed, 'w:p', pNodes)
  if (pNodes.length === 0) return null
  const para: TextboxParaDisplay = { runs: extractRuns(pNodes[0], ctx) }
  const pPr = findChild(pNodes[0], 'w:pPr')
  if (pPr) Object.assign(para, extractParaFormat(pPr))
  if (!para.runs.some((r) => r.text.trim() !== '')) return null
  return {
    paras: [para],
    readOnly: true,
    insetTopPx: 0,
    insetRightPx: 0,
    insetBottomPx: 0,
    insetLeftPx: 0,
  }
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

/** HTML color names VML attributes use ("silver", "blue"…) */
const VML_NAMED_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  silver: 'C0C0C0',
  gray: '808080',
  grey: '808080',
  maroon: '800000',
  olive: '808000',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  fuchsia: 'FF00FF',
  lime: '00FF00',
  aqua: '00FFFF',
  cyan: '00FFFF',
  orange: 'FFA500',
}

/** VML color attr ("#dbe5f1", "#aaa", "#dbe5f1 [3204]", "silver") → hex without '#' */
function vmlColorHex(value: string | undefined): string | undefined {
  if (!value) return undefined
  const v = value.trim()
  const m6 = /^#?([0-9a-fA-F]{6})/.exec(v)
  if (m6) return m6[1]
  const m3 = /^#([0-9a-fA-F]{3})(?![0-9a-fA-F])/.exec(v)
  if (m3) {
    return m3[1]
      .split('')
      .map((c) => c + c)
      .join('')
  }
  return VML_NAMED_COLORS[v.split(/[\s[]/, 1)[0]!.toLowerCase()]
}

/** VML WordArt: a shape carrying its text in a v:textpath string attribute */
const VML_WORDART_RE = /<v:textpath[^>]*\bstring="/

/** px per group-coordinate unit for children of a v:group (drawing canvas) */
interface VmlGroupScale {
  sx: number
  sy: number
}

function vmlGroupScale(group: XNode): VmlGroupScale | null {
  const a = attrsOf(group)
  const style = a['style'] ?? ''
  const wPx = vmlStyleDimPx(style, 'width')
  const hPx = vmlStyleDimPx(style, 'height')
  const cs = /^\s*(-?\d+)[,\s]+(-?\d+)/.exec(a['coordsize'] ?? '')
  const cw = cs ? parseInt(cs[1]!, 10) : NaN
  const ch = cs ? parseInt(cs[2]!, 10) : NaN
  if (!wPx || !hPx || !(cw > 0) || !(ch > 0)) return null
  return { sx: wPx / cw, sy: hPx / ch }
}

/** shape dimension → px: explicit units directly, unitless via the group scale */
function vmlShapeDimPx(
  style: string,
  key: 'width' | 'height',
  scale: VmlGroupScale | null,
): number | undefined {
  const m = new RegExp(`(?:^|;)\\s*${key}:([0-9.]+)(pt|px|in|mm|cm)?`).exec(style)
  if (!m) return undefined
  if (m[2] || !scale) return vmlStyleDimPx(style, key)
  const v = parseFloat(m[1]!)
  if (!Number.isFinite(v) || v <= 0) return undefined
  return Math.round(v * (key === 'width' ? scale.sx : scale.sy))
}

/**
 * WordArt degrade: v:textpath shapes (shapetype 136 family) render their
 * string as plain styled text — no path warp / 3D, but the text is visible at
 * roughly the declared size and position instead of an opaque chip.
 */
function vmlWordArtBox(shape: XNode): TextboxDisplay | null {
  const tp = findChild(shape, 'v:textpath')
  if (!tp) return null
  const text = attrsOf(tp)['string']
  if (!text || text.trim() === '') return null
  const shapeAttrs = attrsOf(shape)
  const style = shapeAttrs['style'] ?? ''
  const box: TextboxDisplay = {
    paras: [],
    readOnly: true,
    insetTopPx: 0,
    insetRightPx: 0,
    insetBottomPx: 0,
    insetLeftPx: 0,
  }
  const w = vmlStyleDimPx(style, 'width')
  if (w) box.widthPx = w
  const h = vmlStyleDimPx(style, 'height')
  if (h) box.heightPx = h
  // floating WordArt keeps the flow like other absolute shapes
  if (/position:\s*absolute/.test(style)) {
    box.floating = true
    const marginPx = (key: string): number => {
      const pt = parseFloat(new RegExp(`(?:^|;)\\s*${key}:(-?[\\d.]+)pt`).exec(style)?.[1] ?? '')
      return Number.isFinite(pt) ? (pt / 72) * 96 : 0
    }
    box.offsetXEmu = Math.round(marginPx('margin-left') * EMU_PER_PX)
    box.offsetYEmu = Math.round(marginPx('margin-top') * EMU_PER_PX)
  }
  const tpStyle = attrsOf(tp)['style'] ?? ''
  const family = /font-family:\s*"?([^;"]+)"?/.exec(tpStyle)?.[1]?.trim()
  const sizePt = parseFloat(/font-size:\s*([\d.]+)pt/.exec(tpStyle)?.[1] ?? '')
  // fill becomes the *text* color: fillcolor, else the v:fill color/color2
  const fillNode = findChild(shape, 'v:fill')
  const fillAttrs = fillNode ? attrsOf(fillNode) : {}
  const fill =
    shapeAttrs['filled'] === 'f'
      ? undefined
      : (vmlColorHex(shapeAttrs['fillcolor']) ??
        vmlColorHex(fillAttrs['color']) ??
        vmlColorHex(fillAttrs['color2']))
  if (shapeAttrs['stroked'] !== 'f') {
    const strokeColor = vmlColorHex(shapeAttrs['strokecolor']) ?? '000000'
    const weightPt = parseFloat(
      /^([\d.]+)(?:pt)?$/.exec(shapeAttrs['strokeweight'] ?? '')?.[1] ?? '',
    )
    box.textOutline = {
      colorHex: strokeColor,
      widthPx:
        Number.isFinite(weightPt) && weightPt > 0
          ? Math.round((weightPt / 72) * 96 * 100) / 100
          : 1,
    }
  }
  const heightPt = h ? (h / 96) * 72 : NaN
  const run: Run = { text }
  // fitshape sizes the glyphs to the box; the declared font-size matches it in
  // practice (box height ≈ font-size × line factor), so prefer the declared pt
  let pt = Number.isFinite(sizePt) && sizePt > 0 ? sizePt : heightPt > 0 ? heightPt / 1.4 : NaN
  // fitpath compresses long strings into the box; approximate by shrinking the
  // font until the single line fits the declared width (~0.62 em per glyph)
  const widthPt = w ? (w / 96) * 72 : NaN
  if (Number.isFinite(pt) && widthPt > 0 && text.length > 0) {
    pt = Math.max(6, Math.min(pt, widthPt / (0.62 * text.length)))
  }
  if (Number.isFinite(pt) && pt > 0) run.sizeHalfPoints = Math.round(pt * 2)
  box.nowrap = true
  if (family) run.fontAscii = family
  if (fill) run.color = fill
  if (/font-weight:\s*bold/.test(tpStyle)) run.bold = true
  if (/font-style:\s*italic/.test(tpStyle)) run.italic = true
  box.paras.push({ runs: [run], align: 'center' })
  return box
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
function lineBoxOf(shape: XNode, theme?: ThemeColors | null): TextboxDisplay | null {
  const spPr = findChild(shape, 'wps:spPr')
  if (!spPr) return null
  const prst = attrsOf(findChild(spPr, 'a:prstGeom') ?? {})['prst']
  if (!prst || !LINE_PRSTS.has(prst)) return null
  const box: TextboxDisplay = { paras: [], readOnly: true }
  const ln = findChild(spPr, 'a:ln')
  const border =
    (ln
      ? attrsOf(findChild(findChild(ln, 'a:solidFill') ?? {}, 'a:srgbClr') ?? {})['val']
      : undefined) ??
    // theme-styled connectors (Word gallery): stroke from wps:style a:lnRef
    colorNodeHex(findChild(findChild(shape, 'wps:style') ?? {}, 'a:lnRef'), theme)
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
  const xfrm = findChild(spPr, 'a:xfrm')
  const xfrmAttrs = attrsOf(xfrm ?? {})
  const ext = findChild(xfrm ?? {}, 'a:ext')
  const cx = ext ? parseInt(attrsOf(ext)['cx'] ?? '', 10) : NaN
  const cy = ext ? parseInt(attrsOf(ext)['cy'] ?? '', 10) : NaN
  if (Number.isFinite(cx) && cx > 0) box.widthPx = Math.round(cx / EMU_PER_PX)
  const straight = box.prst === 'line' || box.prst === 'lineArrow' || box.prst === 'lineArrowDouble'
  // flips apply at any height: Word's horizontal lines carry cy="0", and a
  // flipH there still reverses which tip holds the arrow
  if (straight) {
    if (xfrmAttrs['flipH'] === '1' || xfrmAttrs['flipH'] === 'true') box.flipH = true
    if (xfrmAttrs['flipV'] === '1' || xfrmAttrs['flipV'] === 'true') box.flipV = true
  }
  if (Number.isFinite(cy) && cy > 0) {
    box.heightPx = Math.round(cy / EMU_PER_PX)
    box.minHeightPx = box.heightPx
    // a real vertical extent means the connector runs corner to corner
    // (≤12 px stays level: our own inserted lines keep a 12 px grab band)
    if (straight && (box.heightPx > 12 || box.flipH || box.flipV)) box.lineDiag = true
  } else {
    // zero-height extent = Word's horizontal line; keep a 12 px grab band
    box.heightPx = 12
  }
  // a:headEnd decorates the start point: a head-only arrow renders as the
  // reversed segment so the renderer's single arrowhead lands on the right tip
  if (head && !tail && box.prst === 'lineArrow') {
    const fh = !box.flipH
    const fv = !box.flipV
    delete box.flipH
    delete box.flipV
    if (fh) box.flipH = true
    if (fv) box.flipV = true
  }
  box.insetTopPx = 0
  box.insetRightPx = 0
  box.insetBottomPx = 0
  box.insetLeftPx = 0
  return box
}

/** a:schemeClr val -> ThemeColors slot (DrawingML names; text/bg aliases mapped) */
const SCHEME_CLR_SLOTS: Record<string, keyof ThemeColors> = {
  tx1: 'dk1',
  bg1: 'lt1',
  tx2: 'dk2',
  bg2: 'lt2',
  dk1: 'dk1',
  lt1: 'lt1',
  dk2: 'dk2',
  lt2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
}

/** one a:gs stop -> sRGB triple (srgbClr or theme-resolved schemeClr, lumMod/lumOff applied) */
function gradStopRgb(gs: XNode, theme?: ThemeColors | null): number[] | null {
  const srgb = attrsOf(findChild(gs, 'a:srgbClr') ?? {})['val']
  const scheme = srgb ? undefined : findChild(gs, 'a:schemeClr')
  let base: string | undefined = srgb
  if (!base && scheme) {
    const slot = SCHEME_CLR_SLOTS[attrsOf(scheme)['val'] ?? '']
    if (!slot) return null
    base =
      (theme?.[slot] as string | undefined) ??
      (slot === 'dk1' ? '000000' : slot === 'lt1' ? 'FFFFFF' : undefined)
  }
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return null
  let rgb = [0, 2, 4].map((i) => parseInt(base!.slice(i, i + 2), 16))
  if (scheme) {
    const pct = (name: string): number | null => {
      const v = parseInt(attrsOf(findChild(scheme, name) ?? {})['val'] ?? '', 10)
      return Number.isFinite(v) ? Math.min(100000, Math.max(0, v)) / 100000 : null
    }
    const lumMod = pct('a:lumMod')
    if (lumMod !== null) rgb = rgb.map((c) => c * lumMod)
    const lumOff = pct('a:lumOff')
    if (lumOff !== null) rgb = rgb.map((c) => c + 255 * lumOff)
    const shade = pct('a:shade')
    if (shade !== null) rgb = rgb.map((c) => c * shade)
    const tint = pct('a:tint')
    if (tint !== null) rgb = rgb.map((c) => c * tint + 255 * (1 - tint))
  }
  return rgb
}

/** color-bearing node (a:solidFill / a:fillRef / a:lnRef …) → hex without '#' */
function colorNodeHex(node: XNode | undefined, theme?: ThemeColors | null): string | undefined {
  if (!node) return undefined
  const rgb = gradStopRgb(node, theme)
  if (!rgb) return undefined
  return rgb
    .map((c) =>
      Math.round(Math.min(255, Math.max(0, c)))
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('')
}

/** anchored-drawing placement of one top-level w:drawing fragment */
interface DrawingAnchorMeta {
  offsetXEmu?: number
  offsetYEmu?: number
  /** wrapNone (front) / behindDoc anchors leave the text flow entirely */
  noWrap?: boolean
  anchored?: boolean
  /** wp:positionH/V relativeFrom */
  relH?: string
  relV?: string
  /** wp:align inside wp:positionH/V */
  alignH?: string
  alignV?: string
  /** wp14:pctPosH/VOffset in 1/1000 of a percent of the reference frame */
  pctH?: number
  pctV?: number
  /** wp:extent (drawing size) */
  extentXEmu?: number
  extentYEmu?: number
}

/**
 * Top-level `<w:drawing>` fragments of a paragraph, balanced (a textbox's
 * txbxContent may nest further drawings, which stay inside their parent
 * fragment). Used to attach each drawing's own anchor placement to the shapes
 * it carries — a paragraph can anchor several drawings at distinct offsets.
 */
function topLevelDrawings(xml: string): string[] {
  const out: string[] = []
  let i = 0
  for (;;) {
    const start = xml.indexOf('<w:drawing>', i)
    if (start === -1) break
    let depth = 0
    let j = start
    for (;;) {
      const open = xml.indexOf('<w:drawing>', j + 1)
      const close = xml.indexOf('</w:drawing>', j + 1)
      if (close === -1) return out // malformed; bail with what we have
      if (open !== -1 && open < close) {
        depth++
        j = open
      } else if (depth > 0) {
        depth--
        j = close
      } else {
        out.push(xml.slice(start, close + '</w:drawing>'.length))
        i = close + '</w:drawing>'.length
        break
      }
    }
  }
  return out
}

function drawingAnchorMeta(frag: string): DrawingAnchorMeta {
  const anchorTag = /<wp:anchor[^>]*>/.exec(frag)?.[0]
  if (!anchorTag) return {}
  const meta: DrawingAnchorMeta = { anchored: true }
  const posOf = (dir: 'H' | 'V'): number | undefined => {
    const m = new RegExp(`<wp:position${dir}[^>]*>\\s*<wp:posOffset>(-?\\d+)</wp:posOffset>`).exec(
      frag,
    )
    const v = m ? parseInt(m[1], 10) : NaN
    return Number.isFinite(v) ? v : undefined
  }
  meta.offsetXEmu = posOf('H')
  meta.offsetYEmu = posOf('V')
  for (const dir of ['H', 'V'] as const) {
    const m = new RegExp(`<wp:position${dir}\\b([^>]*)>([\\s\\S]*?)</wp:position${dir}>`).exec(frag)
    if (!m) continue
    const rel = /relativeFrom="(\w+)"/.exec(m[1])?.[1]
    const align = /<wp:align>(\w+)<\/wp:align>/.exec(m[2])?.[1]
    const pct = parseInt(
      new RegExp(`<wp14:pctPos${dir}Offset[^>]*>(-?\\d+)<`).exec(m[2])?.[1] ?? '',
      10,
    )
    if (dir === 'H') {
      meta.relH = rel
      meta.alignH = align
      if (Number.isFinite(pct)) meta.pctH = pct
    } else {
      meta.relV = rel
      meta.alignV = align
      if (Number.isFinite(pct)) meta.pctV = pct
    }
  }
  const extent = /<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(frag)
  if (extent) {
    meta.extentXEmu = parseInt(extent[1], 10)
    meta.extentYEmu = parseInt(extent[2], 10)
  }
  if (frag.includes('<wp:wrapNone') || /behindDoc="(?:1|true)"/.test(anchorTag)) meta.noWrap = true
  return meta
}

const EMU_PER_TWIP = 635

/** anchor position resolved to EMU offsets from the paragraph flow origin
 *  (column left / body top), page coordinates recoverable via the margins */
interface ResolvedAnchorPos {
  xEmu: number
  yEmu?: number
  /** the drawing lies horizontally outside the body column (sidebar layout) */
  outsideColumn: boolean
}

/**
 * Page/margin-anchored placement (wp14:pctPosH/VOffset or wp:align) resolved
 * against the section's page geometry. Plain posOffset anchors keep the
 * legacy paragraph-relative path untouched — only the features the parser
 * previously ignored resolve here, so in-column wrapSquare boxes are not
 * repositioned. The vertical origin approximates the anchor paragraph sitting
 * at the top of the body (where Word puts these full-height sidebar groups).
 */
function resolveAnchorPagePos(
  meta: DrawingAnchorMeta,
  sect: SectionSettings | undefined,
): ResolvedAnchorPos | null {
  if (!sect) return null
  if (meta.relH !== 'page' && meta.relH !== 'margin') return null
  if (meta.pctH === undefined && meta.alignH === undefined) return null
  const pageW = sect.pageWidth * EMU_PER_TWIP
  const pageH = sect.pageHeight * EMU_PER_TWIP
  const marL = sect.marginLeft * EMU_PER_TWIP
  const marR = sect.marginRight * EMU_PER_TWIP
  const marT = sect.marginTop * EMU_PER_TWIP
  const w = meta.extentXEmu ?? 0
  const refW = meta.relH === 'page' ? pageW : pageW - marL - marR
  const relX =
    meta.pctH !== undefined
      ? Math.round((refW * meta.pctH) / 100000)
      : meta.alignH === 'center'
        ? Math.round((refW - w) / 2)
        : meta.alignH === 'right' || meta.alignH === 'outside'
          ? refW - w
          : 0
  const pageX = meta.relH === 'page' ? relX : marL + relX
  const pos: ResolvedAnchorPos = {
    xEmu: pageX - marL,
    outsideColumn: pageX + w <= marL || pageX >= pageW - marR,
  }
  if (
    (meta.relV === 'page' || meta.relV === 'margin') &&
    (meta.pctV !== undefined || meta.alignV !== undefined)
  ) {
    const marB = sect.marginBottom * EMU_PER_TWIP
    const h = meta.extentYEmu ?? 0
    const refH = meta.relV === 'page' ? pageH : pageH - marT - marB
    const relY =
      meta.pctV !== undefined
        ? Math.round((refH * meta.pctV) / 100000)
        : meta.alignV === 'center'
          ? Math.round((refH - h) / 2)
          : meta.alignV === 'bottom' || meta.alignV === 'outside'
            ? refH - h
            : 0
    pos.yEmu = (meta.relV === 'page' ? relY : marT + relY) - marT
  }
  return pos
}

/**
 * a:gradFill approximated as a solid color (display only): equal-weight sRGB
 * average of all stops — the first stop alone is often white and loses the
 * visible tint entirely.
 */
function gradFillApproxHex(spPr: XNode, theme?: ThemeColors | null): string | undefined {
  const gsLst = findChild(findChild(spPr, 'a:gradFill') ?? {}, 'a:gsLst')
  if (!gsLst) return undefined
  const stops = childrenOf(gsLst)
    .filter((n) => nameOf(n) === 'a:gs')
    .map((gs) => gradStopRgb(gs, theme))
    .filter((rgb): rgb is number[] => rgb !== null)
  if (stops.length === 0) return undefined
  return [0, 1, 2]
    .map((i) => stops.reduce((sum, rgb) => sum + rgb[i], 0) / stops.length)
    .map((c) =>
      Math.round(Math.min(255, Math.max(0, c)))
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('')
}

/** sRGB triple → hex without '#' */
function rgbHex(rgb: number[]): string {
  return rgb
    .map((c) =>
      Math.round(Math.min(255, Math.max(0, c)))
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('')
}

/** w14:srgbClr / w14:schemeClr child → sRGB triple with tint/shade/lum/satMod applied */
function w14ColorRgb(node: XNode, theme?: ThemeColors | null): number[] | null {
  const colorNode = findChild(node, 'w14:srgbClr') ?? findChild(node, 'w14:schemeClr')
  if (!colorNode) return null
  const isScheme = nameOf(colorNode) === 'w14:schemeClr'
  let base: string | undefined = attrsOf(colorNode)['w14:val']
  if (isScheme) {
    const slot = SCHEME_CLR_SLOTS[base ?? '']
    if (!slot) return null
    base =
      (theme?.[slot] as string | undefined) ??
      (slot === 'dk1' ? '000000' : slot === 'lt1' ? 'FFFFFF' : undefined)
  }
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return null
  let rgb = [0, 2, 4].map((i) => parseInt(base!.slice(i, i + 2), 16))
  const pct = (name: string): number | null => {
    const v = parseInt(attrsOf(findChild(colorNode, name) ?? {})['w14:val'] ?? '', 10)
    return Number.isFinite(v) && v >= 0 ? v / 100000 : null
  }
  const lumMod = pct('w14:lumMod')
  if (lumMod !== null) rgb = rgb.map((c) => c * lumMod)
  const lumOff = pct('w14:lumOff')
  if (lumOff !== null) rgb = rgb.map((c) => c + 255 * lumOff)
  const shade = pct('w14:shade')
  if (shade !== null) rgb = rgb.map((c) => c * shade)
  const tint = pct('w14:tint')
  if (tint !== null) rgb = rgb.map((c) => c * tint + 255 * (1 - tint))
  const satMod = pct('w14:satMod')
  if (satMod !== null && satMod !== 1) rgb = saturationModulate(rgb, satMod)
  return rgb
}

/** HSL saturation modulation (a:satMod / w14:satMod semantics, clamped) */
function saturationModulate(rgb: number[], mod: number): number[] {
  const [r, g, b] = rgb.map((c) => Math.min(255, Math.max(0, c)) / 255) as [number, number, number]
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const d = max - min
  if (d === 0) return rgb
  let s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h =
    max === r
      ? ((g - b) / d + (g < b ? 6 : 0)) / 6
      : max === g
        ? ((b - r) / d + 2) / 6
        : ((r - g) / d + 4) / 6
  s = Math.min(1, s * mod)
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const hue = (t: number): number => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)].map((c) => c * 255)
}

/**
 * WordArt-styled run text (w14:textFill): solid fill directly, gradient fill
 * as the equal-weight average of its stops — display approximation only.
 */
function w14TextFillHex(rPr: XNode, theme?: ThemeColors | null): string | undefined {
  const tf = findChild(rPr, 'w14:textFill')
  if (!tf) return undefined
  const solid = findChild(tf, 'w14:solidFill')
  if (solid) {
    const rgb = w14ColorRgb(solid, theme)
    return rgb ? rgbHex(rgb) : undefined
  }
  const gsLst = findChild(findChild(tf, 'w14:gradFill') ?? {}, 'w14:gsLst')
  if (!gsLst) return undefined
  const stops = childrenOf(gsLst)
    .filter((n) => nameOf(n) === 'w14:gs')
    .map((gs) => w14ColorRgb(gs, theme))
    .filter((rgb): rgb is number[] => rgb !== null)
  if (stops.length === 0) return undefined
  return rgbHex([0, 1, 2].map((i) => stops.reduce((sum, rgb) => sum + rgb[i], 0) / stops.length))
}

interface ExtractTextboxOpts {
  /** include textless preset shapes (stars, block arrows…) as display boxes */
  shapes?: boolean
  /** include picture-only drawings (pic:pic without a shape) as picture boxes */
  pictures?: boolean
  /** page geometry of the governing section (page/margin-anchored placement) */
  section?: SectionSettings
}

/** child-EMU → anchor-EMU affine transform of a wpg group (X = tx + x·sx) */
interface GroupCtm {
  sx: number
  sy: number
  tx: number
  ty: number
}

const IDENTITY_CTM: GroupCtm = { sx: 1, sy: 1, tx: 0, ty: 0 }

/** wpg:wgp / wpg:grpSp xfrm (off/ext vs chOff/chExt) composed onto `outer` */
function composeGroupCtm(group: XNode, outer: GroupCtm): GroupCtm | null {
  const xfrm = findChild(findChild(group, 'wpg:grpSpPr') ?? {}, 'a:xfrm')
  if (!xfrm) return null
  const num = (node: XNode | undefined, key: string, dflt: number): number => {
    const v = parseInt(attrsOf(node ?? {})[key] ?? '', 10)
    return Number.isFinite(v) ? v : dflt
  }
  const ox = num(findChild(xfrm, 'a:off'), 'x', 0)
  const oy = num(findChild(xfrm, 'a:off'), 'y', 0)
  const ex = num(findChild(xfrm, 'a:ext'), 'cx', 0)
  const ey = num(findChild(xfrm, 'a:ext'), 'cy', 0)
  const chOffX = num(findChild(xfrm, 'a:chOff'), 'x', 0)
  const chOffY = num(findChild(xfrm, 'a:chOff'), 'y', 0)
  const chExtX = num(findChild(xfrm, 'a:chExt'), 'cx', 0)
  const chExtY = num(findChild(xfrm, 'a:chExt'), 'cy', 0)
  const sx = ex > 0 && chExtX > 0 ? ex / chExtX : 1
  const sy = ey > 0 && chExtY > 0 ? ey / chExtY : 1
  return {
    sx: outer.sx * sx,
    sy: outer.sy * sy,
    tx: outer.tx + outer.sx * (ox - chOffX * sx),
    ty: outer.ty + outer.sy * (oy - chOffY * sy),
  }
}

function extractTextboxes(
  xml: string,
  ctx: BuildContext,
  opts?: ExtractTextboxOpts,
): TextboxDisplay[] {
  // wrapSquare gate keeps converter-emitted decorative rules on the thin-rule path
  const hasLineShapes = xml.includes('<wp:wrapSquare') && LINE_PRSTS_RE.test(xml)
  const hasCanvasText = xml.includes('<a:txSp')
  const hasVmlWordArt = VML_WORDART_RE.test(xml)
  if (
    !xml.includes('<w:txbxContent') &&
    !hasLineShapes &&
    !hasCanvasText &&
    !hasVmlWordArt &&
    !(opts?.shapes || opts?.pictures)
  ) {
    return []
  }

  const frags = topLevelDrawings(xml)
  // a paragraph anchoring several drawings never stacks them: every anchored
  // shape floats at its own offset (single wrapSquare boxes keep the flow
  // placement the editor has always used)
  const multiDrawing = frags.length > 1

  const out: TextboxDisplay[] = []

  /** one wps:wsp (already parsed) → display box; null when it renders nothing */
  const buildWpsBox = (shape: XNode): TextboxDisplay | null => {
    const contents: XNode[] = []
    collectNodes(childrenOf(shape), 'w:txbxContent', contents)
    const spPr = findChild(shape, 'wps:spPr')
    const prstOf = spPr ? attrsOf(findChild(spPr, 'a:prstGeom') ?? {})['prst'] : undefined
    if (contents.length === 0) {
      if (prstOf && LINE_PRSTS.has(prstOf)) {
        if (hasLineShapes) return lineBoxOf(shape, ctx.themeColors)
        if (!opts?.shapes) return null
        // anchored gallery connector: a real vertical extent, flips, or arrow
        // ends mark a drawn connector; plain near-flat lines stay on the
        // decorative thin-rule path
        const xfrm = findChild(spPr ?? {}, 'a:xfrm')
        const cy = parseInt(attrsOf(findChild(xfrm ?? {}, 'a:ext') ?? {})['cy'] ?? '', 10)
        const xa = attrsOf(xfrm ?? {})
        const flipped = ['flipH', 'flipV'].some((k) => xa[k] === '1' || xa[k] === 'true')
        const ln = findChild(spPr ?? {}, 'a:ln')
        const arrowed = ['a:headEnd', 'a:tailEnd'].some((name) => {
          const type = attrsOf(findChild(ln ?? {}, name) ?? {})['type']
          return !!type && type !== 'none'
        })
        return (Number.isFinite(cy) && cy > 130000) || flipped || arrowed
          ? lineBoxOf(shape, ctx.themeColors)
          : null
      }
      if (hasLineShapes && !opts?.shapes) return null
      // textless preset shape (star/arrow/…): visible geometry, no text body.
      // rect stays on its existing paths (thin rules, background boxes)
      if (!opts?.shapes || !prstOf || prstOf === 'rect') return null
    }
    const box: TextboxDisplay = { paras: [] }
    if (spPr) {
      if (!findChild(spPr, 'a:noFill')) {
        const fill =
          colorNodeHex(findChild(spPr, 'a:solidFill'), ctx.themeColors) ??
          gradFillApproxHex(spPr, ctx.themeColors)
        if (fill) box.fill = fill
        const blipFill = findChild(spPr, 'a:blipFill')
        if (blipFill) {
          const rId = attrsOf(findChild(blipFill, 'a:blip') ?? {})['r:embed']
          const dataUrl = rId ? ctx.mediaByRid?.get(rId) : undefined
          if (dataUrl) {
            box.fillImageDataUrl = dataUrl
            if (findChild(blipFill, 'a:tile')) box.fillTile = true
          }
        }
      }
      const ln = findChild(spPr, 'a:ln')
      if (ln && !findChild(ln, 'a:noFill')) {
        const border = colorNodeHex(findChild(ln, 'a:solidFill'), ctx.themeColors)
        if (border) box.borderColor = border
        const w = parseInt(attrsOf(ln)['w'] ?? '', 10)
        if (Number.isFinite(w) && w > 0) {
          box.borderWidthPx = Math.round((w / EMU_PER_PX) * 100) / 100
        }
        const dash = attrsOf(findChild(ln, 'a:prstDash') ?? {})['val']
        if (dash) box.borderDash = /dot/i.test(dash) ? 'dotted' : 'dashed'
      }
      // shape-style references (wps:style): theme fill/line for shapes whose
      // spPr declares no explicit color (Word gallery shapes)
      const styleNode = findChild(shape, 'wps:style')
      if (styleNode) {
        if (!box.fill && !box.fillImageDataUrl && !findChild(spPr, 'a:noFill')) {
          const ref = findChild(styleNode, 'a:fillRef')
          if (parseInt(attrsOf(ref ?? {})['idx'] ?? '0', 10) > 0) {
            const fill = colorNodeHex(ref, ctx.themeColors)
            if (fill) box.fill = fill
          }
        }
        if (!box.borderColor && !(ln && findChild(ln, 'a:noFill'))) {
          const ref = findChild(styleNode, 'a:lnRef')
          if (parseInt(attrsOf(ref ?? {})['idx'] ?? '0', 10) > 0) {
            const border = colorNodeHex(ref, ctx.themeColors)
            if (border) box.borderColor = border
          }
        }
      }
      const prst = prstOf
      if (prst && prst !== 'rect') box.prst = prst
      const xfrm = findChild(spPr, 'a:xfrm')
      const rot = parseInt(attrsOf(xfrm ?? {})['rot'] ?? '', 10)
      if (Number.isFinite(rot) && rot !== 0) box.rotDeg = Math.round(rot / 60000)
      const ext = findChild(xfrm ?? {}, 'a:ext')
      const cx = ext ? parseInt(attrsOf(ext)['cx'] ?? '', 10) : NaN
      if (Number.isFinite(cx) && cx > 0) box.widthPx = Math.round(cx / EMU_PER_PX)
      // Word clips overflowing text unless the shape auto-fits its content;
      // carrying the fixed height keeps tall sparse boxes from exploding layout
      const bodyPrNode = findChild(shape, 'wps:bodyPr')
      const autoFit = bodyPrNode ? !!findChild(bodyPrNode, 'a:spAutoFit') : false
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
      if (attrs['anchor'] === 'b') box.vAlign = 'bottom'
      else if (attrs['anchor'] === 'ctr') box.vAlign = 'center'
    }
    for (const content of contents) box.paras.push(...txbxContentParas(content, ctx))
    if (contents.some(txbxHasStructuredContent)) box.readOnly = true
    if (contents.length === 0) {
      // textless preset shape: keep it if the geometry has any visible ink
      if (!box.fill && !box.borderColor && !box.fillImageDataUrl) return null
      box.readOnly = true
      return box
    }
    return box.paras.some((p) => p.runs.length > 0) ? box : null
  }

  const applyAnchor = (
    box: TextboxDisplay,
    meta: DrawingAnchorMeta,
    pagePos: ResolvedAnchorPos | null,
  ): void => {
    if (!meta.anchored) return
    // page/margin-anchored drawing sitting outside the body column (Word
    // resume sidebars): absolute placement at the resolved position instead
    // of stacking in the flow, wrap kind notwithstanding
    if (pagePos?.outsideColumn) {
      // group children already carry a group-relative offset: the anchor adds
      box.offsetXEmu = (box.offsetXEmu ?? 0) + pagePos.xEmu
      box.offsetYEmu = (box.offsetYEmu ?? 0) + (pagePos.yEmu ?? meta.offsetYEmu ?? 0)
      box.floating = true
      return
    }
    if (meta.offsetXEmu !== undefined) box.offsetXEmu = (box.offsetXEmu ?? 0) + meta.offsetXEmu
    if (meta.offsetYEmu !== undefined) box.offsetYEmu = (box.offsetYEmu ?? 0) + meta.offsetYEmu
    if (meta.noWrap || multiDrawing) box.floating = true
  }

  for (const frag of frags) {
    let parsedFrag: XNode[]
    try {
      parsedFrag = xmlParser.parse(frag) as XNode[]
    } catch {
      continue
    }
    const meta = drawingAnchorMeta(frag)
    const pagePos = resolveAnchorPagePos(meta, opts?.section)
    let wspCount = 0
    const pushShape = (shape: XNode, ctm: GroupCtm | null): void => {
      wspCount++
      const box = buildWpsBox(shape)
      if (!box) return
      if (ctm) {
        const xfrm = findChild(findChild(shape, 'wps:spPr') ?? {}, 'a:xfrm')
        const off = attrsOf(findChild(xfrm ?? {}, 'a:off') ?? {})
        const x = parseInt(off['x'] ?? '', 10)
        const y = parseInt(off['y'] ?? '', 10)
        if (Number.isFinite(x)) box.offsetXEmu = Math.round(ctm.tx + x * ctm.sx)
        if (Number.isFinite(y)) box.offsetYEmu = Math.round(ctm.ty + y * ctm.sy)
        if (ctm.sx !== 1 && box.widthPx) box.widthPx = Math.round(box.widthPx * ctm.sx)
        if (ctm.sy !== 1) {
          if (box.heightPx) box.heightPx = Math.round(box.heightPx * ctm.sy)
          if (box.minHeightPx) box.minHeightPx = Math.round(box.minHeightPx * ctm.sy)
        }
        // grouped shapes place absolutely at their mapped offset like Word
        box.floating = true
      }
      applyAnchor(box, meta, pagePos)
      out.push(box)
    }
    // document-order walk (box order must match w:txbxContent order for the
    // patch-save mapping), carrying the enclosing wpg group transform
    const walkShapes = (nodes: XNode[], ctm: GroupCtm | null): void => {
      for (const node of nodes) {
        const name = nameOf(node)
        if (name === 'wps:wsp') pushShape(node, ctm)
        if (name === 'wpg:wgp' || name === 'wpg:grpSp') {
          walkShapes(childrenOf(node), composeGroupCtm(node, ctm ?? IDENTITY_CTM) ?? ctm)
        } else {
          walkShapes(childrenOf(node), ctm)
        }
      }
    }
    walkShapes(parsedFrag, null)
    // picture-only drawing sharing a multi-drawing paragraph: a photo box
    if (opts?.pictures && wspCount === 0 && frag.includes('<pic:pic')) {
      const rId =
        /<a:blip[^>]*r:embed="([^"]+)"/.exec(frag)?.[1] ??
        /<a:blip[^>]*r:link="([^"]+)"/.exec(frag)?.[1]
      const dataUrl = rId ? ctx.mediaByRid?.get(rId) : undefined
      const extent = /<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(frag)
      if (dataUrl && extent) {
        const box: TextboxDisplay = {
          paras: [],
          readOnly: true,
          fillImageDataUrl: dataUrl,
          widthPx: Math.round(parseInt(extent[1], 10) / EMU_PER_PX),
          heightPx: Math.round(parseInt(extent[2], 10) / EMU_PER_PX),
          insetTopPx: 0,
          insetRightPx: 0,
          insetBottomPx: 0,
          insetLeftPx: 0,
        }
        applyAnchor(box, meta, pagePos)
        out.push(box)
      }
    }
  }

  // legacy VML shapes (w:pict, outside w:drawing fragments)
  let parsed: XNode[] | null = null
  if (/<v:(?:shape|rect|roundrect)\b/.test(xml)) {
    try {
      parsed = xmlParser.parse(xml) as XNode[]
    } catch {
      parsed = null
    }
  }
  if (parsed) {
    // Children of a v:group (drawing canvas) position/size in the group's own
    // coordinate space (unitless style values); the scale maps them to px.
    const vmlBox = (shape: XNode, scale: VmlGroupScale | null): void => {
      const shapeAttrs = attrsOf(shape)
      const style = shapeAttrs['style'] ?? ''
      const contents: XNode[] = []
      collectNodes(childrenOf(shape), 'w:txbxContent', contents)
      if (contents.length === 0) {
        const wordArt = vmlWordArtBox(shape)
        if (wordArt) out.push(wordArt)
        return
      }
      const box: TextboxDisplay = { paras: [] }
      // VML shape: geometry from the style attribute, colors from fillcolor/strokecolor
      const w = vmlShapeDimPx(style, 'width', scale)
      if (w) box.widthPx = w
      const h = vmlShapeDimPx(style, 'height', scale)
      if (h) {
        box.heightPx = h
        box.minHeightPx = h
      }
      const fill = vmlColorHex(shapeAttrs['fillcolor'])
      if (fill && shapeAttrs['filled'] !== 'f') box.fill = fill
      const stroke = vmlColorHex(shapeAttrs['strokecolor'])
      if (stroke && shapeAttrs['stroked'] !== 'f') box.borderColor = stroke
      // canvas textboxes: VML strokes default to on/black, and Word draws them
      else if (!stroke && shapeAttrs['stroked'] !== 'f' && scale) box.borderColor = '000000'
      for (const content of contents) box.paras.push(...txbxContentParas(content, ctx))
      if (contents.some(txbxHasStructuredContent)) box.readOnly = true
      if (box.paras.some((p) => p.runs.length > 0)) out.push(box)
    }
    const walkVml = (nodes: XNode[], scale: VmlGroupScale | null): void => {
      for (const node of nodes) {
        const name = nameOf(node)
        if (name === 'v:group') {
          walkVml(childrenOf(node), vmlGroupScale(node) ?? scale)
          continue
        }
        if (name === 'v:shape' || name === 'v:rect' || name === 'v:roundrect') {
          vmlBox(node, scale)
        }
        walkVml(childrenOf(node), scale)
      }
    }
    walkVml(parsed, null)
  }

  // lockedCanvas text shapes (a:txSp): DrawingML a:p/a:r/a:t text bodies inside
  // a drawing canvas (logo lockups etc.) that neither the wps nor the VML path
  // sees. Display-only, so all characters at least remain visible.
  if (hasCanvasText) {
    let parsedAll: XNode[]
    try {
      parsedAll = xmlParser.parse(xml) as XNode[]
    } catch {
      return out
    }
    const txSps: XNode[] = []
    collectNodes(parsedAll, 'a:txSp', txSps)
    for (const sp of txSps) {
      const body = findChild(sp, 'a:txBody')
      if (!body) continue
      const paras: TextboxParaDisplay[] = []
      for (const p of childrenOf(body)) {
        if (nameOf(p) !== 'a:p') continue
        const runs: Run[] = []
        for (const r of childrenOf(p)) {
          if (nameOf(r) !== 'a:r') continue
          const t = findChild(r, 'a:t')
          const text = t ? decodeNumericCharRefs(textOf(t)) : ''
          if (text === '') continue
          const run: Run = { text }
          const rPr = findChild(r, 'a:rPr')
          if (rPr) {
            const a = attrsOf(rPr)
            const sz = parseInt(a['sz'] ?? '', 10)
            // a:sz is in hundredths of a point
            if (Number.isFinite(sz) && sz > 0) run.sizeHalfPoints = Math.round(sz / 50)
            if (a['b'] === '1') run.bold = true
            const latin = attrsOf(findChild(rPr, 'a:latin') ?? {})['typeface']
            if (latin && !latin.startsWith('+')) run.font = latin
          }
          runs.push(run)
        }
        if (runs.length > 0) {
          const algn = attrsOf(findChild(p, 'a:pPr') ?? {})['algn']
          paras.push({
            runs,
            ...(algn === 'ctr' ? { align: 'center' as const } : {}),
          })
        }
      }
      if (paras.length > 0) out.push({ paras, readOnly: true })
    }
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

/** w:autoSpaceDE/DN (Word default on): false only when both are explicitly off */
function autoSpaceOf(pPr: XNode): boolean | undefined {
  const de = onOffOf(pPr, 'w:autoSpaceDE')
  const dn = onOffOf(pPr, 'w:autoSpaceDN')
  if (de === false && dn === false) return false
  if (de === true || dn === true) return true
  return undefined
}

/** pPr <w:tabs> → TabStop[] (shared by direct paragraph format and style display) */
function tabStopsOf(pPr: XNode): import('./types').TabStop[] | undefined {
  const tabsEl = findChild(pPr, 'w:tabs')
  if (!tabsEl) return undefined
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
  return stops.length > 0 ? stops : undefined
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
    } else if (line === 0 && rule === 'atLeast') {
      // w:line="0" atLeast: layout no-op (natural line height) BUT it opts the
      // paragraph out of docGrid snapping (LO probe); keep the rule so the
      // renderer can tell it apart from an undeclared (snapping) paragraph
      format.lineRule = 'atLeast'
      format.lineRawTwips = 0
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
  // snapToGrid: default ON; only store when explicitly set to OFF
  const snapEl = findChild(pPr, 'w:snapToGrid')
  if (snapEl) {
    const v = attrsOf(snapEl)['w:val']
    if (v === '0' || v === 'false') format.snapToGrid = false
  }
  // widowControl: Word default is ON; only store when explicitly set to OFF
  const wcEl = findChild(pPr, 'w:widowControl')
  if (wcEl) {
    const wcVal = attrsOf(wcEl)['w:val']
    // w:widowControl/ (no val) = on; w:widowControl w:val="0" or "false" = off
    if (wcVal === '0' || wcVal === 'false') format.widowControl = false
  }
  if (boolProp(pPr, 'w:contextualSpacing')) format.contextualSpacing = true
  const autoSpace = autoSpaceOf(pPr)
  if (autoSpace !== undefined) format.autoSpace = autoSpace
  const shd = findChild(pPr, 'w:shd')
  if (shd) {
    const fill = attrsOf(shd)['w:fill']
    if (fill && fill !== 'auto') format.shadingFill = fill
  }
  const pBdrs = findChildren(pPr, 'w:pBdr')
  if (pBdrs.length > 0) {
    let borders = ''
    const lines: NonNullable<ParaFormat['borderLines']> = {}
    for (const [side, ch] of [
      ['top', 't'],
      ['bottom', 'b'],
      ['left', 'l'],
      ['right', 'r'],
    ] as const) {
      // duplicated w:pBdr: later containers override earlier ones per side
      let el: XNode | undefined
      for (const pBdr of pBdrs) el = findChild(pBdr, `w:${side}`) ?? el
      // ST_Border spells "no border" both ways, and Word writes nil whenever a style-level
      // border is reset, so treating it as present stamps a rule the document never had
      const attrs = el ? attrsOf(el) : undefined
      const val = attrs?.['w:val']
      if (!el || val === 'none' || val === 'nil') continue
      borders += ch
      const line: import('./types').ParaBorderLine = {}
      if (attrs?.['w:color'] && attrs['w:color'] !== 'auto') line.color = attrs['w:color']
      const sz = parseInt(attrs?.['w:sz'] ?? '', 10)
      if (Number.isFinite(sz) && sz > 0) line.szPt = sz / 8
      if (line.color !== undefined || line.szPt !== undefined) lines[ch] = line
    }
    if (borders) {
      format.borders = borders
      if (Object.keys(lines).length > 0) format.borderLines = lines
    }
  }
  const stops = tabStopsOf(pPr)
  if (stops) format.tabStops = stops
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

/** HYPERLINK "url" (optional \o "tip"): the only field form folded into an editable link run;
 * any other switch (\l bookmark, \t frame...) keeps the protected-passthrough path */
function convertibleHyperlink(instr: string): { href: string; tooltip?: string } | null {
  const m = /^\s*HYPERLINK\s+"([^"\\]+)"\s*(?:\\o\s+"([^"]*)"\s*)?$/.exec(
    decodeNumericCharRefs(instr),
  )
  if (!m) return null
  return { href: m[1], ...(m[2] ? { tooltip: m[2] } : {}) }
}

/** every field instruction is EMBED/LINK (the object-field forms of legacy OLE) */
function onlyOleFields(xml: string): boolean {
  if (xml.includes('<w:fldSimple')) return false
  const instrs = xml.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? []
  if (instrs.length === 0) return false
  return instrs.every((fragment) =>
    /^\s*(EMBED|LINK)\b/.test(decodeEntities(fragment.replace(/<[^>]+>/g, ''))),
  )
}

function onlyXeFields(xml: string): boolean {
  if (xml.includes('<w:fldSimple')) return false
  const instrs = xml.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? []
  if (instrs.length === 0) return false
  let checkboxInstrs = 0
  const ok = instrs.every((fragment) => {
    const text = decodeEntities(fragment.replace(/<[^>]+>/g, ''))
    if (/^\s*FORMCHECKBOX\s*$/.test(text)) {
      checkboxInstrs++
      return true
    }
    return (
      /^\s*XE[\s"]/.test(text) ||
      /^\s*REF\s/.test(text) ||
      SIMPLE_INLINE_FIELD_RE.test(text) ||
      convertibleHyperlink(text) !== null
    )
  })
  if (!ok) return false
  // a FORMCHECKBOX without its w:checkBox definition can't fold into a glyph
  // run; keep the paragraph on the byte-preserving passthrough path instead
  // of silently dropping the field on save
  const checkBoxDefs = (xml.match(/<w:checkBox[\s/>]/g) ?? []).length
  return checkboxInstrs <= checkBoxDefs
}

/** Checked state of a FORMCHECKBOX begin run (w:fldChar/w:ffData/w:checkBox):
 * w:checked wins over w:default; either element without w:val means true */
function checkboxStateOf(beginRun: XNode | null): { checked: boolean } | null {
  if (!beginRun) return null
  const ffData = findChild(findChild(beginRun, 'w:fldChar') ?? {}, 'w:ffData')
  const box = ffData ? findChild(ffData, 'w:checkBox') : undefined
  if (!box) return null
  const state = findChild(box, 'w:checked') ?? findChild(box, 'w:default')
  if (!state) return { checked: false }
  const val = attrsOf(state)['w:val']
  return { checked: val === undefined || val === '1' || val === 'true' || val === 'on' }
}

/**
 * w:sz governing a run-less paragraph's line height: the paragraph-mark rPr
 * (pPr/w:rPr), else the last run's rPr — those runs are all empty and get
 * dropped, but Word still sizes the empty line by them (1pt spacer lines).
 */
function emptyParaSizeHalfPoints(pNode: XNode, pPr: XNode | undefined): number | undefined {
  let sz = pPr
    ? attrsOf(findChild(findChild(pPr, 'w:rPr') ?? {}, 'w:sz') ?? {})['w:val']
    : undefined
  if (!sz) {
    for (const r of findChildren(pNode, 'w:r')) {
      const v = attrsOf(findChild(findChild(r, 'w:rPr') ?? {}, 'w:sz') ?? {})['w:val']
      if (v) sz = v
    }
  }
  const n = sz ? parseInt(sz, 10) : NaN
  return Number.isFinite(n) && n > 0 ? n : undefined
}

function extractRuns(
  pNode: XNode,
  ctx: BuildContext,
  mathFragments: string[] = [],
  rubyFragments: string[] = [],
  withImages = false,
): Run[] {
  const runs: Run[] = []
  let mathIndex = 0
  let rubyIndex = 0
  // paragraph-style rtl inherits into runs without their own flag
  const pStyleId = attrsOf(findChild(findChild(pNode, 'w:pPr') ?? {}, 'w:pStyle') ?? {})['w:val']
  const paraRtl = pStyleId ? ctx.styles?.get(pStyleId)?.display?.rtl : undefined
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
  let fieldCachedRuns: Run[] = []
  let fieldBeginRun: XNode | null = null
  // reference-only comments (bare w:commentReference, LibreOffice style) anchor
  // on the nearest run; refs seen before any run attach to the next one
  let pendingRefIds: string[] = []
  const addCommentIds = (run: Run, ids: string[]) => {
    run.commentIds = [...new Set([...(run.commentIds ?? []), ...ids])].sort()
  }
  const pushRun = (run: Run, rev?: RevCtx) => {
    if (activeComments.size > 0) run.commentIds = [...activeComments].sort()
    if (pendingRefIds.length > 0) {
      addCommentIds(run, pendingRefIds)
      pendingRefIds = []
    }
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
          fieldCachedRuns = []
          fieldBeginRun = node
        }
      } else if (type === 'separate') {
        if (fieldDepth === 1) fieldSeparated = true
      } else if (type === 'end') {
        fieldDepth = Math.max(0, fieldDepth - 1)
        if (fieldDepth === 0) {
          const xe = /^\s*XE\s+(?:"([^"]*)"|(\S+))/.exec(fieldInstr)
          const ref = /^\s*REF\s+(?:"([^"]+)"|([^\s\\]+))/.exec(fieldInstr)
          const hyper = convertibleHyperlink(fieldInstr)
          if (xe) pushRun({ text: '', xeTerm: xe[1] ?? xe[2] }, rev)
          else if (ref) {
            const name = ref[1] ?? ref[2]
            pushRun({ text: fieldCached || name, refField: name, refInstr: fieldInstr }, rev)
          } else if (hyper) {
            // fold the field into plain link runs (the cached result keeps its
            // formatting); regeneration emits w:hyperlink + a fresh rel
            const linkVal: Run['link'] = {
              href: hyper.href,
              ...(hyper.tooltip ? { tooltip: hyper.tooltip } : {}),
            }
            if (fieldCachedRuns.length > 0) {
              for (const cached of fieldCachedRuns) pushRun({ ...cached, link: linkVal }, rev)
            } else pushRun({ text: hyper.href, link: linkVal }, rev)
          } else if (/^\s*FORMCHECKBOX\s*$/.test(fieldInstr)) {
            // Legacy checkbox form field: no cached result — Word draws the box
            // from ffData. Display a glyph; write the begin run back verbatim.
            const state = checkboxStateOf(fieldBeginRun)
            if (state) {
              pushRun(
                {
                  text: state.checked ? '☒' : '☐',
                  instrField: 'FORMCHECKBOX',
                  fldBeginXml: serializeXNode(fieldBeginRun!),
                },
                rev,
              )
            }
          } else if (SIMPLE_INLINE_FIELD_RE.test(fieldInstr)) {
            pushRun({ text: fieldCached || ' ', instrField: fieldInstr.trim() }, rev)
          }
          fieldInstr = ''
          fieldSeparated = false
          fieldCached = ''
          fieldCachedRuns = []
          fieldBeginRun = null
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
        // REF/HYPERLINK cached result is the display text; other fields' caches are dropped
        const cached = buildRun(
          node,
          link,
          ctx.themeColors,
          ctx.themeFonts,
          undefined,
          ctx.styles,
          paraRtl,
        )
        if (cached) {
          fieldCached += cached.text
          fieldCachedRuns.push(cached)
        }
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
    const commentRef = findChild(node, 'w:commentReference')
    if (commentRef) {
      const id = attrsOf(commentRef)['w:id']
      if (id && ctx.referenceOnlyComments?.has(id)) {
        const prev = runs[runs.length - 1]
        if (prev) addCommentIds(prev, [id])
        else pendingRefIds.push(id)
      }
    }
    const run = buildRun(
      node,
      link,
      ctx.themeColors,
      ctx.themeFonts,
      withImages ? ctx.mediaByRid : undefined,
      ctx.styles,
      paraRtl,
    )
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
      } else if (name === 'w:br') {
        // Word honors a <w:br> sitting outside any <w:r> (direct child of w:p / w:ins)
        pushRun({ text: attrsOf(node)['w:type'] === 'page' ? '\f' : '\n' }, rev)
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
      if (nameOf(c) === 'w:t') text += decodeNumericCharRefs(textOf(c))
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

/** Word's face for an East Asian theme slot whose typeface is empty (<a:ea typeface=""/>) */
const EMPTY_EA_THEME_FONT = 'DengXian'

/** empty EA slot faces by settings.xml w:themeFontLang w:eastAsia (zh-CN / missing → DengXian) */
const EMPTY_EA_SLOT_BY_LANG: Record<string, { major: string; minor: string }> = {
  ja: { major: 'Yu Gothic', minor: 'Yu Mincho' },
  // Word probe + fontTable: ko-KR empty EA slot substitutes Malgun Gothic
  ko: { major: 'Malgun Gothic', minor: 'Malgun Gothic' },
}

function emptyEaSlotFont(fonts: ThemeFonts, eaRef: string | undefined): string {
  const lang = fonts.eaLang?.toLowerCase().split('-')[0]
  const byLang = lang ? EMPTY_EA_SLOT_BY_LANG[lang] : undefined
  if (!byLang) return EMPTY_EA_THEME_FONT
  return eaRef === 'majorEastAsia' ? byLang.major : byLang.minor
}

/** w:rFonts with theme references resolved: theme attrs supersede same-slot literal
 * values (ECMA-376 §17.3.2.26). Unresolvable references fall back to the literal,
 * except an empty eastAsia theme slot: Word keeps the theme's authority and renders
 * the theme language's default face, never the leftover literal name (eaSlotEmpty marks this). */
function themedRFonts(
  attrs: Record<string, string | undefined>,
  fonts: ThemeFonts | null | undefined,
): { ascii?: string; hAnsi?: string; eastAsia?: string; cs?: string; eaSlotEmpty?: boolean } {
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
  const eaRef = attrs['w:eastAsiaTheme']
  const themedEa = themeVal(eaRef)
  const eaSlotEmpty =
    !themedEa && !!fonts && (eaRef === 'majorEastAsia' || eaRef === 'minorEastAsia')
  return {
    ascii: themeVal(attrs['w:asciiTheme']) ?? attrs['w:ascii'],
    hAnsi: themeVal(attrs['w:hAnsiTheme']) ?? attrs['w:hAnsi'],
    eastAsia: themedEa ?? (eaSlotEmpty ? emptyEaSlotFont(fonts!, eaRef) : attrs['w:eastAsia']),
    cs: themeVal(attrs['w:cstheme']) ?? attrs['w:cs'],
    ...(eaSlotEmpty ? { eaSlotEmpty } : {}),
  }
}

function buildRun(
  rNode: XNode,
  link?: Run['link'],
  theme?: ThemeColors | null,
  themeFonts?: ThemeFonts | null,
  mediaByRid?: Map<string, string>,
  styles?: Map<string, StyleInfo>,
  paraRtl?: boolean,
): Run | null {
  let text = ''
  for (const child of childrenOf(rNode)) {
    const name = nameOf(child)
    if (name === 'w:t' || name === 'w:delText') text += decodeNumericCharRefs(textOf(child))
    else if (name === 'w:tab' || name === 'w:ptab') text += '\t'
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
  let image: Run['image']
  if (mediaByRid) {
    const drawing = findChild(rNode, 'w:drawing')
    if (drawing) {
      const drawingXml = serializeXNode(drawing)
      const rId = /<a:blip[^>]*r:(?:embed|link)="([^"]+)"/.exec(drawingXml)?.[1]
      const dataUrl = rId ? mediaByRid.get(rId) : undefined
      if (dataUrl) {
        image = { dataUrl, xml: drawingXml }
        const extent = /<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(drawingXml)
        const cx = Number(extent?.[1])
        const cy = Number(extent?.[2])
        if (cx > 0) image.widthPx = Math.round(cx / EMU_PER_PX)
        if (cy > 0) image.heightPx = Math.round(cy / EMU_PER_PX)
        // anchored (floating) picture kept as a run image: carry the wrap kind
        // and anchor offsets so the editor can float it instead of inlining
        if (/<wp:anchor[\s>]/.test(drawingXml)) {
          const meta = imageMeta(drawingXml)
          if (meta.imageWrap) image.wrap = meta.imageWrap
          if (meta.imageOffsetXEmu !== undefined) image.offsetXEmu = meta.imageOffsetXEmu
          if (meta.imageOffsetYEmu !== undefined) image.offsetYEmu = meta.imageOffsetYEmu
        }
      }
    }
    // legacy VML picture (w:pict + v:imagedata, Word 2003 era / stamps) or an
    // inline OLE embed (w:object, whose preview is also a v:imagedata): same
    // run-level image treatment; the fragment round-trips verbatim on save
    if (!image) {
      const pict = findChild(rNode, 'w:pict') ?? findChild(rNode, 'w:object')
      if (pict) {
        const pictXml = serializeXNode(pict)
        const rId = /<v:imagedata[^>]*r:id="([^"]+)"/.exec(pictXml)?.[1]
        const dataUrl = rId ? mediaByRid.get(rId) : undefined
        if (dataUrl) {
          image = { dataUrl, xml: pictXml }
          const style = /<v:shape [^>]*style="([^"]*)"/.exec(pictXml)?.[1] ?? ''
          const w = parseFloat(/(?:^|;)width:([\d.]+)pt/.exec(style)?.[1] ?? '')
          const h = parseFloat(/(?:^|;)height:([\d.]+)pt/.exec(style)?.[1] ?? '')
          // w:object declares its size in twips when the v:shape style is absent
          const objAttrs = /<w:object\b[^>]*>/.exec(pictXml)?.[0] ?? ''
          const wTw = parseInt(/w:dxaOrig="(\d+)"/.exec(objAttrs)?.[1] ?? '', 10)
          const hTw = parseInt(/w:dyaOrig="(\d+)"/.exec(objAttrs)?.[1] ?? '', 10)
          if (w > 0) image.widthPx = Math.round((w / 72) * 96)
          else if (wTw > 0) image.widthPx = Math.round(wTw / 15)
          if (h > 0) image.heightPx = Math.round((h / 72) * 96)
          else if (hTw > 0) image.heightPx = Math.round(hTw / 15)
        }
      }
    }
  }
  if (text === '' && !image) return null

  const rPr = findChild(rNode, 'w:rPr')
  const run: Run = { text }
  if (image) run.image = image
  if (link) run.link = link
  // a bare run under an rtl style still selects the Cs set from its style chain
  if (!rPr && paraRtl) run.cs = true
  if (rPr) {
    run.rawRPr = serializeXNode(rPr)
    const rStyle = attrsOf(findChild(rPr, 'w:rStyle') ?? {})['w:val']
    if (rStyle && rStyle !== 'Hyperlink') run.styleId = rStyle
    // Word picks the whole property set by w:rtl (probed, Word for Mac 2026-08):
    // rtl runs read w:bCs/w:iCs/w:szCs with no fallback to w:b/w:i/w:sz; non-rtl
    // runs read the base props and ignore the Cs twins entirely. Script content
    // and paragraph w:bidi play no part. Font slot choice is separate. The style
    // chain's w:rtl (character style, then paragraph style) is only the inherited
    // value for runs without an explicit flag.
    const inheritedRtl = (rStyle ? styles?.get(rStyle)?.display?.rtl : undefined) ?? paraRtl
    const cs = (onOffOf(rPr, 'w:rtl') ?? inheritedRtl) === true
    if (cs) run.cs = true
    const bold = onOffOf(rPr, cs ? 'w:bCs' : 'w:b')
    if (bold !== undefined) run.bold = bold
    const italic = onOffOf(rPr, cs ? 'w:iCs' : 'w:i')
    if (italic !== undefined) run.italic = italic
    if (underlineProp(rPr)) run.underline = true
    else if (attrsOf(findChild(rPr, 'w:u') ?? {})['w:val'] === 'none') run.underline = false
    const strike = onOffOf(rPr, 'w:strike')
    if (strike !== undefined) run.strike = strike
    const color = colorFrom(rPr, theme) ?? w14TextFillHex(rPr, theme)
    if (color) run.color = color
    const sz = attrsOf(findChild(rPr, cs ? 'w:szCs' : 'w:sz') ?? {})['w:val']
    if (sz) run.sizeHalfPoints = parseInt(sz, 10) || undefined
    const rf = themedRFonts(attrsOf(findChild(rPr, 'w:rFonts') ?? {}), themeFonts)
    const font = rf.eastAsia ?? rf.ascii ?? rf.hAnsi
    if (font) run.font = font
    if (rf.eaSlotEmpty && font && font === rf.eastAsia) run.eaSlotEmpty = true
    const fontAscii = rf.ascii ?? rf.hAnsi
    if (fontAscii) run.fontAscii = fontAscii
    if (rf.cs) run.csFont = rf.cs
    const spc = parseInt(attrsOf(findChild(rPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
    if (spc) run.charSpacingTwips = spc
    // w:caps wins over w:smallCaps when both are on (Word)
    if (onOffOf(rPr, 'w:caps')) run.caps = 'all'
    else if (onOffOf(rPr, 'w:smallCaps')) run.caps = 'small'
    const wScale = parseInt(attrsOf(findChild(rPr, 'w:w') ?? {})['w:val'] ?? '', 10)
    if (wScale > 0 && wScale !== 100) run.charScalePct = wScale
    const highlight = attrsOf(findChild(rPr, 'w:highlight') ?? {})['w:val']
    if (highlight && highlight !== 'none') run.highlight = highlight
    const shdFill = attrsOf(findChild(rPr, 'w:shd') ?? {})['w:fill']
    if (shdFill && shdFill !== 'auto') run.shading = shdFill
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
        // the pre-revision snapshot decodes under the same rtl selection
        const ocs = (onOffOf(oldRPr, 'w:rtl') ?? inheritedRtl) === true
        if (boolProp(oldRPr, ocs ? 'w:bCs' : 'w:b')) old.bold = true
        if (boolProp(oldRPr, ocs ? 'w:iCs' : 'w:i')) old.italic = true
        if (underlineProp(oldRPr)) old.underline = true
        if (boolProp(oldRPr, 'w:strike')) old.strike = true
        const oc = colorFrom(oldRPr, theme)
        if (oc) old.color = oc
        const osz = attrsOf(findChild(oldRPr, ocs ? 'w:szCs' : 'w:sz') ?? {})['w:val']
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
  if (a.image || b.image) return false
  return (
    (a.rawRPr ?? '') === (b.rawRPr ?? '') &&
    a.styleId === b.styleId &&
    !!a.cs === !!b.cs &&
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    !!a.strike === !!b.strike &&
    a.color === b.color &&
    a.sizeHalfPoints === b.sizeHalfPoints &&
    a.font === b.font &&
    a.fontAscii === b.fontAscii &&
    a.csFont === b.csFont &&
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
  // whole try: hostile depth inside a cell paragraph can overflow the
  // run-extraction recursion — degrade to a protected block, not a failed document
  try {
    const parsed = deepXmlParser.parse(xml) as XNode[]
    const tbl = parsed.find((n) => nameOf(n) === 'w:tbl')
    if (!tbl) return undefined
    const model = extractTableModel(tbl, ctx)
    if (!model) return undefined
    const rawTrPrs: Array<string | null> = model.rows.map(() => null)
    attachRawTablePr(xml, model.rows, rawTrPrs)
    if (rawTrPrs.some((r) => r !== null)) model.rawTrPrs = rawTrPrs
    return model
  } catch {
    return undefined
  }
}

/** One w:tbl node → display model (shared by top-level tables and tables nested in cells) */
/**
 * Per-column widths reconstructed from cell w:tcW (dxa) across all rows: the widest
 * un-spanned cell per grid slot wins (Word widens a column to fit later rows, never
 * narrows it). Undefined unless every column got a value — partial data would skew
 * the ratio worse than the tblGrid fallback.
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
      // duplicated w:tcW: Word keeps the last occurrence (generators leave stale first values)
      const a = attrsOf(findChildren(tcPr ?? {}, 'w:tcW').at(-1) ?? {})
      const w = !a['w:type'] || a['w:type'] === 'dxa' ? Number(a['w:w']) || 0 : 0
      if (span === 1 && w > 0) cols[idx] = Math.max(cols[idx] || 0, w)
      idx += span
    }
    colCount = Math.max(colCount, idx)
  }
  if (colCount === 0) return undefined
  for (let i = 0; i < colCount; i++) if (!(cols[i] > 0)) return undefined
  return cols.slice(0, colCount)
}

/** Real documents rarely nest past 3-4 levels; below the cap the subtree flattens so stress files (POI nests 5000) cannot blank the page. */
const MAX_TABLE_NEST_DEPTH = 8

/** Whole subtree → 1×1 sub-table of plain paragraph texts (iterative: the subtree can be thousands of levels deep). */
function flattenedTableModel(tbl: XNode): TableModel | undefined {
  const paras: string[] = []
  const PARA_END: XNode = {}
  let buf: string | null = null
  const stack: XNode[] = [tbl]
  while (stack.length > 0) {
    const node = stack.pop()!
    if (node === PARA_END) {
      paras.push(buf ?? '')
      buf = null
      continue
    }
    if ('#text' in node) {
      if (buf !== null) buf += String(node['#text'])
      continue
    }
    if (buf === null && nameOf(node) === 'w:p') {
      buf = ''
      stack.push(PARA_END)
    }
    const kids = childrenOf(node)
    for (let i = kids.length - 1; i >= 0; i--) stack.push(kids[i])
  }
  if (paras.length === 0) return undefined
  const cell: TableCell = {
    paras,
    richParas: paras.map((text) => ({ runs: text === '' ? [] : [{ text }] })),
  }
  // autofit: without it the synthetic table collapses to a sliver and wraps every word
  return { rows: [[cell]], autoLayout: true }
}

function extractTableModel(tbl: XNode, ctx: BuildContext, depth = 1): TableModel | undefined {
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
  const tblPrNode = findChild(tbl, 'w:tblPr')
  const fixedLayout = attrsOf(findChild(tblPrNode ?? {}, 'w:tblLayout') ?? {})['w:type'] === 'fixed'
  // Cell-level w:tcW is Word's actual layout input for auto tables; generators often
  // leave a stale evenly-split tblGrid behind. When the two disagree, tcW wins.
  const tcwWidths = tcwColumnWidths(tbl)
  if (tcwWidths) {
    const tcwTotal = tcwWidths.reduce((a, b) => a + b, 0)
    const tcwPct = tcwWidths.map((w) => (w / tcwTotal) * 100)
    // Fixed layout sizes columns from tcW alone, so a matching ratio with a
    // different absolute sum still means the grid is stale.
    const gridTotal = (colWidthsTwips ?? []).reduce((a, b) => a + b, 0)
    const disagree =
      !colWidthsPct ||
      colWidthsPct.length !== tcwPct.length ||
      colWidthsPct.some((w, i) => Math.abs(w - tcwPct[i]) > 2) ||
      (fixedLayout && Math.abs(gridTotal - tcwTotal) > tcwWidths.length)
    if (disagree) {
      colWidthsPct = tcwPct
      colWidthsTwips = tcwWidths
    }
  }
  const tblWNode = findChild(tblPrNode ?? {}, 'w:tblW')
  const tblW = attrsOf(tblWNode ?? {})
  let widthPct: number | undefined
  if (tblW['w:type'] === 'pct') {
    const raw = String(tblW['w:w'] ?? '')
    // The pct unit is 1/50 of a percentage point; some generators write a literal "NN%"
    const pct = raw.endsWith('%') ? parseFloat(raw) : Number(raw) / 50
    if (Number.isFinite(pct) && pct > 0 && pct <= 100) widthPct = pct
  }
  const cellMar = cellMarginsOf(findChild(tblPrNode ?? {}, 'w:tblCellMar'))
  const tblBorders = mergedBorderLinesOf(tblPrNode, 'w:tblBorders', true)
  const tblJc = attrsOf(findChild(tblPrNode ?? {}, 'w:jc') ?? {})['w:val']
  const tblAlign =
    tblJc === 'center' ? 'center' : tblJc === 'right' || tblJc === 'end' ? 'right' : undefined
  const tblInd = attrsOf(findChild(tblPrNode ?? {}, 'w:tblInd') ?? {})
  const tblIndTwips = !tblInd['w:type'] || tblInd['w:type'] === 'dxa' ? Number(tblInd['w:w']) : NaN

  const rows: TableCell[][] = []
  const rowHeightsTwips: Array<number | null> = []
  const rowHeightRules: NonNullable<TableModel['rowHeightRules']> = []
  const rowRevisions: NonNullable<TableModel['rowRevisions']> = []
  for (const tr of childrenThroughSdt(tbl, 'w:tr')) {
    const cells: TableCell[] = []
    for (const tc of childrenThroughSdt(tr, 'w:tc')) {
      const cell = extractCell(tc, ctx, depth)
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
      const trH = trPr ? attrsOf(findChild(trPr, 'w:trHeight') ?? {}) : {}
      const h = Number(trH['w:val'])
      const hasH = Number.isFinite(h) && h > 0
      // Word clamps trHeight to 31680 twips / 22in (MS-OI29500 2.1.51); some generators leak EMU-scale values here
      rowHeightsTwips.push(hasH ? Math.min(h, 31680) : null)
      rowHeightRules.push(hasH ? (trH['w:hRule'] === 'exact' ? 'exact' : 'atLeast') : null)
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
  const tblWType = tblW['w:type']
  const autoWidth =
    !tblWNode ||
    tblWType === 'auto' ||
    ((!tblWType || tblWType === 'dxa') && !(Number(tblW['w:w']) > 0))
  if (!fixedLayout && autoWidth && !widthPct) model.autoLayout = true
  if (effCellMar) model.cellMarTwips = effCellMar
  if (effBorders) model.borders = effBorders
  if (tblAlign) model.align = tblAlign
  if (Number.isFinite(tblIndTwips) && tblIndTwips !== 0) model.indentTwips = tblIndTwips
  const tblStyle = attrsOf(findChild(findChild(tbl, 'w:tblPr') ?? {}, 'w:tblStyle') ?? {})['w:val']
  if (tblStyle) model.tblStyleId = tblStyle
  if (tblPrNode && boolProp(tblPrNode, 'w:bidiVisual')) model.bidiVisual = true
  if (rowHeightsTwips.some((h) => h !== null)) {
    model.rowHeightsTwips = rowHeightsTwips
    model.rowHeightRules = rowHeightRules
  }
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
  const lastRowOn = flag('w:lastRow', 0x40, false)
  const firstColOn = flag('w:firstColumn', 0x80, true)
  const lastColOn = flag('w:lastColumn', 0x100, false)
  const hBandOn = !flag('w:noHBand', 0x200, false)

  const totalCols = Math.max(
    ...rows.map((row) => row.reduce((sum, c) => sum + (c.colSpan ?? 1), 0)),
  )
  rows.forEach((row, r) => {
    const isFirst = firstRowOn && r === 0
    const isLast = lastRowOn && r === rows.length - 1
    const bandRow = firstRowOn ? r - 1 : r
    const bandFill =
      hBandOn && bandRow >= 0 ? (bandRow % 2 === 0 ? ts.band1Fill : ts.band2Fill) : undefined
    let col = 0
    for (const cell of row) {
      const span = cell.colSpan ?? 1
      // Word's conditional-format precedence: rows beat columns, all beat bands/whole-table
      const conds = [
        isFirst ? ts.firstRow : undefined,
        isLast ? ts.lastRow : undefined,
        firstColOn && col === 0 ? ts.firstCol : undefined,
        lastColOn && col + span === totalCols ? ts.lastCol : undefined,
      ]
      col += span
      if (cell.fill === undefined) {
        cell.fill = conds.find((c) => c?.fill)?.fill ?? bandFill ?? ts.fill ?? undefined
      }
      const bold = conds.some((c) => c?.bold) || ts.wholeTable?.bold
      if (bold && cell.bold === undefined) cell.bold = true
      const color = conds.find((c) => c?.color)?.color ?? ts.wholeTable?.color
      if (color && !cell.color) cell.color = color
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

/** Duplicated border containers (two w:tcBorders in one tcPr etc.): Word merges per side, later wins */
function mergedBorderLinesOf(
  parent: XNode | undefined,
  tag: string,
  withInside: true,
): TableBorders | undefined
function mergedBorderLinesOf(
  parent: XNode | undefined,
  tag: string,
  withInside: false,
): CellBorders | undefined
function mergedBorderLinesOf(
  parent: XNode | undefined,
  tag: string,
  withInside: boolean,
): TableBorders | undefined {
  if (!parent) return undefined
  let merged: TableBorders | undefined
  for (const node of findChildren(parent, tag)) {
    const b = borderLinesOf(node, withInside as true)
    if (b) merged = { ...merged, ...b }
  }
  return merged
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

const ANCHOR_HOSTS = new Set(['w:drawing', 'w:pict'])

/** depth-first search for any descendant with one of the given names */
function hasDeepChild(node: XNode, names: Set<string>): boolean {
  for (const child of childrenOf(node)) {
    const name = nameOf(child)
    if (name && names.has(name)) return true
    if (hasDeepChild(child, names)) return true
  }
  return false
}

function extractCell(tc: XNode, ctx: BuildContext, depth: number): TableCell {
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
    const borders = mergedBorderLinesOf(tcPr, 'w:tcBorders', false)
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
  // outer table's responsibility); anchors record their position among the paragraphs
  const nested: TableModel[] = []
  const nestedAnchors: number[] = []
  let sawBold = false
  let sawNonBold = false
  const runColors = new Set<string>()
  const textParaJcs = new Set<string>()
  for (const block of childrenThroughSdt(tc, ['w:p', 'w:tbl'])) {
    if (nameOf(block) === 'w:tbl') {
      const model =
        depth >= MAX_TABLE_NEST_DEPTH
          ? flattenedTableModel(block)
          : extractTableModel(block, ctx, depth + 1)
      if (model) {
        nested.push(model)
        nestedAnchors.push(cell.paras.length)
      }
      continue
    }
    let p = block
    // anchored shapes/textboxes in cell paragraphs (blip images already ride the
    // runs): Word renders them inside the cell and grows the row to hold them
    // (tdf134277). Their w:txbxContent is stripped from the paragraph so the box
    // text does not additionally render as plain cell text.
    if (hasDeepChild(p, ANCHOR_HOSTS)) {
      // Word pairs every DrawingML shape with a VML twin in mc:Fallback: strip the
      // fallback like buildBlock's detect, or each shape extracts twice
      const rawPXml = serializeXNode(p)
      const pXml = rawPXml.includes('<mc:Fallback')
        ? rawPXml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '')
        : rawPXml
      if (pXml.includes('<wp:anchor') || pXml.includes('<w:pict')) {
        const boxes = extractTextboxes(pXml, ctx, { shapes: true })
        if (boxes.length > 0) {
          cell.anchoredBoxes = [...(cell.anchoredBoxes ?? []), ...boxes]
          // the boxes now display separately: drop the anchored drawings (and
          // textbox picts) from the paragraph node so their inner text/offsets
          // don't leak into the cell's own runs; inline drawings stay for images
          try {
            let txml = pXml
            for (const frag of topLevelDrawings(txml)) {
              // keep anchored pictures: they are not extracted as shape boxes
              // and must stay on the run-image path
              if (frag.includes('<wp:anchor') && !frag.includes('<pic:pic')) {
                txml = txml.split(frag).join('')
              }
            }
            txml = txml.replace(
              /<w:pict>(?:(?!<\/w:pict>)[\s\S])*?<w:txbxContent>[\s\S]*?<\/w:pict>/g,
              '',
            )
            const stripped = (xmlParser.parse(txml) as XNode[])[0]
            if (stripped && nameOf(stripped) === 'w:p') p = stripped
          } catch {
            /* keep the original paragraph node */
          }
        }
      }
    }
    const paraText = textOf(p)
    cell.paras.push(paraText)
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
    const runs = extractRuns(p, ctx, [], [], true)
    const emptySz = runs.length === 0 ? emptyParaSizeHalfPoints(p, pPr) : undefined
    richParas.push({
      ...format,
      ...(cellStyleId ? { styleId: cellStyleId } : {}),
      ...(emptySz ? { emptyRunSizeHalfPoints: emptySz } : {}),
      ...(list ? { list } : {}),
      runs,
    })
    if (paraText !== '') textParaJcs.add(attrsOf(findChild(pPr ?? {}, 'w:jc') ?? {})['w:val'] ?? '')
    for (const r of findChildren(p, 'w:r')) {
      const rPr = findChild(r, 'w:rPr')
      if (rPr && boolProp(rPr, 'w:b')) sawBold = true
      else sawNonBold = true
      if (textOf(r) !== '') runColors.add((rPr && colorFrom(rPr, ctx.themeColors)) ?? 'none')
    }
  }
  // cell.align only when every text paragraph declares the same jc; a first-wins
  // td-level text-align would leak onto the cell's jc-less paragraphs
  if (textParaJcs.size === 1) {
    const jc = textParaJcs.values().next().value
    if (jc === 'center' || jc === 'right' || jc === 'left' || jc === 'justify') cell.align = jc
  }
  cell.richParas = richParas
  if (nested.length > 0) {
    cell.nestedTables = nested
    cell.nestedTableAnchors = nestedAnchors.map((a) => Math.min(a, cell.paras.length))
  }
  if (sawBold && !sawNonBold) cell.bold = true
  // cell.color only when every text run agrees (mixed colors stay run-level)
  if (runColors.size === 1) {
    const only = runColors.values().next().value as string
    if (only !== 'none') cell.color = only
  }
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
  // inline pictures in layout tables (nested ones included) live on their cell
  // runs (hfTableRowParagraphs); floating ones still position through this
  // part-level list
  const tbls = hfTblRanges(partXml)
  const onCellRun = (at: number) => tbls.some(([s, e]) => at > s && at < e)
  const out: HfImage[] = []
  /** w:jc of the paragraph containing offset `at` (inline images follow it) */
  const paraAlignAt = (at: number): HfImage['align'] => {
    const pStart = Math.max(partXml.lastIndexOf('<w:p ', at), partXml.lastIndexOf('<w:p>', at))
    if (pStart < 0) return undefined
    const jc = /<w:jc w:val="(\w+)"/.exec(partXml.slice(pStart, at))?.[1]
    return jc === 'left' || jc === 'center' || jc === 'right' ? jc : undefined
  }
  for (const m of partXml.matchAll(/<w:drawing>[\s\S]*?<\/w:drawing>/g)) {
    const frag = m[0]
    if (!/<wp:anchor[\s>]/.test(frag) && onCellRun(m.index!)) continue
    // mc:AlternateContent may hold several blips (mac Word: PDF Choice + PNG
    // Fallback); use the first one whose media resolves
    let dataUrl: string | null = null
    for (const b of frag.matchAll(/<a:blip[^>]*r:embed="([^"]+)"/g)) {
      dataUrl = await mediaDataUrl(zip, rels, b[1])
      if (dataUrl) break
    }
    if (!dataUrl) continue
    const image: HfImage = { dataUrl }
    const extent = /<wp:extent[^>]*\/?>/.exec(frag)?.[0] ?? ''
    const cx = parseInt(/cx="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    const cy = parseInt(/cy="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    if (Number.isFinite(cx) && cx > 0) image.widthPx = Math.round(cx / EMU_PER_PX)
    if (Number.isFinite(cy) && cy > 0) image.heightPx = Math.round(cy / EMU_PER_PX)
    if (/<wp:anchor[\s>]/.test(frag)) {
      image.floating = true
      const anchorTag = /<wp:anchor[^>]*>/.exec(frag)?.[0] ?? ''
      if (/behindDoc="(?:1|true)"/.test(anchorTag)) image.behind = true
      readAnchorPos(frag, image)
    } else {
      const align = paraAlignAt(m.index!)
      if (align) image.align = align
    }
    out.push(image)
  }
  for (const m of partXml.matchAll(/<w:pict>[\s\S]*?<\/w:pict>/g)) {
    const frag = m[0]
    if (frag.includes('<v:textpath')) continue
    if (!/position:\s*absolute/.test(frag) && onCellRun(m.index!)) continue
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
    // absolute-positioned shapes (picture watermarks): must not stack into the
    // header strip nor count toward the header height estimate
    if (/position:absolute/.test(style)) {
      image.floating = true
      if (/z-index:\s*-/.test(style)) image.behind = true
      const posH = /mso-position-horizontal:(\w+)/.exec(style)?.[1]
      const posV = /mso-position-vertical:(\w+)/.exec(style)?.[1]
      if (posH === 'left' || posH === 'center' || posH === 'right') image.posH = posH
      if (posV === 'top' || posV === 'center' || posV === 'bottom') image.posV = posV
    } else {
      const align = paraAlignAt(m.index!)
      if (align) image.align = align
    }
    if (/<v:imagedata[^>]*(?:gain|blacklevel)="/.test(frag)) image.washout = true
    out.push(image)
  }
  return out
}

/** wp:anchor wp:positionH/V of a header/footer image: wp:align keeps the VML-style
 *  alignment fields, wp:posOffset (EMU) becomes a px offset from the page edge or
 *  margin box (paragraph/column/character origins approximate to margin). */
function readAnchorPos(frag: string, image: HfImage): void {
  for (const axis of ['H', 'V'] as const) {
    const m = new RegExp(
      `<wp:position${axis}[^>]*relativeFrom="([^"]+)"[^>]*>([\\s\\S]*?)</wp:position${axis}>`,
    ).exec(frag)
    if (!m) continue
    const align = /<wp:align>(\w+)<\/wp:align>/.exec(m[2])?.[1]
    const offset = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(m[2])?.[1]
    if (align) {
      if (axis === 'H' && (align === 'left' || align === 'center' || align === 'right')) {
        image.posH = align
      }
      if (axis === 'V' && (align === 'top' || align === 'center' || align === 'bottom')) {
        image.posV = align
      }
    } else if (offset != null) {
      const px = Math.round(Number(offset) / EMU_PER_PX)
      const rel = m[1] === 'page' ? 'page' : 'margin'
      if (axis === 'H') {
        image.posXPx = px
        image.posHRel = rel
      } else {
        image.posYPx = px
        image.posVRel = rel
      }
    }
  }
}

/** [start, end) spans of top-level w:tbl elements (nesting-aware) */
function hfTblRanges(xml: string): Array<[number, number]> {
  const out: Array<[number, number]> = []
  const re = /<w:tbl[\s>]|<\/w:tbl>/g
  let depth = 0
  let start = 0
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (m[0] === '</w:tbl>') {
      if (depth > 0 && --depth === 0) out.push([start, m.index + m[0].length])
    } else {
      if (depth === 0) start = m.index
      depth++
    }
  }
  return out
}

/** Pre-resolved image media (rId -> data/external URL) for pictures inside a
 *  header/footer part's layout tables, so the sync cell-run extraction can
 *  attach them (mirrors tableBlipMedia). */
async function hfTableMedia(
  zip: JSZip,
  partPath: string,
  partXml: string,
): Promise<Map<string, string> | undefined> {
  if (!partXml.includes('<w:tbl')) return undefined
  if (!partXml.includes('<a:blip') && !partXml.includes('<v:imagedata')) return undefined
  const rels = await parseRels(zip, partPath.replace(/([^/]+)$/, '_rels/$1.rels'))
  const out = new Map<string, string>()
  for (const [start, end] of hfTblRanges(partXml)) {
    const slice = partXml.slice(start, end)
    const refs = [
      ...slice.matchAll(/<a:blip[^>]*r:(?:embed|link)="([^"]+)"/g),
      ...slice.matchAll(/<v:imagedata[^>]*r:id="([^"]+)"/g),
    ]
    for (const m of refs) {
      const rId = m[1]
      if (out.has(rId)) continue
      const rel = rels.get(rId)
      if (!rel) continue
      if (rel.targetMode === 'External' || /^https?:\/\//i.test(rel.target)) {
        out.set(rId, rel.target)
        continue
      }
      const dataUrl = await mediaDataUrl(zip, rels, rId)
      if (dataUrl) out.set(rId, dataUrl)
    }
  }
  return out.size > 0 ? out : undefined
}

/** settings.xml compatSetting compatibilityMode (0 when absent = legacy layout) */
async function parseCompatibilityMode(zip: JSZip): Promise<number> {
  const file = zip.file('word/settings.xml')
  if (!file) return 0
  const xml = await file.async('string')
  const m = /<w:compatSetting[^>]*w:name="compatibilityMode"[^>]*w:val="(\d+)"/.exec(xml)
  return m ? parseInt(m[1], 10) : 0
}

/** settings.xml <w:evenAndOddHeaders/> (w:val="0|false" counts as off) */
async function parseEvenAndOddHeaders(zip: JSZip): Promise<boolean> {
  const file = zip.file('word/settings.xml')
  if (!file) return false
  return xmlFlagOn(await file.async('string'), 'w:evenAndOddHeaders')
}

/** header/footer part XML -> display content (PAGE fields shown as PAGE_MARK) */
function hfContentFromXml(
  xml: string,
  kind: 'header' | 'footer',
  theme?: ThemeColors | null,
  styles?: Map<string, StyleInfo>,
  tableMedia?: Map<string, string>,
): { text: string; hasPageNumber: boolean; watermark: string | null; paras: HfParagraph[] } {
  // Rewrite each field span (begin..end) for display. PAGE and NUMPAGES become
  // private-use markers (the renderer substitutes real numbers; a literal '#'
  // in the part text must never be mistaken for the field), dropping their
  // stale cached results; other fields (DATE, STYLEREF, ...) keep their cached
  // result runs (Word refreshes them on open). fldChar attribute matching is
  // tolerant (Pages writes w:fldLock="0" etc.).
  // hasPageNumber is set by the same match that emits PAGE_MARK, so the two
  // can't drift (Word may split "PAGE" across several instrText runs).
  let hasPageNumber = false
  // mc:AlternateContent carries the same textbox twice (DrawingML Choice + VML
  // Fallback); keeping both prints the content twice (e.g. duplicated "— PAGE —"
  // page numbers), so only the Choice branch feeds text/paragraph extraction
  let cleaned = xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '')
  cleaned = cleaned.replace(
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
        return emit(`<w:r>${rPr}<w:t>${PAGE_MARK}</w:t></w:r>`)
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
        return `<w:r>${rPr}<w:t>${PAGE_MARK}</w:t></w:r>`
      }
      return inner ?? whole
    },
  )
  return {
    text: plainText(cleaned),
    hasPageNumber,
    watermark: kind === 'header' ? readWatermarkText(xml) : null,
    // strip leftover field chars so the page marker parses as plain text
    paras: hfParagraphs(cleaned.replace(/<w:fldChar[^>]*\/>/g, ''), theme, styles, tableMedia),
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
  styles?: Map<string, StyleInfo>,
): Promise<{
  text: string
  hasPageNumber: boolean
  watermark: string | null
  paras: HfParagraph[]
  images?: HfImage[]
} | null> {
  const refs = documentXml.match(new RegExp(`<w:${kind}Reference[^>]*/>`, 'g')) ?? []
  const typed = refs.find((r) => r.includes(`w:type="${hfType}"`))
  // untyped references count as default (w:type is technically required but often
  // omitted); non-schema w:type="odd" is Word's default (odd-page) part too
  const ref =
    hfType === 'default'
      ? (typed ??
        refs.find((r) => r.includes('w:type="odd"')) ??
        refs.find((r) => !/w:type="/.test(r)))
      : typed
  if (!ref) return null
  const rId = /r:id="([^"]+)"/.exec(ref)?.[1]
  const target = rId ? rels.get(rId)?.target : undefined
  if (!target) return null
  const path = target.startsWith('/') ? target.slice(1) : `word/${target}`
  const file = zip.file(path)
  if (!file) return null
  const xml = await file.async('string')
  const content = hfContentFromXml(xml, kind, theme, styles, await hfTableMedia(zip, path, xml))
  const images = await hfImages(zip, path, xml)
  return images.length > 0 ? { ...content, images } : content
}

/** All header/footer parts by rId (multi-section docs look them up via each section's sectPr refs) */
async function parseAllHfParts(
  zip: JSZip,
  rels: Map<string, RelInfo>,
  styles: Map<string, StyleInfo> | undefined,
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
    const content = hfContentFromXml(xml, kind, theme, styles, await hfTableMedia(zip, path, xml))
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
function hfParagraphs(
  partXml: string,
  theme?: ThemeColors | null,
  styles?: Map<string, StyleInfo>,
  tableMedia?: Map<string, string>,
): HfParagraph[] {
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
    styles,
  } as unknown as BuildContext
  const out: HfParagraph[] = []
  // floating tables (w:tblpPr) anchor to the paragraph that follows them in
  // markup; Word draws that paragraph first, so their rows are deferred past it
  const deferred: HfParagraph[] = []
  const flushDeferred = () => {
    out.push(...deferred)
    deferred.length = 0
  }
  // paragraphs may sit inside (nested) w:sdt content controls (OpenXML SDK footers)
  for (const node of childrenThroughSdt(root, ['w:tbl', 'w:p'])) {
    const name = nameOf(node)
    if (name === 'w:tbl') {
      // layout tables (logo | title | date rows): one display paragraph per row
      const rows = hfTableRowParagraphs(
        node,
        tableMedia ? ({ ...ctx, mediaByRid: tableMedia } as BuildContext) : ctx,
      )
      if (findChild(findChild(node, 'w:tblPr') ?? {}, 'w:tblpPr')) deferred.push(...rows)
      else out.push(...rows)
      continue
    }
    if (name !== 'w:p') continue
    const pNode = node
    const runs = extractRuns(pNode, ctx)
    if (runs.length === 0 && (findChild(pNode, 'w:r') || findChild(pNode, 'w:pict'))) {
      // Government-style footers keep their text (e.g. the "— PAGE —" page number)
      // inside a VML textbox shape; surface those inner paragraphs instead of dropping
      // the content. Watermark / decorative drawing paragraphs still skip.
      out.push(...textboxParagraphs(pNode, ctx))
      flushDeferred()
      continue
    }
    const pPr = findChild(pNode, 'w:pPr')
    const direct = pPr ? extractParaFormat(pPr) : undefined
    // display-only style layer: Word's built-in Header/Footer styles carry the
    // center/right tab stops (and sometimes w:jc); direct pPr wins per property
    const styleId = pPr ? attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] : undefined
    const d = styleId ? styles?.get(styleId)?.display : undefined
    // absolute position tabs (w:ptab): carry their own alignment, ignore stops.
    // Indexed by overall tab order — regular w:tab occupies a slot as undefined,
    // so mixed tab/ptab paragraphs keep their alignments on the right segment.
    const ptabAligns: Array<'left' | 'center' | 'right' | undefined> = []
    let sawPtab = false
    const walkTabs = (n: XNode): void => {
      for (const c of childrenOf(n)) {
        const name = nameOf(c)
        if (name === 'w:pPr') continue
        if (name === 'w:tab') ptabAligns.push(undefined)
        else if (name === 'w:ptab') {
          sawPtab = true
          const a = attrsOf(c)['w:alignment']
          ptabAligns.push(a === 'center' ? 'center' : a === 'right' ? 'right' : 'left')
        } else walkTabs(c)
      }
    }
    walkTabs(pNode)
    // w:framePr frames (page-number "1" floated at the right margin): the frame
    // shares the following paragraph's flow line instead of stacking above it
    const framePr = pPr ? findChild(pPr, 'w:framePr') : undefined
    const frameAttrs = framePr ? attrsOf(framePr) : undefined
    const xAlign = frameAttrs && !frameAttrs['w:dropCap'] ? frameAttrs['w:xAlign'] : undefined
    const frameXAlign =
      xAlign === 'right' || xAlign === 'outside'
        ? ('right' as const)
        : xAlign === 'center'
          ? ('center' as const)
          : xAlign === 'left' || xAlign === 'inside'
            ? ('left' as const)
            : undefined
    out.push({
      ...(d?.align && d.align !== 'justify' ? { align: d.align } : {}),
      ...(d?.tabStops ? { tabStops: d.tabStops } : {}),
      ...direct,
      ...(sawPtab ? { ptabAligns } : {}),
      ...(frameXAlign ? { frameXAlign } : {}),
      runs,
    })
    flushDeferred()
  }
  flushDeferred()
  // trailing all-empty paragraphs are layout noise
  while (out.length > 0 && out[out.length - 1].runs.length === 0 && !out[out.length - 1].cells) {
    out.pop()
  }
  return out
}

/** header/footer top-level table → one paragraph per row, cells as width-proportioned columns */
function hfTableRowParagraphs(tbl: XNode, ctx: BuildContext): HfParagraph[] {
  const grid = findChild(tbl, 'w:tblGrid')
  const gridCols = grid
    ? findChildren(grid, 'w:gridCol').map((g) => Number(attrsOf(g)['w:w']) || 0)
    : []
  const out: HfParagraph[] = []
  for (const tr of childrenThroughSdt(tbl, 'w:tr')) {
    const tcs = childrenThroughSdt(tr, 'w:tc')
    const widths = tcs.map((tc) => {
      const a = attrsOf(findChild(findChild(tc, 'w:tcPr') ?? {}, 'w:tcW') ?? {})
      const v = Number(a['w:w'])
      return a['w:type'] !== 'pct' && Number.isFinite(v) && v > 0 ? v : 0
    })
    if (widths.some((w) => w <= 0) && gridCols.some((w) => w > 0)) {
      let col = 0
      tcs.forEach((tc, i) => {
        const span =
          Number(attrsOf(findChild(findChild(tc, 'w:tcPr') ?? {}, 'w:gridSpan') ?? {})['w:val']) ||
          1
        widths[i] = gridCols.slice(col, col + span).reduce((s, w) => s + w, 0)
        col += span
      })
    }
    const total = widths.reduce((s, w) => s + w, 0)
    const cells = tcs.map((tc, i) => {
      const content = hfCellContent(tc, ctx)
      return {
        ...content,
        ...(total > 0 && widths[i] > 0 ? { widthPct: (widths[i] / total) * 100 } : {}),
      }
    })
    // shaded-only rows still render (banner bars with no text)
    if (cells.some((c) => c.paras.some((rs) => rs.some((r) => r.text !== '' || r.image)) || c.fill))
      out.push({ runs: [], cells })
  }
  return out
}

/** Cell content in document order, nested layout tables flattened into the cell
 *  (their cell text, alignment and shading would otherwise be dropped). */
function hfCellContent(
  tc: XNode,
  ctx: BuildContext,
): { paras: Run[][]; align?: HfTableCell['align']; fill?: string } {
  const paras: Run[][] = []
  let align: HfTableCell['align']
  const shd = attrsOf(findChild(findChild(tc, 'w:tcPr') ?? {}, 'w:shd') ?? {})['w:fill']
  let fill = shd && shd !== 'auto' ? shd : undefined
  let sawNested = false
  for (const node of childrenThroughSdt(tc, ['w:p', 'w:tbl'])) {
    if (nameOf(node) === 'w:tbl') {
      sawNested = true
      for (const tr of childrenThroughSdt(node, 'w:tr')) {
        for (const inner of childrenThroughSdt(tr, 'w:tc')) {
          const c = hfCellContent(inner, ctx)
          paras.push(...c.paras)
          align ??= c.align
          fill ??= c.fill
        }
      }
      continue
    }
    // anchored pictures stay in the part-level image list (page positioning);
    // a run carrying both text and an anchored drawing keeps its text
    paras.push(
      extractRuns(node, ctx, [], [], true).flatMap((r) => {
        if (
          !r.image ||
          (!/<wp:anchor[\s>]/.test(r.image.xml) && !/position:\s*absolute/.test(r.image.xml))
        )
          return r
        const { image: _image, ...rest } = r
        return rest.text === '' ? [] : rest
      }),
    )
    const pPr = findChild(node, 'w:pPr')
    if (!align && pPr) align = extractParaFormat(pPr)?.align
  }
  // the mandatory empty paragraph after a nested table is layout noise
  if (sawNested) while (paras.length > 0 && paras[paras.length - 1].length === 0) paras.pop()
  return { paras, ...(align ? { align } : {}), ...(fill ? { fill } : {}) }
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
  const fonts = readThemeFonts(xml)
  if (fonts) {
    const eaLang = await readThemeFontLangEa(zip)
    if (eaLang) fonts.eaLang = eaLang
  }
  return { fonts, colors: readThemeColors(xml) }
}

/** settings.xml w:themeFontLang w:eastAsia */
async function readThemeFontLangEa(zip: JSZip): Promise<string | undefined> {
  const file = zip.file('word/settings.xml')
  if (!file) return undefined
  return /<w:themeFontLang\b[^>]*\bw:eastAsia="([^"]+)"/.exec(await file.async('string'))?.[1]
}

function plainText(xml: string): string {
  const texts: string[] = []
  // a space at cell boundaries keeps table text from gluing together
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<\/w:tc>/g
  let m: RegExpExecArray | null
  let pendingGap = false
  while ((m = re.exec(xml)) !== null) {
    if (m[0] === '</w:tc>') {
      pendingGap = texts.length > 0
      continue
    }
    if (pendingGap) {
      texts.push(' ')
      pendingGap = false
    }
    texts.push(m[1])
  }
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

/**
 * Numeric character references only (`&#65;` / `&#xF0B7;`), which fast-xml-parser
 * leaves alone. Text and attribute values read off the parse tree have already had
 * the five named entities resolved, so they must NOT be run through the named-entity
 * replacements again: a document whose visible text is literally `&lt;` is stored as
 * `&amp;lt;`, and a second decode would turn it into `<`.
 */
function decodeNumericCharRefs(text: string): string {
  return text.replace(
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
}

/** Full decode, for raw XML slices that never passed through the parser. */
function decodeEntities(text: string): string {
  return decodeNumericCharRefs(text)
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
  | 'imageRotDeg'
  | 'imageFlipH'
  | 'imageFlipV'
  | 'imageCrop'
  | 'imageFillRect'
>

/** a:srcRect / a:fillRect attribute (1000ths of a percent; some writers emit decimals) → fraction */
function rectFrac(tag: string, name: string): number {
  const v = parseFloat(new RegExp(`\\b${name}="(-?[\\d.]+)"`).exec(tag)?.[1] ?? '')
  return Number.isFinite(v) ? v / 100000 : 0
}

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
  // rotation/flip live on the pic's own xfrm (an anchored textbox sibling has its own wps xfrm)
  const picXfrm = /<pic:spPr[^>]*>[\s\S]*?<a:xfrm([^>]*)>/.exec(xml)?.[1]
  if (picXfrm) {
    const rot = parseInt(/\brot="(-?\d+)"/.exec(picXfrm)?.[1] ?? '', 10)
    if (Number.isFinite(rot) && rot !== 0) {
      meta.imageRotDeg = ((Math.round(rot / 60000) % 360) + 360) % 360
    }
    if (/\bflipH="(?:1|true)"/.test(picXfrm)) meta.imageFlipH = true
    if (/\bflipV="(?:1|true)"/.test(picXfrm)) meta.imageFlipV = true
  }
  // source crop (a:srcRect) and fill placement (a:stretch/a:fillRect): applied
  // by the renderer as an overflow-hidden window over a scaled/offset image
  const srcRect = /<a:srcRect\s[^>]*\/>/.exec(xml)?.[0]
  if (srcRect) {
    const crop = {
      l: rectFrac(srcRect, 'l'),
      t: rectFrac(srcRect, 't'),
      r: rectFrac(srcRect, 'r'),
      b: rectFrac(srcRect, 'b'),
    }
    if (crop.l || crop.t || crop.r || crop.b) meta.imageCrop = crop
  }
  const fillRect = /<a:stretch>\s*<a:fillRect\s[^>]*\/>/.exec(xml)?.[0]
  if (fillRect) {
    const fr = {
      l: rectFrac(fillRect, 'l'),
      t: rectFrac(fillRect, 't'),
      r: rectFrac(fillRect, 'r'),
      b: rectFrac(fillRect, 'b'),
    }
    if (fr.l || fr.t || fr.r || fr.b) meta.imageFillRect = fr
  }
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
    } else if ((posHFrom === 'margin' || posHFrom === 'page') && alignH && !alignV) {
      // mixed positioning (H aligned, V by offset): keep the horizontal preset
      // so no-wrap images at least center like Word/LO
      meta.imagePosH = alignH as ImageMeta['imagePosH']
    }
  }
  return meta
}

const EMU_PER_PT = 12700

/**
 * VML picture geometry (v:shape style) → the DrawingML ImageMeta model, so
 * legacy stamps/watermark pictures reuse the floating-image render path.
 * Only the high-frequency subset: width/height, absolute position offsets,
 * behind-text z-index, and margin-relative centering.
 */
function vmlImageMeta(xml: string): ImageMeta {
  const meta: ImageMeta = {}
  const style = /<v:shape [^>]*style="([^"]*)"/.exec(xml)?.[1] ?? ''
  const w = parseFloat(/(?:^|;)width:([\d.]+)pt/.exec(style)?.[1] ?? '')
  const h = parseFloat(/(?:^|;)height:([\d.]+)pt/.exec(style)?.[1] ?? '')
  if (w > 0) meta.imageWidthPx = Math.round((w / 72) * 96)
  if (h > 0) meta.imageHeightPx = Math.round((h / 72) * 96)
  const jc = /<w:jc w:val="([^"]+)"/.exec(xml)?.[1]
  if (jc === 'center') meta.imageAlign = 'center'
  else if (jc === 'right' || jc === 'end') meta.imageAlign = 'right'
  if (/position:absolute/.test(style)) {
    meta.imageWrap = /z-index:\s*-/.test(style) ? 'behind' : 'front'
    const mx = parseFloat(/margin-left:(-?[\d.]+)pt/.exec(style)?.[1] ?? '')
    const my = parseFloat(/margin-top:(-?[\d.]+)pt/.exec(style)?.[1] ?? '')
    if (Number.isFinite(mx)) meta.imageOffsetXEmu = Math.round(mx * EMU_PER_PT)
    if (Number.isFinite(my)) meta.imageOffsetYEmu = Math.round(my * EMU_PER_PT)
    const posH = /mso-position-horizontal:(\w+)/.exec(style)?.[1]
    const posV = /mso-position-vertical:(\w+)/.exec(style)?.[1]
    const relH = /mso-position-horizontal-relative:(\w+)/.exec(style)?.[1]
    const relV = /mso-position-vertical-relative:(\w+)/.exec(style)?.[1]
    if (
      (relH === 'margin' || relH === 'page') &&
      (relV === 'margin' || relV === 'page') &&
      (posH === 'left' || posH === 'center' || posH === 'right') &&
      (posV === 'top' || posV === 'center' || posV === 'bottom')
    ) {
      meta.imagePosH = posH
      meta.imagePosV = posV
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
  const partPath = path.replace(/^word\/\.\.\//, '')
  const file = zip.file(partPath)
  if (!file) return null
  const mime = await imagePartMime(zip, partPath)
  if (!mime) return null
  if (isMetafileMime(mime)) return metafileToDataUrl(await file.async('arraybuffer'), mime)
  if (isTiffMime(mime)) return tiffToDataUrl(await file.async('arraybuffer'))
  return `data:${mime};base64,${await file.async('base64')}`
}

/**
 * Pre-resolve blip rIds found inside w:tbl (extractCell is sync, media reads
 * are async). Scoped to tables to bound memory.
 */
async function tableBlipMedia(
  elements: BodyElement[],
  documentXml: string,
  zip: JSZip,
  rels: Map<string, RelInfo>,
): Promise<Map<string, string>> {
  const out = new Map<string, string>()
  const rIds = new Set<string>()
  for (const el of elements) {
    if (el.name !== 'w:tbl' && el.name !== 'w:sdt') continue
    let slice = documentXml.slice(el.start, el.end)
    if (el.name === 'w:sdt') {
      const from = slice.indexOf('<w:tbl')
      if (from === -1) continue
      slice = slice.slice(from, slice.lastIndexOf('</w:tbl>') + '</w:tbl>'.length)
    }
    for (const m of slice.matchAll(/<a:blip[^>]*r:(?:embed|link)="([^"]+)"/g)) rIds.add(m[1])
  }
  for (const rId of rIds) {
    const rel = rels.get(rId)
    if (!rel) continue
    if (rel.targetMode === 'External' || /^https?:\/\//i.test(rel.target)) {
      out.set(rId, rel.target)
      continue
    }
    const dataUrl = await mediaDataUrl(zip, rels, rId)
    if (dataUrl) out.set(rId, dataUrl)
  }
  return out
}

/** resolve a paragraph's blip/VML-imagedata rIds into ctx.mediaByRid so extractRuns (sync) can build image runs */
async function resolveBlipMedia(xml: string, ctx: BuildContext): Promise<void> {
  const media = ctx.mediaByRid
  if (!media) return
  const refs = [
    ...xml.matchAll(/<a:blip[^>]*r:(?:embed|link)="([^"]+)"/g),
    ...xml.matchAll(/<v:imagedata[^>]*r:id="([^"]+)"/g),
  ]
  for (const m of refs) {
    const rId = m[1]
    if (media.has(rId)) continue
    const rel = ctx.rels.get(rId)
    if (!rel) continue
    if (rel.targetMode === 'External' || /^https?:\/\//i.test(rel.target)) {
      media.set(rId, rel.target)
      continue
    }
    const dataUrl = await mediaDataUrl(ctx.zip, ctx.rels, rId)
    if (dataUrl) media.set(rId, dataUrl)
  }
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

  const path = (rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`).replace(
    /^word\/\.\.\//,
    '',
  )
  const file = ctx.zip.file(path)
  if (!file) return null
  const mime = await imagePartMime(ctx.zip, path)
  if (!mime) return null
  if (isMetafileMime(mime)) return metafileToDataUrl(await file.async('arraybuffer'), mime)
  if (isTiffMime(mime)) return tiffToDataUrl(await file.async('arraybuffer'))
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
 * Texts are ordered by walking the diagram's parent-child connections
 * (dgm:cxn srcOrd) — the file order of dgm:pt nodes is arbitrary, while
 * renderers lay shapes out in tree order.
 */
async function extractDiagramText(xml: string, ctx: BuildContext): Promise<string | null> {
  const path = relPartPath(ctx, /r:dm="([^"]+)"/.exec(xml)?.[1])
  const file = path ? ctx.zip.file(path) : null
  if (!file) return null
  const dataXml = await file.async('string')
  // node texts by modelId (pres/transition points carry no content text)
  const ptTexts = new Map<string, string>()
  for (const m of dataXml.matchAll(/<dgm:pt modelId="([^"]+)"([^>]*)>([\s\S]*?)<\/dgm:pt>/g)) {
    if (/type="(?:pres|parTrans|sibTrans)"/.test(m[2])) continue
    const s = (m[3].match(/<a:t>[^<]*<\/a:t>/g) ?? [])
      .map((t) => decodeEntities(t.slice(5, -6)))
      .join('')
      .trim()
    if (s) ptTexts.set(m[1], s)
  }
  // parent-child edges: cxns without an explicit type (default parOf); typed
  // cxns (presOf/presParOf...) are layout wiring, not the content tree
  const children = new Map<string, Array<{ ord: number; id: string }>>()
  const hasParent = new Set<string>()
  for (const m of dataXml.matchAll(/<dgm:cxn [^>]*\/?>/g)) {
    const tag = m[0]
    if (/type="(?!parOf")/.test(tag)) continue
    const src = /srcId="([^"]+)"/.exec(tag)?.[1]
    const dst = /destId="([^"]+)"/.exec(tag)?.[1]
    if (!src || !dst) continue
    const ord = parseInt(/srcOrd="(\d+)"/.exec(tag)?.[1] ?? '0', 10)
    if (!children.has(src)) children.set(src, [])
    children.get(src)!.push({ ord, id: dst })
    hasParent.add(dst)
  }
  const texts: string[] = []
  const seen = new Set<string>()
  const visit = (id: string) => {
    if (seen.has(id)) return
    seen.add(id)
    const t = ptTexts.get(id)
    if (t) texts.push(t)
    for (const c of (children.get(id) ?? []).sort((a, b) => a.ord - b.ord)) visit(c.id)
  }
  for (const root of children.keys()) if (!hasParent.has(root)) visit(root)
  // isolated points (no tree info) keep file order
  for (const [id, t] of ptTexts) if (!seen.has(id)) texts.push(t)
  return texts.length > 0 ? texts.join('\n') : null
}

/**
 * SmartArt visual degrade: Word saves the resolved layout in a companion
 * diagrams/drawingN.xml (dsp:sp shapes with absolute geometry, fills incl.
 * pictures, and text bodies). Rendering those shapes reproduces the diagram
 * without a layout engine. Returns null when the part is absent (then only
 * the text degrade shows).
 */
/**
 * Drawing canvas (a:graphicData lockedCanvas): children carry absolute EMU
 * geometry in the canvas child-coordinate space (grpSpPr chOff/chExt), mapped
 * onto the drawing's wp:extent. Geometry scales through that mapping; text
 * keeps its raw font size and overflows the scaled boxes (LO behavior — the
 * canvas is usually authored far larger than the extent it is placed at).
 */
function extractLockedCanvas(xml: string, ctx: BuildContext): DiagramDisplay | null {
  const ci = xml.indexOf('<lc:lockedCanvas')
  if (ci === -1) return null
  const end = xml.indexOf('</lc:lockedCanvas>', ci)
  if (end === -1) return null
  // display size: the nearest wp:extent before the canvas (its own inline/anchor)
  const extM = [...xml.slice(0, ci).matchAll(/<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/g)].pop()
  const extCx = extM ? parseInt(extM[1], 10) : NaN
  const extCy = extM ? parseInt(extM[2], 10) : NaN
  if (!Number.isFinite(extCx) || !Number.isFinite(extCy) || extCx <= 0 || extCy <= 0) return null
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(xml.slice(ci, end + '</lc:lockedCanvas>'.length)) as XNode[]
  } catch {
    return null
  }
  const canvasNode = parsed.find((n) => nameOf(n) === 'lc:lockedCanvas')
  if (!canvasNode) return null
  const grpXfrm = findChild(findChild(canvasNode, 'a:grpSpPr') ?? {}, 'a:xfrm')
  const emuAttr = (node: XNode | undefined, name: string, fallback: number): number => {
    const v = parseInt(attrsOf(node ?? {})[name] ?? '', 10)
    return Number.isFinite(v) ? v : fallback
  }
  const chOffX = emuAttr(findChild(grpXfrm ?? {}, 'a:chOff'), 'x', 0)
  const chOffY = emuAttr(findChild(grpXfrm ?? {}, 'a:chOff'), 'y', 0)
  const chExtCx = emuAttr(findChild(grpXfrm ?? {}, 'a:chExt'), 'cx', extCx)
  const chExtCy = emuAttr(findChild(grpXfrm ?? {}, 'a:chExt'), 'cy', extCy)
  const scaleX = chExtCx > 0 ? extCx / chExtCx : 1
  const scaleY = chExtCy > 0 ? extCy / chExtCy : 1
  const shapes: DiagramShape[] = []
  for (const child of childrenOf(canvasNode)) {
    const name = nameOf(child)
    if (name !== 'a:sp' && name !== 'a:pic') continue
    const spPr = findChild(child, 'a:spPr')
    const xfrm = findChild(spPr ?? {}, 'a:xfrm')
    const off = findChild(xfrm ?? {}, 'a:off')
    const ext = findChild(xfrm ?? {}, 'a:ext')
    if (!off || !ext) continue
    const shape: DiagramShape = {
      xPx: Math.round(((emuAttr(off, 'x', 0) - chOffX) * scaleX) / EMU_PER_PX),
      yPx: Math.round(((emuAttr(off, 'y', 0) - chOffY) * scaleY) / EMU_PER_PX),
      wPx: Math.round((emuAttr(ext, 'cx', 0) * scaleX) / EMU_PER_PX),
      hPx: Math.round((emuAttr(ext, 'cy', 0) * scaleY) / EMU_PER_PX),
    }
    if (shape.wPx <= 0 || shape.hPx <= 0) continue
    const rot = parseInt(attrsOf(xfrm!)['rot'] ?? '', 10)
    if (Number.isFinite(rot) && rot !== 0) shape.rotDeg = Math.round(rot / 60000)
    const prst = attrsOf(findChild(spPr ?? {}, 'a:prstGeom') ?? {})['prst']
    if (prst && prst !== 'rect') shape.prst = prst
    if (name === 'a:pic') {
      const rId = attrsOf(findChild(findChild(child, 'a:blipFill') ?? {}, 'a:blip') ?? {})[
        'r:embed'
      ]
      const dataUrl = rId ? ctx.mediaByRid?.get(rId) : undefined
      if (dataUrl) shape.imageDataUrl = dataUrl
    } else if (spPr && !findChild(spPr, 'a:noFill')) {
      const fill =
        colorNodeHex(findChild(spPr, 'a:solidFill'), ctx.themeColors) ??
        gradFillApproxHex(spPr, ctx.themeColors)
      if (fill) shape.fillHex = fill
    }
    const body = findChild(findChild(child, 'a:txSp') ?? {}, 'a:txBody')
    if (body) {
      const texts: string[] = []
      let sizePt: number | undefined
      let color: string | undefined
      for (const p of findChildren(body, 'a:p')) {
        const parts: string[] = []
        for (const r of findChildren(p, 'a:r')) {
          const t = findChild(r, 'a:t')
          if (t) parts.push(decodeNumericCharRefs(textOf(t)))
          const rPr = findChild(r, 'a:rPr')
          if (rPr && sizePt === undefined) {
            const sz = parseInt(attrsOf(rPr)['sz'] ?? '', 10)
            // raw canvas font size (hundredths of a point), deliberately unscaled
            if (Number.isFinite(sz) && sz > 0) sizePt = Math.round(sz / 100)
            const c = colorNodeHex(findChild(rPr, 'a:solidFill'), ctx.themeColors)
            if (c) color = c
          }
        }
        if (parts.join('').trim() !== '') texts.push(parts.join(''))
      }
      if (texts.length > 0) {
        shape.texts = texts
        if (sizePt) shape.fontSizePt = sizePt
        if (color) shape.textColorHex = color
      }
    }
    shapes.push(shape)
  }
  if (shapes.length === 0) return null
  // LO parity for severely overflowing canvas text (the canvas is authored
  // several times larger than its placed extent, so raw-size text dwarfs the
  // scaled boxes): LO stacks the text shapes as columns, each next column
  // starting about halfway down the previous column's text run. Reproducing
  // that stacking keeps the glyph reading order identical to LO's render.
  const textShapes = shapes.filter((s) => s.texts?.length && s.fontSizePt)
  const colGeom = (s: DiagramShape): { lines: number; pitchPx: number } => {
    const fontPx = (s.fontSizePt! * 96) / 72
    const charsPerLine = Math.max(1, Math.floor(s.wPx / (fontPx * 0.72)))
    const chars = (s.texts ?? []).join('').length
    return { lines: Math.ceil(chars / charsPerLine), pitchPx: fontPx * 1.2 }
  }
  const overflowing = textShapes.filter((s) => {
    const g = colGeom(s)
    return g.lines * g.pitchPx > 2 * s.hPx
  })
  if (overflowing.length > 0) {
    // the first overflowing column anchors at the canvas top (LO ignores the
    // scaled box offset once the text no longer fits the box)
    overflowing[0].yPx = 0
    for (let i = 1; i < overflowing.length; i++) {
      const prev = overflowing[i - 1]
      const g = colGeom(prev)
      overflowing[i].yPx = Math.round(prev.yPx + Math.ceil(g.lines / 2) * g.pitchPx - 23)
    }
    // single-letter-per-line columns explode into per-letter shapes ordered
    // top-down: the exported PDF then emits glyphs in painter order like LO,
    // and text extraction reads the interleaved columns identically
    for (const s of overflowing) {
      const g = colGeom(s)
      const chars = [...(s.texts ?? []).join('')]
      if (g.lines < chars.length) continue // wraps more than one char per line
      const at = shapes.indexOf(s)
      if (at === -1) continue
      const letters: DiagramShape[] = chars.map((ch, i) => ({
        xPx: s.xPx,
        yPx: Math.round(s.yPx + i * g.pitchPx),
        wPx: s.wPx,
        hPx: Math.ceil(g.pitchPx),
        texts: [ch],
        ...(s.fontSizePt ? { fontSizePt: s.fontSizePt } : {}),
        ...(s.textColorHex ? { textColorHex: s.textColorHex } : {}),
      }))
      shapes.splice(at, 1, ...letters)
    }
    shapes.sort((a, b) => a.yPx - b.yPx)
  }
  return {
    widthPx: Math.round(extCx / EMU_PER_PX),
    heightPx: Math.round(extCy / EMU_PER_PX),
    shapes,
    canvas: true,
  }
}

async function extractDiagramDrawing(
  xml: string,
  ctx: BuildContext,
): Promise<DiagramDisplay | null> {
  const dmPath = relPartPath(ctx, /r:dm="([^"]+)"/.exec(xml)?.[1])
  if (!dmPath) return null
  const drawingPath = dmPath.replace(/data(\d*)\.xml$/, 'drawing$1.xml')
  const file = drawingPath !== dmPath ? ctx.zip.file(drawingPath) : null
  if (!file) return null
  const extent = /<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(xml)
  const widthPx = extent ? Math.round(parseInt(extent[1], 10) / EMU_PER_PX) : 0
  const heightPx = extent ? Math.round(parseInt(extent[2], 10) / EMU_PER_PX) : 0
  if (!widthPx || !heightPx) return null
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(await file.async('string')) as XNode[]
  } catch {
    return null
  }
  // picture fills resolve through the drawing part's own rels (../media/...)
  const relsPath = drawingPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const rels = await parseRels(ctx.zip, relsPath)
  const dir = drawingPath.replace(/[^/]+$/, '')
  const mediaOf = async (rId: string): Promise<string | null> => {
    const rel = rels.get(rId)
    if (!rel || rel.targetMode === 'External') return null
    const parts: string[] = []
    for (const seg of (rel.target.startsWith('/') ? rel.target.slice(1) : dir + rel.target).split(
      '/',
    )) {
      if (seg === '..') parts.pop()
      else if (seg !== '.') parts.push(seg)
    }
    const path = parts.join('/')
    const f = ctx.zip.file(path)
    if (!f) return null
    const mime = await imagePartMime(ctx.zip, path)
    if (!mime) return null
    if (isMetafileMime(mime)) return metafileToDataUrl(await f.async('arraybuffer'), mime)
    return `data:${mime};base64,${await f.async('base64')}`
  }
  const sps: XNode[] = []
  collectNodes(parsed, 'dsp:sp', sps)
  const shapes: DiagramShape[] = []
  for (const sp of sps) {
    const spPr = findChild(sp, 'dsp:spPr')
    if (!spPr) continue
    const xfrm = findChild(spPr, 'a:xfrm')
    const off = xfrm ? findChild(xfrm, 'a:off') : undefined
    const ext = xfrm ? findChild(xfrm, 'a:ext') : undefined
    if (!off || !ext) continue
    const shape: DiagramShape = {
      xPx: Math.round(parseInt(attrsOf(off)['x'] ?? '0', 10) / EMU_PER_PX),
      yPx: Math.round(parseInt(attrsOf(off)['y'] ?? '0', 10) / EMU_PER_PX),
      wPx: Math.round(parseInt(attrsOf(ext)['cx'] ?? '0', 10) / EMU_PER_PX),
      hPx: Math.round(parseInt(attrsOf(ext)['cy'] ?? '0', 10) / EMU_PER_PX),
    }
    if (shape.wPx <= 0 || shape.hPx <= 0) continue
    const rot = parseInt(attrsOf(xfrm!)['rot'] ?? '', 10)
    if (Number.isFinite(rot) && rot !== 0) shape.rotDeg = Math.round(rot / 60000)
    const prst = attrsOf(findChild(spPr, 'a:prstGeom') ?? {})['prst']
    if (prst) shape.prst = prst
    const blipFill = findChild(spPr, 'a:blipFill')
    if (blipFill) {
      const rId = attrsOf(findChild(blipFill, 'a:blip') ?? {})['r:embed']
      const dataUrl = rId ? await mediaOf(rId) : null
      if (dataUrl) {
        shape.imageDataUrl = dataUrl
        const fr = /<a:fillRect\s[^>]*\/>/.exec(serializeXNode(blipFill))?.[0]
        if (fr) {
          const rect = {
            l: rectFrac(fr, 'l'),
            t: rectFrac(fr, 't'),
            r: rectFrac(fr, 'r'),
            b: rectFrac(fr, 'b'),
          }
          if (rect.l || rect.t || rect.r || rect.b) shape.fillRect = rect
        }
      }
    } else {
      const solid = findChild(spPr, 'a:solidFill')
      const srgb = solid ? attrsOf(findChild(solid, 'a:srgbClr') ?? {})['val'] : undefined
      const scheme = solid ? attrsOf(findChild(solid, 'a:schemeClr') ?? {})['val'] : undefined
      if (srgb) shape.fillHex = srgb
      else if (scheme) {
        const theme = ctx.themeColors as Record<string, string> | null | undefined
        shape.fillHex = (theme && theme[scheme]) || '9AB5E4'
      }
    }
    const body = findChild(sp, 'dsp:txBody')
    if (body) {
      const texts: string[] = []
      let sizePt: number | undefined
      let color: string | undefined
      for (const p of findChildren(body, 'a:p')) {
        const parts: string[] = []
        for (const r of findChildren(p, 'a:r')) {
          const t = findChild(r, 'a:t')
          if (t) parts.push(decodeNumericCharRefs(textOf(t)))
          const rPr = findChild(r, 'a:rPr')
          if (rPr && sizePt === undefined) {
            const sz = parseInt(attrsOf(rPr)['sz'] ?? '', 10)
            if (Number.isFinite(sz) && sz > 0) sizePt = Math.round(sz / 100)
            const fill = findChild(rPr, 'a:solidFill')
            const c = fill ? attrsOf(findChild(fill, 'a:srgbClr') ?? {})['val'] : undefined
            if (c) color = c
          }
        }
        if (parts.join('').trim() !== '') texts.push(parts.join(''))
      }
      if (texts.length > 0) {
        shape.texts = texts
        if (sizePt) shape.fontSizePt = sizePt
        if (color) shape.textColorHex = color
      }
    }
    shapes.push(shape)
  }
  return shapes.length > 0 ? { widthPx, heightPx, shapes } : null
}

/**
 * OLE embed degrade: the original packages a preview picture
 * (v:imagedata) and declares its kind (o:OLEObject ProgID) — surface both
 * instead of a bare type label.
 */
async function oleDisplay(
  xml: string,
  ctx: BuildContext,
): Promise<
  Pick<Block, 'imageDataUrl' | 'oleProgId' | 'imageWidthPx' | 'imageHeightPx' | 'imageAlign'>
> {
  const out: Pick<
    Block,
    'imageDataUrl' | 'oleProgId' | 'imageWidthPx' | 'imageHeightPx' | 'imageAlign'
  > = {}
  const progId = /<o:OLEObject[^>]*ProgID="([^"]+)"/.exec(xml)?.[1]
  if (progId) out.oleProgId = progId
  const path = relPartPath(ctx, /<v:imagedata[^>]*r:id="([^"]+)"/.exec(xml)?.[1])
  const file = path ? ctx.zip.file(path) : null
  if (file && path) {
    const mime = await imagePartMime(ctx.zip, path)
    if (isMetafileMime(mime)) {
      const converted = await metafileToDataUrl(await file.async('arraybuffer'), mime)
      if (converted) out.imageDataUrl = converted
    } else if (mime) {
      out.imageDataUrl = `data:${mime};base64,${await file.async('base64')}`
    }
  }
  // Word draws the preview at the declared size — v:shape style (pt), falling
  // back to w:object dxaOrig/dyaOrig (twips). Without it the metafile preview's
  // intrinsic pixels (2x dpiScale) blow up to the full content width.
  const style = /<v:shape [^>]*style="([^"]*)"/.exec(xml)?.[1] ?? ''
  const wPt = parseFloat(/(?:^|;)width:([\d.]+)pt/.exec(style)?.[1] ?? '')
  const hPt = parseFloat(/(?:^|;)height:([\d.]+)pt/.exec(style)?.[1] ?? '')
  const objAttrs = /<w:object\b[^>]*>/.exec(xml)?.[0] ?? ''
  const wTw = parseInt(/w:dxaOrig="(\d+)"/.exec(objAttrs)?.[1] ?? '', 10)
  const hTw = parseInt(/w:dyaOrig="(\d+)"/.exec(objAttrs)?.[1] ?? '', 10)
  const w = wPt > 0 ? (wPt / 72) * 96 : wTw > 0 ? wTw / 15 : 0
  const h = hPt > 0 ? (hPt / 72) * 96 : hTw > 0 ? hTw / 15 : 0
  if (w > 0) out.imageWidthPx = Math.round(w)
  if (h > 0) out.imageHeightPx = Math.round(h)
  const jc = /<w:jc w:val="([^"]+)"/.exec(xml)?.[1]
  if (jc === 'center') out.imageAlign = 'center'
  else if (jc === 'right' || jc === 'end') out.imageAlign = 'right'
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
    // chartex (cx:) parts are display-only degrades: the data-edit patcher
    // speaks classic c: syntax, so keeping them out makes edits no-ops
    // instead of corruption
    if (!partXml.includes('<cx:chartSpace')) ctx.chartParts[path] = partXml
    const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml)
    const cx = extent ? parseInt(extent[1]!, 10) : NaN
    const cy = extent ? parseInt(extent[2]!, 10) : NaN
    if (Number.isFinite(cx) && cx > 0) display.widthPx = Math.round(cx / EMU_PER_PX)
    if (Number.isFinite(cy) && cy > 0) display.heightPx = Math.round(cy / EMU_PER_PX)
  }
  return display
}

/** Word's substitution face when the East Asian font slot is empty, by w:lang w:eastAsia */
const EA_LANG_DEFAULT_FONT: Record<string, string> = {
  // Word probe + fontTable: ko-KR empty EA slot substitutes Malgun Gothic, not Batang
  ko: 'Malgun Gothic',
  'ko-kr': 'Malgun Gothic',
  ja: 'MS Mincho',
  'ja-jp': 'MS Mincho',
  'zh-cn': 'SimSun',
  'zh-tw': 'PMingLiU',
  'zh-hk': 'PMingLiU',
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
    // docDefaults keeps the lang-based backfill below for the empty-slot case
    if (ddRf.eastAsia && !ddRf.eaSlotEmpty) dd.eastAsiaFont = ddRf.eastAsia
    // Empty EA slot + explicit w:lang w:eastAsia: Word substitutes the locale's
    // default face (e.g. ko-KR theme with <a:ea typeface=""/> renders Malgun Gothic)
    const eaLang = rPr ? attrsOf(findChild(rPr, 'w:lang') ?? {})['w:eastAsia'] : undefined
    if (!dd.eastAsiaFont && eaLang) {
      const eaDefault = EA_LANG_DEFAULT_FONT[eaLang.toLowerCase()]
      if (eaDefault) {
        dd.eastAsiaFont = eaDefault
        if (ddRf.eaSlotEmpty) dd.eaSlotEmpty = true
      }
    }
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
        if (numId === '0') {
          numPr = 'none'
        } else if (numId) {
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
    'boldCs',
    'italicCs',
    'sizeCsHalfPoints',
    'rtl',
    'underline',
    'strike',
    'font',
    'fontAscii',
    'csFont',
    'caps',
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
  const merged: TableStyleDisplay = { ...parent, ...(child ?? {}) }
  const DEEP = ['wholeTable', 'firstRow', 'firstCol', 'lastCol', 'lastRow', 'paraSpacing'] as const
  for (const key of DEEP) {
    if (parent[key] || child?.[key]) {
      merged[key] = { ...(parent[key] ?? {}), ...(child?.[key] ?? {}) } as never
    }
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
  const styleRPr = findChild(styleNode, 'w:rPr')
  if (styleRPr) {
    const wholeTable: NonNullable<TableStyleDisplay['wholeTable']> = {}
    const color = colorFrom(styleRPr, theme)
    if (color) wholeTable.color = color
    if (boolProp(styleRPr, 'w:b')) wholeTable.bold = true
    if (Object.keys(wholeTable).length > 0) display.wholeTable = wholeTable
  }
  for (const cond of findChildren(styleNode, 'w:tblStylePr')) {
    const type = attrsOf(cond)['w:type']
    const tcPr = findChild(cond, 'w:tcPr')
    const fill = shdFill(tcPr)
    if (type === 'firstRow' || type === 'firstCol' || type === 'lastCol' || type === 'lastRow') {
      const rPr = findChild(cond, 'w:rPr')
      const fmt: NonNullable<TableStyleDisplay['firstRow']> = {}
      if (fill) fmt.fill = fill
      if (rPr && boolProp(rPr, 'w:b')) fmt.bold = true
      const color = colorFrom(rPr, theme)
      if (color) fmt.color = color
      if (Object.keys(fmt).length > 0) display[type] = fmt
    } else if (type === 'band1Horz' && fill) {
      display.band1Fill = fill
    } else if (type === 'band2Horz' && fill) {
      display.band2Fill = fill
    }
  }
  const styleTblPr = findChild(styleNode, 'w:tblPr')
  const borders = mergedBorderLinesOf(styleTblPr, 'w:tblBorders', true)
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

/**
 * Style-chain run props under Word's rtl selection (probed, Word for Mac 2026-08):
 * rtl runs read only the Cs twins (w:bCs/w:iCs/w:szCs), non-rtl runs read only the
 * base props — no cross-fallback. Pass the run's cs flag; callers without run
 * context (style gallery previews, caret defaults) pass false (= non-rtl).
 */
export function styleRunFormat(
  display: StyleDisplay | undefined,
  cs: boolean,
): Pick<StyleDisplay, 'bold' | 'italic' | 'sizeHalfPoints'> {
  if (!display) return {}
  return cs
    ? { bold: display.boldCs, italic: display.italicCs, sizeHalfPoints: display.sizeCsHalfPoints }
    : { bold: display.bold, italic: display.italic, sizeHalfPoints: display.sizeHalfPoints }
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
    // Cs twins carried separately: the consuming run picks the set by its rtl flag
    // (styleRunFormat); consumers without run context read the base props (= non-rtl)
    const boldCs = onOffOf(rPr, 'w:bCs')
    if (boldCs !== undefined) display.boldCs = boldCs
    const italicCs = onOffOf(rPr, 'w:iCs')
    if (italicCs !== undefined) display.italicCs = italicCs
    const szCs = attrsOf(findChild(rPr, 'w:szCs') ?? {})['w:val']
    if (szCs) display.sizeCsHalfPoints = parseInt(szCs, 10) || undefined
    const rtl = onOffOf(rPr, 'w:rtl')
    if (rtl !== undefined) display.rtl = rtl
    const u = attrsOf(findChild(rPr, 'w:u') ?? {})['w:val']
    if (u) display.underline = u !== 'none'
    const strike = onOffOf(rPr, 'w:strike')
    if (strike !== undefined) display.strike = strike
    const rf = themedRFonts(attrsOf(findChild(rPr, 'w:rFonts') ?? {}), themeFonts)
    const font = rf.eastAsia ?? rf.ascii ?? rf.hAnsi
    const fontAscii = rf.ascii ?? rf.hAnsi
    if (fontAscii) display.fontAscii = fontAscii
    if (rf.cs) display.csFont = rf.cs
    if (font) display.font = font
    if (rf.eaSlotEmpty && font && font === rf.eastAsia) display.eaSlotEmpty = true
    const spc = parseInt(attrsOf(findChild(rPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
    if (spc) display.charSpacingTwips = spc
    if (onOffOf(rPr, 'w:caps')) display.caps = 'all'
    else if (onOffOf(rPr, 'w:smallCaps')) display.caps = 'small'
    // w:specVanish marks a style separator, not hidden text
    const vanish = onOffOf(rPr, 'w:vanish')
    if (vanish !== undefined && onOffOf(rPr, 'w:specVanish') !== true) display.vanish = vanish
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
    if (boolProp(pPr, 'w:pageBreakBefore')) display.pageBreakBefore = true
    if (boolProp(pPr, 'w:contextualSpacing')) display.contextualSpacing = true
    const autoSpace = autoSpaceOf(pPr)
    if (autoSpace !== undefined) display.autoSpace = autoSpace
    const jc = attrsOf(findChild(pPr, 'w:jc') ?? {})['w:val']
    if (jc === 'center' || jc === 'right' || jc === 'left' || jc === 'justify') display.align = jc
    else if (jc === 'both' || jc === 'distribute') display.align = 'justify'
    const stops = tabStopsOf(pPr)
    if (stops) display.tabStops = stops
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

/** w:numFmt of a w:lvl, including w14 custom formats hidden in mc:AlternateContent
 *  (mc:Choice carries val="custom" + w:format; mc:Fallback the standard substitute) */
function numFmtOfLevel(lvlNode: XNode): { numFmt?: string; customFormat?: string } {
  const direct = findChild(lvlNode, 'w:numFmt')
  if (direct) return { numFmt: attrsOf(direct)['w:val'] }
  const alt = childrenOf(lvlNode).find((c) => nameOf(c)?.endsWith(':AlternateContent'))
  if (!alt) return {}
  const pick = (local: string) => childrenOf(alt).find((c) => nameOf(c)?.endsWith(`:${local}`))
  const choice = attrsOf(findChild(pick('Choice') ?? {}, 'w:numFmt') ?? {})
  const format = choice['w:format']
  if (choice['w:val'] === 'custom' && format && customEnumItems(format))
    return { numFmt: 'custom', customFormat: format }
  if (choice['w:val'] && choice['w:val'] !== 'custom') return { numFmt: choice['w:val'] }
  return { numFmt: attrsOf(findChild(pick('Fallback') ?? {}, 'w:numFmt') ?? {})['w:val'] }
}

function parseNumberingLevel(lvlNode: XNode): NumberingLevel {
  // ECMA-376: a w:lvl without w:start starts at 0 (Word renders "0.")
  const start = parseInt(attrsOf(findChild(lvlNode, 'w:start') ?? {})['w:val'] ?? '0', 10)
  const { numFmt, customFormat } = numFmtOfLevel(lvlNode)
  const level: NumberingLevel = {
    numFmt: numFmt ?? 'decimal',
    lvlText: decodeNumericCharRefs(attrsOf(findChild(lvlNode, 'w:lvlText') ?? {})['w:val'] ?? ''),
    start: Number.isFinite(start) ? start : 0,
  }
  if (customFormat) level.customFormat = customFormat
  const suff = attrsOf(findChild(lvlNode, 'w:suff') ?? {})['w:val']
  if (suff === 'space' || suff === 'nothing' || suff === 'tab') level.suff = suff
  const lvlPPr = findChild(lvlNode, 'w:pPr')
  const ind = lvlPPr ? findChild(lvlPPr, 'w:ind') : undefined
  if (ind) {
    const attrs = attrsOf(ind)
    const left = parseInt(attrs['w:left'] ?? attrs['w:start'] ?? '', 10)
    if (left > 0) level.indentLeft = left
    const hanging = parseInt(attrs['w:hanging'] ?? '', 10)
    if (hanging > 0) level.hanging = hanging
    const firstLine = parseInt(attrs['w:firstLine'] ?? '', 10)
    if (!level.hanging && firstLine > 0) level.firstLine = firstLine
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
  const numStyleLinks = new Map<string, string>()
  const styleLinkAbs = new Map<string, string>()
  for (const abs of findChildren(root, 'w:abstractNum')) {
    const absId = attrsOf(abs)['w:abstractNumId']
    if (!absId) continue
    const levels: Record<number, NumberingLevel> = {}
    for (const lvl of findChildren(abs, 'w:lvl')) {
      const ilvl = parseInt(attrsOf(lvl)['w:ilvl'] ?? '', 10)
      if (Number.isFinite(ilvl)) levels[ilvl] = parseNumberingLevel(lvl)
    }
    absLevels.set(absId, levels)
    const numStyleLink = attrsOf(findChild(abs, 'w:numStyleLink') ?? {})['w:val']
    if (numStyleLink) numStyleLinks.set(absId, numStyleLink)
    const styleLink = attrsOf(findChild(abs, 'w:styleLink') ?? {})['w:val']
    if (styleLink) styleLinkAbs.set(styleLink, absId)
  }
  // w:numStyleLink indirection: the abstractNum is a reference to a numbering style;
  // the real levels live on the abstractNum carrying the matching w:styleLink
  for (const absId of numStyleLinks.keys()) {
    const seen = new Set([absId])
    let target = absId
    for (;;) {
      const styleId = numStyleLinks.get(target)
      const next = styleId !== undefined ? styleLinkAbs.get(styleId) : undefined
      if (next === undefined || seen.has(next)) break
      seen.add(next)
      target = next
    }
    if (target !== absId) {
      absLevels.set(absId, { ...absLevels.get(target), ...absLevels.get(absId) })
    }
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
