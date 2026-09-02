import JSZip from 'jszip'
import { parseCustGeom } from '@genoffice/pptx-engine/custgeom'
import { parseChartPartXml } from './chart'
import { findInkRuns, stripInkRuns } from './ink'
import { isMetafileMime, metafileToDataUrl } from './metafile'
import { isTiffMime, tiffToDataUrl } from './tiff'
import { ommlFragmentsOf, ommlToLatex, ommlToMathML } from './math'
import { splitXmlChildren } from './generate'
import { NOTE_PART_PATH, parseNotesXml } from './notes'
import { scanBody, type BodyElement } from './scan'
import { sectionSettingsFromXml, xmlFlagOn } from './section'
import { findSourcesPart, parseSourcesXml } from './sources'
import { decodeSymbolChar, decodeSymbolText } from './symbol-fonts'
import { FONT_TABLE_PART_PATH, parseFontTable } from './font-table'
import { DEFAULT_THEME_COLORS, THEME_PART_PATH, readThemeColors, readThemeFonts } from './theme'
import { PAGE_MARK, TOTAL_PAGES_MARK } from './types'
import { assertZipWithinLimits, loadDocxZip } from './zip-load'
import type {
  Block,
  ChartDisplay,
  DiagramDisplay,
  DiagramShape,
  DocDefaults,
  InkInfo,
  NoteInfo,
  HfImage,
  HfParagraph,
  HfTableCell,
  HfPartInfo,
  NumberingDef,
  ParaFormat,
  ParsedDoc,
  CellMargins,
  RevisionInfo,
  Run,
  SectionSettings,
  SourceInfo,
  StrayIndent,
  StyleInfo,
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
import {
  applyBalancedDbcsSpacing,
  applyProtectedLeadingBreaks,
  blockRunGroups,
  normalizeImageZOrders,
} from './parse-block-passes'
import { applyTocEntryNumbers, fieldDisplayOf, fieldLabel, tocLevelOf } from './parse-fields'
import {
  EMU_PER_PT,
  EMU_PER_PX,
  colorFrom,
  decodeEntities,
  decodeNumericCharRefs,
  lineTwipsOf,
  mathTokens,
  onOffOf,
  plainText,
  stripHash,
  type RelInfo,
} from './parse-xml-text'
import {
  EMU_PER_TWIP,
  IDENTITY_CTM,
  LINE_PRSTS,
  LINE_PRSTS_RE,
  MIN_WRAP_SLIVER_EMU,
  colorNodeHex,
  composeGroupCtm,
  drawingAnchorMeta,
  gradFillApproxHex,
  lineBoxOf,
  resolveAnchorPagePos,
  topLevelDrawings,
  w14TextFillHex,
  type DrawingAnchorMeta,
  type ExtractTextboxOpts,
  type GroupCtm,
  type ResolvedAnchorPos,
} from './parse-drawing-geometry'
import {
  IMAGE_RUN_CHILDREN,
  JC_ALIGN,
  SIMPLE_INLINE_FIELD_RE,
  activeCharIndents,
  autoSpaceOf,
  bookmarkNamesOf,
  cellMarginsOf,
  charIndentsOf,
  checkboxStateOf,
  collectNodes,
  collectTopNodes,
  convertibleHyperlink,
  crossParaCommentMarkers,
  emptyParaMarkFont,
  emptyParaSizeHalfPoints,
  hasLayoutRunContent,
  hostPageBreak,
  isInvisibleEmptyShape,
  isInvisibleVmlPict,
  isThinRule,
  mergeCharIndents,
  mergeRuns,
  mergedBorderLinesOf,
  onlyOleFields,
  onlyXeFields,
  partXmlSpacePreserve,
  ptabDisplayStops,
  rawPPrOf,
  resolveCharIndents,
  type CharUnits,
  rubyFragmentsOf,
  rubyPartText,
  ruleDisplayOf,
  shdDisplayFill,
  splitImageRun,
  staysVanished,
  stripTextboxes,
  tabStopsOf,
  tableLookOf,
  themedRFonts,
  txbxHasStructuredContent,
} from './parse-props'
import {
  parseComments,
  parseNumbering,
  parseProtection,
  parseRels,
  parseRemovePersonalInfo,
  parseWriteProtection,
  resolveMainDocumentPath,
} from './parse-package'
import { parseSdtBlock, sdtMeta, sdtTableXml, splitSdtParts } from './parse-sdt'
import { parseStyles } from './parse-styles'
import {
  VML_PICT_RID_RE,
  VML_WORDART_RE,
  vmlColorHex,
  vmlCoordPx,
  vmlGroupScale,
  vmlPathToNormD,
  vmlShapeDimPx,
  vmlWordArtBox,
  type VmlGroupScale,
  type VmlOrigin,
} from './parse-vml'

export { resolveMainDocumentPath } from './parse-package'
export { styleRunFormat } from './parse-styles'

export { assertZipWithinLimits }

/** w:br w:type → run-text control char (\f page, \v column; else soft \n) */
const BREAK_CHAR: Record<string, string> = { page: '\f', column: '\v' }

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

export interface ParseExtras {
  /** ranges of top-level body elements, aligned with docxIndex */
  elements: BodyElement[]
  /** original XML of chart parts referenced by chart blocks (partPath -> xml) */
  chartParts: Record<string, string>
}

export async function parseDocx(bytes: Uint8Array): Promise<ParsedDoc & { extras: ParseExtras }> {
  const zip = await loadDocxZip(bytes)
  assertZipWithinLimits(zip)
  const docPath = await resolveMainDocumentPath(zip)
  if (!docPath) {
    const mime = (await zip.file('mimetype')?.async('string'))?.trim()
    if (mime?.startsWith('application/vnd.oasis.opendocument'))
      throw new Error(`OpenDocument file (${mime}), not OOXML — save as .docx to open`)
    throw new Error('not a docx: missing word/document.xml')
  }
  const documentXml = await zip.file(docPath)!.async('string')

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

  const rels = await parseRels(zip, docPath.replace(/([^/]+)$/, '_rels/$1.rels'))
  const { formats: numFormats, defs: numbering } = await parseNumbering(zip)
  const comments = await parseComments(zip)
  const protection = await parseProtection(zip)
  const writeProtection = await parseWriteProtection(zip)
  const removePersonalInfo = await parseRemovePersonalInfo(zip)
  const footnotes = await parseNotesPart(zip, 'footnote')
  const endnotes = await parseNotesPart(zip, 'endnote')
  const sources = await parseSources(zip)
  const fontTableFile = zip.file(FONT_TABLE_PART_PATH)
  const fontTable = fontTableFile ? parseFontTable(await fontTableFile.async('string')) : []

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
  const externalTxbxByRid = await externalTxbxParts(documentXml, zip, rels)
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
    externalTxbxByRid,
    referenceOnlyComments,
    sectionAt,
    docDefaults,
    defaultParaStyle: [...styles.values()].find((s) => s.type === 'paragraph' && s.isDefault),
    xmlSpacePreserve: partXmlSpacePreserve(documentXml, 'w:document'),
    // explicit breaks in any attribute order, plus Word's rendered-page hint
    // (catches natural pages in single-section docs); a false positive only
    // turns page-pinning off, which is the conservative direction
    firstPageBreakAt: (() => {
      const br = documentXml.search(
        /<w:br [^>]*w:type="page"|<w:pageBreakBefore[^>]*\/>|<w:lastRenderedPageBreak[^>]*\/>/,
      )
      const sect = documentXml.indexOf('</w:sectPr>')
      const cands = [br, sect].filter((i) => i !== -1)
      return cands.length > 0 ? Math.min(...cands) : undefined
    })(),
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
  normalizeImageZOrders(blocks)
  applyProtectedLeadingBreaks(blocks)

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
  const layoutSettings = await parseLayoutSettings(zip)
  const hfParts = await parseAllHfParts(zip, rels, styles, theme.colors)
  if (layoutSettings.balanceDbcsSpacing) {
    const hfGroups: Array<Run[] | undefined> = []
    for (const part of [header, footer, headerFirst, footerFirst, headerEven, footerEven]) {
      for (const para of part?.paras ?? []) {
        hfGroups.push(para.runs)
        for (const cell of para.cells ?? []) hfGroups.push(...cell.paras)
      }
    }
    for (const part of Object.values(hfParts ?? {})) {
      for (const para of part.paras) {
        hfGroups.push(para.runs)
        for (const cell of para.cells ?? []) hfGroups.push(...cell.paras)
      }
    }
    applyBalancedDbcsSpacing([...blockRunGroups(blocks), ...hfGroups])
  }

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
    writeProtection,
    removePersonalInfo,
    footnotes,
    endnotes,
    sources,
    inks,
    themeFonts: theme.fonts,
    themeColors: theme.colors,
    ...(fontTable.length > 0 ? { fontTable } : {}),
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
    ...layoutSettings,
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
  /** live palette for w:themeColor resolution (built-in Office palette when the doc has no theme part) */
  themeColors?: ThemeColors | null
  /** theme font scheme for w:asciiTheme/... resolution */
  themeFonts?: ThemeFonts | null
  /** pre-resolved a:blip rId -> data/external URL for pictures inside w:tbl (extractCell is sync, media reads are async) */
  mediaByRid?: Map<string, string>
  /** pre-fetched external textbox parts (wps:txbx r:txbx -> word/txbx*.xml content; extractTextboxes is sync) */
  externalTxbxByRid?: Map<string, string>
  /** comment ids anchored only by a bare w:commentReference (no range markers anywhere in document.xml, LibreOffice style) */
  referenceOnlyComments?: Set<string>
  /** page geometry of the section governing a document.xml byte offset (page/margin-anchored drawings) */
  sectionAt?: (docOffset: number) => SectionSettings
  /** styles.xml w:docDefaults (the Normal font size behind character-unit indents) */
  docDefaults?: DocDefaults
  /** the w:default paragraph style (Normal): its size is the unit of leftChars/rightChars */
  defaultParaStyle?: StyleInfo
  /** byte offset of the first explicit page/section break; content before it
   *  is on the first page (page-pinned cover art placement) */
  firstPageBreakAt?: number
  /** part root declares xml:space="preserve" (inherited XML scope; PDF converters
   *  rely on it instead of per-w:t attributes) */
  xmlSpacePreserve?: boolean
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
    return {
      ...base,
      type: 'table',
      ...tableSummary(xml),
      table: extractTable(xml, ctx, el.start),
    }
  }

  // --- SDT (structured document tag): extract sdtContent paragraph as editable ---
  if (el.name === 'w:sdt') {
    // sdtContent that starts with a table (research-report templates wrap whole
    // tables in content controls): display as a real table. Untouched
    // it saves byte-identical; cell-text edits patch inside the sdt shell.
    const tblXml = sdtTableXml(xml)
    if (tblXml) {
      return {
        ...base,
        type: 'table',
        ...tableSummary(tblXml),
        table: extractTable(tblXml, ctx, el.start),
      }
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
    if (!sdtPreview.trim() && !xml.includes('<w:drawing') && !/<w:pict[\s>]/.test(xml)) {
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
    xml.includes('<mc:Fallback')
      ? xml.replace(/<mc:Fallback[^>]*>[\s\S]*?<\/mc:Fallback>/g, '')
      : xml,
  )

  // Paragraph: certain constructs are protected as whole passthrough blocks.
  // Regenerating them would silently drop structure (section breaks, fields,
  // footnote anchors...), which is exactly the kind of damage patch-save exists
  // to prevent.
  // Only a content-less section-break paragraph is protected: one with visible
  // text renders as a normal paragraph (Word shows its content on the section's
  // last page); its w:sectPr rides along in rawPPr, which mergePPrFormat keeps
  // verbatim on edits, so the section survives regeneration.
  if (detect.includes('<w:sectPr') && !plainText(detect).trim()) {
    return {
      ...base,
      type: 'passthrough',
      label: 'Section break paragraph',
      previewText: '',
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
        fieldDisplay: fieldDisplayOf(xml, ctx.styles),
        ...(pStyle ? { styleId: pStyle } : {}),
      }
    }
  }
  // TOC-styled paragraphs are part of a TOC field result even when they carry
  // no field chars themselves (entries with literal page numbers). Editing
  // them individually would corrupt the field, so they stay protected.
  // Word writes styleIds "TOC1".."TOC9"; Pages exports "TOC 1"/"TOC 2" (with
  // space); html2docx-style exports use opaque ids with the name "toc 1".
  const tocStyleId = /<w:pStyle w:val="([^"]+)"/.exec(xml)?.[1]
  if (tocStyleId && tocLevelOf(tocStyleId, ctx.styles) !== null) {
    return {
      ...base,
      type: 'passthrough',
      label: 'TOC entry',
      previewText: plainText(xml),
      fieldDisplay: fieldDisplayOf(xml, ctx.styles),
      styleId: tocStyleId,
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
  if (detect.includes('<w:object') || /<w:pict[\s>]/.test(detect)) {
    // Legacy VML textboxes (v:shape/v:textbox) and WordArt (v:textpath) in
    // w:pict: extract the structured display model instead of flattening every
    // nested paragraph and table into one unreadable run — or degrading the
    // whole paragraph to an opaque chip. w:object (OLE) keeps the plain preview.
    if (
      !detect.includes('<w:object') &&
      (detect.includes('<w:txbxContent') || VML_WORDART_RE.test(detect))
    ) {
      // DrawingML pictures anchored next to the VML shape: resolve their media
      // and open the pictures gate, or the photos silently vanish from the page.
      // Host-level XML only — a picture nested in the textbox's own content
      // must not be lifted into a page-level photo box.
      const hostXml = stripTextboxes(detect)
      const anchoredPics =
        hostXml.includes('<w:drawing') &&
        (hostXml.includes('<pic:pic') || hostXml.includes('<a:blip'))
      if (anchoredPics || VML_PICT_RID_RE.test(hostXml)) await resolveBlipMedia(detect, ctx)
      const textboxes = extractTextboxes(
        detect,
        ctx,
        anchoredPics
          ? {
              shapes: true,
              pictures: true,
              section: ctx.sectionAt?.(el.start),
              docOffset: el.start,
              firstPage:
                index > 0 &&
                (ctx.firstPageBreakAt === undefined || el.start < ctx.firstPageBreakAt),
            }
          : { docOffset: el.start },
      )
      const strayText = plainText(stripTextboxes(detect)).trim()
      // paragraphs mixing an inline VML picture with real text stay on the
      // editable run-image path below; boxes would drop the picture. Only a
      // v:imagedata with an r:id is a picture — WPS stamps a bare
      // <v:imagedata o:title=""/> on plain textbox shapes.
      const keepForImages = strayText !== '' && VML_PICT_RID_RE.test(detect)
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
    if (!detect.includes('<w:object') && VML_PICT_RID_RE.test(detect)) {
      if (plainText(stripTextboxes(detect)).trim() !== '') {
        // picture shares the paragraph with real text: keep the text editable
        // with run-level images (the pict fragment round-trips verbatim)
        await resolveBlipMedia(detect, ctx)
        return buildTextParagraph(base, xml, ctx, true, el.start)
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
    // OLE embed sharing the paragraph with real text or another drawing: an
    // "Embedded object" block renders only the preview picture, silently
    // dropping the rest. Keep the paragraph on the run-image path (the
    // w:object fragment round-trips verbatim). Only when every object's
    // preview resolved — otherwise the passthrough below at least keeps the
    // object bytes safe.
    if (
      detect.includes('<w:object') &&
      (plainText(stripTextboxes(detect)).trim() !== '' || detect.includes('<w:drawing'))
    ) {
      await resolveBlipMedia(detect, ctx)
      const objects = detect.match(/<w:object[\s\S]*?<\/w:object>/g) ?? []
      const displayable = objects.every((o) => {
        const rId = /<v:imagedata[^>]*r:id="([^"]+)"/.exec(o)?.[1]
        return rId !== undefined && ctx.mediaByRid?.has(rId)
      })
      if (displayable && objects.length > 0)
        return buildTextParagraph(base, xml, ctx, true, el.start)
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
    if (
      detect.includes('<c:chart') ||
      detect.includes('<cx:chart') ||
      detect.includes('r:dm=') ||
      detect.includes('<dgm:')
    ) {
      const isChart = detect.includes('<c:chart') || detect.includes('<cx:chart')
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
          docOffset: el.start,
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
    // wps shapes would vanish on the plain-paragraph and single-image paths:
    // paragraphs carrying them prefer the box-extraction route below
    const hasWsp = detect.includes('<wps:wsp')
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
        if ((hasText && !hasWsp) || (multiPic && !detect.includes('<wp:anchor'))) {
          await resolveBlipMedia(detect, ctx)
          return buildTextParagraph(base, xml, ctx, true, el.start)
        }
        // several separately-anchored pictures in one paragraph (photo walls):
        // the single-image block keeps only the first blip and drops the other
        // photos with their anchor offsets, so take the photo-box route below
        const anchoredDrawings = topLevelDrawings(detect).filter((f) =>
          f.includes('<wp:anchor'),
        ).length
        if (!hasWsp && anchoredDrawings <= 1) {
          return {
            ...base,
            type: 'image',
            label: 'Image',
            imageDataUrl: image,
            ...imageMeta(detect),
          }
        }
      }
    }
    // picture fills inside shapes (a:blipFill) and sibling/grouped pic:pic
    // blips resolve from pre-fetched media
    if (detect.includes('<a:blip')) await resolveBlipMedia(detect, ctx)
    // shapes on: textless preset shapes (stars, triangles, block arrows,
    // anchored connectors) render instead of degrading to a chip; pictures on:
    // grouped/sibling pic:pic drawings render as photo boxes
    const textboxes = extractTextboxes(detect, ctx, {
      shapes: true,
      pictures: true,
      section: ctx.sectionAt?.(el.start),
      docOffset: el.start,
      // page-pinning needs content ABOVE the anchor paragraph to matter: a
      // first-block anchor is already exact under the paragraph-origin path
      // (and stays aligned with the body text around it)
      firstPage:
        index > 0 && (ctx.firstPageBreakAt === undefined || el.start < ctx.firstPageBreakAt),
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
    // (the stray text joins the preview instead of vanishing). Same for
    // wps-shape paragraphs that extracted visible boxes: dropping them loses
    // real page furniture (cards, dividers), so the text rides along as
    // display-only strayRuns instead.
    if (
      strayText !== '' &&
      !boxTexts.some((t) => t.trim() !== '') &&
      !(hasWsp && textboxes.length > 0)
    ) {
      return buildTextParagraph(base, xml, ctx, true, el.start)
    }
    // Anchored textboxes (code boxes, callout cards): all visible text lives in
    // w:txbxContent. Extract a display-only model so content and box styling
    // render; the block stays protected and saves byte-identical. Text the
    // paragraph itself carries next to the boxes is kept as display-only runs
    // (strayRuns) so it doesn't vanish from the page.
    if (textboxes.length > 0) {
      // inline pictures sharing an anchored-drawing paragraph flow as stray run
      // images: a non-floating photo box would knock the whole block out of the
      // zero-height floating overlay (Word keeps them on the paragraph's line)
      const inlineRunPics =
        detect.includes('<wp:anchor') &&
        topLevelDrawings(detect).some(
          (f) =>
            !f.includes('<wp:anchor') && f.includes('<pic:pic') && !f.includes('<w:txbxContent'),
        )
      const stray =
        strayText !== '' || inlineRunPics ? strayParaRuns(detect, ctx, inlineRunPics) : null
      return {
        ...base,
        type: 'passthrough',
        label: 'Text box',
        previewText: (strayText !== '' ? [strayText, ...boxTexts] : boxTexts).join('\n'),
        textboxes,
        ...(hostPageBreak(detect) ? { fieldDisplay: { kind: 'pageBreak' as const } } : {}),
        ...(stray && stray.runs.length > 0
          ? {
              strayRuns: stray.runs,
              ...(stray.styleId ? { strayStyleId: stray.styleId } : {}),
              ...(stray.indent ? { strayIndent: stray.indent } : {}),
            }
          : {}),
        ...imageMeta(detect),
      }
    }
    // wps paragraph whose shapes all turned out invisible: fall back to the
    // plain image block the pre-shape path would have produced
    if (image) {
      return { ...base, type: 'image', label: 'Image', imageDataUrl: image, ...imageMeta(detect) }
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

  return buildTextParagraph(base, xml, ctx, false, el.start)
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

const defaultParaVanishCache = new WeakMap<Map<string, StyleInfo>, boolean>()

/** Default paragraph style's w:vanish (cached per styles map): inherited by style-less paragraphs */
function defaultParaVanish(styles?: Map<string, StyleInfo>): boolean | undefined {
  if (!styles) return undefined
  let v = defaultParaVanishCache.get(styles)
  if (v === undefined) {
    v = false
    for (const info of styles.values()) {
      if (info.isDefault && info.type === 'paragraph') {
        v = info.display?.vanish === true
        break
      }
    }
    defaultParaVanishCache.set(styles, v)
  }
  return v || undefined
}

/** parse a w:p as editable text content (paragraph / heading / listItem) */
/** Word's built-in text size when neither docDefaults nor the default style sets one */
const WORD_DEFAULT_SIZE_HALF_POINTS = 20

/**
 * Twips per "character" for a paragraph's character-unit indents (Word for Mac
 * probe, 2026-09-02). The special indents follow the first text run — its w:sz
 * plus its w:spacing letter spacing; an empty run does not count and the
 * paragraph mark never does (a run-less paragraph falls back to its mark, then
 * its style). Left/right follow the default paragraph style's size, whatever
 * the paragraph's own style or runs say. Under a linesAndChars document grid
 * every character advances by charSpace/4096 pt more, and so do both units
 * (snapToGrid off included).
 */
function charUnitsOf(
  ctx: BuildContext,
  runs: Run[],
  pNode: XNode,
  pPr: XNode | undefined,
  styleId: string | undefined,
  docOffset: number | undefined,
): CharUnits {
  const grid = docOffset !== undefined ? ctx.sectionAt?.(docOffset).docGrid : undefined
  const gridDelta =
    grid?.type === 'linesAndChars' && grid.charSpace ? (grid.charSpace / 4096) * 20 : 0
  const normal =
    ctx.defaultParaStyle?.display?.sizeHalfPoints ??
    ctx.docDefaults?.sizeHalfPoints ??
    WORD_DEFAULT_SIZE_HALF_POINTS
  const styleSize = styleId
    ? ctx.styles.get(styleId)?.display?.sizeHalfPoints
    : ctx.defaultParaStyle?.display?.sizeHalfPoints
  const first = runs.find((r) => r.text !== '')
  const run =
    first?.sizeHalfPoints ??
    (first?.styleId ? ctx.styles.get(first.styleId)?.display?.sizeHalfPoints : undefined) ??
    (first ? undefined : emptyParaSizeHalfPoints(pNode, pPr)) ??
    styleSize ??
    normal
  return {
    run: run * 10 + (first?.charSpacingTwips ?? 0) + gridDelta,
    normal: normal * 10 + gridDelta,
  }
}

/**
 * Fold a paragraph's character-unit indents (the ones its style chain declares,
 * layered under the direct w:ind *Chars attributes) into the format's twips
 * fields. Only a *Chars attribute — an explicit zero included — replaces a
 * style's character indent; a direct twips w:firstLine / w:left leaves it in
 * force (probed: Word lays a style firstLineChars paragraph with a direct
 * w:firstLine="420" at two characters). Style character indents stay off list
 * items, like the style-indent CSS: numbering owns their indent.
 */
function withCharIndents(
  format: ParaFormat | undefined,
  ctx: BuildContext,
  pNode: XNode,
  pPr: XNode | undefined,
  styleId: string | undefined,
  runs: Run[],
  opts: { list: boolean; docOffset?: number },
): ParaFormat | undefined {
  const direct = charIndentsOf(pPr ? findChild(pPr, 'w:ind') : undefined)
  const style = opts.list
    ? undefined
    : (styleId ? ctx.styles.get(styleId) : ctx.defaultParaStyle)?.display?.indentChars
  const chars = activeCharIndents(mergeCharIndents(style, direct))
  if (!chars) return format
  return {
    ...resolveCharIndents(
      format,
      chars,
      charUnitsOf(ctx, runs, pNode, pPr, styleId, opts.docOffset),
    ),
    // the save path cancels these when it rebuilds w:ind in twips
    charIndents: chars,
  }
}

function buildTextParagraph(
  base: Pick<Block, 'id' | 'docxIndex' | 'originalXml'>,
  xml: string,
  ctx: BuildContext,
  withImages = false,
  docOffset?: number,
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
  // style-less paragraph under a vanish-carrying default paragraph style: the
  // mark inherits hidden, so Word gives the line no height (real_run2/93's
  // trailing paragraph pushed a phantom second page)
  if (!styleId && defaultParaVanish(ctx.styles) === true && staysVanished(xml)) {
    return { ...base, type: 'passthrough', label: 'Hidden paragraph', invisibleMarker: true }
  }
  // empty paragraph whose mark is hidden by direct w:vanish (label/card
  // templates end on one): Word gives it no height — rendering it as a normal
  // empty line can push a trailing blank page
  const markRPr = pPr ? findChild(pPr, 'w:rPr') : undefined
  if (
    markRPr &&
    onOffOf(markRPr, 'w:vanish') === true &&
    onOffOf(markRPr, 'w:specVanish') !== true &&
    plainText(xml).trim() === '' &&
    staysVanished(xml) &&
    !hasLayoutRunContent(pNode)
  ) {
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
    xml.includes('<mc:Fallback')
      ? xml.replace(/<mc:Fallback[^>]*>[\s\S]*?<\/mc:Fallback>/g, '')
      : xml,
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
    const emptyFont = emptyParaMarkFont(pNode, pPr)
    if (emptyFont) format = { ...(format ?? {}), emptyRunFontFamily: emptyFont }
  }

  // w:ptab absolute-position tabs (TOC dot leaders to the right margin): the run
  // text carries them as '\t'; surface each as a display-only margin-relative
  // stop so the renderer can align at the column center/edge and draw the leader
  const ptabStops = ptabDisplayStops(pNode)
  if (ptabStops.length > 0) {
    const stops = format?.tabStops ? [...format.tabStops] : []
    for (const st of ptabStops) {
      if (!stops.some((s) => s.rel === 'margin' && s.pos === st.pos)) stops.push(st)
    }
    format = { ...(format ?? {}), tabStops: stops }
  }

  // allowOverlap="0" colliding with a sibling anchor (tdf#134114): Word displaces
  // the object out of the other's box instead of stacking the wrapped floats
  // vertically; approximate with a front overlay at the collider's bottom (it may
  // hang into the margin, like Word). Horizontal overlap is presumed — the align
  // gallery drops the X, so only the declared vertical ranges are compared.
  {
    const wrapped = runs.filter(
      (r) => r.image?.wrap && r.image.wrap !== 'front' && r.image.wrap !== 'behind',
    )
    if (wrapped.length > 1) {
      for (const r of wrapped) {
        const img = r.image!
        if (!img.noOverlap) continue
        const top = (img.offsetYEmu ?? 0) / EMU_PER_PX
        const hit = wrapped.find((o) => {
          if (o === r || o.image!.noOverlap) return false
          const oTop = (o.image!.offsetYEmu ?? 0) / EMU_PER_PX
          return top < oTop + (o.image!.heightPx ?? 0) && oTop < top + (img.heightPx ?? 0)
        })
        if (!hit) continue
        img.wrap = 'front'
        img.offsetXEmu = 0
        // the collider float renders at its line top (run floats ignore posOffset)
        img.offsetYEmu = ((hit.image!.heightPx ?? 0) + 2) * EMU_PER_PX
      }
    }
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

  // --- deleted paragraph mark (w:pPr/w:rPr/w:del) ---
  let paraMarkDel: Block['paraMarkDel']
  {
    const pRPr = pPr ? findChild(pPr, 'w:rPr') : undefined
    const delEl = pRPr ? findChild(pRPr, 'w:del') : undefined
    if (delEl) {
      const a = attrsOf(delEl)
      paraMarkDel = { author: a['w:author'] ?? '' }
      if (a['w:date']) paraMarkDel.date = a['w:date']
      if (a['w:id']) paraMarkDel.id = a['w:id']
    }
  }

  // list item?
  const listRef = listRefOf(ctx, pPr, styleId)
  format = withCharIndents(format, ctx, pNode, pPr, styleId, runs, { list: !!listRef, docOffset })
  /** extra revision fields shared across all return paths */
  const revExtras = {
    ...(moveRevision ? { moveRevision } : {}),
    ...(pPrChangeInfo ? { pPrChangeInfo } : {}),
    ...(paraMarkDel ? { paraMarkDel } : {}),
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
 * Text runs the anchor paragraph carries alongside content textboxes
 * (heading text sharing its paragraph with an anchored sidebar). Display-only:
 * the block still saves byte-identical.
 */
function strayParaRuns(
  paragraphXml: string,
  ctx: BuildContext,
  withImages = false,
): { runs: Run[]; styleId?: string; indent?: StrayIndent } | null {
  try {
    let strayXml = stripTextboxes(paragraphXml)
    // inline run images only — anchored drawings already render as boxes
    if (withImages) {
      for (const frag of topLevelDrawings(strayXml)) {
        if (frag.includes('<wp:anchor')) strayXml = strayXml.replace(frag, '')
      }
    }
    const parsed = xmlParser.parse(strayXml) as XNode[]
    const pNode = parsed.find((n) => nameOf(n) === 'w:p')
    if (!pNode) return null
    const runs = extractRuns(pNode, ctx, [], [], withImages).filter((r) => r.text !== '' || r.image)
    if (runs.length === 0) return null
    const pPr = findChild(pNode, 'w:pPr')
    const styleId = pPr ? attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] : undefined
    // the anchor paragraph's own w:ind shapes the stray line (a wrap column
    // carved out with a large right indent must not run under the box)
    let indent: StrayIndent | undefined
    const ind = pPr ? findChild(pPr, 'w:ind') : undefined
    if (ind) {
      const a = attrsOf(ind)
      const num = (v: string | undefined) => {
        const n = parseInt(v ?? '', 10)
        return Number.isFinite(n) && n !== 0 ? n : undefined
      }
      const left = num(a['w:left'] ?? a['w:start'])
      const right = num(a['w:right'] ?? a['w:end'])
      const hanging = num(a['w:hanging'])
      const firstLine = hanging !== undefined && hanging > 0 ? -hanging : num(a['w:firstLine'])
      if (left !== undefined || right !== undefined || firstLine !== undefined) {
        indent = {
          ...(left !== undefined ? { leftTwips: left } : {}),
          ...(right !== undefined ? { rightTwips: right } : {}),
          ...(firstLine !== undefined ? { firstLineTwips: firstLine } : {}),
        }
      }
    }
    return {
      runs,
      ...(styleId ? { styleId } : {}),
      ...(indent ? { indent } : {}),
    }
  } catch {
    return null
  }
}

/** paragraphs (and tables, one display line per row) of a w:txbxContent node */
/** @param docOffset host paragraph's document.xml offset (section grid for character indents) */
function txbxContentParas(
  content: XNode,
  ctx: BuildContext,
  docOffset?: number,
): TextboxParaDisplay[] {
  const out: TextboxParaDisplay[] = []
  for (const child of childrenOf(content)) {
    const name = nameOf(child)
    if (name === 'w:p') {
      // withImages: inline drawings in textbox paragraphs (form checkboxes)
      // must become run images or they are dropped — and lost on first edit
      const para: TextboxParaDisplay = { runs: extractRuns(child, ctx, [], [], true) }
      const pPr = findChild(child, 'w:pPr')
      const styleId = pPr ? attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] : undefined
      // pPr-less paragraphs too: the Normal style's character indents apply to them
      // in a textbox just as in the body
      Object.assign(
        para,
        withCharIndents(
          pPr ? extractParaFormat(pPr) : undefined,
          ctx,
          child,
          pPr,
          styleId,
          para.runs,
          {
            list: !!listRefOf(ctx, pPr, styleId),
            docOffset,
          },
        ),
      )
      if (styleId) para.styleId = styleId
      out.push(para)
    } else if (name === 'w:tbl') {
      out.push(...txbxTableParas(child, ctx))
    } else if (name === 'w:sdt') {
      const inner = findChild(child, 'w:sdtContent')
      if (inner) out.push(...txbxContentParas(inner, ctx, docOffset))
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
        const pRuns = extractRuns(p, ctx, [], [], true)
        if (pRuns.every((r) => r.text.trim() === '' && !r.image)) continue
        if (cellRuns.length > 0) cellRuns.push({ text: ' ' })
        cellRuns.push(...pRuns)
      }
      if (!cellRuns.every((r) => r.text.trim() === '' && !r.image)) {
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
  // every w:txbxContent consumes one save-path ordinal (box emitted or not),
  // mirroring how patchTextboxParas counts the non-fallback segments
  let txbxOrdinal = 0

  /** one wps:wsp (already parsed) → display box; null when it renders nothing */
  const buildWpsBox = (
    shape: XNode,
    groupFill?: string,
    nested?: boolean,
  ): TextboxDisplay | null => {
    const contents: XNode[] = []
    collectNodes(childrenOf(shape), 'w:txbxContent', contents)
    // only top-level contents of a top-level shape consume a save-path
    // ordinal: xmlSegments never descends into a segment, so neither a
    // textbox nested inside another nor a shape living inside some other
    // box's w:txbxContent may advance the counter
    const topContents: XNode[] = []
    collectTopNodes(childrenOf(shape), 'w:txbxContent', topContents)
    const ordinal = txbxOrdinal
    if (!nested) txbxOrdinal += topContents.length
    // external textbox part (wps:txbx r:txbx, word/txbx*.xml): its root
    // w14:txbx carries the w:p list directly. Display-only — it stays out of
    // topContents, so the box takes no save ordinal and turns readOnly below.
    if (contents.length === 0) {
      const extRid = attrsOf(findChild(shape, 'wps:txbx') ?? {})['r:txbx']
      const extXml = extRid ? ctx.externalTxbxByRid?.get(extRid) : undefined
      if (extXml) {
        try {
          const nodes = xmlParser.parse(extXml) as XNode[]
          const root = nodes.find((n) => nameOf(n)?.endsWith(':txbx'))
          if (root) contents.push(root)
        } catch {
          /* unreadable part: fall through to the textless-shape path */
        }
      }
    }
    const spPr = findChild(shape, 'wps:spPr')
    const prstOf = spPr ? attrsOf(findChild(spPr, 'a:prstGeom') ?? {})['prst'] : undefined
    const custGeomNode = spPr ? findChild(spPr, 'a:custGeom') : undefined
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
      if (!opts?.shapes || (!prstOf && !custGeomNode)) return null
      if (prstOf === 'rect') {
        // near-flat rects stay on the decorative thin-rule path; a real
        // rectangle (pattern/solid-filled swatch, tdf dml-shape-fillpattern)
        // renders like any other textless preset shape
        const xfrm = findChild(spPr ?? {}, 'a:xfrm')
        const cy = parseInt(attrsOf(findChild(xfrm ?? {}, 'a:ext') ?? {})['cy'] ?? '', 10)
        if (!Number.isFinite(cy) || cy <= 130000) return null
      }
    }
    const box: TextboxDisplay = { paras: [] }
    if (!nested && topContents.length > 0) box.txbxIndex = ordinal
    if (nested) box.readOnly = true
    const shapeId = attrsOf(findChild(shape, 'wps:cNvPr') ?? {})['id']
    if (!nested && shapeId) box.shapeId = shapeId
    if (spPr) {
      if (!findChild(spPr, 'a:noFill')) {
        // a:pattFill approximates to its foreground color (same tradeoff as gradFill)
        const pattFill = findChild(spPr, 'a:pattFill')
        const fill =
          colorNodeHex(findChild(spPr, 'a:solidFill'), ctx.themeColors) ??
          gradFillApproxHex(spPr, ctx.themeColors) ??
          (pattFill ? colorNodeHex(findChild(pattFill, 'a:fgClr'), ctx.themeColors) : undefined) ??
          // a:grpFill inherits the enclosing wpg group's fill
          (findChild(spPr, 'a:grpFill') ? groupFill : undefined)
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
        // a:fontRef is where a gallery shape's text color comes from — the default
        // blue shape references lt1, which is why Word and PowerPoint show white text
        // on it without writing a color on any run. Runs with their own w:color win.
        const fontColor = colorNodeHex(findChild(styleNode, 'a:fontRef'), ctx.themeColors)
        if (fontColor) box.textColor = fontColor
      }
      const prst = prstOf
      if (prst && prst !== 'rect') box.prst = prst
      const xfrm = findChild(spPr, 'a:xfrm')
      if (custGeomNode) {
        const extAttrs = attrsOf(findChild(xfrm ?? {}, 'a:ext') ?? {})
        const geom = parseCustGeom(
          serializeXNode(spPr),
          parseInt(extAttrs['cx'] ?? '', 10) || 0,
          parseInt(extAttrs['cy'] ?? '', 10) || 0,
        )
        if (geom) box.pathData = geom
      }
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
    for (const content of contents)
      box.paras.push(...txbxContentParas(content, ctx, opts?.docOffset))
    if (contents.some(txbxHasStructuredContent)) box.readOnly = true
    // a textbox nested inside this one flattens into paras above: a commit
    // would rewrite the outer w:p list and destroy the nested shape
    if (contents.length > topContents.length) box.readOnly = true
    if (contents.length === 0) {
      // textless preset shape: keep it if the geometry has any visible ink
      if (!box.fill && !box.borderColor && !box.fillImageDataUrl) return null
      // no w:txbxContent to patch — editable only when the cNvPr id gives the
      // save path a shape to inject a fresh wps:txbx into
      if (!box.shapeId) box.readOnly = true
      // Word centers shape text: without an explicit anchor the live preview
      // and the anchor="ctr" the inject path writes agree
      else if (!box.vAlign && !(bodyPr && attrsOf(bodyPr)['anchor'])) box.vAlign = 'center'
      return box
    }
    if (box.paras.some((p) => p.runs.length > 0)) return box
    // text-empty box: keep visible ink (a full-page white box must occupy its
    // extent instead of degrading to a chip); its w:txbxContent stays
    // addressable, so typing into the empty shape still commits
    if (!box.fill && !box.borderColor && !box.fillImageDataUrl) return null
    box.paras = []
    return box
  }

  const applyAnchor = (
    box: TextboxDisplay,
    meta: DrawingAnchorMeta,
    pagePos: ResolvedAnchorPos | null,
    grouped = false,
  ): void => {
    if (!meta.anchored) return
    if (meta.behind) box.behind = true
    if (meta.z !== undefined) box.z = meta.z
    // first-page page-anchored cover art: raw page coordinates, rendered
    // against the page box (doc-protected-pagepinned)
    if (meta.pageXEmu !== undefined) {
      box.offsetXEmu = (box.offsetXEmu ?? 0) + meta.pageXEmu
      box.offsetYEmu = (box.offsetYEmu ?? 0) + (meta.pageYEmu ?? 0)
      box.floating = true
      box.pagePinned = true
      return
    }
    // page/margin-anchored drawing sitting outside the body column (Word
    // resume sidebars): absolute placement at the resolved position instead
    // of stacking in the flow, wrap kind notwithstanding
    // page/margin-relative X is absolute on the page in Word: a
    // column-translated anchor block must not drag the box sideways
    const relXAbsolute = meta.relH === 'page' || meta.relH === 'margin'
    if (pagePos?.outsideColumn) {
      // group children already carry a group-relative offset: the anchor adds
      box.offsetXEmu = (box.offsetXEmu ?? 0) + pagePos.xEmu
      box.offsetYEmu = (box.offsetYEmu ?? 0) + (pagePos.yEmu ?? meta.offsetYEmu ?? 0)
      box.floating = true
      if (relXAbsolute) box.pageRelX = true
      return
    }
    if (meta.offsetXEmu !== undefined) {
      box.offsetXEmu = (box.offsetXEmu ?? 0) + meta.offsetXEmu
      if (relXAbsolute) box.pageRelX = true
    }
    // margin-aligned X (wp:align left/center/right) on a floating drawing (photo
    // rows): resolve against the margin box; in-flow wrapSquare boxes keep the
    // legacy flow placement
    if (
      pagePos !== null &&
      !pagePos.outsideColumn &&
      meta.offsetXEmu === undefined &&
      (meta.topBottom || meta.noWrap || multiDrawing)
    ) {
      box.offsetXEmu = (box.offsetXEmu ?? 0) + pagePos.xEmu
      if (relXAbsolute) box.pageRelX = true
    }
    if (meta.offsetYEmu !== undefined) box.offsetYEmu = (box.offsetYEmu ?? 0) + meta.offsetYEmu
    // page/margin-relative posOffset V is absolute on the anchor's page in
    // Word; rendered from the anchor paragraph it needs the canvas re-pin
    if (
      (meta.relV === 'page' || meta.relV === 'margin') &&
      meta.offsetYEmu !== undefined &&
      !meta.alignV
    ) {
      box.pageRelV = true
    }
    // wrapSquare box spanning (nearly) the whole column: no text fits in the
    // leftover sliver, so Word floats the box at its offset (the anchor line
    // stays under it) and lays text above/below — reserve the band like
    // wrapTopAndBottom instead of stacking the box into the flow (or, in a
    // multi-drawing paragraph, overlaying the flow with zero footprint).
    // A multi-drawing paragraph covers the column with the UNION of its boxes
    // (half-width sign-off frames side by side): the wrapSquare ones band too.
    let ownSpansColumn = false
    let squareSpansColumn = false
    const sect = opts?.section
    if (!meta.noWrap && !meta.topBottom && !grouped && sect && sect.columns <= 1) {
      const colWEmu = (sect.pageWidth - sect.marginLeft - sect.marginRight) * EMU_PER_TWIP
      const wEmu = box.widthPx !== undefined ? box.widthPx * EMU_PER_PX : meta.extentXEmu
      const xEmu = box.offsetXEmu ?? 0
      ownSpansColumn =
        wEmu !== undefined &&
        wEmu > 0 &&
        xEmu < MIN_WRAP_SLIVER_EMU &&
        colWEmu - xEmu - wEmu < MIN_WRAP_SLIVER_EMU
      squareSpansColumn = ownSpansColumn || (multiDrawing && anchorUnionSpansColumn(sect))
    }
    // wrapTopAndBottom (Word): body text is excluded from the box's whole
    // vertical band. The box floats at its offset and the anchor paragraph
    // reserves flow height down to the box bottom (union over its boxes).
    if (
      (meta.topBottom || squareSpansColumn) &&
      (meta.relV === 'paragraph' || meta.relV === 'line')
    ) {
      // wp:extent cy covers the whole drawing: a usable height fallback only
      // for an ungrouped shape (each group child would claim the group height)
      const h =
        box.heightPx ??
        (!grouped && meta.extentYEmu !== undefined
          ? Math.round(meta.extentYEmu / EMU_PER_PX)
          : undefined)
      if (h !== undefined) {
        const top = Math.round((box.offsetYEmu ?? 0) / EMU_PER_PX)
        if (top + h > 0) {
          box.bandTopPx = top
          box.bandBottomPx = top + h
          // only a box that spans the column BY ITSELF may overflow the page
          // bottom (Word keeps it on its anchor's page); a union band (side by
          // side frames) is pushed whole to the next page like Word
          if (ownSpansColumn && !meta.topBottom) box.bandOverflow = true
        }
      }
      box.floating = true
    }
    if (meta.noWrap || multiDrawing) box.floating = true
  }

  // union of the paragraph's anchored drawing extents: when the widest
  // remaining horizontal gap in the column is a sliver, no text fits beside
  // the boxes and Word lays it above/below the row
  let unionSpans: boolean | undefined
  const anchorUnionSpansColumn = (sect: SectionSettings): boolean => {
    if (unionSpans !== undefined) return unionSpans
    const colWEmu = (sect.pageWidth - sect.marginLeft - sect.marginRight) * EMU_PER_TWIP
    const iv = fragMetas
      .filter((m) => m.anchored && m.offsetXEmu !== undefined && (m.extentXEmu ?? 0) > 0)
      .map((m): [number, number] => [m.offsetXEmu!, m.offsetXEmu! + m.extentXEmu!])
      .sort((a, b) => a[0] - b[0])
    let gap = 0
    let cursor = 0
    for (const [a, b] of iv) {
      gap = Math.max(gap, a - cursor)
      cursor = Math.max(cursor, b)
    }
    gap = Math.max(gap, colWEmu - cursor)
    unionSpans = iv.length > 0 && gap < MIN_WRAP_SLIVER_EMU
    return unionSpans
  }

  // first-page cover art with a page-relative V anchor: keep RAW page
  // coordinates and pin the boxes to the page box itself. Any paragraph- or
  // body-top-relative rendering re-adds whatever sits above the anchor
  // paragraph (logo lines, empty leads) and shifts the whole composition.
  // All-or-nothing per paragraph: the un-positioned pagepinned wrapper would
  // break any sibling drawing still positioned from the paragraph origin.
  const pinnable = (m: DrawingAnchorMeta): boolean =>
    m.anchored === true &&
    (m.noWrap === true || multiDrawing) &&
    m.relV === 'page' &&
    (m.relH === 'page' || m.relH === 'margin')
  const fragMetas = frags.map(drawingAnchorMeta)
  const anchoredMetas = fragMetas.filter((m) => m.anchored)
  const pinAll =
    opts?.firstPage === true &&
    opts.section !== undefined &&
    anchoredMetas.length > 0 &&
    anchoredMetas.every(pinnable)
  // page-relative posOffset measures from the page edge; boxes render from
  // the column/paragraph origin, so keeping the raw value double-counts the
  // margins (X is exact; Y approximates the anchor paragraph at body top,
  // strictly closer than the raw page offset). One pass up front keeps every
  // consumer -- applyAnchor, the column-span test, the anchor union -- in the
  // same column space; pinned metas keep raw page coordinates.
  if (opts?.section) {
    for (const m of fragMetas) {
      if (pinAll && pinnable(m)) continue
      if (m.relH === 'page' && m.offsetXEmu !== undefined && !m.alignH) {
        m.offsetXEmu -= opts.section.marginLeft * EMU_PER_TWIP
      }
      if (m.relV === 'page' && m.offsetYEmu !== undefined && !m.alignV) {
        m.offsetYEmu -= opts.section.marginTop * EMU_PER_TWIP
      }
    }
  }

  for (const [fragIndex, frag] of frags.entries()) {
    let parsedFrag: XNode[]
    try {
      parsedFrag = xmlParser.parse(frag) as XNode[]
    } catch {
      continue
    }
    const meta = fragMetas[fragIndex]
    if (pinAll && pinnable(meta) && opts?.section) {
      const marL = opts.section.marginLeft * EMU_PER_TWIP
      const marT = opts.section.marginTop * EMU_PER_TWIP
      const aligned = resolveAnchorPagePos(meta, opts.section)
      meta.pageXEmu =
        aligned !== null
          ? aligned.xEmu + marL
          : (meta.relH === 'margin' ? marL : 0) + (meta.offsetXEmu ?? 0)
      meta.pageYEmu = aligned?.yEmu !== undefined ? aligned.yEmu + marT : (meta.offsetYEmu ?? 0)
    }
    const pagePos = meta.pageXEmu === undefined ? resolveAnchorPagePos(meta, opts?.section) : null
    let wspCount = 0
    const pushShape = (
      shape: XNode,
      ctm: GroupCtm | null,
      groupFill?: string,
      nested?: boolean,
    ): void => {
      wspCount++
      const box = buildWpsBox(shape, groupFill, nested)
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
      applyAnchor(box, meta, pagePos, ctm !== null)
      out.push(box)
    }
    // grouped pic:pic (photo inside a wpg group): a photo box at the mapped
    // child offset, painted in document order between its sibling shapes
    const pushPic = (node: XNode, ctm: GroupCtm | null): void => {
      const blip = findChild(findChild(node, 'pic:blipFill') ?? {}, 'a:blip')
      const blipAttrs = attrsOf(blip ?? {})
      const rId = blipAttrs['r:embed'] ?? blipAttrs['r:link']
      const dataUrl = rId ? ctx.mediaByRid?.get(rId) : undefined
      if (!dataUrl) return
      const xfrm = findChild(findChild(node, 'pic:spPr') ?? {}, 'a:xfrm')
      const ext = attrsOf(findChild(xfrm ?? {}, 'a:ext') ?? {})
      const cx = parseInt(ext['cx'] ?? '', 10)
      const cy = parseInt(ext['cy'] ?? '', 10)
      if (!(cx > 0) || !(cy > 0)) return
      const rot = parseInt(attrsOf(xfrm ?? {})['rot'] ?? '', 10)
      const box: TextboxDisplay = {
        paras: [],
        readOnly: true,
        fillImageDataUrl: dataUrl,
        widthPx: Math.round((cx * (ctm?.sx ?? 1)) / EMU_PER_PX),
        heightPx: Math.round((cy * (ctm?.sy ?? 1)) / EMU_PER_PX),
        insetTopPx: 0,
        insetRightPx: 0,
        insetBottomPx: 0,
        insetLeftPx: 0,
      }
      if (Number.isFinite(rot) && rot !== 0) box.rotDeg = Math.round(rot / 60000)
      if (ctm) {
        const off = attrsOf(findChild(xfrm ?? {}, 'a:off') ?? {})
        const x = parseInt(off['x'] ?? '', 10)
        const y = parseInt(off['y'] ?? '', 10)
        if (Number.isFinite(x)) box.offsetXEmu = Math.round(ctm.tx + x * ctm.sx)
        if (Number.isFinite(y)) box.offsetYEmu = Math.round(ctm.ty + y * ctm.sy)
        box.floating = true
      }
      applyAnchor(box, meta, pagePos, ctm !== null)
      out.push(box)
    }
    // document-order walk (box order must match w:txbxContent order for the
    // patch-save mapping), carrying the enclosing wpg group transform and
    // fill; anything below a wps:wsp lives inside its txbx → nested
    const walkShapes = (
      nodes: XNode[],
      ctm: GroupCtm | null,
      groupFill?: string,
      nested = false,
    ): void => {
      for (const node of nodes) {
        const name = nameOf(node)
        if (name === 'wps:wsp') pushShape(node, ctm, groupFill, nested)
        if (name === 'pic:pic' && opts?.pictures && ctm) pushPic(node, ctm)
        if (name === 'wpg:wgp' || name === 'wpg:grpSp') {
          const fill =
            colorNodeHex(
              findChild(findChild(node, 'wpg:grpSpPr') ?? {}, 'a:solidFill'),
              ctx.themeColors,
            ) ?? groupFill
          walkShapes(
            childrenOf(node),
            composeGroupCtm(node, ctm ?? IDENTITY_CTM) ?? ctm,
            fill,
            nested,
          )
        } else {
          walkShapes(childrenOf(node), ctm, groupFill, nested || name === 'wps:wsp')
        }
      }
    }
    walkShapes(parsedFrag, null)
    // picture-only drawing sharing a multi-drawing paragraph: a photo box.
    // Inline pics in a paragraph that also anchors drawings stay out — they
    // flow as stray run images, keeping every box floating (overlay layout)
    if (
      opts?.pictures &&
      wspCount === 0 &&
      frag.includes('<pic:pic') &&
      (meta.anchored || !xml.includes('<wp:anchor'))
    ) {
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
        // rotation lives on the pic's own xfrm (same read as imageMeta)
        const picXfrm = /<pic:spPr[^>]*>[\s\S]*?<a:xfrm([^>]*)>/.exec(frag)?.[1]
        const rot = parseInt(/\brot="(-?\d+)"/.exec(picXfrm ?? '')?.[1] ?? '', 10)
        if (Number.isFinite(rot) && rot !== 0) box.rotDeg = Math.round(rot / 60000)
        applyAnchor(box, meta, pagePos)
        out.push(box)
      }
    }
  }

  // legacy VML shapes (w:pict, outside w:drawing fragments)
  let parsed: XNode[] | null = null
  if (/<v:(?:shape|rect|roundrect|oval)\b/.test(xml)) {
    try {
      parsed = xmlParser.parse(xml) as XNode[]
    } catch {
      parsed = null
    }
  }
  if (parsed) {
    // shared placement: top-level absolute shapes float at their pt margins;
    // group children float at the group origin plus scaled group coordinates
    const placeVmlBox = (
      box: TextboxDisplay,
      style: string,
      scale: VmlGroupScale | null,
      origin: VmlOrigin | null,
    ): void => {
      if (!scale && /position:absolute/.test(style)) {
        box.floating = true
        const mx = parseFloat(/margin-left:(-?[\d.]+)pt/.exec(style)?.[1] ?? '')
        const my = parseFloat(/margin-top:(-?[\d.]+)pt/.exec(style)?.[1] ?? '')
        if (Number.isFinite(mx)) box.offsetXEmu = Math.round(mx * EMU_PER_PT)
        if (Number.isFinite(my)) box.offsetYEmu = Math.round(my * EMU_PER_PT)
      } else if (scale && origin) {
        box.floating = true
        box.offsetXEmu = Math.round(vmlCoordPx(style, 'left', scale, origin) * EMU_PER_PX)
        box.offsetYEmu = Math.round(vmlCoordPx(style, 'top', scale, origin) * EMU_PER_PX)
      }
    }
    // VML picture shape (v:imagedata with an r:id): a photo box, like the
    // DrawingML pushPic path — dropping it loses real page content
    const vmlPicBox = (
      shape: XNode,
      scale: VmlGroupScale | null,
      origin: VmlOrigin | null,
    ): boolean => {
      const rId = attrsOf(findChild(shape, 'v:imagedata') ?? {})['r:id']
      const dataUrl = rId ? ctx.mediaByRid?.get(rId) : undefined
      if (!dataUrl) return false
      const style = attrsOf(shape)['style'] ?? ''
      const box: TextboxDisplay = {
        paras: [],
        readOnly: true,
        fillImageDataUrl: dataUrl,
        insetTopPx: 0,
        insetRightPx: 0,
        insetBottomPx: 0,
        insetLeftPx: 0,
      }
      const w = vmlShapeDimPx(style, 'width', scale)
      if (w) box.widthPx = w
      const h = vmlShapeDimPx(style, 'height', scale)
      if (h) box.heightPx = h
      placeVmlBox(box, style, scale, origin)
      out.push(box)
      return true
    }
    // textless VML geometry with a visible fill or stroke (callout tabs,
    // path-drawn table backdrops): rendered, not silently dropped
    const vmlGeomBox = (
      shape: XNode,
      scale: VmlGroupScale | null,
      origin: VmlOrigin | null,
    ): void => {
      const a = attrsOf(shape)
      const style = a['style'] ?? ''
      if (a['o:hr'] === 't' || /visibility:\s*hidden/.test(style)) return
      // a picture shape (t75) whose image did not resolve draws no frame
      if (a['o:spt'] === '75' || /_x0000_t75\b/.test(a['type'] ?? '')) return
      const fill = a['filled'] === 'f' ? undefined : vmlColorHex(a['fillcolor'])
      // VML strokes default on/black; like the textbox path, only canvas
      // (group) children get that default — top-level shapes need an explicit
      // strokecolor so converter placeholder furniture stays invisible
      const stroke =
        a['stroked'] === 'f'
          ? undefined
          : (vmlColorHex(a['strokecolor']) ?? (scale ? '000000' : undefined))
      // white unstroked placeholders draw nothing in Word (isInvisibleVmlPict)
      if ((!fill || fill.toUpperCase() === 'FFFFFF') && !stroke) return
      const w = vmlShapeDimPx(style, 'width', scale)
      const h = vmlShapeDimPx(style, 'height', scale)
      if (!w || !h) return
      const box: TextboxDisplay = {
        paras: [],
        readOnly: true,
        widthPx: w,
        heightPx: h,
        minHeightPx: h,
        insetTopPx: 0,
        insetRightPx: 0,
        insetBottomPx: 0,
        insetLeftPx: 0,
      }
      if (fill) box.fill = fill
      if (stroke) box.borderColor = stroke
      const name = nameOf(shape)
      if (name === 'v:roundrect') box.prst = 'roundRect'
      else if (name === 'v:oval') box.prst = 'ellipse'
      const path = a['path']
      if (path) {
        const cs = /^\s*(\d+)[,\s]+(\d+)/.exec(a['coordsize'] ?? '')
        const d = cs ? vmlPathToNormD(path, parseInt(cs[1]!, 10), parseInt(cs[2]!, 10)) : undefined
        // an unconvertible path must not degrade to a solid bounding box
        if (!d) return
        box.pathData = { path: d }
      }
      placeVmlBox(box, style, scale, origin)
      out.push(box)
    }
    // Children of a v:group (drawing canvas) position/size in the group's own
    // coordinate space (unitless style values); the scale maps them to px.
    const vmlBox = (
      shape: XNode,
      scale: VmlGroupScale | null,
      origin: VmlOrigin | null,
      nested: boolean,
    ): void => {
      const shapeAttrs = attrsOf(shape)
      const style = shapeAttrs['style'] ?? ''
      const contents: XNode[] = []
      collectNodes(childrenOf(shape), 'w:txbxContent', contents)
      if (contents.length === 0) {
        const wordArt = vmlWordArtBox(shape)
        if (wordArt) {
          out.push(wordArt)
          return
        }
        // a shape nested inside some textbox's content must not escape onto
        // the page as a sibling box
        if (nested) return
        if (vmlPicBox(shape, scale, origin)) return
        vmlGeomBox(shape, scale, origin)
        return
      }
      const topContents: XNode[] = []
      collectTopNodes(childrenOf(shape), 'w:txbxContent', topContents)
      const box: TextboxDisplay = { paras: [] }
      if (!nested) {
        box.txbxIndex = txbxOrdinal
        txbxOrdinal += topContents.length
      }
      if (nested || contents.length > topContents.length) box.readOnly = true
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
      // absolutely positioned VML shape: leaves the flow like a wp:anchor box
      // (page banners stack at full height otherwise, tdf 1194 family).
      // Canvas (v:group) children float at their scaled group coordinates.
      placeVmlBox(box, style, scale, origin)
      for (const content of contents)
        box.paras.push(...txbxContentParas(content, ctx, opts?.docOffset))
      if (contents.some(txbxHasStructuredContent)) box.readOnly = true
      if (box.paras.some((p) => p.runs.length > 0)) out.push(box)
    }
    const walkVml = (
      nodes: XNode[],
      scale: VmlGroupScale | null,
      origin: VmlOrigin | null,
      nested = false,
    ): void => {
      for (const node of nodes) {
        const name = nameOf(node)
        if (name === 'v:group') {
          const gScale = vmlGroupScale(node, scale)
          const gAttrs = attrsOf(node)
          const gStyle = gAttrs['style'] ?? ''
          let gOrigin: VmlOrigin | null = null
          if (gScale && scale && origin) {
            // nested group: placed in the parent group's coordinate space
            gOrigin = {
              x: vmlCoordPx(gStyle, 'left', scale, origin),
              y: vmlCoordPx(gStyle, 'top', scale, origin),
            }
          } else if (gScale && /position:absolute/.test(gStyle)) {
            const mx = parseFloat(/margin-left:(-?[\d.]+)pt/.exec(gStyle)?.[1] ?? '')
            const my = parseFloat(/margin-top:(-?[\d.]+)pt/.exec(gStyle)?.[1] ?? '')
            gOrigin = {
              x: Number.isFinite(mx) ? (mx * 96) / 72 : 0,
              y: Number.isFinite(my) ? (my * 96) / 72 : 0,
            }
          } else if (gScale && !nested) {
            // inline canvas: reserve its flow footprint so the floated
            // children overlay it instead of collapsing the paragraph
            const w = vmlShapeDimPx(gStyle, 'width', scale)
            const h = vmlShapeDimPx(gStyle, 'height', scale)
            if (w && h) {
              out.push({
                paras: [],
                readOnly: true,
                widthPx: w,
                heightPx: h,
                minHeightPx: h,
                insetTopPx: 0,
                insetRightPx: 0,
                insetBottomPx: 0,
                insetLeftPx: 0,
              })
              gOrigin = { x: 0, y: 0 }
            }
          }
          // children position from coordorigin, not 0,0 (drawing canvases)
          if (gScale && gOrigin) {
            const co = /^\s*(-?\d+)[,\s]+(-?\d+)/.exec(gAttrs['coordorigin'] ?? '')
            if (co) {
              gOrigin = {
                x: gOrigin.x - parseInt(co[1]!, 10) * gScale.sx,
                y: gOrigin.y - parseInt(co[2]!, 10) * gScale.sy,
              }
            }
          }
          walkVml(childrenOf(node), gScale ?? scale, gScale ? gOrigin : null, nested)
          continue
        }
        if (
          name === 'v:shape' ||
          name === 'v:rect' ||
          name === 'v:roundrect' ||
          name === 'v:oval'
        ) {
          vmlBox(node, scale, origin, nested)
        }
        // below a w:txbxContent = inside some box: consumes no save ordinal
        walkVml(childrenOf(node), scale, origin, nested || name === 'w:txbxContent')
      }
    }
    walkVml(parsed, null, null)
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
    const line = lineTwipsOf(attrs['w:line'])
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
    // autospacing=1: Word ignores the literal and uses its HTML auto value (14pt,
    // measured font-size independent; adjacent auto margins collapse, and they
    // collapse to 0 between two list items — the renderer applies those rules).
    // Tri-state: an explicit "0" overrides a style-chain auto. The literal is
    // still modeled so unchanged spacing round-trips byte-for-byte.
    const autoOf = (v: string | undefined): boolean | undefined =>
      v === undefined ? undefined : v === '1' || v === 'true'
    const autoBefore = autoOf(attrs['w:beforeAutospacing'])
    if (autoBefore !== undefined) format.spaceBeforeAuto = autoBefore
    const autoAfter = autoOf(attrs['w:afterAutospacing'])
    if (autoAfter !== undefined) format.spaceAfterAuto = autoAfter
    const before = parseInt(attrs['w:before'] ?? '', 10)
    if (before >= 0 && attrs['w:before'] !== undefined) format.spaceBefore = before
    const after = parseInt(attrs['w:after'] ?? '', 10)
    if (after >= 0 && attrs['w:after'] !== undefined) format.spaceAfter = after
  }
  const ind = findChild(pPr, 'w:ind')
  if (ind) {
    const attrs = attrsOf(ind)
    const left = parseInt(attrs['w:left'] ?? attrs['w:start'] ?? '', 10)
    // Negative indent (hanging past the left margin) is legal, keep it; an
    // explicit 0 is kept too — it overrides a numbering-level indent in Word
    if (Number.isFinite(left)) format.indentLeft = left
    const right = parseInt(attrs['w:right'] ?? attrs['w:end'] ?? '', 10)
    if (Number.isFinite(right) && right !== 0) format.indentRight = right
    const firstLine = parseInt(attrs['w:firstLine'] ?? '', 10)
    const hanging = parseInt(attrs['w:hanging'] ?? '', 10)
    if (hanging > 0) format.indentFirstLine = -hanging
    else if (firstLine > 0) format.indentFirstLine = firstLine
  }
  {
    // tri-state: an explicit w:val="0" must override a style-chain true
    const pbb = onOffOf(pPr, 'w:pageBreakBefore')
    if (pbb !== undefined) format.pageBreakBefore = pbb
  }
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
  {
    // tri-state: an explicit w:val="0" must override a style-chain true (Word
    // honors the direct spacing again — deliberate blank-page/-1-page driver)
    const ctx = onOffOf(pPr, 'w:contextualSpacing')
    if (ctx !== undefined) format.contextualSpacing = ctx
  }
  const autoSpace = autoSpaceOf(pPr)
  if (autoSpace !== undefined) format.autoSpace = autoSpace
  const shd = findChild(pPr, 'w:shd')
  if (shd) {
    const fill = attrsOf(shd)['w:fill']
    if (fill && fill !== 'auto') format.shadingFill = stripHash(fill)
    // pattern shading (pctNN/stripes): white 20pt text on w:shd pct70 was
    // invisible with the fill-only read (real_run2/61 Font Cascade)
    const display = shdDisplayFill(shd)
    if (display && display !== format.shadingFill) format.shadingDisplay = display
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
      if (attrs?.['w:color'] && attrs['w:color'] !== 'auto')
        line.color = stripHash(attrs['w:color'])
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
  // paragraph-style hidden text (w:vanish) inherits into runs without their own
  // flag; style-less paragraphs read the default paragraph style (real_run2/93:
  // a default style carrying vanish hides every run without an explicit off)
  const paraVanish = pStyleId
    ? ctx.styles?.get(pStyleId)?.display?.vanish
    : defaultParaVanish(ctx.styles)
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
          } else if (fieldCachedRuns.length > 0) {
            // any other field (HYPERLINK with switches, MERGEFIELD mail-merge
            // labels...): the cached result is still the visible text — keep it
            // as plain runs. Body paragraphs with these fields stay on the
            // passthrough path (onlyXeFields rejects them), so this only feeds
            // display-only contexts such as table cells and textbox content.
            for (const cached of fieldCachedRuns) pushRun(cached, rev)
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
          ctx.xmlSpacePreserve,
          paraVanish,
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
    // Run.image is singular: a run carrying several drawings/picts must split
    // into one run per picture, or every picture past the first is dropped —
    // and permanently lost from the file once the paragraph is edited
    const nodes =
      withImages &&
      childrenOf(node).filter((c) => IMAGE_RUN_CHILDREN.has(nameOf(c) ?? '')).length > 1
        ? splitImageRun(node)
        : [node]
    for (const part of nodes) {
      const run = buildRun(
        part,
        link,
        ctx.themeColors,
        ctx.themeFonts,
        withImages ? ctx.mediaByRid : undefined,
        ctx.styles,
        paraRtl,
        ctx.xmlSpacePreserve,
        paraVanish,
      )
      if (run) pushRun(run, rev)
    }
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
      } else if (name === 'w:fldSimple') {
        // single-element field form: the children are the cached result runs
        // (Word shows them until the field refreshes) — MERGEFIELD address
        // labels, DATE stamps... Dropping them blanks mail-merge documents.
        // Skip only inside an enclosing complex field's instruction phase.
        if (fieldDepth === 0 || fieldSeparated) walk(childrenOf(node), link, rev)
      } else if (name === 'w:br') {
        // Word honors a <w:br> sitting outside any <w:r> (direct child of w:p / w:ins)
        pushRun({ text: BREAK_CHAR[attrsOf(node)['w:type'] ?? ''] ?? '\n' }, rev)
      }
    }
  }
  walk(childrenOf(pNode))
  return mergeRuns(runs)
}

function buildRun(
  rNode: XNode,
  link?: Run['link'],
  theme?: ThemeColors | null,
  themeFonts?: ThemeFonts | null,
  mediaByRid?: Map<string, string>,
  styles?: Map<string, StyleInfo>,
  paraRtl?: boolean,
  partPreserve?: boolean,
  paraVanish?: boolean,
): Run | null {
  let text = ''
  for (const child of childrenOf(rNode)) {
    const name = nameOf(child)
    if (name === 'w:t' || name === 'w:delText') {
      const raw = decodeNumericCharRefs(textOf(child))
      // without xml:space="preserve" in scope (own attribute or part root) Word
      // drops the element's leading/trailing XML whitespace (pretty-printed
      // documents carry literal newlines + tabs)
      const own = attrsOf(child)['xml:space']
      text +=
        own === 'preserve' || (own === undefined && partPreserve)
          ? raw
          : raw.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '')
    } else if (name === 'w:tab' || name === 'w:ptab') text += '\t'
    // In-paragraph page breaks (w:br w:type="page") are encoded as \f and preserved; column/soft breaks become \n
    else if (name === 'w:br') text += BREAK_CHAR[attrsOf(child)['w:type'] ?? ''] ?? '\n'
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
    // Word wraps modern drawings in mc:AlternateContent (Choice + VML twin)
    const choice = findChild(findChild(rNode, 'mc:AlternateContent') ?? {}, 'mc:Choice')
    const drawing = findChild(rNode, 'w:drawing') ?? (choice && findChild(choice, 'w:drawing'))
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
        const border = picBorderOf(drawingXml)
        if (border) image.border = border
        // anchored (floating) picture kept as a run image: carry the wrap kind
        // and anchor offsets so the editor can float it instead of inlining
        if (/<wp:anchor[\s>]/.test(drawingXml)) {
          const meta = imageMeta(drawingXml)
          if (meta.imageWrap) image.wrap = meta.imageWrap
          if (meta.imageOffsetXEmu !== undefined) image.offsetXEmu = meta.imageOffsetXEmu
          if (meta.imageOffsetYEmu !== undefined) image.offsetYEmu = meta.imageOffsetYEmu
          if (meta.imageWrapDistTopEmu !== undefined)
            image.wrapDistTopEmu = meta.imageWrapDistTopEmu
          if (meta.imageWrapDistBottomEmu !== undefined)
            image.wrapDistBottomEmu = meta.imageWrapDistBottomEmu
          if (meta.imageWrapDistLeftEmu !== undefined)
            image.wrapDistLeftEmu = meta.imageWrapDistLeftEmu
          if (meta.imageWrapDistRightEmu !== undefined)
            image.wrapDistRightEmu = meta.imageWrapDistRightEmu
          // Word centers the object on the anchor line (tdf#162551: the picture
          // juts above the line rather than hanging below it)
          if (
            /<wp:positionV[^>]*relativeFrom="line"[^>]*>\s*<wp:align>center<\/wp:align>/.test(
              drawingXml,
            )
          )
            image.lineCenterV = true
          if (meta.imageNoOverlap) image.noOverlap = true
        }
      }
    }
    // legacy VML picture (w:pict + v:imagedata, Word 2003 era / stamps) or an
    // inline OLE embed (w:object, whose preview is also a v:imagedata): same
    // run-level image treatment; the fragment round-trips verbatim on save
    if (!image) {
      // a run can carry several (an empty w:pict next to the real w:object):
      // take the first with a resolvable preview picture
      const picts = [
        findChild(rNode, 'w:pict'),
        findChild(rNode, 'w:object'),
        ...(choice ? [findChild(choice, 'w:pict'), findChild(choice, 'w:object')] : []),
      ].filter((n): n is XNode => n !== undefined)
      for (const pict of picts) {
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
          break
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
  if (!rPr && paraVanish) run.vanish = true
  if (rPr) {
    run.rawRPr = serializeXNode(rPr)
    const rStyle = attrsOf(findChild(rPr, 'w:rStyle') ?? {})['w:val']
    if (rStyle && rStyle !== 'Hyperlink') run.styleId = rStyle
    // hidden text: an explicit run w:vanish wins over the character/paragraph
    // style chain; w:specVanish (style separator) keeps the run visible
    const vanishOwn = onOffOf(rPr, 'w:specVanish') === true ? undefined : onOffOf(rPr, 'w:vanish')
    const vanish =
      vanishOwn ?? (rStyle ? styles?.get(rStyle)?.display?.vanish : undefined) ?? paraVanish
    if (vanish === true) run.vanish = true
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
    const rfAttrs = attrsOf(findChild(rPr, 'w:rFonts') ?? {})
    const rf = themedRFonts(rfAttrs, themeFonts)
    const font = rf.eastAsia ?? rf.ascii ?? rf.hAnsi
    if (font) run.font = font
    if (rf.eaSlotEmpty && font && font === rf.eastAsia) run.eaSlotEmpty = true
    const fontAscii = rf.ascii ?? rf.hAnsi
    if (fontAscii) run.fontAscii = fontAscii
    // record which resolved values came from theme refs (per winning slot)
    const fontThemed =
      rf.eastAsia !== undefined
        ? rf.themed?.eastAsia
        : rf.ascii !== undefined
          ? rf.themed?.ascii
          : rf.themed?.hAnsi
    const fontAsciiThemed = rf.ascii !== undefined ? rf.themed?.ascii : rf.themed?.hAnsi
    if ((fontThemed && font) || (fontAsciiThemed && fontAscii)) {
      run.themeRFonts = {
        ...(fontThemed && font ? { font } : {}),
        ...(fontAsciiThemed && fontAscii ? { fontAscii } : {}),
      }
    }
    // complex-script slot: literal attribute only — theme refs (w:cstheme) stay in
    // rawRPr so untouched runs keep their original bytes
    if (rfAttrs['w:cs']) run.fontCs = rfAttrs['w:cs']
    // theme-resolved cs font for display consumers
    if (rf.cs) run.csFont = rf.cs
    const rtl = onOffOf(rPr, 'w:rtl')
    if (rtl !== undefined) run.rtl = rtl
    const spc = parseInt(attrsOf(findChild(rPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
    if (spc) run.charSpacingTwips = spc
    // w:caps wins over w:smallCaps when both are on (Word)
    const capsOn = onOffOf(rPr, 'w:caps')
    const smallCapsOn = onOffOf(rPr, 'w:smallCaps')
    if (capsOn) run.caps = 'all'
    else if (smallCapsOn) run.caps = 'small'
    else if (capsOn === false || smallCapsOn === false) run.caps = 'none'
    const wScale = parseInt(attrsOf(findChild(rPr, 'w:w') ?? {})['w:val'] ?? '', 10)
    if (wScale > 0 && wScale !== 100) run.charScalePct = wScale
    const highlight = attrsOf(findChild(rPr, 'w:highlight') ?? {})['w:val']
    if (highlight && highlight !== 'none') run.highlight = highlight
    const shdFill = attrsOf(findChild(rPr, 'w:shd') ?? {})['w:fill']
    if (shdFill && shdFill !== 'auto') run.shading = stripHash(shdFill)
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
      delete run.fontCs
      if (run.rawRPr) run.rawRPr = run.rawRPr.replace(/<w:rFonts[^>]*\/>/, '')
    }
  }
  return run
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
 *
 * @param docOffset the table's document.xml offset: cell paragraphs resolve
 *   their character-unit indents under that section's document grid
 */
function extractTable(xml: string, ctx: BuildContext, docOffset?: number): TableModel | undefined {
  // whole try: hostile depth inside a cell paragraph can overflow the
  // run-extraction recursion — degrade to a protected block, not a failed document
  try {
    const parsed = deepXmlParser.parse(xml) as XNode[]
    const tbl = parsed.find((n) => nameOf(n) === 'w:tbl')
    if (!tbl) return undefined
    const model = extractTableModel(tbl, ctx, 1, docOffset)
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
    const edges = rowGridEdges(tr)
    let idx = edges.before
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
    colCount = Math.max(colCount, idx + edges.after)
  }
  if (colCount === 0) return undefined
  for (let i = 0; i < colCount; i++) if (!(cols[i] > 0)) return undefined
  return cols.slice(0, colCount)
}

/** trPr w:gridBefore/w:gridAfter column counts plus their w:wBefore/w:wAfter widths (twips) */
function rowGridEdges(tr: XNode): {
  before: number
  after: number
  wBefore?: number
  wAfter?: number
} {
  const trPr = findChild(tr, 'w:trPr')
  if (!trPr) return { before: 0, after: 0 }
  const count = (name: string) => {
    const v = Number(attrsOf(findChild(trPr, name) ?? {})['w:val'])
    return Number.isFinite(v) && v > 0 ? Math.floor(v) : 0
  }
  const width = (name: string) => {
    const a = attrsOf(findChild(trPr, name) ?? {})
    const w = !a['w:type'] || a['w:type'] === 'dxa' ? Number(a['w:w']) : NaN
    return w > 0 ? w : undefined
  }
  return {
    before: count('w:gridBefore'),
    after: count('w:gridAfter'),
    wBefore: width('w:wBefore'),
    wAfter: width('w:wAfter'),
  }
}

/** boundary positions within TOL twips fuse into one grid line (rounding drift across generator-written rows) */
const GRID_SNAP_TOL = 20

/**
 * Repair grids whose rows do not all span the declared column count (legacy
 * generators emit gridSpan/tblGrid pairs that disagree with the cells' w:tcW).
 * Word lays each such row out from the cells' preferred widths, so the true
 * grid is the union of every row's tcW boundaries; without this, a cell mapped
 * onto a leftover sliver column collapses to one character per line.
 * Mutates the cells' colSpan and returns the rebuilt column widths, or
 * undefined when the grid is already consistent or a width is unresolvable.
 */
function reconcileGridColumns(
  rows: TableCell[][],
  rowTcws: Array<Array<number | undefined>>,
  gridCols: number[] | undefined,
): number[] | undefined {
  const spanSums = rows.map((row) => row.reduce((sum, c) => sum + (c.colSpan ?? 1), 0))
  const colCount = gridCols?.length ?? Math.max(...spanSums)
  if (spanSums.every((sum) => sum === colCount)) return undefined
  const rowBounds: number[][] = []
  for (let r = 0; r < rows.length; r++) {
    const bounds: number[] = []
    let x = 0
    let pos = 0
    for (let c = 0; c < rows[r].length; c++) {
      const span = rows[r][c].colSpan ?? 1
      let w = rowTcws[r][c]
      if (!(w !== undefined && w > 0)) {
        w = gridCols?.slice(pos, pos + span).reduce((sum, v) => sum + v, 0)
      }
      if (!(w !== undefined && w > 0)) return undefined
      x += w
      bounds.push(x)
      pos += span
    }
    rowBounds.push(bounds)
  }
  const sorted = rowBounds.flat().sort((a, b) => a - b)
  const reps: number[] = []
  for (const b of sorted) {
    if (reps.length === 0 || b - reps[reps.length - 1] > GRID_SNAP_TOL) reps.push(b)
  }
  // Word caps table grids at 63 columns; a wider union means garbage input
  if (reps.length === 0 || reps.length > 96) return undefined
  const repIndex = (b: number) => {
    let lo = 0
    let hi = reps.length - 1
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1
      if (reps[mid] <= b) lo = mid
      else hi = mid - 1
    }
    return reps[lo] <= b && b - reps[lo] <= GRID_SNAP_TOL ? lo : -1
  }
  const newSpans: number[][] = []
  for (const bounds of rowBounds) {
    const spans: number[] = []
    let prev = -1
    for (const b of bounds) {
      const idx = repIndex(b)
      if (idx <= prev) return undefined
      spans.push(idx - prev)
      prev = idx
    }
    newSpans.push(spans)
  }
  rows.forEach((row, r) =>
    row.forEach((cell, c) => {
      if (newSpans[r][c] > 1) cell.colSpan = newSpans[r][c]
      else delete cell.colSpan
    }),
  )
  return reps.map((v, i) => v - (i > 0 ? reps[i - 1] : 0))
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

function extractTableModel(
  tbl: XNode,
  ctx: BuildContext,
  depth = 1,
  docOffset?: number,
): TableModel | undefined {
  const grid = findChild(tbl, 'w:tblGrid')
  let colWidthsPct: number[] | undefined
  let colWidthsTwips: number[] | undefined
  let gridWidthsRaw: number[] | undefined
  if (grid) {
    const widths = findChildren(grid, 'w:gridCol').map((c) => Number(attrsOf(c)['w:w']) || 0)
    const total = widths.reduce((a, b) => a + b, 0)
    if (total > 0) {
      gridWidthsRaw = widths
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
  // Placeholder grids (generators emit 100-twip gridCols without computing a
  // layout) collapse the table to a strip. Word lays autofit tables out from
  // tblW/tcW, so a grid summing below the declared dxa width stretches to it.
  if (!fixedLayout && colWidthsTwips) {
    const tblWDxa = !tblW['w:type'] || tblW['w:type'] === 'dxa' ? Number(tblW['w:w']) : NaN
    const gridTotal = colWidthsTwips.reduce((a, b) => a + b, 0)
    if (tblWDxa > 0 && gridTotal > 0 && gridTotal < tblWDxa - colWidthsTwips.length) {
      const scale = tblWDxa / gridTotal
      colWidthsTwips = colWidthsTwips.map((w) => Math.round(w * scale))
    }
  }
  const cellMar = cellMarginsOf(findChild(tblPrNode ?? {}, 'w:tblCellMar'))
  const tblBorders = mergedBorderLinesOf(tblPrNode, 'w:tblBorders', true)
  const tblJc = attrsOf(findChild(tblPrNode ?? {}, 'w:jc') ?? {})['w:val']
  const tblAlign =
    tblJc === 'center' ? 'center' : tblJc === 'right' || tblJc === 'end' ? 'right' : undefined
  const tblInd = attrsOf(findChild(tblPrNode ?? {}, 'w:tblInd') ?? {})
  const tblIndTwips = !tblInd['w:type'] || tblInd['w:type'] === 'dxa' ? Number(tblInd['w:w']) : NaN
  const tblpPr = attrsOf(findChild(tblPrNode ?? {}, 'w:tblpPr') ?? {})
  let floatSide: TableModel['floatSide']
  let floatPos: TableModel['floatPos']
  if (Object.keys(tblpPr).length > 0) {
    const xSpec = tblpPr['w:tblpXSpec']
    // no alignment keyword: an absolute X past mid-body (~9360 twips of usable
    // width on Letter/A4) means the table hugs the right side
    floatSide =
      xSpec === 'right' || xSpec === 'outside' || (!xSpec && Number(tblpPr['w:tblpX']) > 4680)
        ? 'right'
        : 'left'
    const x = Number(tblpPr['w:tblpX'])
    const y = Number(tblpPr['w:tblpY'])
    const distanceTwips: CellMargins = {}
    const distanceAttrs = {
      top: 'w:topFromText',
      right: 'w:rightFromText',
      bottom: 'w:bottomFromText',
      left: 'w:leftFromText',
    } as const
    for (const [side, attr] of Object.entries(distanceAttrs) as Array<
      [keyof CellMargins, (typeof distanceAttrs)[keyof typeof distanceAttrs]]
    >) {
      const value = Number(tblpPr[attr])
      if (Number.isFinite(value) && value >= 0) distanceTwips[side] = value
    }
    const horzAnchor = tblpPr['w:horzAnchor']
    const vertAnchor = tblpPr['w:vertAnchor']
    floatPos = {
      xTwips: Number.isFinite(x) ? x : floatSide === 'right' ? 9360 : 0,
      yTwips: Number.isFinite(y) ? y : 0,
      ...(horzAnchor === 'page' || horzAnchor === 'margin' || horzAnchor === 'text'
        ? { horzAnchor }
        : {}),
      ...(vertAnchor === 'page' || vertAnchor === 'margin' || vertAnchor === 'text'
        ? { vertAnchor }
        : {}),
      ...(Object.keys(distanceTwips).length > 0 ? { distanceTwips } : {}),
    }
  }

  const spacingOf = (node: XNode | undefined) => {
    const a = attrsOf(findChild(node ?? {}, 'w:tblCellSpacing') ?? {})
    const w = !a['w:type'] || a['w:type'] === 'dxa' ? Number(a['w:w']) : NaN
    return w > 0 ? w : undefined
  }
  let cellSpacing = spacingOf(tblPrNode)
  const tblFill = shdDisplayFill(findChild(tblPrNode ?? {}, 'w:shd'))

  const rows: TableCell[][] = []
  // per-cell w:tcW (dxa twips), aligned with rows: grid-repair input
  const rowTcws: Array<Array<number | undefined>> = []
  const rowEdges: Array<ReturnType<typeof rowGridEdges>> = []
  const rowHeightsTwips: Array<number | null> = []
  const rowHeightRules: NonNullable<TableModel['rowHeightRules']> = []
  const repeatHeaderRows: boolean[] = []
  const rowRevisions: NonNullable<TableModel['rowRevisions']> = []
  for (const tr of childrenThroughSdt(tbl, 'w:tr')) {
    const cells: TableCell[] = []
    const tcws: Array<number | undefined> = []
    for (const tc of childrenThroughSdt(tr, 'w:tc')) {
      const cell = extractCell(tc, ctx, depth, docOffset)
      const a = attrsOf(findChildren(findChild(tc, 'w:tcPr') ?? {}, 'w:tcW').at(-1) ?? {})
      const rawW = !a['w:type'] || a['w:type'] === 'dxa' ? Number(a['w:w']) : NaN
      const tcw = rawW > 0 ? rawW : undefined
      const prev = cells[cells.length - 1]
      // Legacy horizontal merge: continue cells fold into the cell to their left (same effect
      // as gridSpan)
      if (cell.hMerge === 'continue' && prev) {
        prev.colSpan = (prev.colSpan ?? 1) + (cell.colSpan ?? 1)
        const prevW = tcws[tcws.length - 1]
        tcws[tcws.length - 1] = prevW !== undefined && tcw !== undefined ? prevW + tcw : undefined
        continue
      }
      cells.push(cell)
      tcws.push(tcw)
    }
    if (cells.length > 0) {
      rows.push(cells)
      rowTcws.push(tcws)
      rowEdges.push(rowGridEdges(tr))
      const trPr = findChild(tr, 'w:trPr')
      cellSpacing = cellSpacing ?? spacingOf(trPr)
      const trH = trPr ? attrsOf(findChild(trPr, 'w:trHeight') ?? {}) : {}
      const h = Number(trH['w:val'])
      const hasH = Number.isFinite(h) && h > 0
      // Word clamps trHeight to 31680 twips / 22in (MS-OI29500 2.1.51); some generators leak EMU-scale values here
      rowHeightsTwips.push(hasH ? Math.min(h, 31680) : null)
      rowHeightRules.push(hasH ? (trH['w:hRule'] === 'exact' ? 'exact' : 'atLeast') : null)
      repeatHeaderRows.push(trPr ? boolProp(trPr, 'w:tblHeader') : false)
      rowRevisions.push(trPr ? rowRevisionOf(trPr) : null)
    }
  }
  if (rows.length === 0) return undefined
  applyTableStyleDisplay(rows, findChild(tbl, 'w:tblPr'), ctx)
  // w:gridBefore/w:gridAfter offset rows within the grid: placeholder cells keep
  // the offset visible (inserted after style banding so real cells keep their look)
  if (rowEdges.some((e) => e.before > 0 || e.after > 0)) {
    rows.forEach((cells, i) => {
      const e = rowEdges[i]
      if (e.before > 0) {
        cells.unshift({ paras: [], gridGap: true, ...(e.before > 1 ? { colSpan: e.before } : {}) })
        rowTcws[i].unshift(e.wBefore)
      }
      if (e.after > 0) {
        cells.push({ paras: [], gridGap: true, ...(e.after > 1 ? { colSpan: e.after } : {}) })
        rowTcws[i].push(e.wAfter)
      }
    })
  }
  const reconciled = reconcileGridColumns(rows, rowTcws, gridWidthsRaw)
  if (reconciled) {
    colWidthsTwips = reconciled
    const total = reconciled.reduce((a, b) => a + b, 0)
    colWidthsPct = reconciled.map((w) => (w / total) * 100)
  }
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
  // pct widths still lay out with the autofit algorithm (only w:tblLayout
  // fixed switches it off), so their columns keep the min-content floor
  if (!fixedLayout && (autoWidth || widthPct)) model.autoLayout = true
  model.autoFit =
    fixedLayout || (!autoWidth && !widthPct) ? 'fixed' : widthPct === 100 ? 'window' : 'contents'
  if (fixedLayout) model.fixedLayout = true
  if (effCellMar) model.cellMarTwips = effCellMar
  if (cellSpacing) model.cellSpacingTwips = cellSpacing
  if (tblFill) model.fill = tblFill
  if (effBorders) model.borders = effBorders
  if (tblAlign) model.align = tblAlign
  if (floatSide) model.floatSide = floatSide
  if (floatPos) model.floatPos = floatPos
  if (Number.isFinite(tblIndTwips) && tblIndTwips !== 0) model.indentTwips = tblIndTwips
  const tblStyle = attrsOf(findChild(findChild(tbl, 'w:tblPr') ?? {}, 'w:tblStyle') ?? {})['w:val']
  if (tblStyle) model.tblStyleId = tblStyle
  model.tableLook = tableLookOf(tblPrNode)
  if (tblPrNode && boolProp(tblPrNode, 'w:bidiVisual')) model.bidiVisual = true
  if (rowHeightsTwips.some((h) => h !== null)) {
    model.rowHeightsTwips = rowHeightsTwips
    model.rowHeightRules = rowHeightRules
  }
  model.repeatHeaderRows = repeatHeaderRows
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
    // gridGap placeholders have no w:tc in the source: align against real cells only
    const targets = rows[ri].filter((cell) => !cell.gridGap)
    if (tcs.length !== targets.length) return
    tcs.forEach((tc, ci) => {
      const tcOpenEnd = tc.xml.indexOf('>') + 1
      const tcInner = tc.xml.slice(tcOpenEnd, tc.xml.lastIndexOf('</w:tc>'))
      const tcPr = splitXmlChildren(tcInner).find((k) => k.name === 'w:tcPr')
      if (tcPr) targets[ci].rawTcPr = tcPr.xml
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

function extractCell(tc: XNode, ctx: BuildContext, depth: number, docOffset?: number): TableCell {
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
    const fill = shdDisplayFill(findChild(tcPr, 'w:shd'))
    if (fill) cell.fill = fill
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
          : extractTableModel(block, ctx, depth + 1, docOffset)
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
        ? rawPXml.replace(/<mc:Fallback[^>]*>[\s\S]*?<\/mc:Fallback>/g, '')
        : rawPXml
      if (pXml.includes('<wp:anchor') || /<w:pict[\s>]/.test(pXml)) {
        const boxes = extractTextboxes(pXml, ctx, { shapes: true, docOffset })
        if (boxes.length > 0) {
          cell.anchoredBoxes = [...(cell.anchoredBoxes ?? []), ...boxes]
          // positionV relativeFrom="paragraph" measures from the anchor
          // paragraph, not the cell top: remember which paragraph hosts each box
          cell.anchoredBoxAnchors = [
            ...(cell.anchoredBoxAnchors ?? []),
            ...boxes.map(() => cell.paras.length),
          ]
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
    const format = withCharIndents(extractParaFormat(pPr ?? {}), ctx, p, pPr, cellStyleId, runs, {
      list: !!list,
      docOffset,
    })
    const emptySz = runs.length === 0 ? emptyParaSizeHalfPoints(p, pPr) : undefined
    const emptyFont = runs.length === 0 ? emptyParaMarkFont(p, pPr) : undefined
    richParas.push({
      ...format,
      ...(cellStyleId ? { styleId: cellStyleId } : {}),
      ...(emptySz ? { emptyRunSizeHalfPoints: emptySz } : {}),
      ...(emptyFont ? { emptyRunFontFamily: emptyFont } : {}),
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
  if (
    !partXml.includes('<a:blip') &&
    !partXml.includes('<v:imagedata') &&
    !partXml.includes('<wps:wsp')
  ) {
    return []
  }
  const relsPath = partPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const rels = await parseRels(zip, relsPath)
  // inline pictures in layout tables (nested ones included) live on their cell
  // runs (hfTableRowParagraphs); floating ones still position through this
  // part-level list
  const tbls = hfTblRanges(partXml)
  const onCellRun = (at: number) => tbls.some(([s, e]) => at > s && at < e)
  // mc:AlternateContent: the Choice is authoritative; its Fallback re-emits the
  // same picture (an anchored logo shape falls back to an inline copy), doubling
  // the drawn image and inflating the strip's reserved height (prod100r4/43).
  // Fallback content only counts when nothing in the same block's Choice
  // resolved (mac Word PDF Choice + PNG Fallback keeps working: those blips
  // share one w:drawing and are tried in order above).
  const acs = Array.from(
    partXml.matchAll(/<mc:AlternateContent[\s>][\s\S]*?<\/mc:AlternateContent>/g),
    (m) => {
      const fb = /<mc:Fallback>[\s\S]*?<\/mc:Fallback>/.exec(m[0])
      return {
        start: m.index!,
        end: m.index! + m[0].length,
        fbStart: fb ? m.index! + fb.index : -1,
        fbEnd: fb ? m.index! + fb.index + fb[0].length : -1,
      }
    },
  )
  const acAt = (at: number) => acs.findIndex((a) => at > a.start && at < a.end)
  const inFallbackOf = (at: number) => acs.findIndex((a) => at > a.fbStart && at < a.fbEnd)
  const choiceProduced = new Set<number>()
  /** true = skip this match (its Choice sibling already produced an image) */
  const fallbackDup = (at: number): boolean => {
    const fb = inFallbackOf(at)
    return fb >= 0 && choiceProduced.has(fb)
  }
  const recordProduced = (at: number) => {
    const ac = acAt(at)
    if (ac >= 0 && inFallbackOf(at) < 0) choiceProduced.add(ac)
  }
  const out: HfImage[] = []
  /** w:jc of the paragraph containing offset `at` (inline images follow it) */
  const paraAlignAt = (at: number): HfImage['align'] => {
    const pStart = Math.max(partXml.lastIndexOf('<w:p ', at), partXml.lastIndexOf('<w:p>', at))
    if (pStart < 0) return undefined
    const jc = /<w:jc w:val="(\w+)"/.exec(partXml.slice(pStart, at))?.[1]
    return jc === 'left' || jc === 'center' || jc === 'right' ? jc : undefined
  }
  for (const m of partXml.matchAll(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/g)) {
    const frag = m[0]
    if (fallbackDup(m.index!)) continue
    if (!/<wp:anchor[\s>]/.test(frag) && onCellRun(m.index!)) continue
    // mc:AlternateContent may hold several blips (mac Word: PDF Choice + PNG
    // Fallback); use the first one whose media resolves
    let dataUrl: string | null = null
    for (const b of frag.matchAll(/<a:blip[^>]*r:embed="([^"]+)"/g)) {
      dataUrl = await mediaDataUrl(zip, rels, b[1])
      if (dataUrl) break
    }
    // textless vector decorations (wpg group of custGeom shapes) render as one SVG
    if (!dataUrl) dataUrl = hfShapeDrawingSvg(frag)
    if (!dataUrl) continue
    const image: HfImage = { dataUrl }
    const extent = /<wp:extent[^>]*\/?>/.exec(frag)?.[0] ?? ''
    const cx = parseInt(/cx="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    const cy = parseInt(/cy="(\d+)"/.exec(extent)?.[1] ?? '', 10)
    if (Number.isFinite(cx) && cx > 0) image.widthPx = Math.round(cx / EMU_PER_PX)
    if (Number.isFinite(cy) && cy > 0) image.heightPx = Math.round(cy / EMU_PER_PX)
    // a:srcRect source crop: two same-image anchors cropped to different
    // regions read as duplicated pictures without it (prod100r1 sample 90)
    const srcRect = /<a:srcRect\s[^>]*\/>/.exec(frag)?.[0]
    if (srcRect) {
      const crop = {
        l: rectFrac(srcRect, 'l'),
        t: rectFrac(srcRect, 't'),
        r: rectFrac(srcRect, 'r'),
        b: rectFrac(srcRect, 'b'),
      }
      if (crop.l || crop.t || crop.r || crop.b) image.crop = crop
    }
    if (/<wp:anchor[\s>]/.test(frag)) {
      image.floating = true
      const anchorTag = /<wp:anchor[^>]*>/.exec(frag)?.[0] ?? ''
      if (/behindDoc="(?:1|true)"/.test(anchorTag)) image.behind = true
      const wrap = /<wp:wrap(None|Square|Tight|Through|TopAndBottom)[\s/>]/.exec(frag)?.[1]
      if (wrap) {
        image.wrap = wrap === 'TopAndBottom' ? 'topBottom' : (wrap.toLowerCase() as HfImage['wrap'])
      }
      readAnchorPos(frag, image)
    } else {
      const align = paraAlignAt(m.index!)
      if (align) image.align = align
    }
    recordProduced(m.index!)
    out.push(image)
  }
  for (const m of partXml.matchAll(/<w:pict[\s>][\s\S]*?<\/w:pict>/g)) {
    const frag = m[0]
    if (fallbackDup(m.index!)) continue
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
    recordProduced(m.index!)
    out.push(image)
  }
  return out
}

/** wp:anchor wp:positionH/V of a header/footer image: wp:align keeps the VML-style
 *  alignment fields, wp:posOffset (EMU) becomes a px offset from the page edge or
 *  margin box (horizontal paragraph/column/character origins approximate to margin;
 *  vertical paragraph/line keeps 'paragraph' so body push-down can measure from the
 *  header strip top). */
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
      if (axis === 'H') {
        image.posXPx = px
        image.posHRel = m[1] === 'page' ? 'page' : 'margin'
      } else {
        image.posYPx = px
        image.posVRel =
          m[1] === 'page'
            ? 'page'
            : m[1] === 'paragraph' || m[1] === 'line'
              ? 'paragraph'
              : 'margin'
      }
    }
  }
}

/**
 * Textless vector decoration in a header/footer drawing (wpg group of solid-fill
 * custGeom wps shapes, e.g. corner ornament groups) composed into one SVG data
 * URL at the wp:extent size. Bails (null) on any unsupported piece — rotation,
 * flips, text content, missing fill/geometry — so partial art never renders.
 */
function hfShapeDrawingSvg(frag: string): string | null {
  if (!frag.includes('<wps:wsp') || frag.includes('<w:txbxContent')) return null
  // Word pairs DrawingML shapes with a VML twin in mc:Fallback
  const body = frag.replace(/<mc:Fallback[^>]*>[\s\S]*?<\/mc:Fallback>/g, '')
  const attrNum = (tag: string, key: string): number | null => {
    const v = new RegExp(`${key}="(-?\\d+)"`).exec(tag)?.[1]
    return v == null ? null : parseInt(v, 10)
  }
  const extent = /<wp:extent[^>]*\/?>/.exec(body)?.[0] ?? ''
  const extCx = attrNum(extent, 'cx') ?? 0
  const extCy = attrNum(extent, 'cy') ?? 0
  if (extCx <= 0 || extCy <= 0) return null
  if (/rot="-?[1-9]|flipH="(?:1|true)"|flipV="(?:1|true)"/.test(body)) return null
  // group child space -> drawing space (wpg:grpSpPr a:xfrm)
  let sx = 1
  let sy = 1
  let tx = 0
  let ty = 0
  const grpXfrm = /<wpg:grpSpPr[^>]*>[\s\S]*?<a:xfrm[^>]*>([\s\S]*?)<\/a:xfrm>/.exec(body)?.[1]
  if (grpXfrm) {
    const el = (name: string) => new RegExp(`<a:${name}[^>]*/>`).exec(grpXfrm)?.[0] ?? ''
    const ext = { x: attrNum(el('ext'), 'cx') ?? 0, y: attrNum(el('ext'), 'cy') ?? 0 }
    const chExt = { x: attrNum(el('chExt'), 'cx') ?? 0, y: attrNum(el('chExt'), 'cy') ?? 0 }
    sx = ext.x > 0 && chExt.x > 0 ? ext.x / chExt.x : 1
    sy = ext.y > 0 && chExt.y > 0 ? ext.y / chExt.y : 1
    const off = { x: attrNum(el('off'), 'x') ?? 0, y: attrNum(el('off'), 'y') ?? 0 }
    const chOff = { x: attrNum(el('chOff'), 'x') ?? 0, y: attrNum(el('chOff'), 'y') ?? 0 }
    tx = off.x - chOff.x * sx
    ty = off.y - chOff.y * sy
  }
  const px = (emu: number) => Math.round((emu / EMU_PER_PX) * 100) / 100
  /** normalized 0..1 path tokens -> px path inside the shape's rect */
  const placePath = (d: string, x: number, y: number, w: number, h: number): string => {
    let axis = 0
    return d
      .split(' ')
      .map((tok) => {
        const n = Number(tok)
        if (!Number.isFinite(n)) {
          axis = 0
          return tok
        }
        return String(
          axis++ % 2 === 0
            ? Math.round((x + n * w) * 100) / 100
            : Math.round((y + n * h) * 100) / 100,
        )
      })
      .join(' ')
  }
  const paths: string[] = []
  for (const s of body.matchAll(/<wps:wsp[\s>][\s\S]*?<\/wps:wsp>/g)) {
    const wsp = s[0]
    const spPr = /<wps:spPr[\s\S]*?<\/wps:spPr>/.exec(wsp)?.[0] ?? ''
    const xfrm = /<a:xfrm[^>]*>[\s\S]*?<\/a:xfrm>/.exec(spPr)?.[0] ?? ''
    const off = /<a:off[^>]*\/>/.exec(xfrm)?.[0] ?? ''
    const ext = /<a:ext[^>]*\/>/.exec(xfrm)?.[0] ?? ''
    const cx = attrNum(ext, 'cx') ?? 0
    const cy = attrNum(ext, 'cy') ?? 0
    if (cx <= 0 || cy <= 0) return null
    const fill = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(spPr)?.[1]
    if (!fill) return null
    const geom = parseCustGeom(wsp, cx, cy)
    const d = [geom?.path, geom?.fillPath].filter(Boolean).join(' ')
    if (!d) return null
    const x = px((attrNum(off, 'x') ?? 0) * sx + tx)
    const y = px((attrNum(off, 'y') ?? 0) * sy + ty)
    paths.push(`<path d="${placePath(d, x, y, px(cx * sx), px(cy * sy))}" fill="#${fill}"/>`)
  }
  if (paths.length === 0) return null
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${px(extCx)} ${px(extCy)}">` +
    paths.join('') +
    '</svg>'
  return `data:image/svg+xml,${encodeURIComponent(svg)}`
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

/** settings.xml w:autoHyphenation + w:defaultTabStop (absent = Word's 720 twips) */
async function parseLayoutSettings(zip: JSZip): Promise<{
  autoHyphenation?: boolean
  defaultTabStopTwips?: number
  balanceDbcsSpacing?: boolean
  compressPunctuation?: boolean
  adjustLineHeightInTable?: boolean
}> {
  const file = zip.file('word/settings.xml')
  if (!file) return {}
  const xml = await file.async('string')
  const tab = /<w:defaultTabStop[^>]*w:val="(-?\d+)"/.exec(xml)
  const csc = /<w:characterSpacingControl[^>]*w:val="(\w+)"/.exec(xml)
  return {
    ...(xmlFlagOn(xml, 'w:autoHyphenation') ? { autoHyphenation: true } : {}),
    ...(tab ? { defaultTabStopTwips: parseInt(tab[1], 10) } : {}),
    ...(xmlFlagOn(xml, 'w:balanceSingleByteDoubleByteWidth') ? { balanceDbcsSpacing: true } : {}),
    ...(csc && csc[1].startsWith('compressPunctuation') ? { compressPunctuation: true } : {}),
    ...(xmlFlagOn(xml, 'w:adjustLineHeightInTable') ? { adjustLineHeightInTable: true } : {}),
  }
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
  let cleaned = xml.replace(/<mc:Fallback[^>]*>[\s\S]*?<\/mc:Fallback>/g, '')
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
  // legacy <w:pgNum/> run element (pre-field page number) renders as a PAGE field
  cleaned = cleaned.replace(/<w:pgNum\s*\/>/g, () => {
    hasPageNumber = true
    return `<w:t>${PAGE_MARK}</w:t>`
  })
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

/** Word merges style-chain and direct tab stops (direct wins per position; w:val="clear" removes) */
function mergeTabStops(
  style: import('./types').TabStop[] | undefined,
  direct: import('./types').TabStop[] | undefined,
): import('./types').TabStop[] | undefined {
  if (!style?.length || !direct?.length) {
    const only = direct ?? style
    return only?.filter((s) => s.val !== 'clear')
  }
  const merged = [...style.filter((s) => !direct.some((o) => o.pos === s.pos)), ...direct]
    .filter((s) => s.val !== 'clear')
    .sort((a, b) => a.pos - b.pos)
  // duplicated positions inside one list (seen in production pPr) collapse to the first
  const out = merged.filter((s, i) => i === 0 || s.pos !== merged[i - 1].pos)
  return out.length > 0 ? out : undefined
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
    xmlSpacePreserve: attrsOf(root)['xml:space'] === 'preserve',
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
    const mergedStops = mergeTabStops(d?.tabStops, direct?.tabStops)
    out.push({
      ...(d?.align && d.align !== 'justify' ? { align: d.align } : {}),
      ...direct,
      ...(mergedStops ? { tabStops: mergedStops } : {}),
      ...(sawPtab ? { ptabAligns } : {}),
      ...(frameXAlign ? { frameXAlign } : {}),
      runs,
    })
    flushDeferred()
  }
  flushDeferred()
  // an all-empty part collapses to no paragraphs; otherwise trailing empty
  // paragraphs stay — Word reserves their lines (header height pushes the body)
  if (out.every((p) => p.runs.length === 0 && !p.cells)) return []
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
  let fill = shd && shd !== 'auto' ? stripHash(shd) : undefined
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
  const walk = (node: XNode, anchored: boolean) => {
    const name = nameOf(node)
    if (name === 'w:txbxContent') {
      for (const inner of findChildren(node, 'w:p')) {
        const runs = extractRuns(inner, ctx)
        if (runs.length === 0) continue
        const pPr = findChild(inner, 'w:pPr')
        // content of a floating box draws at the anchor, not in the strip flow:
        // it must not push the body down (Word), only display does
        out.push({
          ...(pPr ? extractParaFormat(pPr) : {}),
          runs,
          ...(anchored ? { boxAnchored: true } : {}),
        })
      }
      return
    }
    let next = anchored
    if (name === 'wp:anchor') next = true
    else if (name?.startsWith('v:') && /position:\s*absolute/.test(attrsOf(node)['style'] ?? '')) {
      next = true
    }
    for (const child of childrenOf(node)) walk(child, next)
  }
  walk(pNode, false)
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
  // no theme part: Word still resolves schemeClr/themeColor references
  // against the built-in Office palette, so a missing part must not null out
  // the color context (themeless docx4j/mc test docs, sample real_run2 09/10)
  const file = zip.file(THEME_PART_PATH)
  if (!file) return { fonts: null, colors: { ...DEFAULT_THEME_COLORS } }
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

/** display size (wp:extent), paragraph alignment and wrap mode of an image paragraph */
type ImageMeta = Pick<
  Block,
  | 'imageWidthPx'
  | 'imageHeightPx'
  | 'imageLeadingText'
  | 'imageLeadingFont'
  | 'imageLeadingExplicitSpaceWidthPx'
  | 'imageLeadingImplicitSpaceCount'
  | 'imageParagraphIndentLeft'
  | 'imageParagraphIndentRight'
  | 'imageParagraphIndentFirstLine'
  | 'imageAlign'
  | 'imageWrap'
  | 'imageWrapDistTopEmu'
  | 'imageWrapDistBottomEmu'
  | 'imageWrapDistLeftEmu'
  | 'imageWrapDistRightEmu'
  | 'imageZOrder'
  | 'imageOffsetXEmu'
  | 'imageOffsetYEmu'
  | 'imageAnchorLocked'
  | 'imagePosH'
  | 'imagePosV'
  | 'imageRotDeg'
  | 'imageFlipH'
  | 'imageFlipV'
  | 'imageCrop'
  | 'imageFillRect'
  | 'imageBorder'
> & {
  /** wp:anchor allowOverlap="0": Word displaces the object out of a colliding anchor's box */
  imageNoOverlap?: boolean
}

/** a:srcRect / a:fillRect attribute (1000ths of a percent; some writers emit decimals) → fraction */
function rectFrac(tag: string, name: string): number {
  const v = parseFloat(new RegExp(`\\b${name}="(-?[\\d.]+)"`).exec(tag)?.[1] ?? '')
  return Number.isFinite(v) ? v / 100000 : 0
}

/**
 * Picture outline (a:ln with a solid fill on the pic's own spPr, so a sibling
 * textbox outline in the same drawing is not picked up); rendered as a CSS
 * border, display-only like crop. Theme-colored (schemeClr) outlines stay
 * unrendered.
 */
function picBorderOf(xml: string): { color: string; widthPt: number } | undefined {
  const picSpPr = /<pic:spPr[^>]*>([\s\S]*?)<\/pic:spPr>/.exec(xml)?.[1]
  const picLn = picSpPr ? /<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/.exec(picSpPr)?.[0] : undefined
  if (!picLn || /<a:noFill\s*\/>/.test(picLn)) return undefined
  const color = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(picLn)?.[1]
  if (!color) return undefined
  const w = parseInt(/<a:ln\b[^>]*\bw="(\d+)"/.exec(picLn)?.[1] ?? '', 10)
  // DrawingML default stroke width is 0.75pt (9525 EMU)
  return {
    color: color.toUpperCase(),
    widthPt: Number.isFinite(w) && w > 0 ? w / EMU_PER_PT : 0.75,
  }
}

function imageMeta(xml: string): ImageMeta {
  const meta: ImageMeta = {}
  const drawingAt = xml.search(/<w:(?:drawing|pict)[\s>]/)
  if (drawingAt >= 0 && /^<w:p[\s>]/.test(xml)) {
    const leadingXml = xml.slice(0, drawingAt)
    const leadingText = plainText(leadingXml)
    if (leadingText !== '') meta.imageLeadingText = leadingText
    const pPr = rawPPrOf(xml)
    const leadingFonts = [...leadingXml.matchAll(/<w:rFonts\b[^>]*\/?>/g)].at(-1)?.[0]
    meta.imageLeadingFont =
      /w:eastAsia="([^"]+)"/.exec(leadingFonts ?? '')?.[1] ??
      /w:ascii="([^"]+)"/.exec(leadingFonts ?? '')?.[1]
    let explicitSpaceWidthPx = 0
    let implicitSpaceCount = 0
    for (const runMatch of leadingXml.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g)) {
      const runBody = runMatch[1]
      const text = plainText(`<w:r>${runBody}</w:r>`)
      if (!/^[ ]+$/.test(text)) continue
      const rPr = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(runBody)?.[0]
      const sizeHalfPoints = Number(/<w:sz\b[^>]*w:val="(\d+)"/.exec(rPr ?? '')?.[1] ?? NaN)
      if (Number.isFinite(sizeHalfPoints)) {
        // Word's CJK single-byte spaces are half-em: half-points / 3 = CSS px.
        explicitSpaceWidthPx += (text.length * sizeHalfPoints) / 3
      } else {
        implicitSpaceCount += text.length
      }
    }
    if (explicitSpaceWidthPx > 0) {
      meta.imageLeadingExplicitSpaceWidthPx = explicitSpaceWidthPx
    }
    if (implicitSpaceCount > 0) meta.imageLeadingImplicitSpaceCount = implicitSpaceCount
    const ind = pPr ? /<w:ind\b[^>]*\/?>/.exec(pPr)?.[0] : undefined
    const twips = (name: string): number | undefined => {
      const value = Number(new RegExp(`\\bw:${name}="(-?\\d+)"`).exec(ind ?? '')?.[1] ?? NaN)
      return Number.isFinite(value) ? value : undefined
    }
    meta.imageParagraphIndentLeft = twips('left')
    meta.imageParagraphIndentRight = twips('right')
    meta.imageParagraphIndentFirstLine = twips('firstLine')
    if (meta.imageParagraphIndentFirstLine === undefined) {
      const hanging = twips('hanging')
      if (hanging !== undefined) meta.imageParagraphIndentFirstLine = -hanging
    }
  }
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
  const border = picBorderOf(xml)
  if (border) meta.imageBorder = border
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
    const wrapDistance = (attr: string): number | undefined => {
      const value = Number(new RegExp(`\\b${attr}="(\\d+)"`).exec(anchor)?.[1] ?? NaN)
      return Number.isFinite(value) ? value : undefined
    }
    meta.imageWrapDistTopEmu = wrapDistance('distT')
    meta.imageWrapDistBottomEmu = wrapDistance('distB')
    meta.imageWrapDistLeftEmu = wrapDistance('distL')
    meta.imageWrapDistRightEmu = wrapDistance('distR')
    if (/allowOverlap="(?:0|false)"/.test(anchor)) meta.imageNoOverlap = true
    if (/\blocked="(?:1|true)"/.test(anchor)) meta.imageAnchorLocked = true
    // relativeHeight = 251658240 base + zOrder; keep the delta so overlapping
    // anchors round-trip their paint order and the editor can reorder them
    const relHeight = Number(/relativeHeight="(\d+)"/.exec(anchor)?.[1] ?? NaN)
    if (Number.isFinite(relHeight)) {
      const z = relHeight - 251658240
      if (z !== 0) meta.imageZOrder = z
    }
    // an explicit wrap element wins over behindDoc: Word draws a
    // behindDoc+wrapTight object behind the text and still wraps around it
    if (/behindDoc="1"/.test(anchor) && !/<wp:wrap(Square|Tight|Through|TopAndBottom)/.test(xml))
      meta.imageWrap = 'behind'
    else if (/<wp:wrapTopAndBottom/.test(xml)) meta.imageWrap = 'topBottom'
    else if (/<wp:wrap(Square|Tight|Through)/.test(xml)) {
      const kind = /<wp:wrap(Square|Tight|Through)/.exec(xml)![1]
      const alignRight =
        /<wp:positionH[^>]*>(?:(?!<\/wp:positionH>)[\s\S])*?<wp:align>right<\/wp:align>/.test(xml)
      // column-relative only: margin/page align pairs are the position-gallery
      // presets and keep their square wrap (imagePosH/V round-trip)
      const alignCenter =
        /<wp:positionH[^>]*relativeFrom="column"[^>]*>(?:(?!<\/wp:positionH>)[\s\S])*?<wp:align>center<\/wp:align>/.test(
          xml,
        )
      // wrapText names the side the text goes on — the object floats opposite;
      // with bothSides, the object's CENTER past mid-body (~4680 twips usable
      // half on Letter/A4) means it hugs the right side: a left-edge test
      // misclassifies wide pictures whose X sits before the midline but whose
      // body fills the right half (public issue #118)
      const wrapText = /<wp:wrap(?:Square|Tight|Through)[^>]*wrapText="([^"]+)"/.exec(xml)?.[1]
      const posH = /<wp:positionH[^>]*>([\s\S]*?)<\/wp:positionH>/.exec(xml)?.[1] ?? ''
      const offX = Number(/<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(posH)?.[1] ?? NaN)
      const extentCx = Number(/<wp:extent[^>]*\bcx="(\d+)"/.exec(xml)?.[1] ?? NaN)
      const centerX = offX + (Number.isFinite(extentCx) ? extentCx / 2 : 0)
      const side =
        alignRight || wrapText === 'left' || (wrapText !== 'right' && centerX > 4680 * 635)
          ? 'right'
          : 'left'
      // centered object wrapping both sides: no both-side wrap in the
      // renderer, so approximate with the topBottom centered slot
      meta.imageWrap = alignCenter
        ? 'topBottom'
        : kind === 'Tight'
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
 * Pre-fetch external textbox parts referenced by `<wps:txbx r:txbx="…"/>`
 * (older Word builds store the w:p list in word/txbx*.xml instead of an
 * inline w:txbxContent; extractTextboxes is sync).
 */
async function externalTxbxParts(
  documentXml: string,
  zip: JSZip,
  rels: Map<string, RelInfo>,
): Promise<Map<string, string>> {
  const out = new Map<string, string>()
  for (const m of documentXml.matchAll(/<wps:txbx\b[^>]*\br:txbx="([^"]+)"/g)) {
    const rId = m[1]
    if (out.has(rId)) continue
    const rel = rels.get(rId)
    if (!rel || rel.targetMode === 'External') continue
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`
    const file = zip.file(path.replace(/^word\/\.\.\//, ''))
    if (file) out.set(rId, await file.async('string'))
  }
  return out
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
    // legacy VML pictures and OLE previews inside cells (w:pict / w:object)
    for (const m of slice.matchAll(/<v:imagedata[^>]*r:id="([^"]+)"/g)) rIds.add(m[1])
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
  // embed and link matched separately: on a blip carrying both, the greedy
  // combined pattern only captured the link and left the embedded part unresolved
  const refs = [
    ...xml.matchAll(/<a:blip[^>]*r:embed="([^"]+)"/g),
    ...xml.matchAll(/<a:blip[^>]*r:link="([^"]+)"/g),
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
  // the owning drawing's extent is the last one before its dgm:relIds (a
  // paragraph can hold several drawings)
  const dmAt = xml.indexOf('r:dm="')
  const extents = [
    ...(dmAt >= 0 ? xml.slice(0, dmAt) : xml).matchAll(/<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/g),
  ]
  const extent = extents[extents.length - 1]
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
    const prst = attrsOf(findChild(spPr, 'a:prstGeom') ?? {})['prst']
    if (prst) shape.prst = prst
    // connectors carry cy="0"/cx="0"; only they may be zero-extent
    const isLine = prst === 'line' || (prst?.includes('Connector') ?? false)
    if ((shape.wPx <= 0 || shape.hPx <= 0) && !(isLine && (shape.wPx > 0 || shape.hPx > 0)))
      continue
    const rot = parseInt(attrsOf(xfrm!)['rot'] ?? '', 10)
    if (Number.isFinite(rot) && rot !== 0) shape.rotDeg = Math.round(rot / 60000)
    const ln = findChild(spPr, 'a:ln')
    if (ln && !findChild(ln, 'a:noFill')) {
      const lnHex = colorNodeHex(findChild(ln, 'a:solidFill'), ctx.themeColors)
      if (lnHex) {
        shape.lnHex = lnHex
        const w = parseInt(attrsOf(ln)['w'] ?? '', 10)
        shape.lnWPx = Number.isFinite(w) && w > 0 ? Math.max(1, Math.round(w / EMU_PER_PX)) : 1
      }
    }
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
  const rId = /<cx?:chart [^>]*r:id="([^"]+)"/.exec(xml)?.[1]
  const rel = rId ? ctx.rels.get(rId) : undefined
  if (!rel || rel.targetMode === 'External') return null
  const path = (rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`).replace(
    /^word\/\.\.\//,
    '',
  )
  const file = ctx.zip.file(path)
  if (!file) return null
  const partXml = await file.async('string')
  const display = parseChartPartXml(partXml, path, ctx.themeColors)
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
