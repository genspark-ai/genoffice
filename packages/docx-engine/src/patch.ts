import JSZip from 'jszip'
import { applyImageWrap, generateParagraphXml, inlineRunsXml, splitXmlChildren } from './generate'
import {
  NOTE_CONTENT_TYPE,
  NOTE_PART_PATH,
  NOTE_REL_TYPE,
  buildNotesXml,
  type NoteKind,
} from './notes'
import {
  INK_MEDIA_PATH_RE,
  INK_MEDIA_PREFIX,
  INK_REL_RE,
  anchoredInkRunXml,
  injectInkRunsIntoParagraph,
  stripInkRuns,
} from './ink'
import { assertZipWithinLimits, type ParseExtras } from './parse'
import { BLANK_NUMBERING_XML, abstractNumXml, type CustomNumberingLevel } from './blank'
import { applyPageNumType, applySectionSettings, applySectionStartType } from './section'
import {
  CUSTOM_XML_REL_TYPE,
  buildSourcesItemPropsXml,
  buildSourcesXml,
  findSourcesPart,
} from './sources'
import {
  THEME_CONTENT_TYPE,
  THEME_PART_PATH,
  THEME_REL_TYPE,
  applyThemeColors,
  applyThemeFonts,
  buildThemeXml,
} from './theme'
import { buildChartPartXml, buildChartWorkbookXlsxBase64, CHART_WORKBOOK_REL_TYPE } from './chart'
import type {
  CommentInfo,
  DocProtection,
  GeneratedBlock,
  HeaderFooter,
  NewChart,
  NewImage,
  NewInkImage,
  NoteInfo,
  ParsedDoc,
  SectionSettings,
  SourceInfo,
  ThemeColors,
  ThemeFonts,
} from './types'
import { TOTAL_PAGES_MARK } from './types'
import { patchParagraphTexts } from './text-patch'
import { WATERMARK_NS, watermarkParagraphXml } from './watermark'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

export type ParsedDocFull = ParsedDoc & { extras: ParseExtras }

/** Body content in final editor order (hidden trailing elements are appended automatically). */
export type SaveBlock = (
  | { kind: 'original'; docxIndex: number }
  | { kind: 'generated'; block: GeneratedBlock }
  /** self-contained OOXML fragment created by the editor (e.g. a new table);
   *  docxIndex marks the source block (kept when a section-break paragraph is
   *  rewritten, used to inject per-section header references) */
  | { kind: 'xml'; xml: string; docxIndex?: number }
  /** a new inline image; bytes become word/media/... + relationship */
  | { kind: 'image'; image: NewImage }
  /** a new embedded chart; data becomes word/charts/chartN.xml + relationship */
  | { kind: 'chart'; chart: NewChart; extentPx?: { w: number; h: number } }
) & {
  /** Top-level tracked insertion/deletion wrapper. */
  revision?: { kind: 'ins' | 'del'; author: string; date?: string; id?: string }
}

export interface SaveOptions {
  /** save timestamp (ISO), written to docProps/core.xml dcterms:modified; default = now */
  savedAt?: string
  /** rewrite page size / margins in the trailing w:sectPr */
  section?: SectionSettings
  /** last-section start type (w:type); rewrites the trailing sectPr when inserting a continuous section break; undefined = keep */
  sectionStartType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'
  /** last-section page numbering (w:pgNumType): both fmt/start unset = remove the tag; undefined = keep */
  pgNumType?: { fmt?: string; start?: number }
  /** page color: hex without '#' to set, null to remove, undefined to keep as-is */
  pageColor?: string | null
  /** replace/create the default page header (single centered line); undefined = keep */
  header?: HeaderFooter
  /** replace/create the default page footer; undefined = keep */
  footer?: HeaderFooter
  /** first-page header/footer parts (w:type="first"); undefined = keep */
  headerFirst?: HeaderFooter
  footerFirst?: HeaderFooter
  /** even-page header/footer parts (w:type="even"); undefined = keep */
  headerEven?: HeaderFooter
  footerEven?: HeaderFooter
  /** "different first page": set/remove w:titlePg in the trailing sectPr */
  titlePg?: boolean
  /**
   * Per-section header/footer edits (non-last sections of multi-section docs, default
   * variant). lastBlockIndex locates the section's break paragraph: if the section
   * already has a matching reference, the referenced part is rewritten (earlier sections
   * sharing the part change with it — Word's "same as previous" semantics); if there is
   * no reference (inherited from the previous section), a new part is created and the
   * reference injected into this section's sectPr (the section becomes independent,
   * earlier sections are unaffected).
   */
  sectionHf?: Array<{ lastBlockIndex: number; kind: 'header' | 'footer'; hf: HeaderFooter }>
  /** "different odd & even pages": set/remove settings.xml w:evenAndOddHeaders */
  evenAndOddHeaders?: boolean
  /**
   * Append numbering definitions to word/numbering.xml (when the part is missing, it is
   * created from the blank template, including rel/ContentType). newDefs = brand-new
   * abstractNum + w:num (new lists); restartNums = new w:num pointing at an existing
   * abstractNum + startOverride (restart numbering). Append-only: existing entries keep
   * their original bytes.
   */
  numbering?: {
    /** with levels, generates the abstractNum from custom levels (multilevel list / bullet library); otherwise uses the blank-template style */
    newDefs?: Array<{ numId: string; kind: 'bullet' | 'ordered'; levels?: CustomNumberingLevel[] }>
    restartNums?: Array<{
      numId: string
      abstractNumId: string
      startOverrides: Record<number, number>
    }>
  }
  /** create/modify styles: surgical upsert of word/styles.xml by styleId (replace when present, else append) */
  styleUpserts?: StyleUpsert[]
  /**
   * Replace whole zip parts by path (e.g. patched chart parts from
   * patchChartPartXml). Only paths that already exist in the package are
   * rewritten; unknown paths are ignored.
   */
  partXml?: Record<string, string>
  /**
   * Replace whole zip parts with binary data (base64) — used to update
   * embedded xlsx workbooks alongside patched chart parts.
   * Only paths already present in the package are rewritten; unknown paths are
   * ignored.
   */
  partBinary?: Record<string, string>
  /**
   * Full desired comment list; word/comments.xml is regenerated from it
   * (plain-text bodies). undefined = keep the part byte-identical.
   */
  comments?: CommentInfo[]
  /** editing restriction; null removes w:documentProtection, undefined keeps */
  protection?: DocProtection | null
  /**
   * Full desired footnote / endnote lists; the part is regenerated from them
   * (separator entries preserved). undefined = keep byte-identical.
   */
  footnotes?: NoteInfo[]
  endnotes?: NoteInfo[]
  /**
   * Text watermark in the default page header: a string sets it, null removes
   * it, undefined keeps whatever the header already has.
   */
  watermark?: string | null
  /**
   * Full desired ink-annotation list (freehand strokes), as floating anchored pictures.
   * Existing aidocs-ink runs are stripped and re-emitted from this list, so
   * passing [] removes all our ink. undefined = keep whatever is in the file.
   */
  inks?: NewInkImage[]
  /** full bibliography source list; regenerates the b:Sources customXml part */
  sources?: SourceInfo[]
  /** theme font pair; patches (or creates) word/theme/theme1.xml */
  themeFonts?: ThemeFonts
  /** theme color scheme; patches (or creates) word/theme/theme1.xml */
  themeColors?: ThemeColors
}

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const HF_REL_TYPE = {
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
} as const
/** Model for creating/modifying a style (used by styleUpserts) */
export interface StyleUpsert {
  styleId: string
  type: 'paragraph' | 'character'
  name: string
  basedOn?: string
  rPr?: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    strike?: boolean
    /** hex without '#' */
    color?: string
    sizeHalfPoints?: number
    font?: string
  }
  pPr?: {
    align?: 'left' | 'center' | 'right' | 'justify'
    spaceBeforeTwips?: number
    spaceAfterTwips?: number
    /** line spacing as a multiple (auto) */
    lineSpacing?: number
  }
}

function buildStyleXml(up: StyleUpsert): string {
  const rPr: string[] = []
  if (up.rPr?.font) {
    const f = escapeXmlAttr(up.rPr.font)
    rPr.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}"/>`)
  }
  if (up.rPr?.bold) rPr.push('<w:b/>')
  if (up.rPr?.italic) rPr.push('<w:i/>')
  if (up.rPr?.strike) rPr.push('<w:strike/>')
  if (up.rPr?.color) rPr.push(`<w:color w:val="${escapeXmlAttr(up.rPr.color)}"/>`)
  if (up.rPr?.sizeHalfPoints) {
    rPr.push(`<w:sz w:val="${up.rPr.sizeHalfPoints}"/><w:szCs w:val="${up.rPr.sizeHalfPoints}"/>`)
  }
  if (up.rPr?.underline) rPr.push('<w:u w:val="single"/>')
  const pPr: string[] = []
  const sp = up.pPr
  if (
    sp &&
    (sp.spaceBeforeTwips !== undefined ||
      sp.spaceAfterTwips !== undefined ||
      sp.lineSpacing !== undefined)
  ) {
    const attrs = [
      sp.spaceBeforeTwips !== undefined ? ` w:before="${sp.spaceBeforeTwips}"` : '',
      sp.spaceAfterTwips !== undefined ? ` w:after="${sp.spaceAfterTwips}"` : '',
      sp.lineSpacing !== undefined
        ? ` w:line="${Math.round(sp.lineSpacing * 240)}" w:lineRule="auto"`
        : '',
    ].join('')
    pPr.push(`<w:spacing${attrs}/>`)
  }
  if (sp?.align) pPr.push(`<w:jc w:val="${sp.align === 'justify' ? 'both' : sp.align}"/>`)
  return (
    `<w:style w:type="${up.type}" w:styleId="${escapeXmlAttr(up.styleId)}" w:customStyle="1">` +
    `<w:name w:val="${escapeXmlAttr(up.name)}"/>` +
    (up.basedOn ? `<w:basedOn w:val="${escapeXmlAttr(up.basedOn)}"/>` : '') +
    '<w:qFormat/>' +
    (pPr.length > 0 ? `<w:pPr>${pPr.join('')}</w:pPr>` : '') +
    (rPr.length > 0 ? `<w:rPr>${rPr.join('')}</w:rPr>` : '') +
    '</w:style>'
  )
}

const NUMBERING_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
const COMMENTS_EXT_REL_TYPE =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'
const COMMENTS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const SETTINGS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'
const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

const IMAGE_EXT: Record<NewImage['mime'], string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
}

const EMU_PER_PX = 9525

/**
 * Given the original docx bytes and a chart part path (e.g.
 * "word/charts/chart1.xml"), returns the zip-relative path of the embedded
 * workbook (e.g. "word/charts/embeddings/workbook1.xlsx") by reading the
 * chart's rels file, or null when no workbook relationship exists.
 */
export async function findChartWorkbookPath(
  docxBytes: Uint8Array,
  chartPath: string,
): Promise<string | null> {
  try {
    const zip = await JSZip.loadAsync(docxBytes)
    // chart path: word/charts/chart1.xml → rels: word/charts/_rels/chart1.xml.rels
    const dir = chartPath.substring(0, chartPath.lastIndexOf('/'))
    const file = chartPath.substring(chartPath.lastIndexOf('/') + 1)
    const relsPath = `${dir}/_rels/${file}.rels`
    const relsFile = zip.file(relsPath)
    if (!relsFile) return null
    const relsXml = await relsFile.async('text')
    // find Relationship with Type ending in /package
    const m = relsXml.match(/Type="[^"]*\/package"[^/]*Target="([^"]+)"/)
    if (!m) return null
    // Target is relative to dir (word/charts/)
    const target = m[1]
    if (target.startsWith('/')) return target.slice(1)
    return `${dir}/${target}`
  } catch {
    return null
  }
}

/**
 * Read the raw base64 bytes of a zip part from a docx file.
 * Returns null if the part doesn't exist.
 */
export async function readDocxPartBase64(
  docxBytes: Uint8Array,
  path: string,
): Promise<string | null> {
  try {
    const zip = await JSZip.loadAsync(docxBytes)
    const file = zip.file(path)
    if (!file) return null
    const bytes = await file.async('uint8array')
    let binary = ''
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i])
    return btoa(binary)
  } catch {
    return null
  }
}

const CORE_PROPS_PATH = 'docProps/core.xml'

/**
 * docProps/core.xml: only a real save to disk updates dcterms:modified and cp:revision;
 * all other fields stay as-is. Missing tags are not injected (avoids touching the root
 * element namespaces); returns null = no change.
 */
function patchCoreProps(xml: string, savedAt?: string): string | null {
  const iso = (savedAt ?? new Date().toISOString()).replace(/\.\d{3}Z$/, 'Z')
  let out = xml.replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${iso}$2`)
  out = out.replace(/(<cp:revision>)(\d+)(<\/cp:revision>)/, (_m, open, n, close) => {
    const next = parseInt(n, 10) + 1
    return Number.isFinite(next) ? `${open}${next}${close}` : `${open}${n}${close}`
  })
  return out === xml ? null : out
}

/**
 * Paragraph-patch save.
 *
 * - Blocks marked 'original' are copied as the exact substring of the original
 *   word/document.xml (byte-for-byte after UTF-8 re-encode).
 * - Blocks marked 'generated' become fresh OOXML fragments referencing only
 *   styles that already exist in the document.
 * - 'xml' blocks are self-contained fragments inserted verbatim; 'image' blocks
 *   additionally add media entries and relationships.
 * - Every other zip entry is copied without modification.
 * - If nothing changed at all, the original file bytes are returned untouched.
 */
export async function saveDocx(
  parsed: ParsedDocFull,
  finalBlocks: SaveBlock[],
  options: SaveOptions = {},
): Promise<Uint8Array> {
  const { documentXml, originalBytes, bodyInnerStart, bodyInnerEnd } = parsed.internal
  const elements = parsed.extras.elements

  const visibleOriginalOrder = parsed.blocks.filter((b) => !b.hidden).map((b) => b.docxIndex)
  const isUnchanged =
    finalBlocks.length === visibleOriginalOrder.length &&
    finalBlocks.every(
      (fb, i) =>
        fb.kind === 'original' &&
        fb.docxIndex === visibleOriginalOrder[i] &&
        fb.revision === undefined,
    ) &&
    options.section === undefined &&
    options.sectionStartType === undefined &&
    options.pgNumType === undefined &&
    options.pageColor === undefined &&
    options.header === undefined &&
    options.footer === undefined &&
    options.headerFirst === undefined &&
    options.footerFirst === undefined &&
    options.headerEven === undefined &&
    options.footerEven === undefined &&
    options.titlePg === undefined &&
    (options.sectionHf === undefined || options.sectionHf.length === 0) &&
    options.numbering === undefined &&
    (options.styleUpserts === undefined || options.styleUpserts.length === 0) &&
    options.evenAndOddHeaders === undefined &&
    options.comments === undefined &&
    options.protection === undefined &&
    options.footnotes === undefined &&
    options.endnotes === undefined &&
    options.watermark === undefined &&
    options.inks === undefined &&
    options.sources === undefined &&
    options.themeFonts === undefined &&
    options.themeColors === undefined &&
    (options.partXml === undefined || Object.keys(options.partXml).length === 0) &&
    (options.partBinary === undefined || Object.keys(options.partBinary).length === 0)
  if (isUnchanged) return originalBytes

  const zip = await JSZip.loadAsync(originalBytes)
  assertZipWithinLimits(zip)

  // Relationship allocation for newly created hyperlinks and images.
  const relsPath = 'word/_rels/document.xml.rels'
  const relsFile = zip.file(relsPath)
  // fall back to an empty part so newly allocated rIds are never dangling
  let relsXml = relsFile
    ? await relsFile.async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  const newRels: Array<{ rId: string; type: string; target: string; external: boolean }> = []
  let nextRelNum = maxRelId(relsXml) + 1
  const allocateHyperlinkRel = (href: string): string => {
    const existing = newRels.find((r) => r.external && r.target === href)
    if (existing) return existing.rId
    const rId = `rId${nextRelNum++}`
    newRels.push({ rId, type: HYPERLINK_REL_TYPE, target: href, external: true })
    return rId
  }

  const genCtx = {
    headingStyleIds: parsed.headingStyleIds,
    listParagraphStyleId: parsed.listParagraphStyleId,
    allocateHyperlinkRel,
  }

  const newMedia: Array<{ path: string; base64: string }> = []
  const usedExtensions = new Set<string>()
  let imageSeq = nextImageSeq(zip)
  const embedImage = (image: NewImage): string => {
    const ext = IMAGE_EXT[image.mime]
    const mediaPath = `word/media/aidocs${imageSeq++}.${ext}`
    const rId = `rId${nextRelNum++}`
    newRels.push({
      rId,
      type: IMAGE_REL_TYPE,
      target: mediaPath.replace(/^word\//, ''),
      external: false,
    })
    newMedia.push({ path: mediaPath, base64: image.base64 })
    usedExtensions.add(ext)
    const cx = Math.max(1, Math.round(image.widthPx * EMU_PER_PX))
    const cy = Math.max(1, Math.round(image.heightPx * EMU_PER_PX))
    const docPrId = 9000 + imageSeq
    const pPr =
      image.align && image.align !== 'left' ? `<w:pPr><w:jc w:val="${image.align}"/></w:pPr>` : ''
    const xml =
      `<w:p>${pPr}<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:docPr id="${docPrId}" name="Picture ${docPrId}"/>` +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="Picture ${docPrId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    return image.wrap ? applyImageWrap(xml, image.wrap) : xml
  }

  // ---- new embedded charts: chart part + workbook + relationship + drawing paragraph ----
  const newChartParts: Array<{ path: string; xml: string }> = []
  const newChartWorkbooks: Array<{
    xlsxPath: string
    relsPath: string
    relsXml: string
    base64: string
  }> = []
  let chartDocPrId = 8000
  const embedChart = async (
    chart: NewChart,
    extentPx?: { w: number; h: number },
  ): Promise<string> => {
    let n = 1
    while (
      zip.file(`word/charts/chart${n}.xml`) ||
      newChartParts.some((p) => p.path === `word/charts/chart${n}.xml`)
    ) {
      n++
    }
    const path = `word/charts/chart${n}.xml`
    const rId = `rId${nextRelNum++}`
    newRels.push({ rId, type: CHART_REL_TYPE, target: `charts/chart${n}.xml`, external: false })

    // Build embedded workbook and create chart rels part
    const wbBase64 = await buildChartWorkbookXlsxBase64(chart.categories, chart.series)
    const xlsxPath = `word/charts/embeddings/workbook${n}.xlsx`
    const wbRId = 'rId1'
    const chartRelsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="${wbRId}" Type="${CHART_WORKBOOK_REL_TYPE}" Target="embeddings/workbook${n}.xlsx"/>` +
      '</Relationships>'
    newChartParts.push({ path, xml: buildChartPartXml(chart, wbRId) })
    newChartWorkbooks.push({
      xlsxPath,
      relsPath: `word/charts/_rels/chart${n}.xml.rels`,
      relsXml: chartRelsXml,
      base64: wbBase64,
    })

    const docPrId = chartDocPrId++
    const cx = extentPx ? Math.max(1, Math.round(extentPx.w * 9525)) : 5486400
    const cy = extentPx ? Math.max(1, Math.round(extentPx.h * 9525)) : 3200400
    return (
      '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:docPr id="${docPrId}" name="Chart ${docPrId}"/>` +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/>` +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )
  }

  // ---- ink annotations: floating anchored pictures, re-emitted wholesale ----
  const inksByBlock = new Map<number, NewInkImage[]>()
  for (const ink of options.inks ?? []) {
    const list = inksByBlock.get(ink.blockIndex)
    if (list) list.push(ink)
    else inksByBlock.set(ink.blockIndex, [ink])
  }
  // Ink media get their own name prefix: every ink save strips + re-emits the
  // whole layer, so old word/media/aidocsink*.png parts (and their rels) are
  // dropped from the output instead of accumulating as orphans.
  let inkSeq = 1
  /** allocate media + relationship for one ink PNG, return its anchored run */
  const inkRunXml = (ink: NewInkImage): string => {
    const mediaPath = `word/media/${INK_MEDIA_PREFIX}${inkSeq++}.png`
    const rId = `rId${nextRelNum++}`
    newRels.push({
      rId,
      type: IMAGE_REL_TYPE,
      target: mediaPath.replace(/^word\//, ''),
      external: false,
    })
    newMedia.push({ path: mediaPath, base64: ink.base64 })
    usedExtensions.add('png')
    return anchoredInkRunXml(ink, rId, 9000 + inkSeq)
  }

  // ---- header / footer: overwrite the existing default part or create one ----
  const sectBlock = parsed.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))
  const trailingSectPr = sectBlock?.originalXml ?? ''
  const relTargets = new Map<string, string>()
  if (relsXml) {
    for (const tag of relsXml.match(/<Relationship [^>]*\/>/g) ?? []) {
      const id = /Id="([^"]+)"/.exec(tag)?.[1]
      const target = /Target="([^"]+)"/.exec(tag)?.[1]
      if (id && target) relTargets.set(id, target)
    }
  }
  const hfParts: Array<{ path: string; xml: string }> = []
  const hfRefTags: string[] = []
  const hfOverrides: string[] = []
  const planHeaderFooter = async (
    kind: 'header' | 'footer',
    hf: HeaderFooter | undefined,
    watermark: string | null = null,
    hfType: 'default' | 'first' | 'even' = 'default',
    watermarkOnly = false,
  ) => {
    if (hf === undefined) return
    const refs = trailingSectPr.match(new RegExp(`<w:${kind}Reference[^>]*/>`, 'g')) ?? []
    const existing =
      refs.find((r) => r.includes(`w:type="${hfType}"`)) ??
      (hfType === 'default' ? refs.find((r) => !/w:type="/.test(r)) : undefined)
    const rId = existing ? /r:id="([^"]+)"/.exec(existing)?.[1] : undefined
    const target = rId ? relTargets.get(rId) : undefined
    if (target) {
      const path = target.startsWith('/') ? target.slice(1) : `word/${target}`
      const file = zip.file(path)
      const originalXml = file ? await file.async('string') : null
      // A watermark-only change patches the original part in place (tables/logos/fields
      // all preserved); header text edits use paragraph replace-merge (non-paragraph
      // children preserved).
      let partXml: string | null = null
      if (watermarkOnly && kind === 'header' && originalXml) {
        partXml = patchWatermarkInPart(originalXml, watermark)
      }
      if (partXml === null) partXml = headerFooterPartXml(kind, hf, watermark, originalXml)
      hfParts.push({ path, xml: partXml })
    } else {
      const partXml = headerFooterPartXml(kind, hf, watermark)
      let n = 1
      while (
        zip.file(`word/${kind}${n}.xml`) ||
        hfParts.some((p) => p.path === `word/${kind}${n}.xml`)
      )
        n++
      const filename = `${kind}${n}.xml`
      const newRId = `rId${nextRelNum++}`
      newRels.push({ rId: newRId, type: HF_REL_TYPE[kind], target: filename, external: false })
      hfParts.push({ path: `word/${filename}`, xml: partXml })
      hfRefTags.push(`<w:${kind}Reference w:type="${hfType}" r:id="${newRId}"/>`)
      hfOverrides.push(
        `<Override PartName="/word/${filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`,
      )
    }
  }
  // A watermark change forces a header-part rewrite even when the header text
  // itself is untouched; conversely a header rewrite must carry the existing
  // watermark through (the part is regenerated wholesale).
  const effectiveHeader =
    options.header ??
    (options.watermark !== undefined ? { text: parsed.headerText ?? '' } : undefined)
  const effectiveWatermark =
    options.watermark !== undefined ? options.watermark : (parsed.watermarkText ?? null)
  await planHeaderFooter(
    'header',
    effectiveHeader,
    effectiveWatermark,
    'default',
    options.header === undefined,
  )
  await planHeaderFooter('footer', options.footer)
  await planHeaderFooter('header', options.headerFirst, null, 'first')
  await planHeaderFooter('footer', options.footerFirst, null, 'first')
  await planHeaderFooter('header', options.headerEven, null, 'even')
  await planHeaderFooter('footer', options.footerEven, null, 'even')

  // ---- Per-section header/footer (non-last sections): with a reference, rewrite the
  // part; without one, create a part + inject the reference ----
  const sectionRefTags = new Map<number, string[]>()
  for (const edit of options.sectionHf ?? []) {
    const block = parsed.blocks.find((b) => b.docxIndex === edit.lastBlockIndex)
    const sectPr =
      block?.originalXml?.match(/<w:sectPr[^>]*\/>|<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0] ?? ''
    const refs = sectPr.match(new RegExp(`<w:${edit.kind}Reference[^>]*/>`, 'g')) ?? []
    const existing =
      refs.find((r) => r.includes('w:type="default"')) ?? refs.find((r) => !/w:type="/.test(r))
    const rId = existing ? /r:id="([^"]+)"/.exec(existing)?.[1] : undefined
    const target = rId ? relTargets.get(rId) : undefined
    if (target) {
      const path = target.startsWith('/') ? target.slice(1) : `word/${target}`
      const file = zip.file(path)
      const originalXml = file ? await file.async('string') : null
      hfParts.push({ path, xml: headerFooterPartXml(edit.kind, edit.hf, null, originalXml) })
    } else {
      const partXml = headerFooterPartXml(edit.kind, edit.hf, null)
      let n = 1
      while (
        zip.file(`word/${edit.kind}${n}.xml`) ||
        hfParts.some((p) => p.path === `word/${edit.kind}${n}.xml`)
      )
        n++
      const filename = `${edit.kind}${n}.xml`
      const newRId = `rId${nextRelNum++}`
      newRels.push({ rId: newRId, type: HF_REL_TYPE[edit.kind], target: filename, external: false })
      hfParts.push({ path: `word/${filename}`, xml: partXml })
      hfOverrides.push(
        `<Override PartName="/word/${filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${edit.kind}+xml"/>`,
      )
      const tags = sectionRefTags.get(edit.lastBlockIndex) ?? []
      tags.push(`<w:${edit.kind}Reference w:type="default" r:id="${newRId}"/>`)
      sectionRefTags.set(edit.lastBlockIndex, tags)
    }
  }

  // ---- numbering: append numbering definitions (create the part from the blank
  // template when missing) ----
  const numberingPath = 'word/numbering.xml'
  let numberingXmlOut: string | null = null
  let numberingIsNew = false
  if (
    (options.numbering?.newDefs?.length ?? 0) > 0 ||
    (options.numbering?.restartNums?.length ?? 0) > 0
  ) {
    const file = zip.file(numberingPath)
    let xml = file ? await file.async('string') : null
    if (xml === null) {
      xml = BLANK_NUMBERING_XML
      numberingIsNew = true
      newRels.push({
        rId: `rId${nextRelNum++}`,
        type: NUMBERING_REL_TYPE,
        target: 'numbering.xml',
        external: false,
      })
    }
    // New abstractNum ids are assigned by the engine (the App only ever sees abstracts
    // referenced by a w:num)
    let maxAbs = -1
    for (const m of xml.matchAll(/<w:abstractNum [^>]*w:abstractNumId="(\d+)"/g)) {
      maxAbs = Math.max(maxAbs, parseInt(m[1], 10))
    }
    const absXmls: string[] = []
    const numXmls: string[] = []
    for (const def of options.numbering?.newDefs ?? []) {
      const absId = String(++maxAbs)
      absXmls.push(abstractNumXml(absId, def.kind, def.levels))
      numXmls.push(`<w:num w:numId="${def.numId}"><w:abstractNumId w:val="${absId}"/></w:num>`)
    }
    for (const r of options.numbering?.restartNums ?? []) {
      const overrides = Object.entries(r.startOverrides)
        .map(
          ([ilvl, v]) =>
            `<w:lvlOverride w:ilvl="${ilvl}"><w:startOverride w:val="${v}"/></w:lvlOverride>`,
        )
        .join('')
      numXmls.push(
        `<w:num w:numId="${r.numId}"><w:abstractNumId w:val="${r.abstractNumId}"/>${overrides}</w:num>`,
      )
    }
    // Schema order: abstractNum* comes before num* — insert new abstracts before the first w:num
    if (absXmls.length > 0) {
      xml = /<w:num[\s>]/.test(xml)
        ? xml.replace(/<w:num[\s>]/, (m) => absXmls.join('') + m)
        : xml.replace('</w:numbering>', `${absXmls.join('')}</w:numbering>`)
    }
    xml = xml.replace('</w:numbering>', `${numXmls.join('')}</w:numbering>`)
    numberingXmlOut = xml
  }

  // ---- styles: surgical upsert of word/styles.xml (create/modify styles) ----
  const stylesPath = 'word/styles.xml'
  let stylesXmlOut: string | null = null
  if ((options.styleUpserts?.length ?? 0) > 0) {
    const file = zip.file(stylesPath)
    let xml = file
      ? await file.async('string')
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>'
    for (const up of options.styleUpserts ?? []) {
      const styleXml = buildStyleXml(up)
      const existing = new RegExp(
        `<w:style [^>]*w:styleId="${up.styleId.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')}"[\\s\\S]*?</w:style>`,
      )
      xml = existing.test(xml)
        ? xml.replace(existing, styleXml)
        : xml.replace('</w:styles>', `${styleXml}</w:styles>`)
    }
    stylesXmlOut = xml
  }

  // ---- comments: regenerate word/comments.xml from the full desired list ----
  const commentsPath = 'word/comments.xml'
  const commentsExtPath = 'word/commentsExtended.xml'
  let commentsXml: string | null = null
  let commentsIsNew = false
  let commentsExtXml: string | null = null
  let commentsExtIsNew = false
  if (options.comments) {
    // Ensure every comment has a paraId (the commentsExtended link key; old comments
    // carry theirs in the original bytes)
    let paraSeq = 1
    const withParaIds = options.comments.map((c) =>
      c.paraId
        ? c
        : {
            ...c,
            paraId: (0x10000000 + paraSeq++ * 0x1111 + parseInt(c.id, 10))
              .toString(16)
              .toUpperCase()
              .padStart(8, '0'),
          },
    )
    const commentsFile = zip.file(commentsPath)
    commentsXml = buildCommentsXml(
      withParaIds,
      commentsFile ? await commentsFile.async('string') : null,
    )
    if (!zip.file(commentsPath)) {
      commentsIsNew = true
      newRels.push({
        rId: `rId${nextRelNum++}`,
        type: COMMENTS_REL_TYPE,
        target: 'comments.xml',
        external: false,
      })
    }
    // Regenerate commentsExtended.xml in step when replies/resolved flags exist, or the
    // part already exists
    const needExt =
      zip.file(commentsExtPath) !== null ||
      withParaIds.some((c) => c.parentId !== undefined || c.done !== undefined)
    if (needExt) {
      commentsExtXml = buildCommentsExtendedXml(withParaIds)
      if (!zip.file(commentsExtPath)) {
        commentsExtIsNew = true
        newRels.push({
          rId: `rId${nextRelNum++}`,
          type: COMMENTS_EXT_REL_TYPE,
          target: 'commentsExtended.xml',
          external: false,
        })
      }
    }
  }

  // ---- footnotes / endnotes: regenerate the part from the full desired list ----
  const notesParts: Array<{ path: string; xml: string; isNew: boolean; kind: NoteKind }> = []
  const planNotes = async (kind: NoteKind, notes: NoteInfo[] | undefined) => {
    if (!notes) return
    const path = NOTE_PART_PATH[kind]
    const file = zip.file(path)
    const originalXml = file ? await file.async('string') : null
    notesParts.push({ path, xml: buildNotesXml(kind, notes, originalXml), isNew: !file, kind })
    if (!file) {
      newRels.push({
        rId: `rId${nextRelNum++}`,
        type: NOTE_REL_TYPE[kind],
        target: path.replace(/^word\//, ''),
        external: false,
      })
    }
  }
  await planNotes('footnote', options.footnotes)
  await planNotes('endnote', options.endnotes)

  // ---- bibliography sources: the b:Sources customXml part ----
  let sourcesPart: { path: string; propsPath: string; xml: string; isNew: boolean } | null = null
  if (options.sources) {
    const existing = await findSourcesPart(zip)
    const xml = buildSourcesXml(
      options.sources,
      existing ? await zip.file(existing)!.async('string') : null,
    )
    if (existing) {
      const n = /item(\d+)\.xml$/.exec(existing)?.[1] ?? '1'
      sourcesPart = { path: existing, propsPath: `customXml/itemProps${n}.xml`, xml, isNew: false }
    } else {
      let n = 1
      while (zip.file(`customXml/item${n}.xml`)) n++
      sourcesPart = {
        path: `customXml/item${n}.xml`,
        propsPath: `customXml/itemProps${n}.xml`,
        xml,
        isNew: true,
      }
      newRels.push({
        rId: `rId${nextRelNum++}`,
        type: CUSTOM_XML_REL_TYPE,
        target: `../customXml/item${n}.xml`,
        external: false,
      })
    }
  }

  // ---- theme fonts / colors: patch or create word/theme/theme1.xml ----
  let themePart: { xml: string; isNew: boolean } | null = null
  if (options.themeFonts || options.themeColors) {
    const themeFile = zip.file(THEME_PART_PATH)
    if (themeFile) {
      let xml = await themeFile.async('string')
      if (options.themeFonts) xml = applyThemeFonts(xml, options.themeFonts)
      if (options.themeColors) xml = applyThemeColors(xml, options.themeColors)
      themePart = { xml, isNew: false }
    } else {
      themePart = {
        xml: buildThemeXml(
          options.themeFonts ?? { major: 'Calibri Light', minor: 'Calibri', eastAsia: '' },
          options.themeColors ?? {},
        ),
        isNew: true,
      }
      newRels.push({
        rId: `rId${nextRelNum++}`,
        type: THEME_REL_TYPE,
        target: 'theme/theme1.xml',
        external: false,
      })
    }
  }

  const parts: string[] = []
  for (let i = 0; i < finalBlocks.length; i++) {
    const fb = finalBlocks[i]
    let xml: string
    let fbDocxIndex: number | undefined
    if (fb.kind === 'original') {
      const el = elements[fb.docxIndex]
      if (!el) throw new Error(`invalid docxIndex ${fb.docxIndex}`)
      xml = documentXml.slice(el.start, el.end)
      fbDocxIndex = fb.docxIndex
    } else if (fb.kind === 'generated') {
      xml = generateParagraphXml(fb.block, genCtx)
      // If the original paragraph was inside a w:sdt shell, re-wrap it
      if (fb.block.sdtShell) {
        xml = fb.block.sdtShell.openXml + xml + fb.block.sdtShell.closeXml
      }
    } else if (fb.kind === 'xml') {
      xml = fb.xml
      fbDocxIndex = fb.docxIndex
    } else if (fb.kind === 'chart') {
      xml = await embedChart(fb.chart, fb.extentPx)
    } else {
      xml = embedImage(fb.image)
    }
    // Inject newly created per-section header/footer references into the section's sectPr
    // (the reference must be the first sectPr child)
    const refTags = fbDocxIndex !== undefined ? sectionRefTags.get(fbDocxIndex) : undefined
    if (refTags && refTags.length > 0) {
      xml = xml.replace(/(<w:sectPr[^>]*>)/, `$1${refTags.join('')}`)
    }
    // The ink list is authoritative: old aidocs-ink runs go away, the desired
    // set is re-injected at its (possibly new) anchor paragraphs.
    if (options.inks !== undefined) xml = stripInkRuns(xml)
    // check anchor viability BEFORE allocating media, so a non-paragraph
    // anchor doesn't leave orphan relationship/media entries behind
    const blockInks = inksByBlock.get(i)
    if (blockInks && /^<w:p[\s/>]/.test(xml)) {
      const injected = injectInkRunsIntoParagraph(xml, blockInks.map(inkRunXml).join(''))
      if (injected !== null) xml = injected
    }
    if (fb.revision && !new RegExp(`^<w:${fb.revision.kind}[\\s>]`).test(xml)) {
      const revision = fb.revision
      const attrs =
        ` w:id="${escapeXmlAttr(revision.id ?? '0')}"` +
        ` w:author="${escapeXmlAttr(revision.author)}"` +
        (revision.date ? ` w:date="${escapeXmlAttr(revision.date)}"` : '')
      xml = `<w:${revision.kind}${attrs}>${xml}</w:${revision.kind}>`
    }
    parts.push(xml)
  }
  // Trailing hidden elements (w:sectPr) always keep their original bytes and position,
  // unless the editor changed the page setup.
  for (const block of parsed.blocks) {
    if (block.hidden && block.docxIndex !== null) {
      const el = elements[block.docxIndex]
      let xml = documentXml.slice(el.start, el.end)
      if (xml.includes('<w:sectPr')) {
        if (options.section) xml = applySectionSettings(xml, options.section)
        if (options.sectionStartType) xml = applySectionStartType(xml, options.sectionStartType)
        if (options.pgNumType)
          xml = applyPageNumType(xml, options.pgNumType.fmt, options.pgNumType.start)
        if (options.titlePg !== undefined) xml = applyTitlePg(xml, options.titlePg)
        // headerReference/footerReference must be the first sectPr children
        if (hfRefTags.length > 0) {
          xml = xml.replace(/(<w:sectPr[^>]*>)/, `$1${hfRefTags.join('')}`)
        }
      }
      parts.push(xml)
    }
  }

  let newDocumentXml =
    documentXml.slice(0, bodyInnerStart) + parts.join('') + documentXml.slice(bodyInnerEnd)

  // editor-generated formulas need the math namespace on the document root;
  // docx produced by non-Word generators may not declare it
  if (newDocumentXml.includes('<m:') && !/<w:document[^>]*xmlns:m=/.test(newDocumentXml)) {
    newDocumentXml = newDocumentXml.replace(
      /<w:document /,
      '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ',
    )
  }

  if (options.pageColor !== undefined) {
    newDocumentXml = applyPageColor(newDocumentXml, options.pageColor)
  }

  const settingsPath = 'word/settings.xml'
  let settingsXml: string | null = null
  let settingsIsNew = false
  if (
    options.pageColor ||
    options.protection !== undefined ||
    options.evenAndOddHeaders !== undefined
  ) {
    const file = zip.file(settingsPath)
    let xml: string
    let touched = false
    if (file) {
      xml = await file.async('string')
    } else {
      xml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>'
      settingsIsNew = true
      touched = true
      newRels.push({
        rId: `rId${nextRelNum++}`,
        type: SETTINGS_REL_TYPE,
        target: 'settings.xml',
        external: false,
      })
    }
    // Word only renders w:background when settings.xml opts in.
    if (options.pageColor && !xml.includes('<w:displayBackgroundShape')) {
      xml = xml.replace(/(<w:settings[^>]*>)/, '$1<w:displayBackgroundShape/>')
      touched = true
    }
    if (options.protection !== undefined) {
      xml = applyProtection(xml, options.protection)
      touched = true
    }
    if (options.evenAndOddHeaders !== undefined) {
      xml = applyEvenAndOddHeaders(xml, options.evenAndOddHeaders)
      touched = true
    }
    if (touched) settingsXml = xml
  }

  let relsChanged = false
  // drop relationships of stripped ink runs (their media parts are dropped too)
  if (options.inks !== undefined && relsXml) {
    const cleaned = relsXml.replace(INK_REL_RE, '')
    if (cleaned !== relsXml) {
      relsXml = cleaned
      relsChanged = true
    }
  }
  if (newRels.length > 0 && relsXml) {
    const inserts = newRels
      .map(
        (r) =>
          `<Relationship Id="${escapeXmlAttr(r.rId)}" Type="${r.type}" Target="${escapeXmlAttr(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`,
      )
      .join('')
    relsXml = relsXml.replace('</Relationships>', `${inserts}</Relationships>`)
    relsChanged = true
  }

  const contentTypesPath = '[Content_Types].xml'
  let contentTypesXml: string | null = null
  const hasNewParts =
    usedExtensions.size > 0 ||
    hfOverrides.length > 0 ||
    newChartParts.length > 0 ||
    newChartWorkbooks.length > 0 ||
    settingsIsNew ||
    commentsIsNew ||
    commentsExtIsNew ||
    numberingIsNew ||
    notesParts.some((p) => p.isNew) ||
    sourcesPart?.isNew ||
    themePart?.isNew
  if (hasNewParts) {
    const file = zip.file(contentTypesPath)
    if (file) {
      contentTypesXml = await file.async('string')
      const addOverride = (partName: string, contentType: string) => {
        if (!contentTypesXml!.includes(`PartName="${partName}"`)) {
          contentTypesXml = contentTypesXml!.replace(
            '</Types>',
            `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`,
          )
        }
      }
      for (const ext of usedExtensions) {
        if (!new RegExp(`Extension="${ext}"`).test(contentTypesXml)) {
          const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/jpeg'
          contentTypesXml = contentTypesXml.replace(
            '</Types>',
            `<Default Extension="${ext}" ContentType="${mime}"/></Types>`,
          )
        }
      }
      for (const override of hfOverrides) {
        const partName = /PartName="([^"]+)"/.exec(override)?.[1] ?? ''
        if (!contentTypesXml.includes(`PartName="${partName}"`)) {
          contentTypesXml = contentTypesXml.replace('</Types>', `${override}</Types>`)
        }
      }
      if (commentsIsNew) {
        addOverride(
          '/word/comments.xml',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        )
      }
      if (commentsExtIsNew) {
        addOverride(
          '/word/commentsExtended.xml',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
        )
      }
      if (settingsIsNew) {
        addOverride(
          '/word/settings.xml',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        )
      }
      if (numberingIsNew) {
        addOverride(
          '/word/numbering.xml',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
        )
      }
      for (const part of newChartParts) addOverride(`/${part.path}`, CHART_CONTENT_TYPE)
      for (const wb of newChartWorkbooks) addOverride(`/${wb.xlsxPath}`, XLSX_CONTENT_TYPE)
      for (const part of notesParts) {
        if (part.isNew) addOverride(`/${part.path}`, NOTE_CONTENT_TYPE[part.kind])
      }
      if (sourcesPart?.isNew) {
        addOverride(
          `/${sourcesPart.propsPath}`,
          'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
        )
      }
      if (themePart?.isNew) addOverride(`/${THEME_PART_PATH}`, THEME_CONTENT_TYPE)
    }
  }

  const coreEntry = zip.file(CORE_PROPS_PATH)
  const coreXmlOut = coreEntry
    ? patchCoreProps(await coreEntry.async('string'), options.savedAt)
    : null

  const out = new JSZip()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) {
      out.folder(name)
      continue
    }
    // old ink PNGs are re-emitted from options.inks; don't carry orphans over
    if (options.inks !== undefined && INK_MEDIA_PATH_RE.test(name)) continue
    const hfPart = hfParts.find((p) => p.path === name)
    if (name === 'word/document.xml') {
      out.file(name, newDocumentXml, { date: entry.date })
    } else if (hfPart) {
      out.file(name, hfPart.xml, { date: entry.date })
    } else if (name === relsPath && relsChanged && relsXml) {
      out.file(name, relsXml, { date: entry.date })
    } else if (name === contentTypesPath && contentTypesXml !== null) {
      out.file(name, contentTypesXml, { date: entry.date })
    } else if (name === settingsPath && settingsXml !== null) {
      out.file(name, settingsXml, { date: entry.date })
    } else if (name === commentsPath && commentsXml !== null) {
      out.file(name, commentsXml, { date: entry.date })
    } else if (name === commentsExtPath && commentsExtXml !== null) {
      out.file(name, commentsExtXml, { date: entry.date })
    } else if (name === numberingPath && numberingXmlOut !== null) {
      out.file(name, numberingXmlOut, { date: entry.date })
    } else if (name === stylesPath && stylesXmlOut !== null) {
      out.file(name, stylesXmlOut, { date: entry.date })
    } else if (notesParts.some((p) => p.path === name)) {
      out.file(name, notesParts.find((p) => p.path === name)!.xml, { date: entry.date })
    } else if (sourcesPart && name === sourcesPart.path) {
      out.file(name, sourcesPart.xml, { date: entry.date })
    } else if (themePart && name === THEME_PART_PATH) {
      out.file(name, themePart.xml, { date: entry.date })
    } else if (name === CORE_PROPS_PATH && coreXmlOut !== null) {
      out.file(name, coreXmlOut, { date: entry.date })
    } else if (options.partXml && options.partXml[name] !== undefined) {
      out.file(name, options.partXml[name], { date: entry.date })
    } else if (options.partBinary && options.partBinary[name] !== undefined) {
      out.file(name, options.partBinary[name], { base64: true, date: entry.date })
    } else {
      out.file(name, await entry.async('uint8array'), { date: entry.date })
    }
  }
  for (const media of newMedia) {
    out.file(media.path, media.base64, { base64: true })
  }
  for (const part of hfParts) {
    if (!zip.file(part.path)) out.file(part.path, part.xml)
  }
  if (relsChanged && relsXml && !zip.file(relsPath)) {
    out.file(relsPath, relsXml)
  }
  if (commentsIsNew && commentsXml !== null) {
    out.file(commentsPath, commentsXml)
  }
  if (commentsExtIsNew && commentsExtXml !== null) {
    out.file(commentsExtPath, commentsExtXml)
  }
  if (numberingIsNew && numberingXmlOut !== null) {
    out.file(numberingPath, numberingXmlOut)
  }
  if (stylesXmlOut !== null && !zip.file(stylesPath)) {
    out.file(stylesPath, stylesXmlOut)
  }
  if (settingsIsNew && settingsXml !== null) {
    out.file(settingsPath, settingsXml)
  }
  for (const part of newChartParts) {
    out.file(part.path, part.xml)
  }
  for (const wb of newChartWorkbooks) {
    if (!zip.file(wb.relsPath)) {
      out.file(wb.relsPath, wb.relsXml)
    }
    out.file(wb.xlsxPath, wb.base64, { base64: true })
  }
  for (const part of notesParts) {
    if (part.isNew) out.file(part.path, part.xml)
  }
  if (sourcesPart?.isNew) {
    out.file(sourcesPart.path, sourcesPart.xml)
    out.file(sourcesPart.propsPath, buildSourcesItemPropsXml())
    const relsName = sourcesPart.path.replace(/^customXml\//, '').replace(/\.xml$/, '')
    out.file(
      `customXml/_rels/${relsName}.xml.rels`,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="${sourcesPart.propsPath.replace(/^customXml\//, '')}"/>` +
        '</Relationships>',
    )
  }
  if (themePart?.isNew) {
    out.file(THEME_PART_PATH, themePart.xml)
  }
  return out.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

/**
 * Standalone header/footer part: one centered paragraph, optional PAGE field.
 * A '#' in the text marks where the page number goes (e.g. "- # -"). Headers
 * additionally carry the page watermark shape when one is set.
 */
/** Add the namespaces a VML watermark needs to the original part's root tag (existing ones untouched) */
function withWatermarkNs(openTag: string): string {
  let out = openTag
  for (const ns of [
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:o="urn:schemas-microsoft-com:office:office"',
    'xmlns:w10="urn:schemas-microsoft-com:office:word"',
  ]) {
    if (!out.includes(ns.split('=')[0] + '=')) out = out.replace(/>$/, ` ${ns}>`)
  }
  return out
}

/**
 * In-place patch for a watermark-only change: drop the old watermark paragraph (<w:p>
 * containing v:textpath), insert the new watermark at the start of the part, and keep
 * everything else (paragraph formatting/tables/logos/fields) byte-identical. Returns
 * null when the root tag cannot be recognized so the caller falls back to a full rebuild.
 */
function patchWatermarkInPart(originalXml: string, watermark: string | null): string | null {
  const open = /<w:hdr[^>]*>/.exec(originalXml)?.[0]
  if (!open) return null
  const openIdx = originalXml.indexOf(open)
  const closeIdx = originalXml.lastIndexOf('</w:hdr>')
  if (closeIdx < 0) return null
  const prefix = originalXml.slice(0, openIdx)
  const inner = originalXml.slice(openIdx + open.length, closeIdx)
  const kept = splitXmlChildren(inner).filter(
    (c) => !(c.name === 'w:p' && c.xml.includes('<v:textpath')),
  )
  const wm = watermark ? watermarkParagraphXml(watermark) : ''
  const rootOpen = watermark ? withWatermarkNs(open) : open
  return `${prefix}${rootOpen}${wm}${kept.map((c) => c.xml).join('')}</w:hdr>`
}

function headerFooterPartXml(
  kind: 'header' | 'footer',
  hf: HeaderFooter,
  watermark: string | null = null,
  originalXml: string | null = null,
): string {
  const root = kind === 'header' ? 'w:hdr' : 'w:ftr'
  const textRun = (t: string) =>
    t ? `<w:r><w:t xml:space="preserve">${escapeXmlText(t)}</w:t></w:r>` : ''
  const pageField =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>1</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  const numPagesField =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>1</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  let content: string
  if (hf.paras) {
    // rich paragraphs; the first '#' run becomes the PAGE field when enabled,
    // every TOTAL_PAGES_MARK becomes a NUMPAGES field
    let pageEmitted = !hf.pageNumber
    content = hf.paras
      .map((para) => {
        const jc = para.align
          ? `<w:pPr><w:jc w:val="${para.align === 'justify' ? 'both' : para.align}"/></w:pPr>`
          : ''
        let runs = ''
        for (const run of para.runs) {
          if (run.text.includes(TOTAL_PAGES_MARK) || (!pageEmitted && run.text.includes('#'))) {
            run.text.split(TOTAL_PAGES_MARK).forEach((seg, k) => {
              if (k > 0) runs += numPagesField
              if (!pageEmitted && seg.includes('#')) {
                const [before, ...rest] = seg.split('#')
                runs +=
                  inlineRunsXml(before ? [{ ...run, text: before }] : []) +
                  pageField +
                  inlineRunsXml(rest.join('#') ? [{ ...run, text: rest.join('#') }] : [])
                pageEmitted = true
              } else if (seg) {
                runs += inlineRunsXml([{ ...run, text: seg }])
              }
            })
          } else {
            runs += inlineRunsXml([run])
          }
        }
        return `<w:p>${jc}${runs}</w:p>`
      })
      .join('')
    if (!pageEmitted) content += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${pageField}</w:p>`
  } else {
    const textWithTotal = (t: string) => t.split(TOTAL_PAGES_MARK).map(textRun).join(numPagesField)
    const runs: string[] = []
    if (hf.pageNumber && hf.text.includes('#')) {
      const [before, ...rest] = hf.text.split('#')
      runs.push(textWithTotal(before), pageField, textWithTotal(rest.join('#')))
    } else {
      if (hf.text) runs.push(textWithTotal(hf.text + (hf.pageNumber ? ' ' : '')))
      if (hf.pageNumber) runs.push(pageField)
    }
    content = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${runs.join('')}</w:p>`
  }
  const watermarkXml = kind === 'header' && watermark ? watermarkParagraphXml(watermark) : ''
  const body = `${watermarkXml}${content}`
  // Surgical merge: non-paragraph children of the original part (tables/sdt) and
  // paragraphs containing images/objects (logos etc., which are not in the text-paragraph
  // model) keep their original bytes; only the set of text paragraphs is replaced as a
  // whole at the position of the first text paragraph. The watermark paragraph
  // (v:textpath) is the exception — it is regenerated from watermarkXml.
  if (originalXml) {
    const open = new RegExp(`<${root}[^>]*>`).exec(originalXml)?.[0]
    const closeIdx = originalXml.lastIndexOf(`</${root}>`)
    if (open && closeIdx >= 0) {
      const openIdx = originalXml.indexOf(open)
      const children = splitXmlChildren(originalXml.slice(openIdx + open.length, closeIdx))
      const isProtectedPara = (xml: string) =>
        /<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]/.test(xml) && !xml.includes('<v:textpath')
      const isTextPara = (c: { name: string; xml: string }) =>
        c.name === 'w:p' && !isProtectedPara(c.xml)
      if (children.some((c) => !isTextPara(c))) {
        const parts: string[] = []
        let injected = false
        for (const c of children) {
          if (isTextPara(c)) {
            if (!injected) {
              parts.push(body)
              injected = true
            }
          } else if (c.name === 'w:p' && c.xml.includes('<v:textpath')) {
            // Drop the old watermark paragraph (body already carries the regenerated watermarkXml)
          } else {
            parts.push(c.xml)
          }
        }
        if (!injected) parts.unshift(body)
        const rootOpen = watermarkXml ? withWatermarkNs(open) : open
        return `${originalXml.slice(0, openIdx)}${rootOpen}${parts.join('')}</${root}>`
      }
    }
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    `${WATERMARK_NS}>` +
    `${body}</${root}>`
  )
}

/** word/comments.xml regenerated from the comment list (plain-text bodies). */
/**
 * Surgical rebuild of comments.xml: existing comments whose text is unchanged keep
 * their original bytes (rich formatting/multiple paragraphs/inline hyperlinks are
 * preserved); only new or edited comments fall back to a plain-text rebuild.
 */
function buildCommentsXml(comments: CommentInfo[], originalXml: string | null): string {
  const originals = new Map<string, { text: string; xml: string }>()
  if (originalXml) {
    for (const m of originalXml.match(/<w:comment\s[^>]*>[\s\S]*?<\/w:comment>/g) ?? []) {
      const id = /w:id="([^"]+)"/.exec(m)?.[1]
      if (id) originals.set(id, { text: commentPlainText(m), xml: m })
    }
  }
  const body = comments
    .map((c) => {
      const orig = originals.get(c.id)
      if (orig && orig.text === c.text) return orig.xml
      if (orig) {
        // In-paragraph w:t-level patch: text edits keep the comment's rich formatting,
        // multiple paragraphs, and inline links
        const patched = patchParagraphTexts(orig.xml, c.text)
        if (patched !== null) return patched
      }
      const attrs =
        `w:id="${escapeXmlAttr(c.id)}" w:author="${escapeXmlAttr(c.author)}"` +
        (c.initials ? ` w:initials="${escapeXmlAttr(c.initials)}"` : '') +
        (c.date ? ` w:date="${escapeXmlAttr(c.date)}"` : '')
      const lines = c.text.split('\n')
      const paras = lines
        .map((line, i) => {
          // The last paragraph carries w14:paraId (the commentsExtended link key)
          const pid =
            i === lines.length - 1 && c.paraId ? ` w14:paraId="${escapeXmlAttr(c.paraId)}"` : ''
          return `<w:p${pid}><w:r><w:t xml:space="preserve">${escapeXmlText(line)}</w:t></w:r></w:p>`
        })
        .join('')
      return `<w:comment ${attrs}>${paras}</w:comment>`
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
    `${body}</w:comments>`
  )
}

/** word/commentsExtended.xml: one commentEx per comment (reply parent-child + resolved flag) */
function buildCommentsExtendedXml(comments: CommentInfo[]): string {
  const paraIdOf = new Map(comments.map((c) => [c.id, c.paraId]))
  const body = comments
    .filter((c) => c.paraId)
    .map((c) => {
      const parentParaId = c.parentId ? paraIdOf.get(c.parentId) : undefined
      return (
        `<w15:commentEx w15:paraId="${escapeXmlAttr(c.paraId!)}"` +
        (parentParaId ? ` w15:paraIdParent="${escapeXmlAttr(parentParaId)}"` : '') +
        ` w15:done="${c.done ? '1' : '0'}"/>`
      )
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
    ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15">' +
    `${body}</w15:commentsEx>`
  )
}

/** Aligned with parseComments' textOf: each w:p's w:t text, '\n' between paragraphs */
function commentPlainText(commentXml: string): string {
  const paras: string[] = []
  const pRe = /<w:p[\s>][\s\S]*?<\/w:p>|<w:p\/>/g
  let p: RegExpExecArray | null
  while ((p = pRe.exec(commentXml)) !== null) {
    const texts: string[] = []
    const tRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
    let t: RegExpExecArray | null
    while ((t = tRe.exec(p[0])) !== null) texts.push(t[1])
    paras.push(
      texts
        .join('')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&'),
    )
  }
  return paras.join('\n')
}

/** set or remove <w:documentProtection> right after the settings root opens */
function applyProtection(xml: string, protection: DocProtection | null): string {
  let out = xml.replace(/<w:documentProtection[^>]*\/>/, '')
  if (protection) {
    const crypt = protection.hash
      ? ' w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny"' +
        ` w:cryptAlgorithmSid="${protection.algorithmSid ?? 14}"` +
        ` w:cryptSpinCount="${protection.spinCount ?? 100000}"` +
        ` w:hash="${escapeXmlAttr(protection.hash)}"` +
        (protection.salt ? ` w:salt="${escapeXmlAttr(protection.salt)}"` : '')
      : ''
    const tag =
      `<w:documentProtection w:edit="${escapeXmlAttr(protection.edit)}"` +
      (protection.enforced ? ' w:enforcement="1"' : '') +
      crypt +
      '/>'
    out = out.replace(/(<w:settings[^>]*>)/, `$1${tag}`)
  }
  return out
}

/** set or remove <w:titlePg/> ("different first page"), before w:docGrid per schema order */
function applyTitlePg(sectPrXml: string, on: boolean): string {
  let xml = sectPrXml.replace(/<w:titlePg[^>]*\/>/, '')
  if (on) {
    if (/<w:docGrid/.test(xml)) xml = xml.replace(/<w:docGrid/, '<w:titlePg/><w:docGrid')
    else xml = xml.replace(/<\/w:sectPr>/, '<w:titlePg/></w:sectPr>')
  }
  return xml
}

/** set or remove <w:evenAndOddHeaders/> right after the settings root opens */
function applyEvenAndOddHeaders(xml: string, on: boolean): string {
  const out = xml.replace(/<w:evenAndOddHeaders[^>]*\/>/, '')
  return on ? out.replace(/(<w:settings[^>]*>)/, '$1<w:evenAndOddHeaders/>') : out
}

/** Set, replace or remove <w:background> (must be the first child of w:document). */
function applyPageColor(documentXml: string, color: string | null): string {
  let xml = documentXml.replace(/<w:background[^>]*\/>/, '')
  if (color) {
    xml = xml.replace(/(<w:document[^>]*>)/, `$1<w:background w:color="${escapeXmlAttr(color)}"/>`)
  }
  return xml
}

function maxRelId(relsXml: string | null): number {
  if (!relsXml) return 1000
  let max = 0
  const re = /Id="rId(\d+)"/g
  let m: RegExpExecArray | null
  while ((m = re.exec(relsXml)) !== null) {
    max = Math.max(max, parseInt(m[1], 10))
  }
  return max
}

function nextImageSeq(zip: JSZip): number {
  let max = 0
  for (const name of Object.keys(zip.files)) {
    const m = /^word\/media\/aidocs(\d+)\./.exec(name)
    if (m) max = Math.max(max, parseInt(m[1], 10))
  }
  return max + 1
}
