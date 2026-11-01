/**
 * Write the parsed .doc model as a real .docx.
 *
 * The body is assembled on top of docx-engine's blank template (styles.xml /
 * numbering.xml already present): one generated paragraph per model paragraph,
 * with run-level direct formatting and paragraph formatting. Stage 3b resolves
 * the stylesheet into that formatting; stage 3c-1 replaces the template's
 * trailing sectPr with the document's section geometry and adds header/footer
 * parts (plus settings.xml when even/odd headers are used), tables as
 * <w:tbl> and inline PNG/JPEG pictures as <w:drawing>. Footnotes and text
 * boxes are still omitted and reported in the honest warnings list.
 */
import JSZip from 'jszip'
import {
  buildBlankDocx,
  generateParagraphXml,
  type GenerateContext,
  type GeneratedBlock,
  type ParaFormat,
  type Run,
} from '@genoffice/docx-engine'
import type { FormattedParagraph, FormattedRun, ParagraphFormatting } from './formatting'
import type { HeaderFooterStory, SectionHeaderFooter } from './headers'
import { EMU_PER_TWIP, type ImageBlock } from './images'
import {
  isBulletLevel,
  levelTextForWriter,
  NUMBER_FORMATS,
  type ListDefinition,
  type ListLevel,
} from './lists'
import type { SectionInfo } from './sections'
import type {
  BodyBlock,
  LegacyDocBody,
  ParagraphListInfo,
  TableBorders,
  TableCell,
  TableModel,
  TableRow,
} from './tables'
import type { LegacyDocLayout, LegacyDocModel } from './index'

/** Word color index (sprmCIco 1..16) -> OOXML hex value */
const COLOR_BY_INDEX: Record<number, string> = {
  1: '000000',
  2: '0000FF',
  3: '00FFFF',
  4: '00FF00',
  5: 'FF00FF',
  6: 'FF0000',
  7: 'FFFF00',
  8: 'FFFFFF',
  9: '000080',
  10: '008080',
  11: '008000',
  12: '800080',
  13: '800000',
  14: '808000',
  15: '808080',
  16: 'C0C0C0',
}

/** Word highlight index (sprmCHighlight 1..16) -> OOXML w:highlight name */
const HIGHLIGHT_BY_INDEX: Record<number, string> = {
  1: 'black',
  2: 'blue',
  3: 'cyan',
  4: 'green',
  5: 'magenta',
  6: 'red',
  7: 'yellow',
  8: 'white',
  9: 'darkBlue',
  10: 'darkCyan',
  11: 'darkGreen',
  12: 'darkMagenta',
  13: 'darkRed',
  14: 'darkYellow',
  15: 'darkGray',
  16: 'lightGray',
}

/** plain paragraphs only: no hyperlinks, so this allocator must never be called */
const GENERATE_CONTEXT: GenerateContext = {
  headingStyleIds: new Map(),
  allocateHyperlinkRel: () => {
    throw new Error('no hyperlinks in stage 3a')
  },
}

/** underline kinds that mean "underlined"; the exact kind is a stage-3b nicety */
const UNDERLINE_KINDS = new Set([
  'single',
  'words',
  'double',
  'dotted',
  'thick',
  'dash',
  'dotDash',
  'dotDotDash',
  'wave',
])

/**
 * Map one stage-2 run to a docx-engine Run. Explicit "off" toggles are dropped
 * because the blank template's defaults are already off and docx-engine writes
 * w:b/w:i/w:u only when the value is true (styles that might turn them back on
 * arrive in stage 3b).
 */
export function toDocxRun(run: FormattedRun): Run {
  const formatting = run.formatting
  const out: Run = { text: run.text }
  if (formatting.bold) out.bold = true
  if (formatting.italic) out.italic = true
  if (formatting.strike) out.strike = true
  if (formatting.underline !== undefined && UNDERLINE_KINDS.has(formatting.underline))
    out.underline = true
  if (formatting.fontSizeHalfPoints !== undefined)
    out.sizeHalfPoints = formatting.fontSizeHalfPoints
  if (formatting.verticalAlign === 'superscript' || formatting.verticalAlign === 'subscript')
    out.vertAlign = formatting.verticalAlign
  if (formatting.highlight !== undefined) {
    const highlight = HIGHLIGHT_BY_INDEX[formatting.highlight]
    if (highlight !== undefined) out.highlight = highlight
  }
  if (formatting.colorIndex !== undefined && formatting.colorIndex !== 0) {
    const color = COLOR_BY_INDEX[formatting.colorIndex]
    if (color !== undefined) out.color = color
  }
  if (run.link) {
    out.link = {
      href: run.link.href,
      ...(run.link.tooltip ? { tooltip: run.link.tooltip } : {}),
    }
  }
  return out
}

/** Map stage-2 paragraph formatting to docx-engine's ParaFormat (undefined when empty). */
export function toParaFormat(formatting: ParagraphFormatting): ParaFormat | undefined {
  const out: ParaFormat = {}
  if (formatting.alignment !== undefined) out.align = formatting.alignment
  if (formatting.indentLeftTwips !== undefined) out.indentLeft = formatting.indentLeftTwips
  if (formatting.indentRightTwips !== undefined) out.indentRight = formatting.indentRightTwips
  if (formatting.indentFirstLineTwips !== undefined)
    out.indentFirstLine = formatting.indentFirstLineTwips
  if (formatting.spacingBeforeTwips !== undefined) out.spaceBefore = formatting.spacingBeforeTwips
  if (formatting.spacingAfterTwips !== undefined) out.spaceAfter = formatting.spacingAfterTwips
  if (formatting.keepNext) out.keepNext = true
  if (formatting.keepLines) out.keepLines = true
  if (formatting.pageBreakBefore) out.pageBreakBefore = true
  if (formatting.lineSpacing !== undefined) {
    if (formatting.lineSpacing.multiple) {
      // Word's single line spacing is 240 twips, so dyaLine/240 is the multiplier
      out.lineSpacing = formatting.lineSpacing.dyaLine / 240
    } else {
      out.lineRule = 'exact'
      out.lineRawTwips = formatting.lineSpacing.dyaLine
    }
  }
  return Object.keys(out).length > 0 ? out : undefined
}

function paragraphBlock(entry: {
  paragraph: FormattedParagraph
  list?: ParagraphListInfo
}): GeneratedBlock {
  const paragraph = entry.paragraph
  const block: GeneratedBlock = {
    type: entry.list ? 'listItem' : 'paragraph',
    runs: paragraph.runs.map(toDocxRun),
  }
  if (entry.list) block.list = { ...entry.list }
  const format = toParaFormat(paragraph.formatting)
  if (format) block.format = format
  return block
}

const W_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WORD_REL_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const HEADER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const FOOTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const SETTINGS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const HEADER_REL_TYPE = `${WORD_REL_NAMESPACE}/header`
const FOOTER_REL_TYPE = `${WORD_REL_NAMESPACE}/footer`
const SETTINGS_REL_TYPE = `${WORD_REL_NAMESPACE}/settings`
const IMAGE_REL_TYPE = `${WORD_REL_NAMESPACE}/image`
const DRAWING_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const CONTENT_TYPES_PATH = '[Content_Types].xml'
const DOCUMENT_RELS_PATH = 'word/_rels/document.xml.rels'
const SETTINGS_PATH = 'word/settings.xml'
const STORY_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/** one header/footer paragraph: tabs become w:tab, mapped line breaks become w:br */
function storyParagraphXml(paragraph: string): string {
  let runs = ''
  let text = ''
  const flushText = (): void => {
    if (text === '') return
    runs += `<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`
    text = ''
  }
  for (const char of paragraph) {
    if (char === '\t') {
      flushText()
      runs += '<w:r><w:tab/></w:r>'
    } else if (char === '\n') {
      flushText()
      runs += '<w:r><w:br/></w:r>'
    } else {
      text += char
    }
  }
  flushText()
  return runs === '' ? '<w:p/>' : `<w:p>${runs}</w:p>`
}

/** a word/headerN.xml or word/footerN.xml part; story formatting is plain text for now */
function storyPartXml(kind: 'header' | 'footer', story: HeaderFooterStory): string {
  const tag = kind === 'header' ? 'w:hdr' : 'w:ftr'
  const paragraphs = story.paragraphs.map(storyParagraphXml).join('')
  return `${STORY_XML_DECL}<${tag} xmlns:w="${W_NAMESPACE}">` + `${paragraphs || '<w:p/>'}</${tag}>`
}

interface HeaderFooterPart {
  path: string
  kind: 'header' | 'footer'
  referenceType: 'default' | 'first' | 'even'
  story: HeaderFooterStory
  rId: string
}

/**
 * The header/footer parts of one section. Word 97 references stories by order:
 * odd = default variant, first = first variant (needs w:titlePg), even = even
 * variant (needs settings.xml w:evenAndOddHeaders, added when needed).
 */
function collectHeaderFooterParts(
  stories: SectionHeaderFooter,
  section: SectionInfo,
): { parts: HeaderFooterPart[]; settingsNeeded: boolean } {
  const parts: HeaderFooterPart[] = []
  let headerCount = 0
  let footerCount = 0
  const add = (
    kind: 'header' | 'footer',
    referenceType: HeaderFooterPart['referenceType'],
    story: HeaderFooterStory,
  ): void => {
    const index = kind === 'header' ? ++headerCount : ++footerCount
    parts.push({
      path: `word/${kind === 'header' ? 'header' : 'footer'}${index}.xml`,
      kind,
      referenceType,
      story,
      rId: '',
    })
  }
  if (stories.oddHeader) add('header', 'default', stories.oddHeader)
  if (stories.oddFooter) add('footer', 'default', stories.oddFooter)
  if (section.titlePage) {
    if (stories.firstHeader) add('header', 'first', stories.firstHeader)
    if (stories.firstFooter) add('footer', 'first', stories.firstFooter)
  }
  if (stories.evenHeader) add('header', 'even', stories.evenHeader)
  if (stories.evenFooter) add('footer', 'even', stories.evenFooter)
  return { parts, settingsNeeded: parts.some((part) => part.referenceType === 'even') }
}

/** the generated trailing <w:sectPr> (header/footer references come before pgSz per CT_SectPr) */
function sectPrXml(section: SectionInfo, references: string): string {
  const orient = section.orientation === 'landscape' ? ' w:orient="landscape"' : ''
  const { margins } = section
  return (
    `<w:sectPr>${references}<w:pgSz w:w="${section.pageWidthTwips}" w:h="${section.pageHeightTwips}"${orient}/>` +
    `<w:pgMar w:top="${margins.top}" w:right="${margins.right}" w:bottom="${margins.bottom}"` +
    ` w:left="${margins.left}" w:header="${margins.header}" w:footer="${margins.footer}"` +
    ` w:gutter="${margins.gutter}"/>` +
    (section.titlePage ? '<w:titlePg/>' : '') +
    '</w:sectPr>'
  )
}

const W_BORDER_SIDES = ['top', 'left', 'bottom', 'right'] as const

/** one single-line border element; table borders always use a 4-eighth-point line */
function borderXml(side: string): string {
  return '<w:' + side + ' w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
}

/** the table border set: any visible border becomes the full single-line set (stage 3c-2) */
function tableBordersXml(borders: TableBorders): string {
  const visible =
    borders.top ||
    borders.left ||
    borders.bottom ||
    borders.right ||
    borders.insideHorizontal ||
    borders.insideVertical
  if (!visible) return ''
  const sides = W_BORDER_SIDES.map(borderXml).join('') + borderXml('insideH') + borderXml('insideV')
  return '<w:tblBorders>' + sides + '</w:tblBorders>'
}

function cellBordersXml(cell: TableCell): string {
  if (!cell.borders) return ''
  const visible = W_BORDER_SIDES.filter((side) => cell.borders![side])
  if (visible.length === 0) return ''
  return '<w:tcBorders>' + visible.map(borderXml).join('') + '</w:tcBorders>'
}

function tableCellXml(cell: TableCell, context: GenerateContext): string {
  const properties = ['<w:tcW w:w="' + cell.widthTwips + '" w:type="dxa"/>']
  if (cell.gridSpan > 1) properties.push('<w:gridSpan w:val="' + cell.gridSpan + '"/>')
  if (cell.verticalMerge === 'restart') properties.push('<w:vMerge w:val="restart"/>')
  else if (cell.verticalMerge === 'continue') properties.push('<w:vMerge/>')
  const borders = cellBordersXml(cell)
  if (borders !== '') properties.push(borders)
  const paragraphs =
    cell.paragraphs.length === 0
      ? '<w:p/>'
      : cell.paragraphs
          .map((paragraph) => generateParagraphXml(paragraphBlock({ paragraph }), context))
          .join('')
  return '<w:tc><w:tcPr>' + properties.join('') + '</w:tcPr>' + paragraphs + '</w:tc>'
}

function tableRowXml(row: TableRow, context: GenerateContext): string {
  const properties: string[] = []
  if (row.cantSplit) properties.push('<w:cantSplit/>')
  if (row.heightTwips !== undefined)
    properties.push(
      '<w:trHeight w:val="' +
        row.heightTwips +
        '" w:hRule="' +
        (row.heightExact ? 'exact' : 'atLeast') +
        '"/>',
    )
  if (row.header) properties.push('<w:tblHeader/>')
  const trPr = properties.length > 0 ? '<w:trPr>' + properties.join('') + '</w:trPr>' : ''
  return '<w:tr>' + trPr + row.cells.map((cell) => tableCellXml(cell, context)).join('') + '</w:tr>'
}

/** a complete <w:tbl> for one reconstructed table (fixed layout, explicit grid) */
export function generateTableXml(
  table: TableModel,
  context: GenerateContext = GENERATE_CONTEXT,
): string {
  const tableProperties = [
    '<w:tblW w:w="0" w:type="auto"/>',
    ...(table.alignment ? ['<w:jc w:val="' + table.alignment + '"/>'] : []),
    tableBordersXml(table.borders),
    '<w:tblLayout w:type="fixed"/>',
  ].filter((part) => part !== '')
  const grid =
    '<w:tblGrid>' +
    table.gridWidths.map((width) => '<w:gridCol w:w="' + width + '"/>').join('') +
    '</w:tblGrid>'
  const rows = table.rows.map((row) => tableRowXml(row, context)).join('')
  return '<w:tbl><w:tblPr>' + tableProperties.join('') + '</w:tblPr>' + grid + rows + '</w:tbl>'
}

/** an inline picture paragraph; the blank template declares the wp/a/pic/r namespaces */
function generateImageParagraphXml(image: ImageBlock, rId: string, pictureNumber: number): string {
  const cx = image.widthTwips * EMU_PER_TWIP
  const cy = image.heightTwips * EMU_PER_TWIP
  const name = 'Picture ' + pictureNumber
  return (
    '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="' +
    cx +
    '" cy="' +
    cy +
    '"/>' +
    '<wp:docPr id="' +
    pictureNumber +
    '" name="' +
    name +
    '"/>' +
    '<a:graphic><a:graphicData uri="' +
    DRAWING_URI +
    '">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="' +
    pictureNumber +
    '" name="' +
    name +
    '"/>' +
    '<pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="' +
    rId +
    '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' +
    cx +
    '" cy="' +
    cy +
    '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
  )
}

/** Generated XML for the reconstructed body (paragraphs, tables and images in CP order). */
export function generateBodyBlocksXml(
  blocks: BodyBlock[],
  imageRelIds: Map<number, string> = new Map(),
  context: GenerateContext = GENERATE_CONTEXT,
): string {
  let pictureNumber = 0
  return blocks
    .map((block) => {
      if (block.type === 'table') return generateTableXml(block.table, context)
      if (block.type === 'image') {
        pictureNumber++
        return generateImageParagraphXml(
          block.image,
          imageRelIds.get(block.image.cp) ?? '',
          pictureNumber,
        )
      }
      return generateParagraphXml(paragraphBlock(block), context)
    })
    .join('')
}

const BODY_OPEN = '<w:body>'
const BODY_PLACEHOLDER = '<w:p/>'
const DOCUMENT_PATH = 'word/document.xml'

/** Generated <w:p> XML for every model paragraph. */
export function generateBodyParagraphsXml(paragraphs: FormattedParagraph[]): string {
  return paragraphs
    .map((paragraph) => generateParagraphXml(paragraphBlock({ paragraph }), GENERATE_CONTEXT))
    .join('')
}

const NUMBERING_PATH = 'word/numbering.xml'
const HYPERLINK_REL_TYPE = `${WORD_REL_NAMESPACE}/hyperlink`

/** the w:ind of one numbering level (level props come from the LVL grpprlPapx) */
function numberingIndentXml(level: ListLevel): string {
  const attributes: string[] = []
  if (level.indentLeftTwips !== undefined) attributes.push('w:left="' + level.indentLeftTwips + '"')
  if (level.indentRightTwips !== undefined)
    attributes.push('w:right="' + level.indentRightTwips + '"')
  if (level.indentFirstLineTwips !== undefined) {
    if (level.indentFirstLineTwips >= 0)
      attributes.push('w:firstLine="' + level.indentFirstLineTwips + '"')
    else attributes.push('w:hanging="' + Math.abs(level.indentFirstLineTwips) + '"')
  }
  return attributes.length > 0 ? '<w:ind ' + attributes.join(' ') + '/>' : ''
}

/** one w:lvl of a reconstructed list definition */
function numberingLevelXml(level: ListLevel, ilvl: number): string {
  const bullet = isBulletLevel(level)
  const numberFormat = bullet ? 'bullet' : (NUMBER_FORMATS[level.numberFormat] ?? 'decimal')
  const levelText = escapeXmlText(levelTextForWriter(level))
  const indent = numberingIndentXml(level)
  return (
    '<w:lvl w:ilvl="' +
    ilvl +
    '"><w:start w:val="' +
    level.startAt +
    '"/><w:numFmt w:val="' +
    numberFormat +
    '"/>' +
    (level.legal ? '<w:isLgl/>' : '') +
    '<w:suff w:val="' +
    level.follow +
    '"/><w:lvlText w:val="' +
    levelText +
    '"/><w:lvlJc w:val="' +
    level.alignment +
    '"/>' +
    (indent !== '' ? '<w:pPr>' + indent + '</w:pPr>' : '') +
    '</w:lvl>'
  )
}

/**
 * Insert the reconstructed list definitions into the blank template's
 * numbering.xml: new abstractNum definitions (ids from 10) after the existing
 * abstract numbers and the w:num entries (ids 100 + ilfo) before the closing
 * tag, keeping the CT_Numbering abstractNum* -> num* order.
 */
function reconstructNumberingXml(
  numbering: string,
  used: Array<{ numId: number; ilfo: number }>,
  lists: Map<number, number>,
  definitions: Map<number, ListDefinition>,
): string {
  const abstractIdByLsid = new Map<number, number>()
  const abstractXml: string[] = []
  const numXml: string[] = []
  for (const entry of used) {
    const lsid = lists.get(entry.ilfo)
    const definition = lsid !== undefined ? definitions.get(lsid) : undefined
    if (!definition) continue
    let abstractId = abstractIdByLsid.get(definition.lsid)
    if (abstractId === undefined) {
      abstractId = 10 + abstractIdByLsid.size
      abstractIdByLsid.set(definition.lsid, abstractId)
      abstractXml.push(
        '<w:abstractNum w:abstractNumId="' +
          abstractId +
          '"><w:multiLevelType w:val="hybridMultilevel"/>' +
          definition.levels.map(numberingLevelXml).join('') +
          '</w:abstractNum>',
      )
    }
    numXml.push(
      '<w:num w:numId="' + entry.numId + '"><w:abstractNumId w:val="' + abstractId + '"/></w:num>',
    )
  }
  if (abstractXml.length === 0) return numbering
  const firstNum = numbering.indexOf('<w:num ')
  let out =
    firstNum >= 0
      ? numbering.slice(0, firstNum) + abstractXml.join('') + numbering.slice(firstNum)
      : numbering.replace('</w:numbering>', abstractXml.join('') + '</w:numbering>')
  out = out.replace('</w:numbering>', numXml.join('') + '</w:numbering>')
  return out
}
/**
 * Write the formatting model as a .docx: start from docx-engine's blank
 * template, replace its single placeholder paragraph inside <w:body> with the
 * generated paragraphs, replace the template's trailing sectPr with the
 * document's section geometry and add the header/footer parts. Documents
 * without a section table keep the template sectPr unchanged.
 */
export async function writeLegacyDocx(
  model: LegacyDocModel,
  layout?: LegacyDocLayout,
  body?: LegacyDocBody,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildBlankDocx())
  const documentFile = zip.file(DOCUMENT_PATH)
  if (!documentFile) throw new Error('the blank docx template has no word/document.xml')
  const template = await documentFile.async('string')
  const bodyStart = template.indexOf(BODY_OPEN)
  if (bodyStart < 0) throw new Error('the blank docx template has no <w:body>')
  const placeholderStart = template.indexOf(BODY_PLACEHOLDER, bodyStart)
  if (placeholderStart < 0)
    throw new Error('the blank docx template has no placeholder paragraph in <w:body>')

  const parts: HeaderFooterPart[] = []
  let settingsNeeded = false
  let nextRelId = 3
  let generatedSectPr: string | undefined
  if (layout && layout.sections.length > 0) {
    const section = layout.sections[layout.sections.length - 1]!
    const stories = layout.headerFooter[layout.headerFooter.length - 1] ?? {}
    const collected = collectHeaderFooterParts(stories, section)
    settingsNeeded = collected.settingsNeeded
    for (const part of collected.parts) part.rId = `rId${nextRelId++}`
    parts.push(...collected.parts)
    const references = parts
      .map((part) => {
        const element = part.kind === 'header' ? 'headerReference' : 'footerReference'
        return `<w:${element} w:type="${part.referenceType}" r:id="${part.rId}"/>`
      })
      .join('')
    generatedSectPr = sectPrXml(section, references)
  }

  interface MediaPart {
    path: string
    extension: 'png' | 'jpeg'
    contentType: string
    bytes: Uint8Array
    rId: string
  }
  const mediaParts: MediaPart[] = []
  const imageRelIds = new Map<number, string>()
  if (body) {
    let mediaIndex = 0
    for (const block of body.blocks) {
      if (block.type !== 'image') continue
      mediaIndex++
      const extension = block.image.mediaType === 'image/png' ? 'png' : 'jpeg'
      const mediaPart: MediaPart = {
        path: `word/media/image${mediaIndex}.${extension}`,
        extension,
        contentType: block.image.mediaType,
        bytes: block.image.bytes,
        rId: `rId${nextRelId++}`,
      }
      imageRelIds.set(block.image.cp, mediaPart.rId)
      mediaParts.push(mediaPart)
    }
  }

  // numbered paragraphs: one w:num per used ilfo, definitions patched into numbering.xml
  const usedNumbering: Array<{ numId: number; ilfo: number }> = []
  if (body) {
    const seenNumIds = new Set<number>()
    for (const block of body.blocks) {
      if (block.type !== 'paragraph' || !block.list) continue
      const numId = Number(block.list.numId)
      if (!Number.isFinite(numId) || seenNumIds.has(numId)) continue
      seenNumIds.add(numId)
      usedNumbering.push({ numId, ilfo: numId - 100 })
    }
  }

  // external hyperlink relationships get rIds while the body runs are generated
  const hyperlinkRels: Array<{ rId: string; href: string }> = []
  const generateContext: GenerateContext = {
    headingStyleIds: new Map(),
    allocateHyperlinkRel: (href: string) => {
      const rId = `rId${nextRelId++}`
      hyperlinkRels.push({ rId, href })
      return rId
    },
  }

  const usesBodyBlocks =
    body !== undefined &&
    body.blocks.some(
      (block) =>
        block.type === 'table' ||
        block.type === 'image' ||
        (block.type === 'paragraph' &&
          (block.list !== undefined || block.paragraph.runs.some((run) => run.link !== undefined))),
    )
  const bodyXml = usesBodyBlocks
    ? generateBodyBlocksXml(body!.blocks, imageRelIds, generateContext)
    : generateBodyParagraphsXml(model.paragraphs)
  let documentXml =
    template.slice(0, placeholderStart) +
    bodyXml +
    template.slice(placeholderStart + BODY_PLACEHOLDER.length)
  if (generatedSectPr) {
    const sectStart = documentXml.lastIndexOf('<w:sectPr')
    if (sectStart >= 0) {
      const sectEnd = documentXml.indexOf('</w:sectPr>', sectStart)
      if (sectEnd >= 0) {
        documentXml =
          documentXml.slice(0, sectStart) +
          generatedSectPr +
          documentXml.slice(sectEnd + '</w:sectPr>'.length)
      }
    }
  }
  zip.file(DOCUMENT_PATH, documentXml)

  if (
    parts.length > 0 ||
    settingsNeeded ||
    mediaParts.length > 0 ||
    hyperlinkRels.length > 0 ||
    usedNumbering.length > 0
  ) {
    const contentTypesFile = zip.file(CONTENT_TYPES_PATH)
    if (!contentTypesFile) throw new Error('the blank docx template has no [Content_Types].xml')
    let contentTypes = await contentTypesFile.async('string')
    const extensions = [...new Set(mediaParts.map((part) => part.extension))]
    const defaults = extensions
      .map((extension) => {
        const contentType =
          mediaParts.find((part) => part.extension === extension)?.contentType ?? 'image/png'
        return `<Default Extension="${extension}" ContentType="${contentType}"/>`
      })
      .join('')
    const overrides =
      parts
        .map(
          (part) =>
            `<Override PartName="/${part.path}" ContentType="${
              part.kind === 'header' ? HEADER_CONTENT_TYPE : FOOTER_CONTENT_TYPE
            }"/>`,
        )
        .join('') +
      (settingsNeeded
        ? `<Override PartName="/${SETTINGS_PATH}" ContentType="${SETTINGS_CONTENT_TYPE}"/>`
        : '')
    if (defaults !== '') {
      const firstOverride = contentTypes.indexOf('<Override')
      contentTypes =
        firstOverride >= 0
          ? contentTypes.slice(0, firstOverride) + defaults + contentTypes.slice(firstOverride)
          : contentTypes.replace('</Types>', `${defaults}</Types>`)
    }
    contentTypes = contentTypes.replace('</Types>', `${overrides}</Types>`)
    zip.file(CONTENT_TYPES_PATH, contentTypes)

    const relsFile = zip.file(DOCUMENT_RELS_PATH)
    if (!relsFile) throw new Error('the blank docx template has no document.xml.rels')
    let relationships = await relsFile.async('string')
    const newRelationships =
      parts
        .map(
          (part) =>
            `<Relationship Id="${part.rId}" Type="${
              part.kind === 'header' ? HEADER_REL_TYPE : FOOTER_REL_TYPE
            }" Target="${part.path.slice('word/'.length)}"/>`,
        )
        .join('') +
      mediaParts
        .map(
          (part) =>
            `<Relationship Id="${part.rId}" Type="${IMAGE_REL_TYPE}" Target="${part.path.slice(
              'word/'.length,
            )}"/>`,
        )
        .join('') +
      hyperlinkRels
        .map(
          (rel) =>
            `<Relationship Id="${rel.rId}" Type="${HYPERLINK_REL_TYPE}" Target="${escapeXmlText(
              rel.href,
            )}" TargetMode="External"/>`,
        )
        .join('') +
      (settingsNeeded
        ? `<Relationship Id="rId${nextRelId}" Type="${SETTINGS_REL_TYPE}" Target="settings.xml"/>`
        : '')
    relationships = relationships.replace('</Relationships>', `${newRelationships}</Relationships>`)
    zip.file(DOCUMENT_RELS_PATH, relationships)
    for (const part of parts) zip.file(part.path, storyPartXml(part.kind, part.story))
    for (const part of mediaParts) zip.file(part.path, part.bytes)
    if (settingsNeeded)
      zip.file(
        SETTINGS_PATH,
        `${STORY_XML_DECL}<w:settings xmlns:w="${W_NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
      )
    if (usedNumbering.length > 0 && body?.lists) {
      const numberingFile = zip.file(NUMBERING_PATH)
      if (numberingFile) {
        const numbering = await numberingFile.async('string')
        const patched = reconstructNumberingXml(
          numbering,
          usedNumbering,
          body.lists.overrides,
          body.lists.lists,
        )
        if (patched !== numbering) zip.file(NUMBERING_PATH, patched)
      }
    }
  }

  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}
export interface ConversionWarningOptions {
  /** the model was read through the style resolver, so styles are applied (not a gap) */
  stylesResolved?: boolean
  /** section geometry and header/footer stories read from the document */
  layout?: LegacyDocLayout
  /** tables were reconstructed and written as w:tbl (stage 3c-2) */
  tablesHandled?: boolean
}

/**
 * Honest conversion warnings: what this stage knows it does not carry over yet,
 * plus the non-fatal parse/resolve warnings from earlier stages.
 */
export function conversionWarnings(
  model: LegacyDocModel,
  options: ConversionWarningOptions = {},
): string[] {
  const warnings = [...model.warnings]
  const usesStyles = model.paragraphs.some(
    (paragraph) =>
      (paragraph.formatting.styleIndex ?? 0) !== 0 ||
      paragraph.runs.some((run) => (run.formatting.styleIndex ?? 0) !== 0),
  )
  if (usesStyles && !options.stylesResolved)
    warnings.push(
      'paragraph and character styles are not resolved yet (stage 3b); the text is imported with direct formatting only',
    )
  const flattensTables = model.paragraphs.some(
    (paragraph) =>
      paragraph.formatting.inTable === true || paragraph.formatting.tableRowEnd === true,
  )
  if (flattensTables && !options.tablesHandled)
    warnings.push('tables are flattened to paragraphs (stage 3a); cell structure is not preserved')
  const { fib } = model
  if (fib.ccpFtn > 0) warnings.push('footnotes are not converted yet; their text is omitted')
  if (fib.ccpEdn > 0) warnings.push('endnotes are not converted yet; their text is omitted')
  const layout = options.layout
  if (layout) {
    warnings.push(...layout.warnings)
    if (layout.sections.length > 1)
      warnings.push(
        `${layout.sections.length} sections were flattened to one; only the last section's page setup and header/footer variants are kept`,
      )
  }
  const headersConverted = layout !== undefined && layout.headerFooter.length > 0
  if (fib.ccpHdd > 0 && !headersConverted)
    warnings.push('headers are not converted yet; their text is omitted')
  if (fib.ccpTxbx + fib.ccpHdrTxbx > 0)
    warnings.push('text boxes are not converted yet; their text is omitted')
  return warnings
}

/**
 * Read a legacy .doc and convert it into a .docx. Documents with a stylesheet
 * are read through the style resolver (stage 3b), so style-based formatting is
 * written as direct formatting; styles themselves are not emitted as .docx
 * styles yet. Section geometry and header/footer stories are written for the
 * last section (multi-section documents are flattened with a warning).
 */
export async function convertLegacyDocToDocx(
  bytes: Uint8Array,
): Promise<{ docx: Uint8Array; warnings: string[] }> {
  // dynamic import keeps the module graph acyclic: index.ts re-exports this file
  const { readLegacyDocBody, readLegacyDocLayout, readLegacyDocModel, readResolvedLegacyDocModel } =
    await import('./index')
  const layout = readLegacyDocLayout(bytes)
  const base = readLegacyDocModel(bytes)
  const stylesResolved = base.fib.lcbStshf > 0
  const model = stylesResolved ? readResolvedLegacyDocModel(bytes) : base
  const body = readLegacyDocBody(bytes)
  const tablesHandled = body.blocks.some((block) => block.type === 'table')
  const docx = await writeLegacyDocx(model, layout, body)
  const warnings = conversionWarnings(model, { stylesResolved, layout, tablesHandled })
  for (const warning of body.warnings) if (!warnings.includes(warning)) warnings.push(warning)
  return { docx, warnings }
}
