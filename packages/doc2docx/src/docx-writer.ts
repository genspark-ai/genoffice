/**
 * Stage 3a: write the stage-2 formatting model as a real .docx.
 *
 * The body is assembled on top of docx-engine's blank template (styles.xml /
 * numbering.xml already present): one generated paragraph per model paragraph,
 * with run-level direct formatting and paragraph formatting. Styles are not
 * resolved (stage 3b), tables are flattened to paragraphs, and non-body stories
 * (footnotes, headers, text boxes) are not converted yet; convertLegacyDocToDocx
 * reports each of these honestly in its warnings list.
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
import type { LegacyDocModel } from './index'

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

function paragraphBlock(paragraph: FormattedParagraph): GeneratedBlock {
  const block: GeneratedBlock = {
    type: 'paragraph',
    runs: paragraph.runs.map(toDocxRun),
  }
  const format = toParaFormat(paragraph.formatting)
  if (format) block.format = format
  return block
}

const BODY_OPEN = '<w:body>'
const BODY_PLACEHOLDER = '<w:p/>'
const DOCUMENT_PATH = 'word/document.xml'

/** Generated <w:p> XML for every model paragraph. */
export function generateBodyParagraphsXml(paragraphs: FormattedParagraph[]): string {
  return paragraphs
    .map((paragraph) => generateParagraphXml(paragraphBlock(paragraph), GENERATE_CONTEXT))
    .join('')
}

/**
 * Write the formatting model as a .docx: start from docx-engine's blank
 * template, replace its single placeholder paragraph inside <w:body> with the
 * generated paragraphs (keeping the template's root namespaces and trailing
 * sectPr), and rewrite word/document.xml.
 */
export async function writeLegacyDocx(model: LegacyDocModel): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildBlankDocx())
  const documentFile = zip.file(DOCUMENT_PATH)
  if (!documentFile) throw new Error('the blank docx template has no word/document.xml')
  const template = await documentFile.async('string')
  const bodyStart = template.indexOf(BODY_OPEN)
  if (bodyStart < 0) throw new Error('the blank docx template has no <w:body>')
  const placeholderStart = template.indexOf(BODY_PLACEHOLDER, bodyStart)
  if (placeholderStart < 0)
    throw new Error('the blank docx template has no placeholder paragraph in <w:body>')
  const withBody =
    template.slice(0, placeholderStart) +
    generateBodyParagraphsXml(model.paragraphs) +
    template.slice(placeholderStart + BODY_PLACEHOLDER.length)
  zip.file(DOCUMENT_PATH, withBody)
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

export interface ConversionWarningOptions {
  /** the model was read through the style resolver, so styles are applied (not a gap) */
  stylesResolved?: boolean
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
  if (flattensTables)
    warnings.push('tables are flattened to paragraphs (stage 3a); cell structure is not preserved')
  const { fib } = model
  if (fib.ccpFtn > 0) warnings.push('footnotes are not converted yet; their text is omitted')
  if (fib.ccpEdn > 0) warnings.push('endnotes are not converted yet; their text is omitted')
  if (fib.ccpHdd > 0) warnings.push('headers are not converted yet; their text is omitted')
  if (fib.ccpTxbx + fib.ccpHdrTxbx > 0)
    warnings.push('text boxes are not converted yet; their text is omitted')
  return warnings
}

/**
 * Read a legacy .doc and convert it into a .docx. Documents with a stylesheet
 * are read through the style resolver (stage 3b), so style-based formatting is
 * written as direct formatting; styles themselves are not emitted as .docx
 * styles yet.
 */
export async function convertLegacyDocToDocx(
  bytes: Uint8Array,
): Promise<{ docx: Uint8Array; warnings: string[] }> {
  // dynamic import keeps the module graph acyclic: index.ts re-exports this file
  const { readLegacyDocModel, readResolvedLegacyDocModel } = await import('./index')
  const base = readLegacyDocModel(bytes)
  const stylesResolved = base.fib.lcbStshf > 0
  const model = stylesResolved ? readResolvedLegacyDocModel(bytes) : base
  const docx = await writeLegacyDocx(model)
  return { docx, warnings: conversionWarnings(model, { stylesResolved }) }
}
