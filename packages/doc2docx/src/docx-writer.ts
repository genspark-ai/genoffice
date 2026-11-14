/**
 * Write the parsed .doc model as a real .docx.
 *
 * The body is assembled on top of docx-engine's blank template (styles.xml /
 * numbering.xml already present): one generated paragraph per model paragraph,
 * with run-level direct formatting and paragraph formatting. Stage 3b resolves
 * the stylesheet into that formatting; stage 3c-1 replaces the template's
 * trailing sectPr with the last section's geometry, ends earlier sections in
 * their last paragraph's sectPr and adds header/footer
 * parts (plus settings.xml when even/odd headers are used), tables as
 * <w:tbl>, inline PNG/JPEG pictures as <w:drawing> and text boxes as anchored
 * wps shapes; footnotes/endnotes go to their own parts.
 */
import { compatOptionsOf } from './compat'
import JSZip from 'jszip'
import {
  buildBlankDocx,
  generateParagraphXml,
  mergeRPrModel,
  type GenerateContext,
  type GeneratedBlock,
  type ParaFormat,
  type Run,
} from '@genoffice/docx-engine'
import {
  COLOR_BY_INDEX,
  characterFormattingOf,
  shadingFillOf,
  type CharacterFormatting,
  type FormattedParagraph,
  type Shading,
  type FormattedRun,
  type ParagraphFormatting,
  type ParagraphFrame,
  isPositionedFrame,
} from './formatting'
import type { HeaderFooterStory, SectionHeaderFooter } from './headers'
import { fontNamesFor } from './fonts'
import { legacyStyleId } from './styles'
import type { LegacyStyleDef } from './resolve'
import type { LegacyNote, NoteKind } from './notes'
import { EMU_PER_TWIP, type ImageBlock } from './images'
import {
  PICTURE_EXTENSIONS,
  type FloatingShape,
  type FloatingTextBox,
  type PictureCrop,
  type PictureMediaType,
} from './drawings'
import {
  isBulletLevel,
  levelTextForWriter,
  NUMBER_FORMATS,
  type ListLevel,
  type ListTables,
} from './lists'
import type { PageBorders, SectionBreakType, SectionColumns, SectionInfo } from './sections'
import type { FontEntry } from './fonts'
import type { LegacyComment } from './comments'
import type {
  BodyBlock,
  BorderLine,
  BorderStyles,
  LegacyDocBody,
  ParagraphListInfo,
  ParagraphPictures,
  TableBorders,
  TableCell,
  TableCellPadding,
  TableModel,
  TablePosition,
  TableRow,
} from './tables'
import { forEachTableCell } from './tables'
import type { LegacyDocLayout, LegacyDocModel } from './index'

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
  // the run's character style (w:rStyle) — the resolver stamped the id only
  // when the style exists and is not the null "Default Paragraph Font"
  if (formatting.styleId) out.styleId = formatting.styleId
  if (formatting.bold) out.bold = true
  if (formatting.italic) out.italic = true
  if (formatting.rtl) out.rtl = true
  if (formatting.strike) out.strike = true
  if (formatting.underline !== undefined && UNDERLINE_KINDS.has(formatting.underline))
    out.underline = true
  if (formatting.fontSizeHalfPoints !== undefined)
    out.sizeHalfPoints = formatting.fontSizeHalfPoints
  if (formatting.verticalAlign === 'superscript' || formatting.verticalAlign === 'subscript')
    out.vertAlign = formatting.verticalAlign
  if (formatting.spacingTwips) out.charSpacingTwips = formatting.spacingTwips
  if (formatting.scalePercent !== undefined && formatting.scalePercent !== 100)
    out.charScalePct = formatting.scalePercent
  if (formatting.kerningHalfPoints !== undefined) out.kernHalfPoints = formatting.kerningHalfPoints
  if (formatting.positionHalfPoints) out.positionHalfPoints = formatting.positionHalfPoints
  if (formatting.highlight !== undefined) {
    const highlight = HIGHLIGHT_BY_INDEX[formatting.highlight]
    if (highlight !== undefined) out.highlight = highlight
  }
  if (formatting.colorIndex !== undefined && formatting.colorIndex !== 0) {
    const color = COLOR_BY_INDEX[formatting.colorIndex]
    if (color !== undefined) out.color = color
  }
  if (formatting.colorRgb !== undefined) out.color = formatting.colorRgb
  if (formatting.hidden) out.vanishOwn = true
  if (formatting.caps) out.caps = 'all'
  else if (formatting.smallCaps) out.caps = 'small'
  const fonts = formatting.fontNames
  if (fonts) {
    const primary = fonts.eastAsia ?? fonts.ascii
    if (primary !== undefined) out.font = primary
    if (fonts.ascii !== undefined) out.fontAscii = fonts.ascii
    if (fonts.eastAsia !== undefined) out.eastAsiaFont = fonts.eastAsia
    if (fonts.cs !== undefined) out.fontCs = fonts.cs
  }
  if (run.link) {
    out.link = {
      href: run.link.href,
      ...(run.link.tooltip ? { tooltip: run.link.tooltip } : {}),
      ...(run.link.plain ? { plain: true } : {}),
    }
  }
  if (run.field) out.text = FIELD_OPEN + run.field + FIELD_SPLIT + run.text + FIELD_CLOSE
  if (run.note) {
    const { kind, role, custom, id } = run.note
    out.text =
      NOTE_OPEN + kind[0] + role[0] + (custom ? 'c' : 'a') + id + NOTE_SPLIT + out.text + NOTE_CLOSE
  }
  if (run.commentMarks)
    out.text = run.commentMarks
      .map((mark) => COMMENT_MARK_OPEN + mark.kind[0] + mark.id + COMMENT_MARK_CLOSE)
      .join('')
  if (run.pictureCp !== undefined) out.text = PICTURE_OPEN + run.pictureCp + PICTURE_CLOSE
  return out
}

/* private-use markers standing for an inline picture run (its \x01 CP) inside the paragraph text */
const PICTURE_OPEN = '\ue008'
const PICTURE_CLOSE = '\ue009'
const PICTURE_RUN = new RegExp(
  '<w:r>(?:(?!</w:r>).)*?<w:t xml:space="preserve">' +
    PICTURE_OPEN +
    '(\\d+)' +
    PICTURE_CLOSE +
    '</w:t></w:r>',
  'g',
)

/** replace the picture marker runs by their drawing runs (markers without a picture are dropped) */
function expandPictureRuns(xml: string, runsByCp: Map<number, string>): string {
  if (!xml.includes(PICTURE_OPEN)) return xml
  return xml.replace(PICTURE_RUN, (_match, cp: string) => {
    const run = runsByCp.get(Number(cp))
    if (run === undefined) return ''
    runsByCp.delete(Number(cp))
    return run
  })
}

/* private-use markers standing for comment anchors (kind letter + id) inside a run's text */
const COMMENT_MARK_OPEN = '\ue006'
const COMMENT_MARK_CLOSE = '\ue007'
const COMMENT_MARK_RUN = new RegExp(
  '<w:r>(<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:t xml:space="preserve">((?:' +
    COMMENT_MARK_OPEN +
    '[ser]\\d+' +
    COMMENT_MARK_CLOSE +
    ')+)</w:t></w:r>',
  'g',
)
const COMMENT_MARK = new RegExp(COMMENT_MARK_OPEN + '([ser])(\\d+)' + COMMENT_MARK_CLOSE, 'g')
const COMMENT_REFERENCE_STYLE = '<w:rStyle w:val="CommentReference"/>'

/** expand the comment marker runs into commentRangeStart / commentRangeEnd / commentReference */
function expandCommentMarks(xml: string): string {
  if (!xml.includes(COMMENT_MARK_OPEN)) return xml
  return xml.replace(COMMENT_MARK_RUN, (_match, rPr: string | undefined, marks: string) => {
    const props = rPr
      ? rPr.replace('<w:rPr>', '<w:rPr>' + COMMENT_REFERENCE_STYLE)
      : '<w:rPr>' + COMMENT_REFERENCE_STYLE + '</w:rPr>'
    return marks.replace(COMMENT_MARK, (_mark, kind: string, id: string) => {
      if (kind === 's') return '<w:commentRangeStart w:id="' + id + '"/>'
      if (kind === 'e') return '<w:commentRangeEnd w:id="' + id + '"/>'
      return '<w:r>' + props + '<w:commentReference w:id="' + id + '"/></w:r>'
    })
  })
}

/* private-use markers wrapping a live field's instruction and cached result inside a run's text */
const FIELD_OPEN = '\ue000'
const FIELD_SPLIT = '\ue001'
const FIELD_CLOSE = '\ue002'
/* the optional rPr must stay inside its own run: a lazy `.*?` would reach into earlier runs */
const RUN_RPR = '(<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?'
const FIELD_RUN = new RegExp(
  '<w:r>' +
    RUN_RPR +
    '<w:t xml:space="preserve">' +
    FIELD_OPEN +
    '([^' +
    FIELD_SPLIT +
    ']*)' +
    FIELD_SPLIT +
    '(.*?)' +
    FIELD_CLOSE +
    '</w:t></w:r>',
  'g',
)

/** expand the field marker runs into fldChar begin / instrText / separate / result / end runs */
function expandFieldRuns(xml: string): string {
  if (!xml.includes(FIELD_OPEN)) return xml
  return xml.replace(FIELD_RUN, (_match, rPr: string | undefined, instruction, result) => {
    const props = rPr ?? ''
    return (
      '<w:r>' +
      props +
      '<w:fldChar w:fldCharType="begin"/></w:r><w:r>' +
      props +
      '<w:instrText xml:space="preserve"> ' +
      instruction +
      ' </w:instrText></w:r><w:r>' +
      props +
      '<w:fldChar w:fldCharType="separate"/></w:r><w:r>' +
      props +
      '<w:t xml:space="preserve">' +
      result +
      '</w:t></w:r><w:r>' +
      props +
      '<w:fldChar w:fldCharType="end"/></w:r>'
    )
  })
}

/* private-use markers wrapping a note mark run: kind/role/custom + id, then the custom mark text */
const NOTE_OPEN = '\ue003'
const NOTE_SPLIT = '\ue004'
const NOTE_CLOSE = '\ue005'
const NOTE_RUN = new RegExp(
  '<w:r>' +
    RUN_RPR +
    '<w:t xml:space="preserve">' +
    NOTE_OPEN +
    '([fe])([rs])([ac])(\\d+)' +
    NOTE_SPLIT +
    '(.*?)' +
    NOTE_CLOSE +
    '</w:t></w:r>',
  'g',
)

const NOTE_STYLE_IDS: Record<NoteKind, { reference: string; text: string }> = {
  footnote: { reference: 'FootnoteReference', text: 'FootnoteText' },
  endnote: { reference: 'EndnoteReference', text: 'EndnoteText' },
}

/** the mark's own rPr with the reference character style in front (rStyle leads CT_RPr) */
function noteMarkRPr(kind: NoteKind, rPr: string | undefined): string {
  const inner = rPr ? rPr.slice('<w:rPr>'.length, -'</w:rPr>'.length) : ''
  return `<w:rPr><w:rStyle w:val="${NOTE_STYLE_IDS[kind].reference}"/>${inner}</w:rPr>`
}

/** a note's own number mark run (w:footnoteRef / w:endnoteRef) */
function noteSelfMarkXml(kind: NoteKind, rPr?: string): string {
  return `<w:r>${noteMarkRPr(kind, rPr)}<w:${kind}Ref/></w:r>`
}

/** expand the note marker runs into footnoteReference / footnoteRef runs (endnotes alike) */
function expandNoteRuns(xml: string): string {
  if (!xml.includes(NOTE_OPEN)) return xml
  return xml.replace(
    NOTE_RUN,
    (_match, rPr: string | undefined, kindCode, role, custom, id, mark: string) => {
      const kind: NoteKind = kindCode === 'f' ? 'footnote' : 'endnote'
      if (role === 's') return noteSelfMarkXml(kind, rPr)
      if (custom === 'c')
        return (
          `<w:r>${rPr ?? ''}<w:${kind}Reference w:customMarkFollows="1" w:id="${id}"/>` +
          `<w:t xml:space="preserve">${mark}</w:t></w:r>`
        )
      return `<w:r>${noteMarkRPr(kind, rPr)}<w:${kind}Reference w:id="${id}"/></w:r>`
    },
  )
}

/** Map stage-2 paragraph formatting to docx-engine's ParaFormat (undefined when empty). */
export function toParaFormat(formatting: ParagraphFormatting): ParaFormat | undefined {
  const out: ParaFormat = {}
  if (formatting.bidi) out.bidi = true
  if (formatting.alignment !== undefined) {
    const flip = formatting.bidi && formatting.alignmentLogical
    out.align =
      flip && formatting.alignment === 'left'
        ? 'right'
        : flip && formatting.alignment === 'right'
          ? 'left'
          : formatting.alignment
  }
  if (formatting.indentLeftTwips !== undefined) out.indentLeft = formatting.indentLeftTwips
  if (formatting.indentRightTwips !== undefined) out.indentRight = formatting.indentRightTwips
  if (formatting.indentFirstLineTwips !== undefined)
    out.indentFirstLine = formatting.indentFirstLineTwips
  if (formatting.spacingBeforeTwips !== undefined) out.spaceBefore = formatting.spacingBeforeTwips
  if (formatting.spacingAfterTwips !== undefined) out.spaceAfter = formatting.spacingAfterTwips
  if (formatting.keepNext) out.keepNext = true
  if (formatting.keepLines) out.keepLines = true
  if (formatting.shading) {
    const fill = shadingFillOf(formatting.shading)
    if (fill !== undefined) out.shadingFill = fill
  }
  if (formatting.pageBreakBefore) out.pageBreakBefore = true
  if (formatting.lineSpacing !== undefined) {
    if (formatting.lineSpacing.multiple) {
      // Word's single line spacing is 240 twips, so dyaLine/240 is the multiplier
      out.lineSpacing = formatting.lineSpacing.dyaLine / 240
    } else {
      // a negative dyaLine is an exact height, a positive one a minimum (MS-DOC LSPD)
      out.lineRule = formatting.lineSpacing.dyaLine < 0 ? 'exact' : 'atLeast'
      out.lineRawTwips = Math.abs(formatting.lineSpacing.dyaLine)
    }
  }
  if (formatting.tabStops && formatting.tabStops.length > 0)
    out.tabStops = formatting.tabStops.map((stop) => ({
      pos: stop.positionTwips,
      val: stop.alignment,
      ...(stop.leader ? { leader: stop.leader } : {}),
    }))
  if (formatting.borders) {
    const sides = { top: 't', left: 'l', bottom: 'b', right: 'r' } as const
    let borders = ''
    const borderLines: NonNullable<ParaFormat['borderLines']> = {}
    for (const side of ['top', 'left', 'bottom', 'right'] as const) {
      const border = formatting.borders[side]
      if (!border) continue
      borders += sides[side]
      borderLines[sides[side]] = {
        szPt: border.widthEighths / 8,
        spacePt: border.spacePoints,
        ...(border.color ? { color: border.color } : {}),
      }
    }
    if (borders !== '') {
      out.borders = borders
      out.borderLines = borderLines
    }
  }
  const firstLineChars = formatting.charIndents?.firstLine
  if (firstLineChars !== undefined && firstLineChars > 0)
    out.charIndents = { firstLine: firstLineChars }
  return Object.keys(out).length > 0 ? out : undefined
}

function paragraphBlock(entry: {
  paragraph: FormattedParagraph
  list?: ParagraphListInfo
  styleId?: string
}): GeneratedBlock {
  const paragraph = entry.paragraph
  const block: GeneratedBlock = {
    type: entry.list ? 'listItem' : 'paragraph',
    runs: paragraph.runs.map(toDocxRun),
  }
  if (entry.list) block.list = { ...entry.list }
  // the notes pass their own dedicated style; a body paragraph references the
  // legacy stylesheet entry the resolver stamped onto its formatting
  const styleId = entry.styleId ?? paragraph.formatting.styleId
  if (styleId) block.styleId = styleId
  const format = toParaFormat(paragraph.formatting)
  if (format) block.format = format
  return block
}

const FRAME_X_ALIGNS: Record<number, string> = {
  [-4]: 'center',
  [-8]: 'right',
  [-12]: 'inside',
  [-16]: 'outside',
}
const FRAME_Y_ALIGNS: Record<number, string> = {
  [-4]: 'top',
  [-8]: 'center',
  [-12]: 'bottom',
  [-16]: 'inside',
  [-20]: 'outside',
}

/** w:framePr for a positioned paragraph; a negative alignment code becomes xAlign/yAlign */
export function framePrXml(frame: ParagraphFrame): string {
  const attrs: string[] = []
  if (frame.widthTwips) attrs.push(`w:w="${frame.widthTwips}"`)
  if (frame.heightTwips) {
    attrs.push(`w:h="${frame.heightTwips}"`)
    attrs.push(`w:hRule="${frame.heightExact ? 'exact' : 'atLeast'}"`)
  }
  if (frame.hSpaceTwips) attrs.push(`w:hSpace="${frame.hSpaceTwips}"`)
  if (frame.vSpaceTwips) attrs.push(`w:vSpace="${frame.vSpaceTwips}"`)
  attrs.push(`w:wrap="${frame.wrap ?? 'auto'}"`)
  attrs.push(`w:vAnchor="${frame.vAnchor ?? 'margin'}"`)
  attrs.push(`w:hAnchor="${frame.hAnchor ?? 'text'}"`)
  const x = frame.xTwips ?? 0
  const xAlign = FRAME_X_ALIGNS[x]
  attrs.push(xAlign ? `w:xAlign="${xAlign}"` : `w:x="${x}"`)
  const y = frame.yTwips ?? 0
  const yAlign = FRAME_Y_ALIGNS[y]
  attrs.push(yAlign ? `w:yAlign="${yAlign}"` : `w:y="${y}"`)
  return `<w:framePr ${attrs.join(' ')}/>`
}

/** framePr follows pStyle/keepNext/keepLines/pageBreakBefore in CT_PPrBase */
const PPR_BEFORE_FRAME =
  /^(<w:pPr>)((?:<w:(?:pStyle|keepNext|keepLines|pageBreakBefore)\b[^>]*\/>)*)/

function withFramePr(xml: string, framePr: string): string {
  if (xml === '<w:p/>') return '<w:p><w:pPr>' + framePr + '</w:pPr></w:p>'
  const open = xml.match(/^<w:p(?: [^>]*)?>/)![0]
  const rest = xml.slice(open.length)
  if (!rest.startsWith('<w:pPr>')) return open + '<w:pPr>' + framePr + '</w:pPr>' + rest
  return (
    open + rest.replace(PPR_BEFORE_FRAME, (_m, pPr: string, lead: string) => pPr + lead + framePr)
  )
}

/** generateParagraphXml plus the paragraph mark's rPr, which docx-engine's block model does not carry */
function paragraphXml(
  entry: { paragraph: FormattedParagraph; list?: ParagraphListInfo; styleId?: string },
  context: GenerateContext,
): string {
  let xml = expandCommentMarks(
    expandNoteRuns(expandFieldRuns(generateParagraphXml(paragraphBlock(entry), context))),
  )
  if (isPositionedFrame(entry.paragraph.formatting.frame))
    xml = withFramePr(xml, framePrXml(entry.paragraph.formatting.frame))
  const mark = entry.paragraph.markFormatting
  if (!mark) return xml
  const rPr = mergeRPrModel('', toDocxRun({ text: '', formatting: mark }), false)
  if (rPr === '') return xml
  if (xml.includes('</w:pPr>')) return xml.replace('</w:pPr>', rPr + '</w:pPr>')
  if (xml.startsWith('<w:p/>')) return '<w:p><w:pPr>' + rPr + '</w:pPr></w:p>'
  return xml.replace(/^<w:p(?: [^>]*)?>/, (open) => open + '<w:pPr>' + rPr + '</w:pPr>')
}

const W_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WORD_REL_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const HEADER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const FOOTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
const SETTINGS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const NOTES_CONTENT_TYPE: Record<NoteKind, string> = {
  footnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
}
const NOTES_PATH: Record<NoteKind, string> = {
  footnote: 'word/footnotes.xml',
  endnote: 'word/endnotes.xml',
}
const HEADER_REL_TYPE = `${WORD_REL_NAMESPACE}/header`
const FOOTER_REL_TYPE = `${WORD_REL_NAMESPACE}/footer`
const SETTINGS_REL_TYPE = `${WORD_REL_NAMESPACE}/settings`
const FONT_TABLE_PATH = 'word/fontTable.xml'
const FONT_TABLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'
const FONT_TABLE_REL_TYPE = `${WORD_REL_NAMESPACE}/fontTable`
const NOTES_REL_TYPE: Record<NoteKind, string> = {
  footnote: `${WORD_REL_NAMESPACE}/footnotes`,
  endnote: `${WORD_REL_NAMESPACE}/endnotes`,
}
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
const COMMENTS_REL_TYPE = `${WORD_REL_NAMESPACE}/comments`
const COMMENTS_PATH = 'word/comments.xml'
const IMAGE_REL_TYPE = `${WORD_REL_NAMESPACE}/image`
const DRAWING_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const WPS_NAMESPACE = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
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

const STORY_NAMESPACES =
  `xmlns:w="${W_NAMESPACE}" xmlns:r="${WORD_REL_NAMESPACE}"` +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ` xmlns:pic="${DRAWING_URI}" xmlns:wps="${WPS_NAMESPACE}"`

/** the paragraphs of a header/footer or text box story (plain text when the story has no formatted read) */
function storyParagraphsXml(
  story: HeaderFooterStory,
  pictures: PictureRefs,
  kind: StoryKind,
): string {
  if (story.blocks)
    return generateBodyBlocksXml(story.blocks, pictures, GENERATE_CONTEXT, new Map(), kind)
  const runsFor = (index: number): { leading: string; trailing: string } => {
    let leading = (story.pictures?.get(index) ?? [])
      .map((picture) => anchoredShapeRunXml(picture, pictures, kind))
      .join('')
    let trailing = ''
    for (const entry of story.images?.get(index) ?? []) {
      const run = inlinePictureRunXml(entry.image, pictures)
      if (entry.position === 'start') leading += run
      else trailing += run
    }
    return { leading, trailing }
  }
  const withRuns = (xml: string, index: number): string => {
    const runs = runsFor(index)
    return withPictureRuns(xml, runs.leading, runs.trailing)
  }
  const paragraphs = story.formatted
    ? story.formatted
        .map((paragraph, index) => withRuns(paragraphXml({ paragraph }, GENERATE_CONTEXT), index))
        .join('')
    : story.paragraphs
        .map((paragraph, index) => withRuns(storyParagraphXml(paragraph), index))
        .join('')
  return paragraphs || '<w:p/>'
}

/** a word/headerN.xml or word/footerN.xml part */
function storyPartXml(
  kind: 'header' | 'footer',
  story: HeaderFooterStory,
  pictures: PictureRefs,
): string {
  const tag = kind === 'header' ? 'w:hdr' : 'w:ftr'
  return (
    `${STORY_XML_DECL}<${tag} ${STORY_NAMESPACES}>` +
    storyParagraphsXml(story, pictures, kind) +
    `</${tag}>`
  )
}

const SEPARATOR_PPR = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'

/** word/footnotes.xml or word/endnotes.xml: Word's separator entries (ids -1/0), then the notes */
function notesPartXml(kind: NoteKind, notes: LegacyNote[]): string {
  const root = `w:${kind}s`
  const entry = `w:${kind}`
  const styleId = NOTE_STYLE_IDS[kind].text
  const body = notes
    .map((note) => {
      const paragraphs = note.paragraphs.map((paragraph) =>
        paragraphXml({ paragraph, styleId }, GENERATE_CONTEXT),
      )
      if (paragraphs.length === 0)
        paragraphs.push('<w:p><w:pPr><w:pStyle w:val="' + styleId + '"/></w:pPr></w:p>')
      const hasSelfMark = note.paragraphs.some((paragraph) =>
        paragraph.runs.some((run) => run.note?.role === 'self'),
      )
      if (!note.custom && !hasSelfMark)
        paragraphs[0] = withPictureRuns(paragraphs[0]!, noteSelfMarkXml(kind))
      return `<${entry} w:id="${note.id}">${paragraphs.join('')}</${entry}>`
    })
    .join('')
  return (
    `${STORY_XML_DECL}<${root} ${STORY_NAMESPACES}>` +
    `<${entry} w:type="separator" w:id="-1"><w:p>${SEPARATOR_PPR}<w:r><w:separator/></w:r></w:p></${entry}>` +
    `<${entry} w:type="continuationSeparator" w:id="0"><w:p>${SEPARATOR_PPR}<w:r><w:continuationSeparator/></w:r></w:p></${entry}>` +
    body +
    `</${root}>`
  )
}

/** the Footnote/Endnote Text and Reference styles Word expects the note parts to use */
function noteStylesXml(kind: NoteKind): string {
  const ids = NOTE_STYLE_IDS[kind]
  return (
    `<w:style w:type="paragraph" w:styleId="${ids.text}"><w:name w:val="${kind} text"/>` +
    '<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>' +
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
    `<w:style w:type="character" w:styleId="${ids.reference}"><w:name w:val="${kind} reference"/>` +
    '<w:uiPriority w:val="99"/><w:unhideWhenUsed/>' +
    '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
  )
}

/** word/comments.xml: Word opens every comment with an annotation reference run */
function commentsPartXml(comments: LegacyComment[]): string {
  const annotationRef =
    '<w:r><w:rPr>' + COMMENT_REFERENCE_STYLE + '</w:rPr><w:annotationRef/></w:r>'
  const entries = comments.map((comment) => {
    const paragraphs = comment.formatted
      ? comment.formatted.map((paragraph) => paragraphXml({ paragraph }, GENERATE_CONTEXT))
      : comment.paragraphs.map(storyParagraphXml)
    if (paragraphs.length === 0) paragraphs.push('<w:p/>')
    paragraphs[0] = withPictureRuns(paragraphs[0]!, annotationRef)
    const attributes =
      ` w:id="${comment.id}" w:author="${escapeXmlText(comment.author)}"` +
      (comment.date ? ` w:date="${comment.date}"` : '') +
      ` w:initials="${escapeXmlText(comment.initials)}"`
    return `<w:comment${attributes}>${paragraphs.join('')}</w:comment>`
  })
  return `${STORY_XML_DECL}<w:comments ${STORY_NAMESPACES}>${entries.join('')}</w:comments>`
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
 * variant (needs settings.xml w:evenAndOddHeaders, added when needed). A
 * section without a story of its own inherits the previous section's, which
 * docx expresses by omitting the reference. Part numbers continue across
 * sections (`counts`).
 */
function collectHeaderFooterParts(
  stories: SectionHeaderFooter,
  section: SectionInfo,
  counts = { header: 0, footer: 0 },
): { parts: HeaderFooterPart[]; settingsNeeded: boolean } {
  const parts: HeaderFooterPart[] = []
  const add = (
    kind: 'header' | 'footer',
    referenceType: HeaderFooterPart['referenceType'],
    story: HeaderFooterStory,
  ): void => {
    const index = ++counts[kind]
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

/** a section that declared a grid (sprmSClm) keeps it; the line pitch is what Word snaps lines to */
function docGridXml(section: SectionInfo): string {
  const grid = section.docGrid
  if (!grid || grid.linePitchTwips <= 0) return ''
  return (
    '<w:docGrid' +
    (grid.type ? ` w:type="${grid.type}"` : '') +
    ` w:linePitch="${grid.linePitchTwips}"` +
    (grid.charSpace ? ` w:charSpace="${grid.charSpace}"` : '') +
    '/>'
  )
}

const SECTION_TYPES: Record<SectionBreakType, string> = {
  continuous: 'continuous',
  column: 'nextColumn',
  nextPage: 'nextPage',
  evenPage: 'evenPage',
  oddPage: 'oddPage',
}

function pgNumTypeXml(section: SectionInfo): string {
  const attrs: string[] = []
  if (section.pageNumberFormat && section.pageNumberFormat !== 'decimal')
    attrs.push(`w:fmt="${section.pageNumberFormat}"`)
  if (section.pageNumberStart !== undefined) attrs.push(`w:start="${section.pageNumberStart}"`)
  return attrs.length > 0 ? `<w:pgNumType ${attrs.join(' ')}/>` : ''
}

/** w:cols for a multi-column section; explicit widths are written as w:col children */
function colsXml(columns: SectionColumns | undefined): string {
  if (!columns || columns.count <= 1) return ''
  const widths = columns.widths
  const explicit = widths !== undefined && widths.length === columns.count
  const attrs =
    `w:num="${columns.count}"` +
    ` w:space="${columns.spaceTwips}"` +
    (columns.separator ? ' w:sep="1"' : '') +
    (explicit ? ' w:equalWidth="0"' : '')
  if (!explicit) return `<w:cols ${attrs}/>`
  const cols = widths
    .map((width, index) => {
      const space = columns.spacings?.[index]
      return index < widths.length - 1
        ? `<w:col w:w="${width}" w:space="${space ?? columns.spaceTwips}"/>`
        : `<w:col w:w="${width}"/>`
    })
    .join('')
  return `<w:cols ${attrs}>${cols}</w:cols>`
}

function pgBordersXml(borders: PageBorders | undefined): string {
  if (!borders) return ''
  const sides = W_BORDER_SIDES.filter((side) => borders.sides[side])
    .map((side) => borderXml(side, borders.sides[side], borders.sides[side]!.spacePoints))
    .join('')
  if (!sides) return ''
  return (
    `<w:pgBorders w:offsetFrom="${borders.offsetFrom}"` +
    (borders.display ? ` w:display="${borders.display}"` : '') +
    (borders.zOrder ? ` w:zOrder="${borders.zOrder}"` : '') +
    `>${sides}</w:pgBorders>`
  )
}

/** the generated <w:sectPr> (CT_SectPr order: references, type, pgSz, pgMar, pgBorders, ...) */
function sectPrXml(section: SectionInfo, references: string): string {
  const orient = section.orientation === 'landscape' ? ' w:orient="landscape"' : ''
  const { margins } = section
  const type =
    section.breakType && section.breakType !== 'nextPage'
      ? `<w:type w:val="${SECTION_TYPES[section.breakType]}"/>`
      : ''
  return (
    `<w:sectPr>${references}${type}<w:pgSz w:w="${section.pageWidthTwips}" w:h="${section.pageHeightTwips}"${orient}/>` +
    `<w:pgMar w:top="${margins.top}" w:right="${margins.right}" w:bottom="${margins.bottom}"` +
    ` w:left="${margins.left}" w:header="${margins.header}" w:footer="${margins.footer}"` +
    ` w:gutter="${margins.gutter}"/>` +
    pgBordersXml(section.pageBorders) +
    pgNumTypeXml(section) +
    colsXml(section.columns) +
    (section.titlePage ? '<w:titlePg/>' : '') +
    docGridXml(section) +
    '</w:sectPr>'
  )
}

const W_BORDER_SIDES = ['top', 'left', 'bottom', 'right'] as const

/** one border element: the parsed line style, or Word's default single 0.5pt line */
function borderXml(side: string, line?: BorderLine, spacePoints = 0): string {
  return (
    '<w:' +
    side +
    ' w:val="' +
    (line?.style ?? 'single') +
    '" w:sz="' +
    (line?.widthEighths ?? 4) +
    '" w:space="' +
    spacePoints +
    '" w:color="' +
    (line?.color ?? 'auto') +
    '"/>'
  )
}

/** the table border set: any visible border becomes the full border set (stage 3c-2) */
function tableBordersXml(borders: TableBorders, styles: BorderStyles = {}): string {
  const visible =
    borders.top ||
    borders.left ||
    borders.bottom ||
    borders.right ||
    borders.insideHorizontal ||
    borders.insideVertical
  if (!visible) return ''
  const sides =
    W_BORDER_SIDES.filter((side) => borders[side])
      .map((side) => borderXml(side, styles[side]))
      .join('') +
    (borders.insideHorizontal ? borderXml('insideH', styles.insideHorizontal) : '') +
    (borders.insideVertical ? borderXml('insideV', styles.insideVertical) : '')
  return '<w:tblBorders>' + sides + '</w:tblBorders>'
}

/** the table-level border each side of a cell would inherit (outer edge or inside line) */
interface CellEdges {
  firstRow: boolean
  lastRow: boolean
  firstCell: boolean
  lastCell: boolean
  table: TableBorders
}

function inheritedBorder(side: (typeof W_BORDER_SIDES)[number], edges: CellEdges): boolean {
  const { table } = edges
  if (side === 'top') return edges.firstRow ? table.top : table.insideHorizontal
  if (side === 'bottom') return edges.lastRow ? table.bottom : table.insideHorizontal
  if (side === 'left') return edges.firstCell ? table.left : table.insideVertical
  return edges.lastCell ? table.right : table.insideVertical
}

/** a TC80 side that is off must be written as nil where the table borders would otherwise draw it */
function cellBordersXml(cell: TableCell, edges?: CellEdges): string {
  if (!cell.borders) return ''
  const sides = W_BORDER_SIDES.filter(
    (side) => cell.borders![side] || (edges !== undefined && inheritedBorder(side, edges)),
  )
  if (sides.length === 0) return ''
  return (
    '<w:tcBorders>' +
    sides
      .map((side) =>
        cell.borders![side]
          ? borderXml(side, cell.borderStyles?.[side])
          : '<w:' + side + ' w:val="nil"/>',
      )
      .join('') +
    '</w:tcBorders>'
  )
}

function shadingXml(shading: Shading): string {
  return (
    '<w:shd w:val="' +
    shading.pattern +
    '" w:color="' +
    (shading.color ?? 'auto') +
    '" w:fill="' +
    (shading.fill ?? 'auto') +
    '"/>'
  )
}

/** w:tblCellMar / w:tcMar: the sides that are set, in schema order */
function cellMarginsXml(element: string, padding: TableCellPadding): string {
  const sides = W_BORDER_SIDES.filter((side) => padding[side] !== undefined)
  if (sides.length === 0) return ''
  return (
    '<w:' +
    element +
    '>' +
    sides.map((side) => '<w:' + side + ' w:w="' + padding[side] + '" w:type="dxa"/>').join('') +
    '</w:' +
    element +
    '>'
  )
}

type StoryKind = 'body' | 'header' | 'footer'

function tableCellXml(
  cell: TableCell,
  context: GenerateContext,
  pictures: PictureRefs,
  story: StoryKind = 'body',
  edges?: CellEdges,
): string {
  const properties = ['<w:tcW w:w="' + cell.widthTwips + '" w:type="dxa"/>']
  if (cell.gridSpan > 1) properties.push('<w:gridSpan w:val="' + cell.gridSpan + '"/>')
  if (cell.verticalMerge === 'restart') properties.push('<w:vMerge w:val="restart"/>')
  else if (cell.verticalMerge === 'continue') properties.push('<w:vMerge/>')
  const borders = cellBordersXml(cell, edges)
  if (borders !== '') properties.push(borders)
  if (cell.shading) properties.push(shadingXml(cell.shading))
  if (cell.padding) properties.push(cellMarginsXml('tcMar', cell.padding))
  if (cell.vertAlign) properties.push('<w:vAlign w:val="' + cell.vertAlign + '"/>')
  const nestedBefore = (index: number): string =>
    (cell.nestedTables?.get(index) ?? [])
      .map((table) => generateTableXml(table, context, pictures, story))
      .join('')
  // a cell always ends with a paragraph, also after a trailing nested table
  const paragraphs =
    cell.paragraphs
      .map((paragraph, index) => {
        return (
          nestedBefore(index) +
          withParagraphPictures(
            paragraphXml({ paragraph, list: cell.paragraphLists?.get(index) }, context),
            cell.paragraphPictures?.get(index),
            pictures,
            story,
          )
        )
      })
      .join('') +
    (cell.paragraphs.length === 0 || cell.nestedTables?.has(cell.paragraphs.length)
      ? nestedBefore(cell.paragraphs.length) + '<w:p/>'
      : '')
  return '<w:tc><w:tcPr>' + properties.join('') + '</w:tcPr>' + paragraphs + '</w:tc>'
}

function tableRowXml(
  row: TableRow,
  context: GenerateContext,
  pictures: PictureRefs,
  story: StoryKind = 'body',
  rowEdges?: Pick<CellEdges, 'firstRow' | 'lastRow' | 'table'>,
): string {
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
  return (
    '<w:tr>' +
    trPr +
    row.cells
      .map((cell, index) =>
        tableCellXml(
          cell,
          context,
          pictures,
          story,
          rowEdges && {
            ...rowEdges,
            firstCell: index === 0,
            lastCell: index === row.cells.length - 1,
          },
        ),
      )
      .join('') +
    '</w:tr>'
  )
}

/** w:tblpPr (+ w:tblOverlap) of a floating table; the negative codes map like paragraph frames */
function tablePositionXml(position: TablePosition | undefined): string {
  if (!position) return ''
  const attrs: string[] = []
  if (position.leftFromText) attrs.push(`w:leftFromText="${position.leftFromText}"`)
  if (position.rightFromText) attrs.push(`w:rightFromText="${position.rightFromText}"`)
  if (position.topFromText) attrs.push(`w:topFromText="${position.topFromText}"`)
  if (position.bottomFromText) attrs.push(`w:bottomFromText="${position.bottomFromText}"`)
  attrs.push(`w:vertAnchor="${position.vAnchor ?? 'margin'}"`)
  attrs.push(`w:horzAnchor="${position.hAnchor ?? 'text'}"`)
  const x = position.xTwips ?? 0
  const xAlign = FRAME_X_ALIGNS[x]
  attrs.push(xAlign ? `w:tblpXSpec="${xAlign}"` : `w:tblpX="${x}"`)
  const y = position.yTwips ?? 0
  const yAlign = FRAME_Y_ALIGNS[y]
  attrs.push(yAlign ? `w:tblpYSpec="${yAlign}"` : `w:tblpY="${y}"`)
  return (
    `<w:tblpPr ${attrs.join(' ')}/>` + (position.noOverlap ? '<w:tblOverlap w:val="never"/>' : '')
  )
}

/** a complete <w:tbl> for one reconstructed table (fixed layout, explicit grid) */
export function generateTableXml(
  table: TableModel,
  context: GenerateContext = GENERATE_CONTEXT,
  pictures: PictureRefs = NO_PICTURES,
  story: StoryKind = 'body',
): string {
  const tableProperties = [
    tablePositionXml(table.position),
    '<w:tblW w:w="0" w:type="auto"/>',
    ...(table.alignment ? ['<w:jc w:val="' + table.alignment + '"/>'] : []),
    tableBordersXml(table.borders, table.borderStyles),
    ...(table.shading ? [shadingXml(table.shading)] : []),
    '<w:tblLayout w:type="fixed"/>',
    cellMarginsXml('tblCellMar', table.cellPadding),
  ].filter((part) => part !== '')
  const grid =
    '<w:tblGrid>' +
    table.gridWidths.map((width) => '<w:gridCol w:w="' + width + '"/>').join('') +
    '</w:tblGrid>'
  const rows = table.rows
    .map((row, index) =>
      tableRowXml(row, context, pictures, story, {
        firstRow: index === 0,
        lastRow: index === table.rows.length - 1,
        table: table.borders,
      }),
    )
    .join('')
  return '<w:tbl><w:tblPr>' + tableProperties.join('') + '</w:tblPr>' + grid + rows + '</w:tbl>'
}

/** rId lookups and the docPr id counter shared by every picture of the document */
export interface PictureRefs {
  /** relationship id of a picture's media part, scoped to the part being written */
  relIdFor(media: { bytes: Uint8Array }): string
  nextDocPrId(): number
}

const NO_PICTURES: PictureRefs = {
  relIdFor: () => '',
  nextDocPrId: () => 1,
}

/** a:srcRect in 1/1000 of a percent; omitted when nothing is cropped */
function srcRectXml(crop: PictureCrop | undefined): string {
  if (!crop) return ''
  const sides = (['l', 't', 'r', 'b'] as const)
    .filter((side) => crop[side] !== 0)
    .map((side) => ` ${side}="${crop[side]}"`)
    .join('')
  return sides === '' ? '' : `<a:srcRect${sides}/>`
}

function pictureGraphicXml(
  rId: string,
  docPrId: number,
  cx: number,
  cy: number,
  crop?: PictureCrop,
): string {
  const name = 'Picture ' + docPrId
  return (
    '<wp:docPr id="' +
    docPrId +
    '" name="' +
    name +
    '"/><wp:cNvGraphicFramePr/>' +
    '<a:graphic><a:graphicData uri="' +
    DRAWING_URI +
    '">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="' +
    docPrId +
    '" name="' +
    name +
    '"/>' +
    '<pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="' +
    rId +
    '"/>' +
    srcRectXml(crop) +
    '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' +
    cx +
    '" cy="' +
    cy +
    '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic>'
  )
}

/** an inline picture run; the blank template declares the wp/a/pic/r namespaces */
function inlinePictureRunXml(image: ImageBlock, pictures: PictureRefs): string {
  const cx = image.widthTwips * EMU_PER_TWIP
  const cy = image.heightTwips * EMU_PER_TWIP
  return (
    '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="' +
    cx +
    '" cy="' +
    cy +
    '"/>' +
    pictureGraphicXml(pictures.relIdFor(image), pictures.nextDocPrId(), cx, cy, image.crop) +
    '</wp:inline></w:drawing></w:r>'
  )
}

/** a text box as a wordprocessingShape: rect geometry, fill/line, then the story as w:txbxContent */
function textBoxGraphicXml(
  box: FloatingTextBox,
  pictures: PictureRefs,
  docPrId: number,
  cx: number,
  cy: number,
  story: 'body' | 'header' | 'footer',
): string {
  const name = 'Text Box ' + docPrId
  const fill = box.fill
    ? `<a:solidFill><a:srgbClr val="${box.fill}"/></a:solidFill>`
    : '<a:noFill/>'
  const line = box.line
    ? `<a:ln w="${box.line.widthEmu}"><a:solidFill><a:srgbClr val="${box.line.color}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>'
  const [left, top, right, bottom] = box.insetsEmu
  return (
    `<wp:docPr id="${docPrId}" name="${name}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${WPS_NAMESPACE}"><wps:wsp>` +
    `<wps:cNvPr id="${docPrId}" name="${name}"/><wps:cNvSpPr txBox="1"/>` +
    `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${box.geometry}"><a:avLst/></a:prstGeom>${fill}${line}</wps:spPr>` +
    `<wps:txbx><w:txbxContent>${storyParagraphsXml(box.content, pictures, story)}</w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" vert="${box.textDirection ?? 'horz'}" wrap="square" lIns="${left}" tIns="${top}" rIns="${right}" bIns="${bottom}"` +
    ` anchor="${box.verticalAnchor}" anchorCtr="0" upright="1">` +
    (box.autoFit ? '<a:spAutoFit/>' : '<a:noAutofit/>') +
    '</wps:bodyPr></wps:wsp></a:graphicData></a:graphic>'
  )
}

/** Word's base z-order for anchored drawings */
const ANCHOR_RELATIVE_HEIGHT = 251658240

/** a floating picture or text box run: wp:anchor positioned from the FSPA rectangle */
function anchoredShapeRunXml(
  picture: FloatingShape,
  pictures: PictureRefs,
  story: 'body' | 'header' | 'footer' = 'body',
): string {
  // Word positions a header shape's "text"-relative top from the page top
  const verticalRelative =
    story === 'header' && picture.verticalRelative === 'paragraph'
      ? 'page'
      : picture.verticalRelative
  const cx = picture.widthTwips * EMU_PER_TWIP
  const cy = picture.heightTwips * EMU_PER_TWIP
  const docPrId = pictures.nextDocPrId()
  const wrap =
    picture.wrap === 'topAndBottom'
      ? '<wp:wrapTopAndBottom/>'
      : picture.wrap === 'none'
        ? '<wp:wrapNone/>'
        : '<wp:wrapSquare wrapText="bothSides"/>'
  return (
    '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="' +
    (ANCHOR_RELATIVE_HEIGHT + docPrId) +
    '" behindDoc="' +
    (picture.behindText ? '1' : '0') +
    '" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="' +
    picture.horizontalRelative +
    '"><wp:posOffset>' +
    picture.leftTwips * EMU_PER_TWIP +
    '</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="' +
    verticalRelative +
    '"><wp:posOffset>' +
    picture.topTwips * EMU_PER_TWIP +
    '</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="' +
    cx +
    '" cy="' +
    cy +
    '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    wrap +
    (picture.kind === 'picture'
      ? pictureGraphicXml(pictures.relIdFor(picture.media), docPrId, cx, cy, picture.crop)
      : textBoxGraphicXml(picture, pictures, docPrId, cx, cy, story)) +
    '</wp:anchor></w:drawing></w:r>'
  )
}

/** insert picture runs right after the paragraph properties (or at the end for trailing ones) */
function withPictureRuns(paragraphXml: string, leading: string, trailing = ''): string {
  if (leading === '' && trailing === '') return paragraphXml
  let xml = paragraphXml === '<w:p/>' ? '<w:p></w:p>' : paragraphXml
  const openEnd = xml.indexOf('>') + 1
  const at = xml.startsWith('<w:pPr', openEnd)
    ? xml.indexOf('</w:pPr>') + '</w:pPr>'.length
    : openEnd
  xml = xml.slice(0, at) + leading + xml.slice(at)
  if (trailing !== '') xml = xml.slice(0, xml.lastIndexOf('</w:p>')) + trailing + '</w:p>'
  return xml
}

/** the paragraph with its anchored shapes in front and its inline pictures at their marker runs */
function withParagraphPictures(
  paragraphXml: string,
  pictures: ParagraphPictures | undefined,
  refs: PictureRefs,
  story: StoryKind = 'body',
): string {
  if (!pictures) return paragraphXml
  let leading = (pictures.floating ?? [])
    .map((picture) => anchoredShapeRunXml(picture, refs, story))
    .join('')
  const inline = pictures.inline ?? []
  const runsByCp = new Map(
    inline.map((entry) => [entry.image.cp, inlinePictureRunXml(entry.image, refs)]),
  )
  const xml = expandPictureRuns(paragraphXml, runsByCp)
  let trailing = ''
  for (const entry of inline) {
    const run = runsByCp.get(entry.image.cp)
    if (run === undefined) continue
    if (entry.position === 'start') leading += run
    else trailing += run
  }
  return withPictureRuns(xml, leading, trailing)
}

/**
 * Section ends keyed by the CP of the section mark, i.e. the `endCp` of the
 * paragraph that carries the sectPr; the value is that section's <w:sectPr>.
 */
export type SectionBreaks = Map<number, string>

/** put the sectPr last in the paragraph properties (after the mark rPr, per CT_PPr) */
function withSectPr(xml: string, sectPr: string): string {
  if (xml.includes('</w:pPr>')) return xml.replace('</w:pPr>', sectPr + '</w:pPr>')
  if (xml === '<w:p/>') return '<w:p><w:pPr>' + sectPr + '</w:pPr></w:p>'
  return xml.replace(/^<w:p(?: [^>]*)?>/, (open) => open + '<w:pPr>' + sectPr + '</w:pPr>')
}

/**
 * Places each section's sectPr: inside the paragraph ending at the section
 * mark when there is one, otherwise (the mark sits inside a table or the
 * paragraph was not emitted) in an empty paragraph right after the block
 * covering it.
 */
class SectionEndWriter {
  private readonly pending: number[]
  constructor(private readonly breaks: SectionBreaks) {
    this.pending = [...breaks.keys()].sort((left, right) => left - right)
  }
  /** the block's own paragraph ends a section: carry the sectPr in its pPr */
  paragraph(xml: string, endCp: number | undefined): string {
    if (endCp === undefined) return xml
    const index = this.pending.indexOf(endCp)
    if (index < 0) return xml
    this.pending.splice(index, 1)
    return withSectPr(xml, this.breaks.get(endCp)!)
  }
  /** empty sectPr paragraphs for the pending section marks up to `cp` */
  flush(cp: number): string {
    let out = ''
    while (this.pending.length > 0 && this.pending[0]! <= cp) {
      out += withSectPr('<w:p/>', this.breaks.get(this.pending.shift()!)!)
    }
    return out
  }
}

function blockCpRange(block: BodyBlock): { start: number; end: number } | undefined {
  if (block.type === 'paragraph')
    return { start: block.paragraph.startCp, end: block.paragraph.endCp }
  if (block.type === 'image')
    return block.paragraph
      ? { start: block.paragraph.startCp, end: block.paragraph.endCp }
      : { start: block.image.cp, end: block.image.cp + 1 }
  let start = Infinity
  let end = -Infinity
  forEachTableCell(block.table, (cell) => {
    for (const paragraph of cell.paragraphs) {
      start = Math.min(start, paragraph.startCp)
      end = Math.max(end, paragraph.endCp)
    }
  })
  return start <= end ? { start, end } : undefined
}

/**
 * The body text ends with a paragraph mark, so stage 1 always yields one more
 * (empty) paragraph after the real last one; Word has no such paragraph, and
 * writing it adds a blank line that can spill onto an extra page.
 */
function isTrailingEmptyParagraph(
  paragraph: { startCp: number; endCp: number; text: string },
  all: readonly { endCp: number }[],
): boolean {
  if (paragraph.text !== '' || paragraph.startCp !== paragraph.endCp) return false
  return all.length > 1 && all.every((other) => other.endCp <= paragraph.endCp)
}

export function generateBodyBlocksXml(
  blocks: BodyBlock[],
  pictures: PictureRefs = NO_PICTURES,
  context: GenerateContext = GENERATE_CONTEXT,
  sectionBreaks: SectionBreaks = new Map(),
  story: StoryKind = 'body',
): string {
  const ends = new SectionEndWriter(sectionBreaks)
  let out = ''
  const bodyParagraphs = blocks.flatMap((block) =>
    block.type === 'paragraph'
      ? [block.paragraph]
      : block.type === 'image' && block.paragraph
        ? [block.paragraph]
        : [],
  )
  blocks.forEach((block, index) => {
    // the paragraph after a final table stays: OOXML requires one
    if (
      block.type === 'paragraph' &&
      blocks[index - 1]?.type !== 'table' &&
      isTrailingEmptyParagraph(block.paragraph, bodyParagraphs)
    )
      return
    const range = blockCpRange(block)
    if (range) out += ends.flush(range.start - 1)
    if (block.type === 'table') {
      out += generateTableXml(block.table, context, pictures, story)
    } else if (block.type === 'image') {
      const xml = block.paragraph
        ? withParagraphPictures(
            paragraphXml({ paragraph: block.paragraph, list: block.list }, context),
            { inline: [{ image: block.image, position: 'start' }] },
            pictures,
            story,
          )
        : '<w:p>' + inlinePictureRunXml(block.image, pictures) + '</w:p>'
      out += ends.paragraph(xml, range?.end)
    } else {
      const xml = withParagraphPictures(
        paragraphXml(block, context),
        block.pictures,
        pictures,
        story,
      )
      out += ends.paragraph(xml, block.paragraph.endCp)
    }
    if (range) out += ends.flush(range.end)
  })
  return out + ends.flush(Infinity)
}

/** Windows LCID -> BCP-47 for the languages Word 97-2003 documents commonly carry */
const LANGUAGE_TAGS: Record<number, string> = {
  0x0401: 'ar-SA',
  0x0404: 'zh-TW',
  0x0405: 'cs-CZ',
  0x0406: 'da-DK',
  0x0407: 'de-DE',
  0x0408: 'el-GR',
  0x0409: 'en-US',
  0x040a: 'es-ES',
  0x040b: 'fi-FI',
  0x040c: 'fr-FR',
  0x040d: 'he-IL',
  0x040e: 'hu-HU',
  0x0410: 'it-IT',
  0x0411: 'ja-JP',
  0x0412: 'ko-KR',
  0x0413: 'nl-NL',
  0x0414: 'nb-NO',
  0x0415: 'pl-PL',
  0x0416: 'pt-BR',
  0x0419: 'ru-RU',
  0x041d: 'sv-SE',
  0x041e: 'th-TH',
  0x041f: 'tr-TR',
  0x0421: 'id-ID',
  0x042a: 'vi-VN',
  0x0804: 'zh-CN',
  0x0809: 'en-GB',
  0x080a: 'es-MX',
  0x0816: 'pt-PT',
  0x0c04: 'zh-HK',
  0x0c0a: 'es-ES',
  0x0c0c: 'fr-CA',
  0x1004: 'zh-SG',
  0x0c09: 'en-AU',
  0x1009: 'en-CA',
}

/** the document's Latin / East Asian languages: the run languages weighted by text length */
function dominantLanguages(paragraphs: FormattedParagraph[]): {
  latin?: string
  eastAsia?: string
} {
  const latin = new Map<number, number>()
  const eastAsia = new Map<number, number>()
  for (const paragraph of paragraphs)
    for (const run of paragraph.runs) {
      const weight = run.text.length || 1
      const { languageId, eastAsianLanguageId } = run.formatting
      // neutral LCIDs (0x0000 / 0x0400) and unmapped ones must not outvote a real language
      if (languageId !== undefined && LANGUAGE_TAGS[languageId])
        latin.set(languageId, (latin.get(languageId) ?? 0) + weight)
      if (eastAsianLanguageId !== undefined && LANGUAGE_TAGS[eastAsianLanguageId])
        eastAsia.set(eastAsianLanguageId, (eastAsia.get(eastAsianLanguageId) ?? 0) + weight)
    }
  const top = (counts: Map<number, number>): string | undefined => {
    let best: number | undefined
    for (const [lcid, weight] of counts)
      if (best === undefined || weight > counts.get(best)!) best = lcid
    return best === undefined ? undefined : LANGUAGE_TAGS[best]
  }
  const result = { latin: top(latin), eastAsia: top(eastAsia) }
  return {
    ...(result.latin ? { latin: result.latin } : {}),
    ...(result.eastAsia ? { eastAsia: result.eastAsia } : {}),
  }
}

const STYLES_PATH = 'word/styles.xml'

/**
 * The compatibility options Word itself writes when it saves a Word 97-2003
 * document as .docx: the Dop Copts flags (compat.ts), then compatibilityMode
 * 11 plus the Word 2003 layout switches. Without them Word lays the text out
 * with its current rules, which breaks lines differently (e.g. balanced
 * single/double-byte widths).
 */
const LEGACY_COMPAT_HEAD =
  '<w:useNormalStyleForList/>' +
  '<w:doNotUseIndentAsNumberingTabStop/><w:useAltKinsokuLineBreakRules/>' +
  '<w:allowSpaceOfSameStyleInTable/><w:doNotSuppressIndentation/><w:doNotAutofitConstrainedTables/>' +
  '<w:autofitToFirstFixedWidthCell/>'
const LEGACY_COMPAT_TAIL =
  '<w:displayHangulFixedWidth/><w:splitPgBreakAndParaMark/>' +
  '<w:doNotVertAlignCellWithSp/><w:doNotBreakConstrainedForcedTable/><w:doNotVertAlignInTxbx/>' +
  '<w:useAnsiKerningPairs/><w:cachedColBalance/>' +
  '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/>' +
  '</w:compat>'

/** word/fontTable.xml: one w:font per distinct SttbfFfn name with Word's substitution hints */
function fontTableXml(fonts: FontEntry[]): string {
  const seen = new Set<string>()
  const entries = fonts.filter((font) => {
    if (font.name === '' || seen.has(font.name)) return false
    seen.add(font.name)
    return true
  })
  const hex = (value: number, digits: number): string =>
    value.toString(16).toUpperCase().padStart(digits, '0')
  return (
    `${STORY_XML_DECL}<w:fonts xmlns:w="${W_NAMESPACE}">` +
    entries
      .map(
        (font) =>
          `<w:font w:name="${escapeXmlText(font.name)}">` +
          (font.altName ? `<w:altName w:val="${escapeXmlText(font.altName)}"/>` : '') +
          `<w:panose1 w:val="${[...font.panose].map((byte) => hex(byte, 2)).join('')}"/>` +
          `<w:charset w:val="${hex(font.charset, 2)}"/>` +
          `<w:family w:val="${font.family}"/>` +
          `<w:pitch w:val="${font.pitch}"/>` +
          `<w:sig w:usb0="${hex(font.signature[0]!, 8)}" w:usb1="${hex(font.signature[1]!, 8)}" w:usb2="${hex(font.signature[2]!, 8)}" w:usb3="${hex(font.signature[3]!, 8)}" w:csb0="${hex(font.signature[4]!, 8)}" w:csb1="${hex(font.signature[5]!, 8)}"/>` +
          '</w:font>',
      )
      .join('') +
    '</w:fonts>'
  )
}

/** farEast: Word keeps useFELayout only for documents saved by an East Asian Word (FibBase.fFarEast) */
function compatXml(options: string[] = compatOptionsOf(undefined), farEast = false): string {
  const underline = options.includes('underlineTabInNumList')
  const flags = options.filter((name) => name !== 'underlineTabInNumList')
  return (
    '<w:compat>' +
    flags.map((name) => `<w:${name}/>`).join('') +
    (farEast ? '<w:useFELayout/>' : '') +
    LEGACY_COMPAT_HEAD +
    (underline ? '<w:underlineTabInNumList/>' : '') +
    LEGACY_COMPAT_TAIL
  )
}

function settingsXml(
  evenAndOddHeaders: boolean,
  noteKinds: NoteKind[],
  defaultTabStopTwips?: number,
  compatOptions?: string[],
  farEast = false,
): string {
  return (
    `${STORY_XML_DECL}<w:settings xmlns:w="${W_NAMESPACE}">` +
    (defaultTabStopTwips !== undefined
      ? `<w:defaultTabStop w:val="${defaultTabStopTwips}"/>`
      : '') +
    (evenAndOddHeaders ? '<w:evenAndOddHeaders/>' : '') +
    noteKinds
      .map((kind) => `<w:${kind}Pr><w:${kind} w:id="-1"/><w:${kind} w:id="0"/></w:${kind}Pr>`)
      .join('') +
    compatXml(compatOptions, farEast) +
    '</w:settings>'
  )
}

/** The pPr of a style definition: the chain-resolved formatting serialized
    through the same paragraph formatter the body uses (a synthetic empty
    paragraph, keeping element order and attribute spellings identical). */
function stylePPrXml(formatting: ParagraphFormatting): string | undefined {
  const format = toParaFormat(formatting)
  if (!format) return undefined
  const xml = generateParagraphXml({ type: 'paragraph', runs: [], format }, GENERATE_CONTEXT)
  return /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(xml)?.[1]
}

/** The rPr element of a style definition (paragraph styles' CHPX layer, or a
    character style's whole chain): the same run-model merge the paragraph mark
    uses. Returns the complete <w:rPr> element, or undefined when empty. */
function styleRPrXml(formatting: CharacterFormatting | undefined): string | undefined {
  if (!formatting) return undefined
  const rPr = mergeRPrModel('', toDocxRun({ text: '', formatting }), false)
  return rPr === '' ? undefined : rPr
}

/** w:style definitions for the legacy stylesheet entries. Ids are w97-<istd>
    (styles.ts legacyStyleId), names are the source document's own, and
    basedOn links only point at bases that are emitted too (Normal is the
    template's own and the null character style is never a base reference). */
function styleDefsXml(defs: LegacyStyleDef[]): string {
  const emitted = new Set(defs.map((def) => def.istd))
  return defs
    .map((def) => {
      const parts = [`<w:style w:type="${def.kind}" w:styleId="${legacyStyleId(def.istd)}">`]
      parts.push(`<w:name w:val="${escapeXmlText(def.name)}"/>`)
      if (def.basedOn !== undefined && emitted.has(def.basedOn))
        parts.push(`<w:basedOn w:val="${legacyStyleId(def.basedOn)}"/>`)
      if (def.kind === 'paragraph') {
        const pPr = def.paragraph ? stylePPrXml(def.paragraph) : undefined
        if (pPr) parts.push(`<w:pPr>${pPr}</w:pPr>`)
      }
      const rPr = styleRPrXml(def.character)
      if (rPr) parts.push(rPr)
      parts.push('</w:style>')
      return parts.join('')
    })
    .join('')
}

/**
 * The blank template's docDefaults are docx-engine's authoring defaults; a
 * converted .doc gets Word 97's (10pt, no space after, single spacing) plus
 * the document languages, which Word needs for East Asian line breaking.
 */
function convertedStylesXml(
  styles: string,
  languages: { latin?: string; eastAsia?: string },
  noteKinds: NoteKind[] = [],
  styleDefs: LegacyStyleDef[] = [],
): string {
  let out = styles
    .replace(
      /<w:pPrDefault><w:pPr><w:spacing [^>]*\/><\/w:pPr><\/w:pPrDefault>/,
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>',
    )
    .replace(
      /(<w:rPrDefault><w:rPr>(?:(?!<\/w:rPr>).)*?)<w:sz w:val="\d+"\/><w:szCs w:val="\d+"\/>/,
      '$1<w:sz w:val="20"/><w:szCs w:val="20"/>',
    )
  if (languages.latin || languages.eastAsia) {
    const lang =
      '<w:lang' +
      (languages.latin ? ` w:val="${languages.latin}"` : '') +
      (languages.eastAsia ? ` w:eastAsia="${languages.eastAsia}"` : '') +
      '/>'
    out = out.replace('</w:rPr></w:rPrDefault>', lang + '</w:rPr></w:rPrDefault>')
  }
  const extra =
    noteKinds.map(noteStylesXml).join('') + (styleDefs.length > 0 ? styleDefsXml(styleDefs) : '')
  if (extra !== '') out = out.replace('</w:styles>', extra + '</w:styles>')
  return out
}

const BODY_OPEN = '<w:body>'
const BODY_PLACEHOLDER = '<w:p/>'
const DOCUMENT_PATH = 'word/document.xml'

/** Generated <w:p> XML for every model paragraph. */
export function generateBodyParagraphsXml(
  paragraphs: FormattedParagraph[],
  sectionBreaks: SectionBreaks = new Map(),
): string {
  const ends = new SectionEndWriter(sectionBreaks)
  let out = ''
  for (const paragraph of paragraphs) {
    if (isTrailingEmptyParagraph(paragraph, paragraphs)) continue
    out += ends.flush(paragraph.startCp - 1)
    out += ends.paragraph(paragraphXml({ paragraph }, GENERATE_CONTEXT), paragraph.endCp)
  }
  return out + ends.flush(Infinity)
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

/**
 * The level number's run properties (LVL grpprlChpx): font, bold/italic, size
 * and colour. Fonts resolve through the table without the stylesheet defaults,
 * so a level that names no font inherits the paragraph's.
 */
function numberingRunPropertiesXml(
  level: ListLevel,
  fonts: FontEntry[] | undefined,
): { xml: string; symbolFont: boolean } {
  const formatting = characterFormattingOf(level.characterSprms)
  const names = fonts
    ? fontNamesFor(
        { names: fonts.map((entry) => entry.name), defaults: {} },
        {
          ascii: formatting.ftcAscii,
          hAnsi: formatting.ftcHAnsi,
          eastAsia: formatting.ftcEastAsia,
          cs: formatting.ftcBi,
        },
      )
    : undefined
  const font = names?.ascii ?? names?.eastAsia
  const symbolFont =
    font !== undefined && fonts?.find((entry) => entry.name === font)?.charset === 2
  let xml = ''
  if (font !== undefined)
    xml +=
      `<w:rFonts w:ascii="${escapeXmlText(font)}"` +
      (names?.eastAsia ? ` w:eastAsia="${escapeXmlText(names.eastAsia)}"` : '') +
      ` w:hAnsi="${escapeXmlText(font)}"` +
      (names?.cs ? ` w:cs="${escapeXmlText(names.cs)}"` : '') +
      ' w:hint="default"/>'
  if (formatting.bold) xml += '<w:b/>'
  if (formatting.italic) xml += '<w:i/>'
  if (formatting.colorRgb !== undefined) xml += `<w:color w:val="${formatting.colorRgb}"/>`
  else if (formatting.colorIndex !== undefined && COLOR_BY_INDEX[formatting.colorIndex])
    xml += `<w:color w:val="${COLOR_BY_INDEX[formatting.colorIndex]}"/>`
  if (formatting.fontSizeHalfPoints !== undefined)
    xml += `<w:sz w:val="${formatting.fontSizeHalfPoints}"/><w:szCs w:val="${formatting.fontSizeHalfPoints}"/>`
  return { xml: xml === '' ? '' : '<w:rPr>' + xml + '</w:rPr>', symbolFont }
}

/** one w:lvl of a reconstructed list definition */
function numberingLevelXml(level: ListLevel, ilvl: number, fonts?: FontEntry[]): string {
  const bullet = isBulletLevel(level)
  const numberFormat = bullet ? 'bullet' : (NUMBER_FORMATS[level.numberFormat] ?? 'decimal')
  const runProperties = numberingRunPropertiesXml(level, fonts)
  // a symbol-font bullet keeps its private-use code so the font draws the glyph Word shows
  const symbolCode = level.levelText.charCodeAt(0)
  const levelText =
    bullet && runProperties.symbolFont && symbolCode >= 0xf020 && symbolCode <= 0xf0ff
      ? `&#x${symbolCode.toString(16).toUpperCase()};`
      : escapeXmlText(levelTextForWriter(level))
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
    runProperties.xml +
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
  tables: ListTables,
  fonts?: FontEntry[],
): string {
  const abstractIdByLsid = new Map<number, number>()
  const abstractXml: string[] = []
  const numXml: string[] = []
  for (const entry of used) {
    const lsid = tables.overrides.get(entry.ilfo)
    const definition = lsid !== undefined ? tables.lists.get(lsid) : undefined
    if (!definition) continue
    let abstractId = abstractIdByLsid.get(definition.lsid)
    if (abstractId === undefined) {
      abstractId = 10 + abstractIdByLsid.size
      abstractIdByLsid.set(definition.lsid, abstractId)
      abstractXml.push(
        '<w:abstractNum w:abstractNumId="' +
          abstractId +
          '"><w:multiLevelType w:val="hybridMultilevel"/>' +
          definition.levels.map((level, ilvl) => numberingLevelXml(level, ilvl, fonts)).join('') +
          '</w:abstractNum>',
      )
    }
    const levelOverrides = (tables.levelOverrides.get(entry.ilfo) ?? [])
      .map(
        (override) =>
          '<w:lvlOverride w:ilvl="' +
          override.ilvl +
          '">' +
          (override.startAt !== undefined
            ? '<w:startOverride w:val="' + override.startAt + '"/>'
            : '') +
          (override.level ? numberingLevelXml(override.level, override.ilvl, fonts) : '') +
          '</w:lvlOverride>',
      )
      .join('')
    numXml.push(
      '<w:num w:numId="' +
        entry.numId +
        '"><w:abstractNumId w:val="' +
        abstractId +
        '"/>' +
        levelOverrides +
        '</w:num>',
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
 * generated paragraphs, replace the template's trailing sectPr with the last
 * section's geometry (earlier sections end inside their last paragraph) and
 * add the header/footer parts. Documents
 * without a section table keep the template sectPr unchanged.
 */
export async function writeLegacyDocx(
  model: LegacyDocModel,
  layout?: LegacyDocLayout,
  body?: LegacyDocBody,
  styleDefs: LegacyStyleDef[] = [],
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
  let evenAndOddHeaders = false
  let nextRelId = 3
  let generatedSectPr: string | undefined
  const sectionBreaks: SectionBreaks = new Map()
  if (layout && layout.sections.length > 0) {
    const counts = { header: 0, footer: 0 }
    layout.sections.forEach((section, index) => {
      const collected = collectHeaderFooterParts(layout.headerFooter[index] ?? {}, section, counts)
      evenAndOddHeaders ||= collected.settingsNeeded
      for (const part of collected.parts) part.rId = `rId${nextRelId++}`
      parts.push(...collected.parts)
      const references = collected.parts
        .map((part) => {
          const element = part.kind === 'header' ? 'headerReference' : 'footerReference'
          return `<w:${element} w:type="${part.referenceType}" r:id="${part.rId}"/>`
        })
        .join('')
      const xml = sectPrXml(section, references)
      if (index === layout.sections.length - 1) generatedSectPr = xml
      else sectionBreaks.set(section.endCp - 1, xml)
    })
  }

  // one media part per distinct picture; the body and every header part relate to it with their own rId
  interface MediaPart {
    path: string
    extension: string
    contentType: string
    bytes: Uint8Array
  }
  const mediaParts: MediaPart[] = []
  const mediaByBytes = new Map<Uint8Array, MediaPart>()
  const mediaFor = (media: { mediaType: string; bytes: Uint8Array }): MediaPart => {
    let part = mediaByBytes.get(media.bytes)
    if (!part) {
      const extension = PICTURE_EXTENSIONS[media.mediaType as PictureMediaType] ?? 'bin'
      part = {
        path: `word/media/image${mediaParts.length + 1}.${extension}`,
        extension,
        contentType: media.mediaType,
        bytes: media.bytes,
      }
      mediaParts.push(part)
      mediaByBytes.set(media.bytes, part)
    }
    return part
  }
  let docPrCounter = 0
  const documentImageRels = new Map<MediaPart, string>()
  const bodyPictures: PictureRefs = {
    relIdFor: (media) => {
      const part = mediaFor(media as { mediaType: string; bytes: Uint8Array })
      let rId = documentImageRels.get(part)
      if (!rId) {
        rId = `rId${nextRelId++}`
        documentImageRels.set(part, rId)
      }
      return rId
    },
    nextDocPrId: () => ++docPrCounter,
  }
  const partPictures = (rels: Map<MediaPart, string>): PictureRefs => ({
    relIdFor: (media) => {
      const part = mediaFor(media as { mediaType: string; bytes: Uint8Array })
      let rId = rels.get(part)
      if (!rId) {
        rId = `rId${rels.size + 1}`
        rels.set(part, rId)
      }
      return rId
    },
    nextDocPrId: () => ++docPrCounter,
  })

  // numbered paragraphs: one w:num per used ilfo, definitions patched into numbering.xml
  const usedNumbering: Array<{ numId: number; ilfo: number }> = []
  if (body) {
    const seenNumIds = new Set<number>()
    const use = (list: ParagraphListInfo): void => {
      const numId = Number(list.numId)
      if (!Number.isFinite(numId) || seenNumIds.has(numId)) return
      seenNumIds.add(numId)
      usedNumbering.push({ numId, ilfo: numId - 100 })
    }
    for (const block of body.blocks) {
      if ((block.type === 'paragraph' || block.type === 'image') && block.list) use(block.list)
      if (block.type === 'table')
        forEachTableCell(block.table, (cell) => cell.paragraphLists?.forEach(use))
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
          (block.list !== undefined ||
            block.pictures !== undefined ||
            block.paragraph.runs.some(
              (run) =>
                run.link !== undefined ||
                run.field !== undefined ||
                run.note !== undefined ||
                run.commentMarks !== undefined,
            ))),
    )
  const comments = body?.comments ?? []
  const commentsRelId = comments.length > 0 ? `rId${nextRelId++}` : undefined
  const bodyXml = usesBodyBlocks
    ? generateBodyBlocksXml(body!.blocks, bodyPictures, generateContext, sectionBreaks)
    : generateBodyParagraphsXml(model.paragraphs, sectionBreaks)
  // header/footer parts are generated now so their media relationships exist before the rels are written
  const partRels = new Map<HeaderFooterPart, Map<MediaPart, string>>()
  const partXml = new Map<HeaderFooterPart, string>()
  for (const part of parts) {
    const rels = new Map<MediaPart, string>()
    partRels.set(part, rels)
    partXml.set(part, storyPartXml(part.kind, part.story, partPictures(rels)))
  }
  let documentXml =
    template.slice(0, placeholderStart) +
    bodyXml +
    template.slice(placeholderStart + BODY_PLACEHOLDER.length)
  if (bodyXml.includes('<wps:') && !documentXml.includes('xmlns:wps='))
    documentXml = documentXml.replace('<w:document ', `<w:document xmlns:wps="${WPS_NAMESPACE}" `)
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
  const noteParts: Array<{ kind: NoteKind; notes: LegacyNote[]; rId: string }> = []
  if (body?.notes) {
    const { footnotes, endnotes } = body.notes
    if (footnotes.length > 0) noteParts.push({ kind: 'footnote', notes: footnotes, rId: '' })
    if (endnotes.length > 0) noteParts.push({ kind: 'endnote', notes: endnotes, rId: '' })
    for (const part of noteParts) part.rId = `rId${nextRelId++}`
  }
  const noteKinds = noteParts.map((part) => part.kind)
  const stylesFile = zip.file(STYLES_PATH)
  if (stylesFile)
    zip.file(
      STYLES_PATH,
      convertedStylesXml(
        await stylesFile.async('string'),
        dominantLanguages(model.paragraphs),
        noteKinds,
        styleDefs,
      ),
    )

  const fontTable =
    layout?.fonts && layout.fonts.length > 0 ? fontTableXml(layout.fonts) : undefined
  {
    const contentTypesFile = zip.file(CONTENT_TYPES_PATH)
    if (!contentTypesFile) throw new Error('the blank docx template has no [Content_Types].xml')
    let contentTypes = await contentTypesFile.async('string')
    const extensions = [...new Set(mediaParts.map((part) => part.extension))]
    const defaults = extensions
      .filter((extension) => !contentTypes.includes(`Extension="${extension}"`))
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
      noteParts
        .map(
          (part) =>
            `<Override PartName="/${NOTES_PATH[part.kind]}" ContentType="${NOTES_CONTENT_TYPE[part.kind]}"/>`,
        )
        .join('') +
      `<Override PartName="/${SETTINGS_PATH}" ContentType="${SETTINGS_CONTENT_TYPE}"/>` +
      (fontTable
        ? `<Override PartName="/${FONT_TABLE_PATH}" ContentType="${FONT_TABLE_CONTENT_TYPE}"/>`
        : '') +
      (commentsRelId
        ? `<Override PartName="/${COMMENTS_PATH}" ContentType="${COMMENTS_CONTENT_TYPE}"/>`
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
      [...documentImageRels]
        .map(
          ([part, rId]) =>
            `<Relationship Id="${rId}" Type="${IMAGE_REL_TYPE}" Target="${part.path.slice(
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
      noteParts
        .map(
          (part) =>
            `<Relationship Id="${part.rId}" Type="${NOTES_REL_TYPE[part.kind]}" Target="${NOTES_PATH[
              part.kind
            ].slice('word/'.length)}"/>`,
        )
        .join('') +
      (commentsRelId
        ? `<Relationship Id="${commentsRelId}" Type="${COMMENTS_REL_TYPE}" Target="comments.xml"/>`
        : '') +
      `<Relationship Id="rId${nextRelId++}" Type="${SETTINGS_REL_TYPE}" Target="settings.xml"/>` +
      (fontTable
        ? `<Relationship Id="rId${nextRelId}" Type="${FONT_TABLE_REL_TYPE}" Target="fontTable.xml"/>`
        : '')
    relationships = relationships.replace('</Relationships>', `${newRelationships}</Relationships>`)
    zip.file(DOCUMENT_RELS_PATH, relationships)
    for (const part of parts) {
      zip.file(part.path, partXml.get(part)!)
      const rels = partRels.get(part)!
      if (rels.size === 0) continue
      const name = part.path.slice('word/'.length)
      zip.file(
        `word/_rels/${name}.rels`,
        `${STORY_XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          [...rels]
            .map(
              ([media, rId]) =>
                `<Relationship Id="${rId}" Type="${IMAGE_REL_TYPE}" Target="${media.path.slice(
                  'word/'.length,
                )}"/>`,
            )
            .join('') +
          '</Relationships>',
      )
    }
    for (const part of mediaParts) zip.file(part.path, part.bytes)
    for (const part of noteParts)
      zip.file(NOTES_PATH[part.kind], notesPartXml(part.kind, part.notes))
    if (commentsRelId) zip.file(COMMENTS_PATH, commentsPartXml(comments))
    zip.file(
      SETTINGS_PATH,
      settingsXml(
        evenAndOddHeaders,
        noteKinds,
        layout?.defaultTabStopTwips,
        layout?.compatOptions,
        model.fib.farEast,
      ),
    )
    if (fontTable) zip.file(FONT_TABLE_PATH, fontTable)
    if (usedNumbering.length > 0 && body?.lists) {
      const numberingFile = zip.file(NUMBERING_PATH)
      if (numberingFile) {
        const numbering = await numberingFile.async('string')
        const patched = reconstructNumberingXml(numbering, usedNumbering, body.lists, layout?.fonts)
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
  const layout = options.layout
  for (const warning of layout?.warnings ?? [])
    if (!warnings.includes(warning)) warnings.push(warning)
  const headersConverted = layout !== undefined && layout.headerFooter.length > 0
  if (fib.ccpHdd > 0 && !headersConverted)
    warnings.push('headers are not converted yet; their text is omitted')
  return warnings
}

/**
 * Read a legacy .doc and convert it into a .docx. Documents with a stylesheet
 * are read through the style resolver (stage 3b): style-based formatting is
 * written as direct formatting, and every defined stylesheet entry is emitted
 * as a real w:style the paragraphs and runs reference through w:pStyle /
 * w:rStyle (the direct formatting carries the same properties the style
 * contributes, so the rendering is unchanged while the styles gallery and
 * style edits work on the converted file). Every section's geometry and
 * header/footer stories are written; the sectPr of all but the last section
 * sits in its last paragraph.
 */
export async function convertLegacyDocToDocx(
  bytes: Uint8Array,
): Promise<{ docx: Uint8Array; warnings: string[] }> {
  if (isZipArchive(bytes))
    return { docx: bytes, warnings: ['the file is already Office Open XML; it was kept as is'] }
  // dynamic import keeps the module graph acyclic: index.ts re-exports this file
  const { readLegacyDoc } = await import('./index')
  const { model, layout, body, stylesResolved, styleDefs } = readLegacyDoc(bytes)
  const tablesHandled = body.blocks.some((block) => block.type === 'table')
  const docx = await writeLegacyDocx(model, layout, body, styleDefs)
  const warnings = conversionWarnings(model, { stylesResolved, layout, tablesHandled })
  for (const warning of body.warnings) if (!warnings.includes(warning)) warnings.push(warning)
  return { docx, warnings }
}

function isZipArchive(bytes: Uint8Array): boolean {
  return (
    bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 3 && bytes[3] === 4
  )
}
