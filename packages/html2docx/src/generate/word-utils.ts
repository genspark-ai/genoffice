// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import {
  AlignmentType,
  BorderStyle,
  CheckBox,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  InternalHyperlink,
  PageNumber,
  LineRuleType,
  Paragraph,
  ShadingType,
  Tab,
  TabStopType,
  TextRun,
  UnderlineType,
} from 'docx'
import { bookmarkName } from './bookmarks'
import { mapFont } from './fonts'

function makeRuns(context, runs, images: any = {}) {
  const out = []
  for (const r of runs) {
    if (r.pageNumberField) {
      out.push(
        new TextRun({
          children: [r.pageNumberField === 'total' ? PageNumber.TOTAL_PAGES : PageNumber.CURRENT],
          bold: r.bold || undefined,
          italics: r.italic || undefined,
          color: r.color || undefined,
          size: r.sizePx ? context.pxToHalfPoints(r.sizePx) : undefined,
          font: mapFont(r.fontFamily, r.mono, r.text),
        }),
      )
      continue
    }
    if (r.contentControl === 'checkbox') {
      out.push(new CheckBox({ checked: Boolean(r.checked), alias: r.alias || 'Checkbox' }))
      continue
    }
    if (r.inlineImage) {
      const image = images[r.shotId]
      if (image) {
        out.push(
          new ImageRun({
            type: 'png',
            data: image,
            transformation: { width: r.width, height: r.height },
          }),
        )
      }
      continue
    }
    const underlineType =
      r.underlineStyle === 'dotted'
        ? UnderlineType.DOTTED
        : r.underlineStyle === 'dashed'
          ? (UnderlineType as any).DASHED
          : UnderlineType.SINGLE
    const style = {
      bold: r.bold || undefined,
      italics: r.italic || undefined,
      underline: r.underline
        ? { type: underlineType, color: r.underlineColor || undefined }
        : undefined,
      strike: r.strike || undefined,
      // always explicit: inherited styles (Heading1..6 theme blue) must not win
      color: r.color || undefined,
      size: r.sizePx ? context.pxToHalfPoints(r.sizePx) : undefined,
      // Emoji: leave the font unset so Word's built-in emoji fallback picks the
      // platform color-emoji font. Forcing 'Segoe UI Emoji' renders tofu boxes
      // (and flags as "GB") on Mac Word, which lacks that font.
      font: r.emoji
        ? undefined
        : r.symbol
          ? 'Arial Unicode MS'
          : mapFont(r.fontFamily, r.mono, r.text),
      characterSpacing: r.letterSpacing > 0.5 ? context.pxToTwips(r.letterSpacing) : undefined,
      superScript: r.superscript || undefined,
      subScript: r.subscript || undefined,
      allCaps: r.uppercase || undefined,
      // w:rtl on the run so Arabic/Hebrew glyphs shape and order correctly
      // (docx mirrors size/bold/italics into their complex-script twins)
      rightToLeft: r.rtl || undefined,
      // arbitrary hex needs character shading (w:highlight only takes named colors)
      shading: r.highlight ? { type: ShadingType.CLEAR, fill: r.highlight } : undefined,
      // inline boxed chip -> character border (w:bdr); space:1 keeps a hair of
      // padding between glyph and frame like the authored chip padding
      border: r.charBorder
        ? {
            style: BorderStyle.SINGLE,
            size: context.pxToBorderEighths(r.charBorder.widthPx),
            color: r.charBorder.color,
            space: 1,
          }
        : undefined,
      noProof: true,
    }
    const lines = (r.text || '').split('\n')
    lines.forEach((line, i) => {
      const segments = line.split('\t')
      segments.forEach((segment, j) => {
        const needsBreak = i > 0 && j === 0
        const children = []
        if (j > 0) children.push(new Tab())
        if (segment) children.push(segment)
        // standalone <br> produces an empty segment that must still emit the break
        if (!children.length && !needsBreak) return
        const run = new TextRun({ ...style, children, break: needsBreak ? 1 : undefined })
        if (r.anchor) {
          out.push(new InternalHyperlink({ children: [run], anchor: bookmarkName(r.anchor) }))
        } else if (r.href) {
          out.push(new ExternalHyperlink({ children: [run], link: r.href }))
        } else {
          out.push(run)
        }
      })
    })
  }
  return out
}

function makeNodeRuns(context, node, images: any = {}) {
  const runs = makeRuns(context, node.runs, images)
  const insetPx = node.style?.borderLeftSpacePx || 0
  if (insetPx > 0) {
    const sample = node.runs.find((run) => run.text) || {}
    const fontSizePx = sample.sizePx || 16
    const count = Math.max(1, Math.ceil(insetPx / (fontSizePx * 0.65)))
    runs.unshift(
      new TextRun({
        text: '\u00A0'.repeat(count),
        size: context.pxToHalfPoints(fontSizePx),
        font: mapFont(sample.fontFamily, sample.mono, sample.text),
        noProof: true,
      }),
    )
  }
  return runs
}

// Rebuild tab stops from measured positions: a separator whose following
// content starts far right becomes a right-aligned stop at the margin
// (dates), otherwise a left stop at the measured column position.
function tabStopsFor(context, runs) {
  const tabs = runs.filter((r) => r.text.includes('\t'))
  if (!tabs.length) return []
  const stops = []
  const seen = new Set()
  for (const r of tabs) {
    // Column starting past 75% of the line is a right-aligned tail (dates);
    // otherwise reproduce the measured column position with a left stop.
    const stop =
      r.tabFrac != null && r.tabFrac < 0.75
        ? { type: TabStopType.LEFT, position: Math.round(r.tabFrac * context.contentDxa) }
        : { type: TabStopType.RIGHT, position: context.contentDxa }
    const key = `${stop.type}:${stop.position}`
    if (!seen.has(key)) {
      seen.add(key)
      stops.push(stop)
    }
  }
  return stops
}

function paraOptions(context, style: any = {}) {
  const opts: any = {}
  // w:bidi: paragraph flows right-to-left (default alignment becomes right)
  if (style.rtl) opts.bidirectional = true
  if (style.align === 'center') opts.alignment = AlignmentType.CENTER
  if (style.align === 'right') opts.alignment = AlignmentType.RIGHT
  // explicit physical left (an RTL paragraph author-aligned to the left edge)
  if (style.align === 'left' && style.rtl) opts.alignment = AlignmentType.LEFT
  if (style.align === 'justify') opts.alignment = AlignmentType.JUSTIFIED
  if (style.indentLeftPx || style.indentHangingPx) {
    // hanging: first line carries the literal bullet/number prefix at
    // indentLeftPx, wrapped lines align under the text after the prefix
    const hang = style.indentHangingPx ? context.pxToTwips(style.indentHangingPx) : 0
    opts.indent = {
      left: context.pxToTwips(style.indentLeftPx || 0) + hang,
      ...(hang ? { hanging: hang } : {}),
    }
  } else if (style.firstLineIndentPx) {
    // CSS text-indent (CJK-style opening indent) -> w:ind firstLine
    opts.indent = { firstLine: context.pxToTwips(style.firstLineIndentPx) }
  }
  opts.spacing = {}
  if (style.spacingAfterPx) opts.spacing.after = context.pxToTwips(style.spacingAfterPx)
  if (style.spacingBeforePx) opts.spacing.before = context.pxToTwips(style.spacingBeforePx)
  if (style.lineRatio) {
    // Word's intrinsic leading is larger for CJK fonts than Latin fonts.
    const effective = Math.max(1, style.lineRatio / (style.cjk ? 1.3 : 1.15))
    opts.spacing.line = Math.round(effective * 240)
  }
  if (style.exactLineHeightPx) {
    // Shaded single-line bars: force the browser-rendered height, otherwise
    // Word's CJK leading makes the colored strip visibly taller.
    // Bordered paragraphs (left accent bars) must not use EXACT: Word drops
    // the text to the line bottom while the border hugs the paragraph box,
    // so the accent bar floats visibly above the heading text.
    opts.spacing.line = context.pxToTwips(style.exactLineHeightPx)
    opts.spacing.lineRule =
      style.borderLeft || style.borderRight ? LineRuleType.AT_LEAST : LineRuleType.EXACT
  }
  if (style.shading) opts.shading = { type: ShadingType.CLEAR, fill: style.shading }
  if (style.borderLeft || style.borderRight || style.borderTop || style.borderBottom) {
    opts.border = {}
    if (style.borderLeft) {
      opts.border.left = {
        style: BorderStyle.SINGLE,
        size: context.pxToBorderEighths(style.borderLeft.widthPx),
        color: style.borderLeft.color,
        // w:space pushes the border left, away from the text (in points,
        // capped at 31 by Word)
        space: style.borderLeftOutsetPx
          ? Math.min(
              31,
              Math.max(1, Math.round(style.borderLeftOutsetPx * 0.75 * context.measurementScale)),
            )
          : undefined,
      }
    }
    if (style.borderBottom) {
      opts.border.bottom = {
        style:
          style.borderBottom.style === 'dotted'
            ? BorderStyle.DOTTED
            : style.borderBottom.style === 'dashed'
              ? BorderStyle.DASHED
              : BorderStyle.SINGLE,
        size: context.pxToBorderEighths(style.borderBottom.widthPx),
        color: style.borderBottom.color,
        space: 4,
      }
    }
    if (style.borderRight) {
      opts.border.right = {
        style: BorderStyle.SINGLE,
        size: context.pxToBorderEighths(style.borderRight.widthPx),
        color: style.borderRight.color,
        space: style.borderRightSpacePx
          ? Math.max(1, Math.round(style.borderRightSpacePx * 0.75 * context.measurementScale))
          : undefined,
      }
    }
    if (style.borderTop) {
      opts.border.top = {
        style: BorderStyle.SINGLE,
        size: context.pxToBorderEighths(style.borderTop.widthPx),
        color: style.borderTop.color,
        space: 4,
      }
    }
  }
  return opts
}

// Empty paragraph occupying an exact pixel height (spacing around tables,
// which can't carry before/after spacing themselves).
// keepNext binds the spacer to the following block: needed between a heading
// and its section body (else the heading's own keepNext only reaches the
// spacer and Word orphans the heading at a page bottom), but must NOT be on
// trailing spacers \u2014 a document-wide spacer chain would fuse sections and
// force Word to break pages far too early.
function spacerParagraph(
  context,
  px,
  { keepNext = false, shading = null, borderLeft = null } = {},
) {
  const paragraph = new Paragraph({
    // Word tables already contribute a few pixels of leading. Force only
    // authored block gaps large enough to be visually meaningful.
    children: [new TextRun({ text: px >= 24 ? '\u200B' : '', size: 2, noProof: true })],
    ...(keepNext ? { keepNext: true } : {}),
    ...(shading ? { shading: { type: ShadingType.CLEAR, fill: shading } } : {}),
    // continues a flattened card's left rail through spacing gaps
    ...(borderLeft
      ? {
          border: {
            left: {
              style: BorderStyle.SINGLE,
              size: context.pxToBorderEighths(borderLeft.widthPx),
              color: borderLeft.color,
            },
          },
        }
      : {}),
    spacing: {
      before: 0,
      after: 0,
      line: context.pxToTwips(Math.max(px, 2)),
      lineRule: LineRuleType.EXACT,
    },
  })
  // Marker so Generator.render can swap a rendered block's leading spacer
  // for a keepNext one when it directly follows a heading.
  const marked = paragraph as any
  marked.__h2dSpacerPx = px
  marked.__h2dSpacerShading = shading
  marked.__h2dSpacerBorderLeft = borderLeft
  return paragraph
}

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
const NO_BORDERS = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
}
// Table-level: without it the docx library emits default single borders,
// and Word (Mac) paints them even where cells say val="none" — visible as
// a gray frame around dark panels.
const NO_TABLE_BORDERS = {
  ...NO_BORDERS,
  insideHorizontal: NO_BORDER,
  insideVertical: NO_BORDER,
}

function wordBorder(context, border) {
  if (!border) return NO_BORDER
  const styles = {
    dotted: BorderStyle.DOTTED,
    dashed: BorderStyle.DASHED,
    double: BorderStyle.DOUBLE,
  }
  return {
    style: styles[border.style] || BorderStyle.SINGLE,
    // Floor at 0.75pt: a scaled 1px CSS border falls to 0.5pt and turns
    // nearly invisible in Word at common zoom levels (same floor as cards).
    size: Math.max(6, context.pxToBorderEighths(border.widthPx)),
    color: border.color,
  }
}

const HEADING_LEVELS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
]

export {
  HEADING_LEVELS,
  NO_BORDER,
  NO_BORDERS,
  NO_TABLE_BORDERS,
  makeNodeRuns,
  makeRuns,
  paraOptions,
  spacerParagraph,
  tabStopsFor,
  wordBorder,
}
