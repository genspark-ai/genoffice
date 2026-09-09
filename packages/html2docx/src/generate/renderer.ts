// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import {
  AlignmentType,
  BorderStyle,
  ExternalHyperlink,
  HeightRule,
  ImageRun,
  LineRuleType,
  PageBreak,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  TextWrappingType,
  VerticalPositionAlign,
  WidthType,
} from 'docx'
import { withBookmarks } from './bookmarks'
import { createContentControlFactory } from './content-controls'
import { orderedReference } from './numbering'
import { renderCard, renderColorBar, renderKpiRow, renderTable } from './table-renderers'
import {
  HEADING_LEVELS,
  NO_BORDERS,
  NO_TABLE_BORDERS,
  makeNodeRuns,
  paraOptions,
  spacerParagraph,
  tabStopsFor,
  wordBorder,
} from './word-utils'

function splitVerticalBorderSpacing(style: any = {}) {
  if (!style.borderLeft && !style.borderRight) {
    return { paragraphStyle: style, spacingAfterPx: 0, spacingBeforePx: 0 }
  }
  const spacingBeforePx = style.spacingBeforePx || 0
  const spacingAfterPx = style.spacingAfterPx || 0
  return {
    paragraphStyle: {
      ...style,
      spacingBeforePx: 0,
      spacingAfterPx: 0,
    },
    spacingAfterPx,
    spacingBeforePx,
  }
}

function withExternalBorderSpacing(
  context,
  paragraph,
  spacingBeforePx,
  spacingAfterPx,
  shading = null,
) {
  const output = []
  if (spacingBeforePx) output.push(spacerParagraph(context, spacingBeforePx, { shading }))
  output.push(paragraph)
  if (spacingAfterPx) output.push(spacerParagraph(context, spacingAfterPx, { shading }))
  return output
}

class Generator {
  [key: string]: any
  constructor(images, context) {
    this.images = images // shotId -> { buffer, width, height }
    this.context = context
    this.contentControl = createContentControlFactory()
    // width (dxa) available to the node being rendered; shrinks inside table cells
    this.avail = context.contentDxa
    // page-anchored floating images, attached to the first paragraph
    this.floats = []
    // Repeated from the header so printed pages retain the authored backdrop.
    this.pageBackgroundFloat = null
    this.olInstance = 0
    // Word promotes keepNext on in-cell paragraphs to row-level
    // keep-with-next: rows chain into one unsplittable block that jumps to a
    // fresh page whole. Set while rendering table-cell content.
    this.suppressKeepNext = false
    // Inverse of the above, set by cellChildren({keepWithNext}) for the
    // non-last rows of a fits-on-one-page table: every paragraph rendered
    // while true carries keepNext so the whole table stays together.
    this.forceKeepNext = false
  }

  render(nodes, depth = 0) {
    const out = []
    // Word binds keepNext paragraph-to-paragraph only: a heading followed by
    // a plain spacer is "kept" with just the spacer, and the section body
    // still breaks to the next page leaving the heading orphaned. Extend the
    // keep chain from a heading through its section head — spacers and at
    // most one lede/subtitle paragraph — into the first substantive block.
    // Trailing spacers must stay unbound or sections fuse into one
    // document-wide keep chain and Word breaks pages far too early.
    let chainActive = false
    let ledeAllowance = 0
    for (let index = 0; index < nodes.length; index++) {
      const node = nodes[index]
      let rendered
      if (
        !this.suppressKeepNext &&
        node.type === 'para' &&
        nodes[index + 1]?.type === 'spacer' &&
        nodes[index + 1].bindsNeighbors
      ) {
        // Paragraph before an authored in-flow gap (signature space): bind
        // it through the gap to the following paragraph.
        rendered = this.renderNode({ ...node, style: { ...node.style, keepNext: true } }, depth)
        chainActive = false
      } else if (
        node.type === 'spacer' &&
        nodes[index + 1]?.type === 'card' &&
        (nodes[index + 1].heightPx || 0) > 700
      ) {
        // A near-page-height unsplittable card cannot share its page with
        // the authored gap before it — the card slides to the next page and
        // the invisible gap strands a blank page. Collapse the gap; the
        // page break itself provides the separation.
        rendered = [
          spacerParagraph(this.context, Math.min(node.px, 8), {
            keepNext: chainActive,
            shading: node.shading,
          }),
        ]
      } else if (
        (chainActive || (node.bindsNeighbors && !this.suppressKeepNext)) &&
        node.type === 'spacer'
      ) {
        rendered = [
          spacerParagraph(this.context, node.px, { keepNext: true, shading: node.shading }),
        ]
      } else if (chainActive && node.type === 'para' && ledeAllowance > 0) {
        // Only a real lede continues the chain: the next substantive node
        // must be more section content. When it is another heading (or the
        // end), this para IS the section body — it must terminate the chain,
        // or trailing spacers bind section after section into one
        // document-wide keep block that Word shoves onto a fresh page.
        let next = index + 1
        while (next < nodes.length && nodes[next].type === 'spacer') next += 1
        if (next < nodes.length && nodes[next].type !== 'heading') {
          ledeAllowance -= 1
          rendered = this.renderNode({ ...node, style: { ...node.style, keepNext: true } }, depth)
        } else {
          rendered = this.renderNode(node, depth)
          chainActive = false
        }
      } else {
        rendered = this.renderNode(node, depth)
        if (node.type === 'heading' && !this.suppressKeepNext) {
          // A heading's trailing spacing also sits between it and the body.
          const last = rendered.length - 1
          if (rendered[last]?.__h2dSpacerPx) {
            rendered[last] = spacerParagraph(this.context, rendered[last].__h2dSpacerPx, {
              keepNext: true,
              shading: rendered[last].__h2dSpacerShading,
              borderLeft: rendered[last].__h2dSpacerBorderLeft,
            })
          }
        } else if (chainActive && rendered[0]?.__h2dSpacerPx) {
          // First substantive block: bind its own leading margin spacer.
          rendered[0] = spacerParagraph(this.context, rendered[0].__h2dSpacerPx, {
            keepNext: true,
            shading: rendered[0].__h2dSpacerShading,
            borderLeft: rendered[0].__h2dSpacerBorderLeft,
          })
        }
        if (node.type === 'heading') {
          chainActive = !this.suppressKeepNext
          ledeAllowance = 1
        } else {
          chainActive = false
        }
      }
      // A near-page-height unsplittable card cannot share a page with the
      // gap that precedes it (the card slides on; the invisible gap strands
      // a blank page) — collapse the previous block's trailing spacer.
      if (
        nodes[index + 1]?.type === 'card' &&
        (nodes[index + 1].heightPx || 0) > 700 &&
        !nodes[index + 1].allowSplit
      ) {
        const last = rendered.length - 1
        if ((rendered[last]?.__h2dSpacerPx || 0) > 8) {
          rendered[last] = spacerParagraph(this.context, 8, {
            shading: rendered[last].__h2dSpacerShading,
          })
        }
      }
      out.push(...rendered)
    }
    return out
  }

  renderNode(node, depth) {
    const context = this.context
    switch (node.type) {
      case 'pagebreak':
        return [
          new Paragraph({
            children: [new PageBreak()],
            spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
          }),
        ]

      case 'spacer':
        return [
          spacerParagraph(context, node.px, {
            keepNext: this.forceKeepNext,
            shading: node.shading,
            borderLeft: node.borderLeft,
          }),
        ]

      case 'heading': {
        // shaded single-line section header bar (h2 with background +
        // padding): same treatment as shaded paras, otherwise the text sits
        // baseline-left in the bar with no inset or vertical centering
        if (node.style?.shading && node.style?.colorBar) {
          return renderColorBar(this, node)
        }
        node = this.fitOneLineRuns(node)
        const { paragraphStyle, spacingAfterPx, spacingBeforePx } = splitVerticalBorderSpacing(
          node.style,
        )
        const opts = paraOptions(context, paragraphStyle)
        // Built-in Word heading styles add their own top/bottom spacing.
        // Override both sides explicitly so CSS margin: 0 remains zero.
        opts.spacing = {
          ...opts.spacing,
          before: opts.spacing?.before ?? 0,
          after: opts.spacing?.after ?? 0,
        }
        const stops = tabStopsFor(context, node.runs)
        if (stops.length) opts.tabStops = stops
        return withExternalBorderSpacing(
          context,
          new Paragraph({
            heading: HEADING_LEVELS[Math.min(node.level, 6) - 1],
            children: withBookmarks(makeNodeRuns(context, node, this.images), node.bookmarks),
            keepNext: this.forceKeepNext
              ? true
              : node.style?.keepNext === false || this.suppressKeepNext
                ? undefined
                : true,
            keepLines: true,
            ...opts,
          }),
          spacingBeforePx,
          spacingAfterPx,
          paragraphStyle.shading,
        )
      }

      case 'para': {
        if (!node.runs.length) return []
        if (node.style?.shading && node.style?.colorBar) {
          return renderColorBar(this, node)
        }
        node = this.fitFillInRuns(node)
        node = this.fitOneLineRuns(node)
        const { paragraphStyle, spacingAfterPx, spacingBeforePx } = splitVerticalBorderSpacing(
          node.style,
        )
        const opts = paraOptions(context, paragraphStyle)
        const stops = tabStopsFor(context, node.runs)
        if (stops.length) opts.tabStops = stops
        if (node.style?.bullet) {
          opts.numbering = { reference: 'h2d-ul', level: 0 }
        }
        if (this.forceKeepNext || (node.style?.keepNext && !this.suppressKeepNext)) {
          opts.keepNext = true
        }
        if (node.style?.keepLines) opts.keepLines = true
        return withExternalBorderSpacing(
          context,
          new Paragraph({
            children: withBookmarks(makeNodeRuns(context, node, this.images), node.bookmarks),
            ...opts,
          }),
          spacingBeforePx,
          spacingAfterPx,
          paragraphStyle.shading,
        )
      }

      case 'list': {
        // One numbering instance per sublist, emitted at level 0: docx's
        // startOverride only restarts level 0, so real multi-level lists
        // would inherit a document-wide running counter (nested a-e lists
        // continuing f-j, r-bb...). Level 0 per instance restarts like the
        // browser restarts each <ol>; visual nesting comes from indentLeftPx.
        const instanceByLevel = []
        return node.items.map((entry) => {
          // Older IR snapshots stored runs directly; accept both shapes so
          // callers that persist/debug IR remain compatible.
          const item = Array.isArray(entry) ? { runs: entry, style: {} } : entry
          const opts = paraOptions(context, item.style)
          const markerWidthPx = Math.max(10, Math.min(16, (item.runs[0]?.sizePx || 14) * 0.8))
          opts.indent = {
            left: context.pxToTwips(item.indentLeftPx || node.indentLeftPx || 22),
            hanging: context.pxToTwips(markerWidthPx),
          }
          if (node.rtl) opts.bidirectional = true
          if (this.forceKeepNext) opts.keepNext = true
          const ordered = item.ordered ?? node.ordered
          const level = Math.max(0, Math.min(8, item.level || 0))
          if (level < instanceByLevel.length - 1) instanceByLevel.length = level + 1
          if (instanceByLevel[level] == null) instanceByLevel[level] = this.olInstance++
          const bulletReference =
            item.markerType === 'square'
              ? 'h2d-ul-square'
              : item.markerType === 'circle'
                ? 'h2d-ul-circle'
                : 'h2d-ul'
          if (item.continuation) {
            opts.indent = { left: context.pxToTwips(item.indentLeftPx || node.indentLeftPx || 22) }
          } else {
            opts.numbering = ordered
              ? {
                  reference: orderedReference(item.markerType),
                  level: 0,
                  instance: instanceByLevel[level],
                }
              : { reference: bulletReference, level }
          }
          return new Paragraph({
            children: withBookmarks(
              makeNodeRuns(context, { runs: item.runs, style: item.style }, this.images),
              item.bookmarks,
            ),
            ...opts,
          })
        })
      }

      case 'code':
        return node.lines.map(
          (line) =>
            new Paragraph({
              children: [
                new TextRun({
                  text: line || ' ',
                  font: 'Consolas',
                  size: node.sizePx ? context.pxToHalfPoints(node.sizePx) : 18,
                  color: node.color && node.color !== '000000' ? node.color : undefined,
                  noProof: true,
                }),
              ],
              shading: { type: ShadingType.CLEAR, fill: node.shading },
            }),
        )

      case 'hr': {
        const widthFrac = Math.max(0.05, Math.min(1, node.widthFrac || 1))
        if (node.short) {
          const width = Math.max(120, Math.round(this.avail * widthFrac))
          const output = []
          if (node.spacingBeforePx) {
            output.push(
              spacerParagraph(context, node.spacingBeforePx, { shading: node.ambientShading }),
            )
          }
          output.push(
            new Table({
              borders: NO_TABLE_BORDERS,
              rows: [
                new TableRow({
                  height: {
                    value: context.pxToTwips(node.heightPx || 2),
                    rule: HeightRule.EXACT,
                  },
                  children: [
                    new TableCell({
                      children: [new Paragraph({ spacing: { before: 0, after: 0 } })],
                      shading: { type: ShadingType.CLEAR, fill: node.color || 'CCCCCC' },
                      borders: NO_BORDERS,
                      margins: { top: 0, bottom: 0, left: 0, right: 0 },
                    }),
                  ],
                }),
              ],
              width: { size: width, type: WidthType.DXA },
              columnWidths: [width],
              layout: TableLayoutType.FIXED,
            }),
            spacerParagraph(context, node.spacingAfterPx || 2, { shading: node.ambientShading }),
          )
          return output
        }
        const spare = Math.round(this.avail * (1 - widthFrac))
        const left = node.align === 'center' ? Math.round(spare / 2) : 0
        const right = node.align === 'center' ? spare - left : spare
        return [
          new Paragraph({
            indent: { left, right },
            shading: node.ambientShading
              ? { type: ShadingType.CLEAR, fill: node.ambientShading }
              : undefined,
            // collapse the carrier paragraph to (almost) just the border line;
            // a default-height empty line reads as a stray blank row
            children: [new TextRun({ text: '', size: 2, noProof: true })],
            spacing: {
              before: context.pxToTwips(node.spacingBeforePx || 0),
              after: context.pxToTwips(node.spacingAfterPx || 2),
              line: context.pxToTwips(Math.max(2, node.heightPx || 2)),
              lineRule: LineRuleType.EXACT,
            },
            border: {
              bottom: {
                style: BorderStyle.SINGLE,
                size: node.heightPx ? Math.max(6, Math.round(node.heightPx * 8)) : 6,
                color: node.color || 'CCCCCC',
              },
            },
          }),
        ]
      }

      case 'formfield': {
        if (node.controlType === 'checkbox' || node.controlType === 'radio') {
          return [
            new Paragraph({
              children: [this.contentControl(node)],
              spacing: { before: 0, after: 0 },
              ...paraOptions(context, node.style),
            }),
          ]
        }
        if (node.mode === 'line') {
          return [
            new Paragraph({
              children: [this.contentControl(node)],
              border: { bottom: wordBorder(context, node.borders?.bottom) },
              spacing: { before: 0, after: 0, line: 240 },
            }),
          ]
        }
        const width = Math.min(this.avail, Math.max(300, context.pxToTwips(node.width || 100)))
        const borders: any = {}
        for (const side of ['top', 'bottom', 'left', 'right']) {
          borders[side] = wordBorder(context, node.borders?.[side])
        }
        const pad = node.padPx || {}
        return [
          new Table({
            borders: NO_TABLE_BORDERS,
            rows: [
              new TableRow({
                cantSplit: true,
                height: node.height
                  ? { value: context.pxToTwips(node.height), rule: HeightRule.ATLEAST }
                  : undefined,
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [this.contentControl(node)],
                        spacing: { before: 0, after: 0 },
                      }),
                    ],
                    borders,
                    shading: node.shading
                      ? { type: ShadingType.CLEAR, fill: node.shading }
                      : undefined,
                    margins: {
                      top: context.pxToTwips(pad.top || 4),
                      bottom: context.pxToTwips(pad.bottom || 4),
                      left: context.pxToTwips(pad.left || 6),
                      right: context.pxToTwips(pad.right || 6),
                    },
                  }),
                ],
              }),
            ],
            width: { size: width, type: WidthType.DXA },
            columnWidths: [width],
          }),
          spacerParagraph(context, 2),
        ]
      }

      case 'image':
        return this.renderImage(node)

      case 'floatimg':
        return this.collectFloatingImage(node)

      case 'table':
        return renderTable(this, node, depth)

      case 'card':
        return renderCard(this, node, depth)

      case 'kpirow':
        return renderKpiRow(this, node, depth)

      default:
        return []
    }
  }

  // Display text measured as one browser line but wider than the cell it
  // landed in: Word would wrap it mid-word ("INVOIC/E"). Shrink the font
  // proportionally so it stays on one line. Display sizes only — body text
  // in narrow columns should wrap naturally, not shrink.
  // A fill-in blank is an unbreakable nbsp run sized to its browser width.
  // Inside a narrow table cell it overflows and blows up Word's fixed
  // layout (the whole table widens past the page and clips) — shrink the
  // blank to the width the cell actually has left.
  fitFillInRuns(node) {
    const runs = node.runs || []
    const isBlank = (run) =>
      run.underline && run.preserveWhitespace && /^\u00a0+$/.test(run.text || '')
    if (!runs.some(isBlank)) return node
    const availPx = this.avail / 15 / this.context.measurementScale
    const spaceWidth = (run) => Math.max(3, (run.sizePx || 14) * 0.45)
    const textPx = runs.reduce(
      (sum, run) => (isBlank(run) ? sum : sum + (run.text || '').length * (run.sizePx || 14) * 0.6),
      0,
    )
    const blankPx = runs.reduce(
      (sum, run) => (isBlank(run) ? sum + run.text.length * spaceWidth(run) : sum),
      0,
    )
    if (textPx + blankPx <= availPx * 0.92) return node
    const budget = Math.max(0, availPx * 0.92 - textPx)
    const scale = blankPx > 0 ? Math.min(1, budget / blankPx) : 1
    return {
      ...node,
      runs: runs.map((run) =>
        isBlank(run)
          ? { ...run, text: '\u00a0'.repeat(Math.max(3, Math.floor(run.text.length * scale))) }
          : run,
      ),
    }
  }

  fitOneLineRuns(node) {
    if (!node.style?.oneLineWidthPx) return node
    // Headings never letter-wrap regardless of size; body text only when it
    // is display-scale (small body copy should wrap naturally).
    const sizeFloor = node.type === 'heading' || node.style?.keepLines ? 14 : 24
    if (!node.runs.some((run) => (run.sizePx || 0) >= sizeFloor)) return node
    const availPx = this.avail / 15 / this.context.measurementScale
    if (node.style.oneLineWidthPx <= availPx * 0.94) return node
    // 6% safety: Word's Arial runs slightly wider than browser fonts.
    const fit = Math.max(0.55, (availPx * 0.94) / node.style.oneLineWidthPx)
    return {
      ...node,
      runs: node.runs.map((run) => (run.sizePx ? { ...run, sizePx: run.sizePx * fit } : run)),
    }
  }

  renderImage(node) {
    const img = this.images[node.shotId]
    if (!img) {
      const fallbackRun = new TextRun({
        text: node.fallbackText || node.href || 'Unavailable embedded content',
        color: node.href ? '0563C1' : '666666',
        underline: node.href ? {} : undefined,
        noProof: true,
      })
      return [
        new Paragraph({
          children: node.href
            ? [new ExternalHyperlink({ children: [fallbackRun], link: node.href })]
            : [fallbackRun],
        }),
      ]
    }
    const bleedLeftPx = node.bleedLeftPx || 0
    const bleedRightPx = node.bleedRightPx || 0
    const maxPx = Math.floor((this.avail + this.context.pxToTwips(bleedLeftPx + bleedRightPx)) / 15)
    const pageContentHeightPx =
      (this.context.pageHeightDxa -
        this.context.pageMargins.top -
        this.context.pageMargins.bottom) /
      15
    // Word (Mac) starts the first line ~26px below the page top even at zero
    // margins, and the image line adds baseline descent below. A slice scaled
    // to (content height - 4) therefore clipped its bottom ~25px at the page
    // edge (lost resume contact rows). 45px absorbs the inset plus variance.
    const pageScale = node.pageComposition
      ? Math.max(0.01, pageContentHeightPx - 45) / node.height
      : Infinity
    const scale = Math.min(this.context.measurementScale, maxPx / node.width, pageScale)
    const output = []
    // Full-page composition slides: the page margins already carry the
    // authored offset — an extra spacer overflows the page and strands the
    // slide image on a blank-page boundary.
    if (node.spacingBeforePx > 2 && !node.pageComposition) {
      output.push(spacerParagraph(this.context, node.spacingBeforePx))
    }
    const renderedWidthPx = Math.round(node.width * scale)
    const renderedHeightPx = Math.round(node.height * scale)
    const imageRun = new ImageRun({
      type: 'png',
      data: img,
      transformation: {
        width: renderedWidthPx,
        height: renderedHeightPx,
      },
    })
    output.push(
      new Paragraph({
        children: node.href
          ? [new ExternalHyperlink({ children: [imageRun], link: node.href })]
          : [imageRun],
        // wide short strips are section banners — never strand one at a
        // page bottom with its section content on the next page
        keepNext:
          !this.suppressKeepNext &&
          !node.pageComposition &&
          renderedWidthPx >= (this.avail / 15) * 0.85 &&
          renderedHeightPx <= 220
            ? true
            : undefined,
        alignment:
          // Full-page slides: headroom scaling leaves a margin — split it
          // symmetrically instead of piling white space on one side.
          node.pageComposition || node.align === 'center'
            ? AlignmentType.CENTER
            : node.align === 'right'
              ? AlignmentType.RIGHT
              : undefined,
        indent:
          bleedLeftPx || bleedRightPx
            ? {
                left: -this.context.pxToTwips(bleedLeftPx),
                right: -this.context.pxToTwips(bleedRightPx),
              }
            : undefined,
        spacing: {
          before: 0,
          after: 0,
          line: Math.max(1, renderedHeightPx * 15),
          // Word positions inline images relative to the text baseline.
          // An exact line box equal to the image height clips the image's top
          // and may push it to the next page. "At least" keeps the authored
          // image height while allowing Word to include the full ascent.
          lineRule: LineRuleType.AT_LEAST,
        },
        border: node.topEdgeFill
          ? {
              top: {
                style: BorderStyle.SINGLE,
                size: 48,
                color: node.topEdgeFill,
                space: 0,
              },
            }
          : undefined,
      }),
    )
    if (node.spacingAfterPx > 2 && !node.pageComposition) {
      output.push(spacerParagraph(this.context, node.spacingAfterPx))
    }
    return output
  }

  collectFloatingImage(node) {
    const img = this.images[node.shotId]
    if (!img) return []
    // in-parent decorations beyond the first page reach here only when no
    // card consumed them — page-anchoring would pin them to page 1
    if (node.inParent && node.yPx + node.height > (node.pageHpx || 1123)) return []
    const EMU = 9525 // per px @96dpi
    // rescale browser coordinates: the DOCX content area is usually
    // narrower/shorter than the HTML content box
    const contentHeightDxa = this.context.pageHeightDxa - 2 * this.context.marginDxa
    const sx = node.pageWpx ? this.context.contentDxa / 15 / node.pageWpx : 1
    const sy = node.pageHpx ? contentHeightDxa / 15 / node.pageHpx : 1
    this.floats.push(
      new ImageRun({
        type: 'png',
        data: img,
        transformation: {
          width: Math.round(node.width * sx),
          height: Math.round(node.height * sx),
        },
        floating: {
          // margin origin = start of the content area, matching the
          // body-content-box coordinates measured in the browser
          horizontalPosition: { relative: 'margin', offset: Math.round(node.xPx * sx * EMU) },
          verticalPosition: node.atPageBottom
            ? { relative: 'page', align: VerticalPositionAlign.BOTTOM }
            : { relative: 'margin', offset: Math.round(node.yPx * Math.min(sy, 1) * EMU) },
          // decorations painted under the HTML content layer (z<=0) go
          // behind text; explicit z>0 art stays in front (table-cell shading
          // would paint over a behind-text image, so only when needed)
          behindDocument: Boolean(node.behindText),
          wrap: { type: TextWrappingType.NONE },
        },
      }),
    )
    return []
  }
}

export { Generator }
