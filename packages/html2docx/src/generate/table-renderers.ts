// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import {
  AlignmentType,
  BorderStyle,
  HeightRule,
  ImageRun,
  LineRuleType,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextWrappingType,
  VerticalAlign,
  VerticalMergeType,
  WidthType,
} from 'docx'
import {
  NO_BORDER,
  NO_BORDERS,
  NO_TABLE_BORDERS,
  makeRuns,
  paraOptions,
  spacerParagraph,
  wordBorder,
} from './word-utils'

// Render a cell's content with the available width narrowed to the cell.
// keepWithNext: every paragraph in the cell carries keepNext — Word promotes
// that to row-level keep-with-next, binding this row to the following one.
// Defaults to the ambient flag so a card nested inside a kept row inherits
// the binding instead of silently resetting it.
function cellChildren(
  generator,
  entry,
  depth,
  cellDxa,
  { keepWithNext = generator.forceKeepNext } = {},
) {
  const saved = generator.avail
  const savedKeep = generator.suppressKeepNext
  const savedForce = generator.forceKeepNext
  if (cellDxa) generator.avail = Math.max(300, cellDxa)
  generator.suppressKeepNext = !keepWithNext
  generator.forceKeepNext = keepWithNext
  try {
    return cellChildrenInner(generator, entry, depth)
  } finally {
    generator.avail = saved
    generator.suppressKeepNext = savedKeep
    generator.forceKeepNext = savedForce
  }
}

function cellChildrenInner(generator, entry, depth) {
  let children
  if (entry.children) children = generator.render(entry.children, depth + 1)
  else {
    const opts = paraOptions(generator.context, entry.style)
    const spacing = {
      ...(opts.spacing || {}),
      before: 0,
      after: 0,
    }
    const align =
      entry.align === 'center'
        ? AlignmentType.CENTER
        : entry.align === 'right'
          ? AlignmentType.RIGHT
          : undefined
    children = [
      new Paragraph({
        children: makeRuns(generator.context, entry.runs || [], generator.images),
        alignment: align,
        bidirectional: entry.rtl || undefined,
        ...opts,
        keepNext: generator.forceKeepNext || opts.keepNext || undefined,
        spacing,
      }),
    ]
  }
  // a table cell must contain at least one paragraph
  if (!children.length) {
    children = [new Paragraph({ keepNext: generator.forceKeepNext || undefined })]
  }
  return children
}

function renderTable(generator, node, depth) {
  const context = generator.context
  if (node.splitPerRow && node.rows.length > 1) {
    return node.rows.flatMap((row) =>
      renderTable(generator, { ...node, splitPerRow: false, rows: [row] }, depth),
    )
  }
  if (node.pageColumns) return renderPageColumns(generator, node, depth)
  const tableDxa = generator.avail
  const totalPx = (node.colWidths || []).reduce((a, b) => a + b, 0) || 1
  const colDxa = (node.colWidths || []).map((w) => Math.round((w / totalPx) * tableDxa))
  // Column-width floor from measured one-line cell content: a proportionally
  // shrunk column otherwise wraps "Wave 1" pills letter-by-letter into a fake
  // vertical stack. Conserve the table width by borrowing from the widest
  // column (never below its own floor).
  if (colDxa.length > 1) {
    const floors = colDxa.map(() => 0)
    for (const row of node.rows) {
      for (const cell of row.cells) {
        if (!cell.oneLineWidthPx || (cell.colspan || 1) > 1) continue
        const col = cell.gridStart ?? row.cells.indexOf(cell)
        if (col < 0 || col >= floors.length) continue
        const pad = cell.padPx || {}
        const needDxa =
          context.pxToTwips(Math.round(cell.oneLineWidthPx * 1.1)) +
          context.pxToTwips((pad.left ?? 6) + (pad.right ?? 6)) +
          120
        floors[col] = Math.max(floors[col], needDxa)
      }
    }
    // Readability floor for text-bearing columns. A fixed-layout source whose
    // width rules over-consume the table (a missing .w-16 class) legitimately
    // computes a 0-1px column; the browser lets the text overflow its
    // neighbors, Word cannot and letter-stacks it into a 1-char sliver.
    const cellHasText = (cell, depth = 0) => {
      if (depth > 3 || !cell || typeof cell !== 'object') return false
      if ((cell.runs || []).some((run) => (run.text || '').trim().length >= 2)) return true
      return (cell.children || []).some((child) => cellHasText(child, depth + 1))
    }
    const TEXT_COL_FLOOR_DXA = 800
    for (const row of node.rows) {
      for (const cell of row.cells) {
        if ((cell.colspan || 1) > 1 || !cellHasText(cell)) continue
        const col = cell.gridStart ?? row.cells.indexOf(cell)
        if (col < 0 || col >= floors.length) continue
        floors[col] = Math.max(floors[col], TEXT_COL_FLOOR_DXA)
      }
    }
    floors.forEach((floor, col) => {
      if (!floor || colDxa[col] >= floor) return
      let deficit = floor - colDxa[col]
      // Borrow from every column with slack above its own floor, widest
      // first, so one floor-locked column cannot block the whole rebalance.
      const donors = colDxa
        .map((w, j) => ({ j, slack: j === col ? 0 : w - Math.max(600, floors[j]) }))
        .filter((d) => d.slack > 0)
        .sort((a, b) => b.slack - a.slack)
      for (const donor of donors) {
        if (deficit <= 0) break
        const take = Math.min(deficit, donor.slack)
        colDxa[donor.j] -= take
        deficit -= take
      }
      colDxa[col] = floor - Math.max(0, deficit)
    })
  }
  // layout tables (page columns): invisible, rows break freely across pages
  const layout = Boolean(node.layout)
  // A content table that fits on one page must not split at a row boundary
  // mid-page (a 3x3 canvas losing its last row to the next page): keepNext on
  // every paragraph of the non-last rows promotes to row-level keep-with-next
  // and Word moves the table to a fresh page whole. Taller tables must stay
  // splittable or they wedge against the page and overflow.
  const keepWhole =
    !layout &&
    !node.checklist &&
    node.rows.length > 1 &&
    (node.heightPx || 0) > 0 &&
    node.heightPx <= 850
  const padScale = 1
  // No synthesized default border: a cell without computed CSS borders is
  // borderless in the browser too — a fallback grid paints separators that
  // don't exist (white seams across a solid-color header row).
  const border = NO_BORDER
  const rowGap = node.rowGapPx ? context.pxToTwips(node.rowGapPx) : 0
  const pageContentDxa =
    context.pageHeightDxa - context.pageMargins.top - context.pageMargins.bottom
  const rows = node.rows.map((row, rowIdx) => {
    // A row taller than one page can never be placed whole: cantSplit makes
    // Word shove it to a fresh page (stranding a near-blank one) and clip.
    // Let such rows flow and let their content set the height.
    const tallerThanPage = row.heightPx && context.pxToTwips(row.heightPx) > pageContentDxa - 900
    return new TableRow({
      cantSplit: layout ? node.noSplit || undefined : tallerThanPage ? undefined : true,
      tableHeader: (!node.noHeader && row.header) || undefined,
      // Always ATLEAST: EXACT clips lines when Word wraps text differently
      // than the browser measurement (see kpirow note below).
      // Cap the minimum height below one page: a row whose declared floor
      // exceeds the page can never be placed and strands a blank page
      // before it; the content still sets the real rendered height.
      height:
        row.heightPx && !tallerThanPage
          ? {
              value: Math.min(context.pxToTwips(row.heightPx), pageContentDxa - 900),
              rule: HeightRule.ATLEAST,
            }
          : undefined,
      children: row.cells.map((cell, ci) => {
        const gridStart =
          cell.gridStart ??
          row.cells.slice(0, ci).reduce((sum, previous) => sum + (previous.colspan || 1), 0)
        const cellDxa = colDxa
          .slice(gridStart, gridStart + (cell.colspan || 1))
          .reduce((sum, width) => sum + width, 0)
        if (cell.bold && cell.runs) cell.runs = cell.runs.map((r) => ({ ...r, bold: true }))
        const cellBorders: any = {}
        for (const side of ['top', 'bottom', 'left', 'right']) {
          cellBorders[side] = cell.borders ? wordBorder(context, cell.borders[side]) : border
        }
        if (node.outerBorder) {
          if (rowIdx === 0 && node.outerBorder.top) {
            cellBorders.top = wordBorder(context, node.outerBorder.top)
          }
          if (rowIdx === node.rows.length - 1 && node.outerBorder.bottom) {
            cellBorders.bottom = wordBorder(context, node.outerBorder.bottom)
          }
          if (ci === 0 && node.outerBorder.left) {
            cellBorders.left = wordBorder(context, node.outerBorder.left)
          }
          if (ci === row.cells.length - 1 && node.outerBorder.right) {
            cellBorders.right = wordBorder(context, node.outerBorder.right)
          }
        }
        const pad = cell.padPx || {}
        const tablePad = node.padPx || {}
        return new TableCell({
          children: cellChildren(generator, cell, depth, cellDxa, {
            keepWithNext: keepWhole && rowIdx < node.rows.length - 1,
          }),
          width: cellDxa ? { size: cellDxa, type: WidthType.DXA } : undefined,
          columnSpan: cell.colspan > 1 ? cell.colspan : undefined,
          verticalMerge:
            cell.verticalMerge === 'restart'
              ? VerticalMergeType.RESTART
              : cell.verticalMerge === 'continue'
                ? VerticalMergeType.CONTINUE
                : undefined,
          shading:
            cell.shading || node.shading
              ? { type: ShadingType.CLEAR, fill: cell.shading || node.shading }
              : undefined,
          borders: cellBorders,
          verticalAlign:
            cell.vAlign === 'center'
              ? VerticalAlign.CENTER
              : cell.vAlign === 'bottom'
                ? VerticalAlign.BOTTOM
                : undefined,
          margins: layout
            ? {
                top: cell.shading ? 40 : 0,
                bottom: rowIdx < node.rows.length - 1 ? rowGap : 0,
                left: cell.shading ? 60 : 0,
                // shaded layout cells (planner headers) form contiguous bars
                right: cell.shading ? 60 : 200,
              }
            : {
                top: context.pxToTwips(
                  (pad.top ?? 4) * padScale + (rowIdx === 0 ? tablePad.top || 0 : 0),
                ),
                bottom: context.pxToTwips(
                  (pad.bottom ?? 4) * padScale +
                    (rowIdx === node.rows.length - 1 ? tablePad.bottom || 0 : 0),
                ),
                left: context.pxToTwips(
                  (pad.left ?? 6) * padScale + (ci === 0 ? tablePad.left || 0 : 0),
                ),
                right: context.pxToTwips(
                  (pad.right ?? 6) * padScale +
                    (ci === row.cells.length - 1 ? tablePad.right || 0 : 0),
                ),
              },
        })
      }),
    })
  })
  const output = []
  if (node.spacingBeforePx) {
    output.push(spacerParagraph(context, node.spacingBeforePx, { shading: node.ambientShading }))
  }
  output.push(
    new Table({
      borders: NO_TABLE_BORDERS,
      rows,
      width: { size: tableDxa, type: WidthType.DXA },
      columnWidths: colDxa.length ? colDxa : undefined,
      layout: TableLayoutType.FIXED,
      // RTL source table: first DOM cell renders rightmost, same as browser
      visuallyRightToLeft: node.rtl || undefined,
    }),
    // required separator paragraph after a table
    spacerParagraph(context, node.spacingAfterPx ?? 8, { shading: node.ambientShading }),
  )
  return output
}

// A keepNext on a cell's FIRST paragraph can only bind upward out of the
// cell — meaningless, and Word (Mac) refuses to start a table row whose
// first paragraph carries keepNext mid-page: the whole multi-page row
// slides to a fresh page and strands the current one.
function stripLeadingKeep(cell) {
  const children = cell.children || []
  // leading spacers render as paragraphs too — the first REAL paragraph
  // after them is the one whose keepNext wedges the row start
  const index = children.findIndex((n) => n.type !== 'spacer')
  const first = children[index]
  // headings hardcode keepNext at render time — force an explicit false
  if (!first || (first.type !== 'heading' && !first.style?.keepNext)) return cell
  return {
    ...cell,
    children: [
      ...children.slice(0, index),
      { ...first, style: { ...first.style, keepNext: false } },
      ...children.slice(index + 1),
    ],
  }
}

// Page-level columns (sidebar + main): one borderless single-row table.
// The row may break across pages (cantSplit: false is the docx default),
// which is how Word natively paginates resume-style column layouts.
function renderPageColumns(generator, node, depth) {
  const context = generator.context
  const tableDxa = generator.avail
  const totalPx = (node.colWidths || []).reduce((a, b) => a + b, 0) || 1
  const colDxa = (node.colWidths || []).map((w) => Math.round((w / totalPx) * tableDxa))
  return [
    new Table({
      borders: NO_TABLE_BORDERS,
      // FIXED, or Word autofits: a cell whose image is naturally wider than
      // its column expands the column past the page edge and clips it
      layout: TableLayoutType.FIXED,
      rows: node.rows.map(
        (row) =>
          new TableRow({
            children: row.cells.map((cell, ci) => {
              const pad = cell.padPx || {}
              const leftMargin = context.pxToTwips(pad.left || 0)
              const rightMargin = Math.max(
                context.pxToTwips(pad.right || 0),
                context.pxToTwips(cell.gapAfterPx || 0),
              )
              const columnChildren = cellChildren(
                generator,
                stripLeadingKeep(cell),
                depth,
                Math.max(240, colDxa[ci] - leftMargin - rightMargin),
              )
              // Page-column cells span the whole page: invisible trailing
              // spacers can push past the page bottom and create a phantom
              // background-only last page.
              while (
                columnChildren.length > 1 &&
                columnChildren[columnChildren.length - 1]?.__h2dSpacerPx
              ) {
                columnChildren.pop()
              }
              if (!(columnChildren[columnChildren.length - 1] instanceof Paragraph)) {
                // A cell may not end with a table: the docx layer would append
                // a default-height paragraph, taller than the spacers removed.
                columnChildren.push(
                  new Paragraph({
                    children: [],
                    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
                  }),
                )
              }
              return new TableCell({
                children: columnChildren,
                width: { size: colDxa[ci], type: WidthType.DXA },
                shading: cell.shading ? { type: ShadingType.CLEAR, fill: cell.shading } : undefined,
                borders: {
                  ...NO_BORDERS,
                  ...(cell.borderLeft ? { left: wordBorder(context, cell.borderLeft) } : {}),
                  ...(cell.borderRight ? { right: wordBorder(context, cell.borderRight) } : {}),
                },
                margins: {
                  top: context.pxToTwips(pad.top || 0),
                  bottom: 60,
                  left: leftMargin,
                  right: rightMargin,
                },
              })
            }),
          }),
      ),
      width: { size: tableDxa, type: WidthType.DXA },
      columnWidths: colDxa.length ? colDxa : undefined,
    }),
    spacerParagraph(context, 8),
  ]
}

// A big splittable card that contains tables renders as a single-row table
// wrapping nested tables — a row Word (Mac) never breaks internally, so the
// whole card slides to a fresh page and strands half the previous page
// blank. Flatten it instead: text becomes card-shaded paragraphs, inner
// tables/kpirows are promoted to top level (their rows then paginate
// normally and tblHeader repetition works — it is dead inside nested
// tables). Only light/unshaded borderless top-level cards qualify: tinted
// or framed cards would show visible seams between the flattened pieces.
function isLightShading(hex) {
  if (!hex) return true
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return false
  return [0, 2, 4].every((i) => parseInt(hex.slice(i, i + 2), 16) >= 0xdd)
}

const FLOW_CHILD_TYPES = new Set(['para', 'heading', 'list', 'spacer', 'hr'])

function cardShouldFlatten(node, depth) {
  if (
    depth !== 0 ||
    !node.allowSplit ||
    (node.heightPx || 0) <= 561 || // > half a page: keep-whole wastes more than it protects
    !Array.isArray(node.children) ||
    node.children.length <= 1 ||
    node.border ||
    !isLightShading(node.shading)
  ) {
    return false
  }
  if (node.children.some((child) => child.type === 'table' || child.type === 'kpirow')) {
    return true
  }
  // Pure text/list cards must flatten too: rendered as a single-cell row,
  // keep-with-next on the row's first paragraph (bold pseudo-heading labels)
  // makes Word (Mac) refuse to start the row mid-page — the whole card
  // slides to a fresh page and strands a nearly blank one behind its
  // heading. Flattened paragraphs paginate normally.
  return node.children.every((child) => FLOW_CHILD_TYPES.has(child.type))
}

// floatimg children measured relative to this card's box in the browser —
// consumed by the card renderers, which anchor them to the card itself.
function cardOverlays(generator, node) {
  return (node.children || []).filter(
    (child) => child?.type === 'floatimg' && child.inParent && generator.images[child.shotId],
  )
}

// The flattened card loses its single shaded cell, so every child block (not
// just paras) must carry the card paint — unshaded spacers/tables otherwise
// punch page-background stripes through the section. ambientShading marks the
// surrounding section paint for the child's OWN before/after spacers, which
// sit on the section background rather than inside the child's box.
function withCardShading(child, shading, borderLeft = null) {
  switch (child.type) {
    case 'spacer':
      return {
        ...child,
        shading: child.shading ?? shading,
        borderLeft: child.borderLeft ?? borderLeft,
      }
    case 'list':
      return {
        ...child,
        items: (child.items || []).map((item) =>
          Array.isArray(item)
            ? { runs: item, style: { shading, borderLeft } }
            : {
                ...item,
                style: {
                  ...item.style,
                  shading: item.style?.shading ?? shading,
                  borderLeft: item.style?.borderLeft ?? borderLeft,
                },
              },
        ),
      }
    case 'kpirow':
      return {
        ...child,
        ambientShading: shading,
        gapShading: child.gapShading ?? shading,
        cells: (child.cells || []).map((cell) => ({
          ...cell,
          shading: cell.shading ?? shading,
        })),
      }
    case 'table':
      return { ...child, ambientShading: shading, shading: child.shading ?? shading }
    case 'card':
    case 'hr':
      return { ...child, ambientShading: shading }
    default:
      return child
  }
}

function renderCardFlattened(generator, node, depth) {
  const context = generator.context
  const pad = node.padPx || {}
  const shading = node.shading
  // paragraph shading starts at the left indent, but tables/kpirows sit at
  // the column edge — an inset would serrate the section band's left rail
  const insetPx = shading ? 0 : Math.min(24, pad.left ?? 8)
  const overlays = cardOverlays(generator, node)
  // keep enough of the authored top padding to contain overlays that live in
  // it (section watermark numerals) — compressing it drops them onto the text
  const colPx = generator.avail / 15
  const overlayPadNeed = overlays.reduce((need, ov) => {
    const rel = ov.inParent
    if (rel.yPx >= (pad.top ?? 0)) return need
    const parentContentW = Math.max(1, rel.parentWpx - (pad.left ?? 8) - (pad.right ?? 8))
    const scale = Math.min(1, colPx / parentContentW)
    return Math.max(need, rel.yPx + Math.round(ov.height * scale) + 8)
  }, 0)
  const padTopCapPx = pad.top > 2 ? Math.min(pad.top, Math.max(24, overlayPadNeed)) : 0
  // The card's left rail survives flattening as a paragraph border: Word
  // merges identical adjacent pBdr lefts into one continuous line.
  const railBorder = node.borderLeft || null
  const children = node.children
    .filter((child) => !overlays.includes(child))
    .flatMap((child) => {
      if (child.type === 'para' || child.type === 'heading') {
        const style = {
          ...child.style,
          shading: child.style?.shading ?? shading ?? undefined,
          indentLeftPx: (child.style?.indentLeftPx || 0) + insetPx,
        }
        if (railBorder) {
          style.borderLeft = style.borderLeft ?? railBorder
          // keep the authored padding between rail and text
          style.borderLeftSpacePx = style.borderLeftSpacePx ?? Math.min(24, pad.left ?? 8)
        }
        if (!shading && !railBorder) return [{ ...child, style }]
        // paragraph shading does not cover w:spacing before/after — those
        // gaps would show the page background as seams through the band
        // (and break the rail)
        const out = []
        if (style.spacingBeforePx) {
          out.push({ type: 'spacer', px: style.spacingBeforePx, shading, borderLeft: railBorder })
          style.spacingBeforePx = 0
        }
        const trailing = style.spacingAfterPx
        style.spacingAfterPx = 0
        out.push({ ...child, style })
        if (trailing) out.push({ type: 'spacer', px: trailing, shading, borderLeft: railBorder })
        return out
      }
      return [shading || railBorder ? withCardShading(child, shading, railBorder) : child]
    })
  const output = []
  if (node.spacingBeforePx) output.push(spacerParagraph(context, node.spacingBeforePx))
  if (overlays.length) {
    output.push(cardOverlayParagraph(generator, node, overlays, colPx, { padTopCapPx }))
  }
  if (padTopCapPx) {
    output.push(spacerParagraph(context, padTopCapPx, { shading, borderLeft: railBorder }))
  }
  output.push(...generator.render(children, depth + 1))
  if (pad.bottom > 2) {
    output.push(
      spacerParagraph(context, Math.min(pad.bottom, 24), { shading, borderLeft: railBorder }),
    )
  }
  if (node.spacingAfterPx) output.push(spacerParagraph(context, node.spacingAfterPx))
  return output
}

// Absolutely positioned decorations measured inside the card's box (banner
// leaf icons): anchor them to the cell's own text column instead of the page
// margin, so they stay on the card wherever it lands in Word. Drawn in front
// of the cell shading — a behind-text image would be painted over by it.
function cardOverlayParagraph(generator, node, overlays, colPx, { padTopCapPx = null } = {}) {
  const EMU_PER_PX = 9525
  const pad = node.padPx || {}
  const runs = overlays.map((ov) => {
    const rel = ov.inParent
    const parentContentW = Math.max(1, rel.parentWpx - (pad.left ?? 8) - (pad.right ?? 8))
    const scale = Math.min(1, colPx / parentContentW)
    const width = Math.max(1, Math.round(ov.width * scale))
    const height = Math.max(1, Math.round(ov.height * scale))
    const x = Math.max(
      0,
      Math.min(Math.round((rel.xPx - (pad.left ?? 8)) * scale), Math.round(colPx - width)),
    )
    const padT = pad.top ?? 8
    // flattened cards compress their top padding to padTopCapPx: map offsets
    // inside the original padding zone proportionally into the compressed one
    const y =
      padTopCapPx !== null && padT > 0
        ? rel.yPx <= padT
          ? Math.max(0, Math.round((rel.yPx * padTopCapPx) / padT))
          : Math.round(rel.yPx - padT + padTopCapPx)
        : Math.max(0, Math.round(rel.yPx - padT))
    return new ImageRun({
      type: 'png',
      data: generator.images[ov.shotId],
      transformation: { width, height },
      floating: {
        horizontalPosition: { relative: 'column', offset: x * EMU_PER_PX },
        verticalPosition: { relative: 'paragraph', offset: y * EMU_PER_PX },
        wrap: { type: TextWrappingType.NONE },
        allowOverlap: true,
      },
    })
  })
  return new Paragraph({
    children: runs,
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
  })
}

function renderCard(generator, node, depth) {
  const context = generator.context
  const cardDxa = node.widthPx
    ? Math.min(generator.avail, Math.max(240, context.pxToTwips(node.widthPx)))
    : generator.avail
  // depth guard: deeply nested cards degrade to plain flow
  if (depth >= 3) return generator.render(node.children, depth)
  if (cardShouldFlatten(node, depth)) return renderCardFlattened(generator, node, depth)
  const overlays = cardOverlays(generator, node)
  if (overlays.length) {
    node = { ...node, children: node.children.filter((child) => !overlays.includes(child)) }
  }
  const output = []
  // A near-page-height unsplittable card cannot share its page with the
  // authored gap before it — the card slides on and the invisible gap
  // strands a blank page. Collapse the gap; the page break separates them.
  const spacingBefore =
    (node.heightPx || 0) > 700 && !node.allowSplit
      ? Math.min(node.spacingBeforePx || 0, 8)
      : node.spacingBeforePx
  if (spacingBefore) {
    output.push(spacerParagraph(context, spacingBefore, { shading: node.ambientShading }))
  }
  const borders: any = { ...NO_BORDERS }
  // Mildly rounded cards (≤12px) still read as rectangles: draw their border
  // square rather than dropping it. Heavily rounded ones stay borderless
  // (square corners would visibly contradict the design).
  if (node.border && (node.radiusPx || 0) <= 12) {
    const b = {
      style: BorderStyle.SINGLE,
      // A scaled 1px CSS card border otherwise falls to 0.5pt and becomes
      // nearly invisible in Word at common zoom levels.
      size: Math.max(6, context.pxToBorderEighths(node.border.widthPx)),
      color: node.border.color,
    }
    borders.top = b
    borders.bottom = b
    borders.left = b
    borders.right = b
  }
  if (node.borderLeft) {
    borders.left = {
      style: BorderStyle.SINGLE,
      size: context.pxToBorderEighths(node.borderLeft.widthPx),
      color: node.borderLeft.color,
    }
  }
  if (node.borderRight) {
    borders.right = {
      style: BorderStyle.SINGLE,
      size: context.pxToBorderEighths(node.borderRight.widthPx),
      color: node.borderRight.color,
    }
  }
  if (node.borderTop) {
    borders.top = {
      style: BorderStyle.SINGLE,
      size: context.pxToBorderEighths(node.borderTop.widthPx),
      color: node.borderTop.color,
    }
  }
  const pad = node.padPx || {}
  const cellMargins = {
    top: context.pxToTwips(pad.top ?? 8),
    bottom: context.pxToTwips(pad.bottom ?? 8),
    left: context.pxToTwips(pad.left ?? 8),
    right: context.pxToTwips(pad.right ?? 8),
  }
  // narrow children by the actual cell margins (card padding),
  // not a fixed guess - wide paddings otherwise overflow the card
  const cellKids = cellChildren(
    generator,
    node,
    depth,
    Math.max(240, cardDxa - cellMargins.left - cellMargins.right - 80),
  )
  if (overlays.length) {
    cellKids.unshift(
      cardOverlayParagraph(
        generator,
        node,
        overlays,
        (cardDxa - cellMargins.left - cellMargins.right) / 15,
      ),
    )
  }
  output.push(
    new Table({
      borders: NO_TABLE_BORDERS,
      rows: [
        new TableRow({
          cantSplit: node.allowSplit ? undefined : true,
          children: [
            new TableCell({
              children: cellKids,
              // Explicit tcW: a fixed-layout nested table without it
              // collapses to its minimum width in Word for Mac.
              width: { size: cardDxa, type: WidthType.DXA },
              shading: node.shading ? { type: ShadingType.CLEAR, fill: node.shading } : undefined,
              borders,
              margins: cellMargins,
            }),
          ],
        }),
      ],
      width: { size: cardDxa, type: WidthType.DXA },
      // Fixed layout: with autofit, one unbreakable inner run (fill-in
      // blanks, screenshots) re-widens the card past its container and Word
      // shoves the whole card sideways off the cell.
      layout: TableLayoutType.FIXED,
      columnWidths: [cardDxa],
      // Word hangs a table border 3pt into the left margin even with tblInd=0.
      // Cancel that renderer offset so top-level card borders align with
      // adjacent paragraphs, matching the browser box model.
      indent: depth === 0 ? { size: 60, type: WidthType.DXA } : undefined,
    }),
  )
  if ((node.spacingAfterPx ?? 8) > 0) {
    output.push(
      spacerParagraph(context, node.spacingAfterPx ?? 8, { shading: node.ambientShading }),
    )
  }
  return output
}

// Single-line shaded bar (section headers): a shaded paragraph can't
// reproduce the HTML's exact bar height + text inset (padding / leading
// nbsp), so build a 1x1 borderless table with an exact row height and a
// left cell margin. Table row heights survive far more renderers than
// paragraph exact line spacing does.
function renderColorBar(generator, node) {
  const context = generator.context
  const style = node.style
  const out = []
  if (style.spacingBeforePx) out.push(spacerParagraph(context, style.spacingBeforePx))
  const align =
    style.align === 'center'
      ? AlignmentType.CENTER
      : style.align === 'right'
        ? AlignmentType.RIGHT
        : undefined
  out.push(
    new Table({
      borders: NO_TABLE_BORDERS,
      rows: [
        new TableRow({
          cantSplit: true,
          height: { value: context.pxToTwips(style.exactLineHeightPx), rule: HeightRule.EXACT },
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: makeRuns(context, node.runs, generator.images),
                  alignment: align,
                  bidirectional: style.rtl || undefined,
                  // header bars must not dangle at a page bottom while their
                  // card body starts on the next page
                  keepNext: true,
                  spacing: {
                    before: 0,
                    after: 0,
                    line: context.pxToTwips(style.exactLineHeightPx),
                    lineRule: LineRuleType.EXACT,
                  },
                }),
              ],
              shading: { type: ShadingType.CLEAR, fill: style.shading },
              borders: NO_BORDERS,
              verticalAlign: VerticalAlign.CENTER,
              margins: {
                top: 0,
                bottom: 0,
                left: style.barInsetLeftPx ? context.pxToTwips(style.barInsetLeftPx) : 60,
                right: 60,
              },
            }),
          ],
        }),
      ],
      width: { size: generator.avail, type: WidthType.DXA },
      columnWidths: [generator.avail],
    }),
  )
  out.push(spacerParagraph(context, style.spacingAfterPx || 4, { keepNext: true }))
  return out
}

function renderKpiRow(generator, node, depth) {
  const context = generator.context
  if (depth >= 3) {
    // Too deep for another nested table. If every cell is a single short
    // paragraph (stroke-order rows like "horizontal → vertical hook → rise"), merge them into one
    // inline paragraph instead of stacking each cell on its own line.
    // A cell inlines to one paragraph when it holds a single para, an inline
    // checkbox/radio, or a nested row of such cells (checkbox-item rows:
    // ☐ GASOLINA ☐ DIÉSEL stay on one line like the browser).
    const isInlineBox = (kid) =>
      kid.type === 'formfield' &&
      (kid.controlType === 'checkbox' ||
        kid.controlType === 'radio' ||
        (kid.mode === 'box' && (kid.width || 99) <= 20 && (kid.height || 99) <= 20))
    const inlineCellPara = (cell) => {
      const kids = cell.children || []
      if (kids.length === 2 && isInlineBox(kids[0]) && kids[1].type === 'para') {
        return {
          type: 'para',
          runs: [
            { text: '', contentControl: 'checkbox', checked: kids[0].checked },
            { text: ' ' },
            ...kids[1].runs,
          ],
          style: kids[1].style || {},
        }
      }
      if (kids.length !== 1) return null
      const kid = kids[0]
      if (kid.type === 'para') return kid
      if (isInlineBox(kid)) {
        return {
          type: 'para',
          runs: [{ text: '', contentControl: 'checkbox', checked: kid.checked }],
          style: {},
        }
      }
      if (kid.type === 'kpirow') {
        const parts = (kid.cells || []).map(inlineCellPara)
        if (!parts.length || parts.some((part) => !part)) return null
        const runs = []
        parts.forEach((part, i) => {
          if (i > 0) runs.push({ text: ' ' })
          runs.push(...part.runs)
        })
        const {
          shading: _shading,
          colorBar: _colorBar,
          exactLineHeightPx: _exactLineHeightPx,
          ...style
        } = parts.find((part) => part.style && Object.keys(part.style).length)?.style || {}
        return { type: 'para', runs, style }
      }
      return null
    }
    const singleParas = node.cells.map(inlineCellPara)
    const inlineable =
      node.cells.length > 1 &&
      singleParas.every(
        (p) => p && (p.runs || []).reduce((n, r) => n + (r.text || '').length, 0) <= 30,
      )
    let flow
    if (inlineable) {
      const runs = []
      singleParas.forEach((p, i) => {
        if (i > 0) runs.push({ text: '  ', style: {} })
        runs.push(...p.runs)
      })
      const {
        shading: _shading,
        colorBar: _colorBar,
        exactLineHeightPx: _exactLineHeightPx,
        ...style
      } = singleParas[0].style || {}
      flow = [{ type: 'para', runs, style }]
    } else {
      flow = node.cells.flatMap((cell) => cell.children || [])
      // tiny leading icon cell (emoji chip): merge into the first text
      // paragraph instead of leaving the icon on its own line
      const icon = singleParas[0]
      if (
        icon &&
        (icon.runs || []).reduce((n, r) => n + (r.text || '').length, 0) <= 4 &&
        flow[1] &&
        flow[1].type === 'para'
      ) {
        flow = flow.slice(1)
        flow[0] = { ...flow[0], runs: [...icon.runs, { text: ' ', style: {} }, ...flow[0].runs] }
      }
    }
    // a row that carries its own paint (bordered/shaded card squeezed too
    // deep for a row table) keeps its card visuals via a single-cell card
    if (node.rowBorder || node.rowShading) {
      return renderCard(
        generator,
        {
          type: 'card',
          children: flow,
          border: node.rowBorder,
          shading: node.rowShading,
          padPx: node.padPx,
          spacingBeforePx: node.spacingBeforePx,
          spacingAfterPx: node.spacingAfterPx,
        },
        depth,
      )
    }
    return generator.render(flow, depth)
  }
  const tracks = []
  const useExplicitGaps = Boolean(node.explicitGaps || node.separateCells)
  if (useExplicitGaps) {
    const [leadingGap] = node.outerGapWidths || [0, 0]
    // Leading pad can be authored; trailing leftover flex space must NOT
    // become a column. Proportional scaling against a huge trailing gap
    // crushes short meta cells (a two-word header cell wraps to two lines).
    if (leadingGap >= 2) tracks.push({ gap: true, widthPx: leadingGap })
    node.cells.forEach((cell, index) => {
      tracks.push({ cell, widthPx: node.itemWidths[index] })
      if (index < node.cells.length - 1 && node.gapWidths[index] >= 2) {
        tracks.push({ gap: true, widthPx: node.gapWidths[index] })
      }
    })
    // A right-aligned final cell (INVOICE banners in space-between headers)
    // keeps the same visual position when the flex gap is folded into its
    // width — and gains the room Word's slightly wider Arial needs to keep
    // the display text on one line.
    const last = tracks.length - 1
    if (last >= 1 && tracks[last - 1]?.gap && !tracks[last].gap) {
      const firstChild = tracks[last].cell?.children?.[0]
      if (firstChild?.style?.align === 'right') {
        tracks[last].widthPx += tracks[last - 1].widthPx
        tracks.splice(last - 1, 1)
      }
    }
  } else {
    node.cells.forEach((cell, index) => {
      tracks.push({ cell, widthPx: node.colWidths?.[index] || 1 })
    })
  }
  let colDxa
  if (useExplicitGaps) {
    // Keep browser-measured content widths in twips. Only scale down when
    // the row overflows the page; never stretch/shrink to fill avail.
    // Prefer shrinking gap tracks first — space-between footers otherwise
    // crush email/contact cells and wrap the last glyph onto a new line.
    colDxa = tracks.map((track) => context.pxToTwips(track.widthPx))
    const sum = colDxa.reduce((a, b) => a + b, 0) || 1
    if (sum > generator.avail) {
      const contentSum = tracks.reduce((acc, track, i) => (track.gap ? acc : acc + colDxa[i]), 0)
      const gapSum = sum - contentSum
      if (contentSum <= generator.avail && gapSum > 0) {
        const gapScale = Math.max(0, generator.avail - contentSum) / gapSum
        colDxa = tracks.map((track, i) =>
          track.gap ? Math.max(0, Math.round(colDxa[i] * gapScale)) : colDxa[i],
        )
      } else {
        const scale = generator.avail / sum
        colDxa = colDxa.map((w) => Math.max(40, Math.round(w * scale)))
      }
    }
  } else {
    const totalPx = tracks.reduce((sum, track) => sum + track.widthPx, 0) || 1
    colDxa = tracks.map((track) => Math.round((track.widthPx / totalPx) * generator.avail))
  }
  // Single-line cells must not wrap mid-text ("11." TOC numbers splitting
  // into a phantom item, "Sub Total :" labels breaking in half). Widen each
  // such cell to its browser-measured line width (+10% Arial delta + cell
  // margins; tiny cells get a flat allowance since margins eat their width),
  // borrowing first from flexible gap tracks, then from the widest sibling —
  // always conserving the row's total width so no column clips off-page.
  const oneLineNeedDxa = (track) => {
    if (track.gap || !track.cell) return 0
    const needPx = Math.max(
      0,
      ...(track.cell.children || []).map((child) => child.style?.oneLineWidthPx || 0),
    )
    const cellText = (track.cell.children || [])
      .flatMap((child) => child.runs || [])
      .map((run) => run.text || '')
      .join('')
      .trim()
    // Tiny cells ("11.", "$820") get a larger flat allowance: default cell
    // margins (~216 dxa) eat most of their measured width.
    const marginAllowance = cellText.length > 0 && cellText.length <= 5 ? 420 : 260
    return needPx ? context.pxToTwips(Math.round(needPx * 1.1)) + marginAllowance : 0
  }
  tracks.forEach((track, index) => {
    const needDxa = oneLineNeedDxa(track)
    if (process.env.H2D_DEBUG_KPI && needDxa)
      console.error(
        '[kpi]',
        JSON.stringify({
          index,
          needDxa,
          colDxa,
          avail: generator.avail,
          text: (track.cell?.children || [])
            .flatMap((c) => c.runs || [])
            .map((r) => r.text)
            .join('')
            .slice(0, 20),
        }),
      )
    if (!needDxa || needDxa <= colDxa[index]) return
    let deficit = needDxa - colDxa[index]
    for (let g = 0; g < tracks.length && deficit > 0; g++) {
      if (!tracks[g].gap) continue
      const take = Math.min(deficit, Math.max(0, colDxa[g] - 60))
      colDxa[g] -= take
      deficit -= take
    }
    if (deficit > 0) {
      // Grow the row into unused container width before squeezing siblings —
      // the browser layout had that free space around the flex row anyway.
      const rowSum = colDxa.reduce((a, b) => a + b, 0)
      deficit -= Math.min(deficit, Math.max(0, generator.avail - rowSum))
    }
    if (deficit > 0) {
      let widest = -1
      tracks.forEach((sibling, j) => {
        if (sibling.gap || !sibling.cell || j === index) return
        if (widest < 0 || colDxa[j] > colDxa[widest]) widest = j
      })
      if (widest >= 0) {
        const siblingFloor = Math.max(600, oneLineNeedDxa(tracks[widest]))
        const take = Math.min(deficit, Math.max(0, colDxa[widest] - siblingFloor))
        colDxa[widest] -= take
        deficit -= take
      }
    }
    colDxa[index] += needDxa - colDxa[index] - Math.max(0, deficit)
  })
  const output = []
  if (node.spacingBeforePx) {
    output.push(spacerParagraph(context, node.spacingBeforePx, { shading: node.ambientShading }))
  }
  output.push(
    new Table({
      borders: NO_TABLE_BORDERS,
      rows: [
        new TableRow({
          cantSplit: true,
          // ATLEAST, not EXACT: Word wraps text at slightly different points
          // than the browser; an exact browser-measured height clips the
          // overflowing lines (truncated dates / vanishing second lines).
          height: node.heightPx
            ? {
                value: Math.min(
                  context.pxToTwips(node.heightPx),
                  context.pageHeightDxa -
                    context.pageMargins.top -
                    context.pageMargins.bottom -
                    900,
                ),
                rule: HeightRule.ATLEAST,
              }
            : undefined,
          children: tracks.map((track, ci) => {
            // row-level left accent bar (border-left on the source element)
            const accent =
              ci === 0 && node.rowBorderLeft
                ? {
                    style: BorderStyle.SINGLE,
                    size: context.pxToBorderEighths(node.rowBorderLeft.widthPx),
                    color: node.rowBorderLeft.color,
                  }
                : null
            // full border on the row container (amount boxes): frame the row
            // — top/bottom across every track, left/right on the outer cells
            const frame = node.rowBorder
              ? {
                  style: BorderStyle.SINGLE,
                  size: Math.max(6, context.pxToBorderEighths(node.rowBorder.widthPx)),
                  color: node.rowBorder.color,
                }
              : null
            const frameBorders = (base) => {
              if (!frame) return base
              const out = { ...base, top: frame, bottom: frame }
              if (ci === 0) out.left = frame
              if (ci === tracks.length - 1) out.right = frame
              return out
            }
            if (track.gap) {
              return new TableCell({
                children: [new Paragraph({})],
                width: { size: colDxa[ci], type: WidthType.DXA },
                shading: node.gapShading
                  ? { type: ShadingType.CLEAR, fill: node.gapShading }
                  : undefined,
                borders: frameBorders(accent ? { ...NO_BORDERS, left: accent } : NO_BORDERS),
                margins: { top: 0, bottom: 0, left: 0, right: 0 },
              })
            }
            const cell = track.cell
            // A lone full-width card would nest a table sized to its own
            // content, losing the flex-stretch equal heights of its siblings
            // (day columns). Hoist its paint onto the cell — Word equalizes
            // cell heights within the row.
            const kids = cell.children || []
            const loneCard =
              kids.length === 1 &&
              kids[0].type === 'card' &&
              !kids[0].widthPx &&
              !kids[0].borderLeft &&
              !kids[0].borderRight &&
              !kids[0].borderTop &&
              !cell.shading
                ? kids[0]
                : null
            const cellEntry = loneCard ? { ...cell, children: loneCard.children } : cell
            let borders: any = { ...NO_BORDERS }
            if (accent) borders.left = accent
            const fullBorder = loneCard?.border || cell.border
            if (fullBorder) {
              const b = {
                style: BorderStyle.SINGLE,
                size: context.pxToBorderEighths(fullBorder.widthPx),
                color: fullBorder.color,
              }
              borders.top = b
              borders.bottom = b
              borders.left = b
              borders.right = b
            }
            if (cell.borderTop) {
              borders.top = {
                style: BorderStyle.SINGLE,
                size: context.pxToBorderEighths(cell.borderTop.widthPx),
                color: cell.borderTop.color,
              }
            }
            borders = frameBorders(borders)
            const cellMargins = loneCard
              ? {
                  top: context.pxToTwips(loneCard.padPx?.top ?? 8),
                  bottom: context.pxToTwips(loneCard.padPx?.bottom ?? 8),
                  left: context.pxToTwips(loneCard.padPx?.left ?? 8),
                  right: context.pxToTwips(loneCard.padPx?.right ?? 8),
                }
              : {
                  top: context.pxToTwips(node.padPx?.top || 0),
                  bottom: context.pxToTwips(node.padPx?.bottom || 0),
                  // With explicit gap tracks the cell width equals the child's
                  // rendered box; side margins would steal ~19px and wrap short
                  // labels ("Position:" -> "Positio / n:"). Gap columns already
                  // provide the separation.
                  left: node.compact || node.explicitGaps ? 20 : 140,
                  right: node.compact || node.explicitGaps ? 20 : 140,
                }
            return new TableCell({
              children: cellChildren(
                generator,
                cellEntry,
                depth,
                colDxa[ci] - (loneCard ? cellMargins.left + cellMargins.right : 0),
              ),
              width: { size: colDxa[ci], type: WidthType.DXA },
              shading:
                (loneCard?.shading ?? cell.shading)
                  ? { type: ShadingType.CLEAR, fill: loneCard?.shading ?? cell.shading }
                  : undefined,
              borders,
              verticalAlign:
                node.vAlign === 'center' || cell.vAlign === 'center'
                  ? VerticalAlign.CENTER
                  : undefined,
              margins: cellMargins,
            })
          }),
        }),
      ],
      width: { size: colDxa.reduce((a, b) => a + b, 0), type: WidthType.DXA },
      columnWidths: colDxa,
      layout: TableLayoutType.FIXED,
    }),
    spacerParagraph(context, Math.max(1, node.spacingAfterPx ?? 8), {
      shading: node.ambientShading,
    }),
  )
  return output
}

export { cellChildren, renderCard, renderColorBar, renderKpiRow, renderPageColumns, renderTable }
