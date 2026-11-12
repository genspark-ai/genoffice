/**
 * The column grid a table saves from the editor: the first row's dragged column
 * widths (px) when every column has one, else the grid the document came with,
 * else an even split.
 */
export function gridFromEditorColumns(
  explicitPx: number[],
  columns: number,
  declaredPct: number[] | undefined,
): { colWidthsPct: number[] | undefined; colWidthsTwips: number[] | undefined } {
  if (explicitPx.length === columns && explicitPx.every((width) => width > 0)) {
    const total = explicitPx.reduce((sum, width) => sum + width, 0)
    return {
      colWidthsPct: explicitPx.map((width) => (width / total) * 100),
      colWidthsTwips: explicitPx.map((width) => Math.round(width * 15)),
    }
  }
  if (declaredPct?.length !== columns)
    return {
      colWidthsPct: Array.from({ length: columns }, () => 100 / columns),
      colWidthsTwips: undefined,
    }
  return { colWidthsPct: declaredPct, colWidthsTwips: undefined }
}
