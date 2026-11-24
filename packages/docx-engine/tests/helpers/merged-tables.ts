/**
 * Pairs of tables for comparing by column: a grid with cells merged across
 * and down, and the same grid with a column put in or taken out.
 */

/** cells merged into one: rows r0..r1 and columns c0..c1, both ends included */
export interface Merge {
  r0: number
  r1: number
  c0: number
  c1: number
}

export interface Grid {
  /** text per row and column; a merged cell reads as its top-left text */
  texts: string[][]
  merges: Merge[]
}

/** what a cell is, as far as comparing goes: its text, the columns it covers, its part in a vertical merge */
export type CellShape = [text: string, span: number, merge: 'restart' | 'continue' | null]

export const SHAPES: Record<string, Merge[]> = {
  'a heading cell across the first two columns': [{ r0: 0, r1: 0, c0: 0, c1: 1 }],
  'a cell across the last two columns of a middle row': [{ r0: 1, r1: 1, c0: 2, c1: 3 }],
  'a cell down two rows of the first column': [{ r0: 1, r1: 2, c0: 0, c1: 0 }],
  'a cell down the whole last column': [{ r0: 0, r1: 3, c0: 3, c1: 3 }],
  'a block two columns wide and two rows high': [{ r0: 1, r1: 2, c0: 1, c1: 2 }],
  'a cell across and another down': [
    { r0: 0, r1: 0, c0: 0, c1: 1 },
    { r0: 2, r1: 3, c0: 3, c1: 3 },
  ],
}

const WORDS = [
  ['Item', 'Stock', 'Price', 'Origin'],
  ['Bolt', 'many', 'ten dollars', 'Leeds'],
  ['Nut', 'few', 'two dollars', 'Turin'],
  ['Washer', 'none', 'one dollar', 'Lyon'],
]

export const baseGrid = (merges: Merge[]): Grid => ({
  texts: WORDS.map((row) => [...row]),
  merges: merges.map((merge) => ({ ...merge })),
})

/** the grid with a new column standing at `at`; `reaches` when a merged cell grows across it */
export function withColumn(grid: Grid, at: number): { grid: Grid; reaches: boolean } {
  let reaches = false
  const merges = grid.merges.map((merge) => {
    if (merge.c0 >= at) return { ...merge, c0: merge.c0 + 1, c1: merge.c1 + 1 }
    if (merge.c1 < at) return { ...merge }
    reaches = true
    return { ...merge, c1: merge.c1 + 1 }
  })
  const texts = grid.texts.map((row, r) => [
    ...row.slice(0, at),
    `Added ${['kind', 'steel', 'brass', 'zinc'][r]}`,
    ...row.slice(at),
  ])
  return { grid: { texts, merges }, reaches }
}

/** the grid without column `at`; `reaches` when a merged cell loses one of its columns */
export function withoutColumn(grid: Grid, at: number): { grid: Grid; reaches: boolean } {
  let reaches = false
  const merges: Merge[] = []
  for (const merge of grid.merges) {
    if (merge.c0 > at) merges.push({ ...merge, c0: merge.c0 - 1, c1: merge.c1 - 1 })
    else if (merge.c1 < at) merges.push({ ...merge })
    else if (merge.c0 !== merge.c1) {
      reaches = true
      merges.push({ ...merge, c1: merge.c1 - 1 })
    }
  }
  const texts = grid.texts.map((row) => row.filter((_, c) => c !== at))
  return { grid: { texts, merges }, reaches }
}

/** the cells of each row as the file states them */
export function gridShape(grid: Grid): CellShape[][] {
  return grid.texts.map((row, r) => {
    const cells: CellShape[] = []
    for (let c = 0; c < row.length; c++) {
      const merge = grid.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1)
      if (!merge) {
        cells.push([row[c]!, 1, null])
        continue
      }
      if (c !== merge.c0) continue
      const down = merge.r1 > merge.r0
      cells.push([
        r === merge.r0 ? row[c]! : '',
        merge.c1 - merge.c0 + 1,
        !down ? null : r === merge.r0 ? 'restart' : 'continue',
      ])
    }
    return cells
  })
}

const COL = 1800

export function gridXml(grid: Grid): string {
  const cols = grid.texts[0]!.length
  const rows = gridShape(grid).map(
    (row) =>
      `<w:tr>${row
        .map(
          ([text, span, merge]) =>
            `<w:tc><w:tcPr><w:tcW w:w="${COL * span}" w:type="dxa"/>` +
            (span > 1 ? `<w:gridSpan w:val="${span}"/>` : '') +
            (merge ? `<w:vMerge${merge === 'restart' ? ' w:val="restart"' : ''}/>` : '') +
            `</w:tcPr><w:p>${text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ''}</w:p></w:tc>`,
        )
        .join('')}</w:tr>`,
  )
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
      .join('') +
    `</w:tblBorders></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${COL}"/>`.repeat(cols)}</w:tblGrid>${rows.join('')}</w:tbl>`
  )
}

/** the cells of each row of the first table in `xml`, read back from the file */
export function xmlShape(xml: string): CellShape[][] {
  const table = /<w:tbl>[\s\S]*?<\/w:tbl>/.exec(xml)?.[0] ?? ''
  return [...table.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((row) =>
    [...row[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((cell): CellShape => {
      const props = /<w:tcPr>[\s\S]*?<\/w:tcPr>/.exec(cell[0])?.[0] ?? ''
      const merge = /<w:vMerge\b([^>]*)>/.exec(props)
      return [
        [...cell[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
        Number(/<w:gridSpan\b[^>]*w:val="(\d+)"/.exec(props)?.[1] ?? 1),
        !merge ? null : /w:val="restart"/.test(merge[1]!) ? 'restart' : 'continue',
      ]
    }),
  )
}

/** the grid with a new row standing at `at`; a merge that runs through the place takes the row in */
export function withRow(grid: Grid, at: number): Grid {
  const merges = grid.merges.map((merge) =>
    merge.r0 >= at
      ? { ...merge, r0: merge.r0 + 1, r1: merge.r1 + 1 }
      : merge.r1 < at
        ? { ...merge }
        : { ...merge, r1: merge.r1 + 1 },
  )
  const row = grid.texts[0]!.map(
    (_, c) => `Extra ${['gasket', 'plenty', 'six dollars', 'Graz', 'iron'][c]}`,
  )
  return { texts: [...grid.texts.slice(0, at), row, ...grid.texts.slice(at)], merges }
}

/** the grid without row `at`; a merge keeps the rows it has left */
export function withoutRow(grid: Grid, at: number): Grid {
  const merges: Merge[] = []
  for (const merge of grid.merges) {
    if (merge.r0 > at) merges.push({ ...merge, r0: merge.r0 - 1, r1: merge.r1 - 1 })
    else if (merge.r1 < at) merges.push({ ...merge })
    else if (merge.r0 !== merge.r1) merges.push({ ...merge, r1: merge.r1 - 1 })
  }
  return { texts: grid.texts.filter((_, r) => r !== at), merges }
}
