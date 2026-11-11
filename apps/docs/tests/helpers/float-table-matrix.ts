/** synthetic w:tblpPr matrix shared by the anchor and float-table geometry goldens */
const text = (s: string) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`

export function floatTableMatrixXml(): string {
  const tbl = (tblp: string, w = 4000, nested = '') =>
    `<w:p>${text('anchor paragraph with enough words to wrap beside the floating table on the side')}</w:p>` +
    `<w:tbl><w:tblPr><w:tblpPr ${tblp}/><w:tblW w:w="${w}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${w}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${nested}<w:p>${text('cell')}</w:p></w:tc></w:tr></w:tbl>`
  const out: string[] = []
  for (const h of ['text', 'margin', 'page']) {
    for (const v of ['text', 'margin', 'page']) {
      out.push(
        tbl(
          `w:leftFromText="180" w:rightFromText="180" w:topFromText="120" w:vertAnchor="${v}" w:horzAnchor="${h}" w:tblpX="1200" w:tblpY="600"`,
        ),
      )
      out.push(tbl(`w:vertAnchor="${v}" w:horzAnchor="${h}" w:tblpX="6000" w:tblpY="-300"`))
      for (const xs of ['left', 'center', 'right', 'inside', 'outside']) {
        out.push(tbl(`w:vertAnchor="${v}" w:horzAnchor="${h}" w:tblpXSpec="${xs}" w:tblpY="900"`))
      }
      for (const ys of ['top', 'center', 'bottom', 'inside', 'outside']) {
        out.push(tbl(`w:vertAnchor="${v}" w:horzAnchor="${h}" w:tblpYSpec="${ys}" w:tblpX="2000"`))
      }
    }
  }
  out.push(tbl('w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="0" w:tblpY="0"', 9000))
  out.push(tbl('w:vertAnchor="text" w:horzAnchor="page" w:tblpX="400" w:tblpY="200"', 9800))
  out.push(tbl('w:vertAnchor="text" w:horzAnchor="text" w:tblpX="8000" w:tblpY="200"', 3000))
  const inner =
    '<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="300" w:tblpY="400"/><w:tblW w:w="1500" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p>' +
    text('n') +
    '</w:p></w:tc></w:tr></w:tbl>'
  out.push(
    tbl('w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="500" w:tblpY="300"', 6000, inner),
  )
  return out.join('') + '<w:p/>'
}

/** keyword / over-wide / nested-keyword cases the stage-4 flips move */
export function floatTableFlipMatrixXml(): string {
  const tbl = (tblp: string, w: number, nested = '') =>
    `<w:p>${text('Internationalization anchor paragraph words that keep going beside the float')}</w:p>` +
    `<w:tbl><w:tblPr><w:tblpPr ${tblp}/><w:tblW w:w="${w}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${w}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${nested}<w:p>${text('cell')}</w:p></w:tc></w:tr></w:tbl>`
  const out: string[] = []
  for (const h of ['text', 'margin', 'page'])
    for (const xs of ['left', 'center', 'right', 'inside', 'outside'])
      for (const w of [3000, 11000])
        out.push(tbl(`w:horzAnchor="${h}" w:tblpXSpec="${xs}" w:vertAnchor="text" w:tblpY="1"`, w))
  for (const w of [7320, 8520, 9120])
    out.push(
      tbl(
        `w:horzAnchor="margin" w:tblpX="0" w:vertAnchor="text" w:tblpY="1" w:rightFromText="180"`,
        w,
      ),
    )
  const inner = (tblp: string) =>
    `<w:tbl><w:tblPr><w:tblpPr ${tblp}/><w:tblW w:w="3000" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${text('n')}</w:p></w:tc></w:tr></w:tbl>`
  for (const h of ['margin', 'page'])
    for (const xs of ['center', 'right'])
      out.push(
        `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
          `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${inner(`w:horzAnchor="${h}" w:tblpXSpec="${xs}" w:vertAnchor="text" w:tblpY="1"`)}<w:p>${text('cell words beside the nested float')}</w:p></w:tc></w:tr></w:tbl><w:p/>`,
      )
  return out.join('') + '<w:p/>'
}
