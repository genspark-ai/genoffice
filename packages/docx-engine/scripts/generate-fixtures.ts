import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { buildDocx, buildKitchenSinkDocx } from '../tests/helpers/build-docx'

const outDir = join(dirname(fileURLToPath(import.meta.url)), '../../../fixtures/generated')

const TABLE_BORDERS =
  '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>' +
  '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
  '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>'

/** bordered table with explicit dxa grid columns and deterministic cell text */
function capTableXml(options: {
  tblPrXml: string
  colWidths: number[]
  headers: string[]
  dataRows: number
}): string {
  const { tblPrXml, colWidths, headers, dataRows } = options
  const grid = colWidths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')
  const cell = (w: number, text: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>` +
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`
  const rows = [
    `<w:tr>${headers.map((h, c) => cell(colWidths[c], h)).join('')}</w:tr>`,
    ...Array.from(
      { length: dataRows },
      (_, r) => `<w:tr>${colWidths.map((w, c) => cell(w, `val-${r}-${c}`)).join('')}</w:tr>`,
    ),
  ].join('')
  return `<w:tbl><w:tblPr>${tblPrXml}${TABLE_BORDERS}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`
}

/**
 * Visual-regression corpus (e2e/docs-visual.spec.ts): the cap-table shapes
 * from the floating-table overflow fix (#1111). All widths overshoot the
 * A4 content width (9026 dxa) on purpose — the renderer must clamp them.
 */
async function buildVisualCorpus(): Promise<Array<[string, Uint8Array]>> {
  const eightCols = [2400, 1600, 1800, 1600, 2000, 1800, 2000, 1600]
  const captableHeaders = [
    'Shareholder',
    'Class',
    'Shares',
    'Price',
    'Invested',
    'Ownership %',
    'Fully Diluted %',
    'Board Seats',
    'Liquidation Pref',
    'Anti-dilution',
    'Vesting',
    'Notes',
  ]
  const wrapText =
    'The parties agree that the capitalization set out above reflects all issued and ' +
    'outstanding shares on a fully diluted basis as of the closing date.'

  const inline = await buildDocx({
    bodyXml:
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Series B Capitalization Table</w:t></w:r></w:p>' +
      capTableXml({
        tblPrXml: '<w:tblW w:w="23400" w:type="dxa"/>',
        colWidths: [2400, 1600, 1800, 1600, 2000, 1800, 2000, 1600, 2200, 2000, 1800, 2600],
        headers: captableHeaders,
        dataRows: 3,
      }) +
      `<w:p><w:r><w:t xml:space="preserve">${wrapText}</w:t></w:r></w:p>`,
  })

  const float = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Exhibit A — Ownership Summary</w:t></w:r></w:p>' +
      capTableXml({
        tblPrXml:
          '<w:tblW w:w="14800" w:type="dxa"/>' +
          '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" ' +
          'w:horzAnchor="page" w:tblpX="4000" w:tblpY="200"/>',
        colWidths: eightCols,
        headers: eightCols.map((_, c) => `Col${c}`),
        dataRows: 2,
      }) +
      `<w:p><w:r><w:t xml:space="preserve">${wrapText}</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t xml:space="preserve">${wrapText}</w:t></w:r></w:p>`,
  })

  // vertAnchor="page": the table must land at 6000 twips from the page top
  // (≈400 css px) regardless of where its anchor paragraph sits in the flow
  const pageAnchor = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Exhibit C — Page-Anchored Position</w:t></w:r></w:p>' +
      `<w:p><w:r><w:t xml:space="preserve">${wrapText}</w:t></w:r></w:p>` +
      capTableXml({
        tblPrXml:
          '<w:tblW w:w="9026" w:type="dxa"/>' +
          '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" ' +
          'w:horzAnchor="margin" w:tblpY="6000"/>',
        colWidths: [2306, 2240, 2240, 2240],
        headers: ['Holder', 'Class', 'Shares', 'Notes'],
        dataRows: 2,
      }) +
      `<w:p><w:r><w:t xml:space="preserve">${wrapText}</w:t></w:r></w:p>`,
  })

  const pct = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Exhibit B — Percentage Width</w:t></w:r></w:p>' +
      capTableXml({
        tblPrXml: '<w:tblW w:w="7500" w:type="pct"/>',
        colWidths: eightCols,
        headers: eightCols.map((_, c) => `Col${c}`),
        dataRows: 2,
      }),
  })

  return [
    ['visual-captable-inline.docx', inline],
    ['visual-captable-float.docx', float],
    ['visual-captable-pct.docx', pct],
    ['visual-captable-page-anchor.docx', pageAnchor],
  ]
}

async function main() {
  mkdirSync(outDir, { recursive: true })

  const kitchenSink = await buildKitchenSinkDocx()
  writeFileSync(join(outDir, 'kitchen-sink.docx'), kitchenSink)

  const simple = await buildDocx({
    bodyXml:
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>第一段。</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>第二段。</w:t></w:r></w:p>',
  })
  writeFileSync(join(outDir, 'simple.docx'), simple)

  for (const [name, bytes] of await buildVisualCorpus()) {
    writeFileSync(join(outDir, name), bytes)
  }

  console.log(`fixtures written to ${outDir}`)
}

main().catch((err) => {
  console.error(err)
  process.exit(1)
})
