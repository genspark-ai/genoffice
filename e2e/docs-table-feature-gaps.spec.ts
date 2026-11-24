/**
 * Table commands where Word does more than the first version did: Convert
 * Text to Table keeps paragraph formatting and is reviewed as tracked
 * changes, Insert Cells shifts beside a merged cell and a rejected cell moves
 * the row back, and a formula reads a bookmark by the text it covers. Each is
 * driven through the UI, read off the page, and read back from the saved file.
 */
import { test, expect, type Page } from '@playwright/test'
import { saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  ribbon,
  ribbonTab,
  savedBody,
  selectParagraph,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-table-feature-gaps'

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const p = (content: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.startsWith('<') ? content : run(content)}</w:p>`
const tc = (text: string, tcPr = '', width = 3000) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${tcPr}</w:tcPr>${p(text)}</w:tc>`
const BORDERS =
  '<w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join('') +
  '</w:tblBorders>'

/** the top-level blocks: a paragraph's text, or a table's rows of cell texts */
const blocks = (page: Page) =>
  page.evaluate(() => {
    const { editor } = (window as never as { __aidocs: { editor: any } }).__aidocs
    const out: Array<string | string[][]> = []
    editor.state.doc.forEach((node: any) => {
      if (node.type.name !== 'docTable') {
        out.push(node.textContent)
        return
      }
      const rows: string[][] = []
      node.forEach((row: any) => {
        const cells: string[] = []
        row.forEach((cell: any) => {
          if (!cell.attrs.gridGap) cells.push(cell.textContent)
        })
        rows.push(cells)
      })
      out.push(rows)
    })
    return out
  })

const trackChangesOn = async (page: Page) => {
  await ribbonTab(page, 'Review')
  await ribbon(page, 'Track Changes').click()
  await expect(page.locator('.status-track')).toHaveText('Track Changes: On')
}

const rejectAll = async (page: Page) => {
  await ribbonTab(page, 'Review')
  await page
    .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Reject$/ })
    .click()
  await page
    .locator('[data-rb-panel] button', { hasText: /^Reject All Changes$/ })
    .first()
    .click()
}

test('Convert Text to Table keeps each paragraph formatted, and tracked it is rejected back to the text', async () => {
  test.setTimeout(240_000)
  const s = await open(
    VIDEO_DIR,
    p('Item,Qty', '<w:jc w:val="center"/><w:ind w:left="360"/>') + p('Apple,3') + p('tail'),
    'convert.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    await trackChangesOn(page)
    await selectParagraph(page, 'Item,Qty', 'Apple,3')
    await clickMenu(app, ['Table', 'Convert', 'Convert Text to Table…'])
    const dialog = page.locator('.convert-text-dialog')
    await expect(dialog.locator('input[name="cols"]')).toHaveValue('2')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    // the text stays as a deletion, in its order, in front of the table that came in
    await expect
      .poll(() => blocks(page))
      .toEqual([
        'Item,Qty',
        'Apple,3',
        [
          ['Item', 'Qty'],
          ['Apple', '3'],
        ],
        'tail',
      ])
    const cellParagraph = (text: string) =>
      page.locator('.doc-page table :is(td, th) p', { hasText: new RegExp(`^${text}$`) })
    const align = (text: string) =>
      cellParagraph(text).evaluate((el) => getComputedStyle(el).textAlign)
    expect(await align('Item')).toBe('center')
    expect(await align('Qty')).toBe('center')
    expect(await align('Apple')).toMatch(/^(left|start)$/)
    await page.screenshot({ path: test.info().outputPath('converted-tracked.png') })

    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).toContain('<w:tbl>')
    const xml = await savedBody(docPath)
    const table = /<w:tbl>[\s\S]*<\/w:tbl>/.exec(xml)![0]
    const cells = table.match(/<w:tc>[\s\S]*?<\/w:tc>/g)!
    expect(cells).toHaveLength(4)
    for (const cell of cells.slice(0, 2)) {
      expect(cell).toContain('<w:jc w:val="center"/>')
      expect(cell).toMatch(/<w:ind [^>]*w:left="360"/)
    }
    for (const cell of cells.slice(2)) expect(cell).not.toContain('<w:jc ')
    expect(xml.indexOf('Item,Qty')).toBeLessThan(xml.indexOf('Apple,3'))
    expect(xml.indexOf('Apple,3')).toBeLessThan(xml.indexOf('<w:tbl>'))

    await rejectAll(page)
    await expect.poll(() => blocks(page)).toEqual(['Item,Qty', 'Apple,3', 'tail'])
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).not.toContain('<w:tbl>')
    expect(await savedBody(docPath)).not.toMatch(/<w:(?:ins|del) /)
  } finally {
    await s.close()
  }
})

test('Insert Cells shifts a row beside a merged heading, and rejecting the cell moves the row back', async () => {
  test.setTimeout(240_000)
  const table =
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid>` +
    '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${tc('Parts', '<w:gridSpan w:val="2"/>', 6000)}</w:tr>` +
    `<w:tr>${tc('A2')}${tc('B2')}</w:tr><w:tr>${tc('A3')}${tc('B3')}</w:tr></w:tbl>`
  const s = await open(VIDEO_DIR, p('Stock') + table + p('tail'), 'shift.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    await trackChangesOn(page)
    const dialog = page.locator('.insert-cells-dialog')
    const insert = async (at: string, option: string, shot?: string) => {
      await caretAtEnd(page, at)
      await clickMenu(app, ['Table', 'Insert', 'Cells…'])
      await expect(dialog.getByLabel(option)).toBeEnabled()
      await dialog.getByLabel(option).check()
      if (shot) await page.screenshot({ path: test.info().outputPath(shot) })
      await dialog.locator('button.primary, button[type=submit]').first().click()
      await expect(dialog).toHaveCount(0)
    }
    const rows = async () => (await blocks(page))[1]

    await insert('A2', 'Shift cells right', 'insert-cells-merged.png')
    expect(await rows()).toEqual([['Parts'], ['', 'A2', 'B2'], ['A3', 'B3']])
    await expect(page.locator('.doc-page td.cell-rev-ins')).toHaveCount(1)
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).toContain('<w:cellIns ')
    let xml = await savedBody(docPath)
    expect(xml).toContain('<w:gridSpan w:val="2"/>')
    expect(xml.match(/<w:gridCol /g)).toHaveLength(3)

    await rejectAll(page)
    await expect.poll(rows).toEqual([['Parts'], ['A2', 'B2'], ['A3', 'B3']])
    await expect(page.locator('.doc-page td.cell-rev-ins')).toHaveCount(0)
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).not.toContain('<w:cellIns ')
    xml = await savedBody(docPath)
    expect(xml).not.toContain('gridAfter')
    expect(xml.match(/<w:gridCol /g)).toHaveLength(2)

    // downward: the column moves back up and the row the insertion added goes
    await insert('A2', 'Shift cells down')
    expect(await rows()).toEqual([['Parts'], ['', 'B2'], ['A2', 'B3'], ['A3', '']])
    await rejectAll(page)
    await expect.poll(rows).toEqual([['Parts'], ['A2', 'B2'], ['A3', 'B3']])
    await page.screenshot({ path: test.info().outputPath('shift-rejected.png') })
  } finally {
    await s.close()
  }
})

test('a formula reads a bookmark by the text it covers, and a bookmark named like a cell outside a table', async () => {
  test.setTimeout(240_000)
  const marked = (id: number, name: string, text: string) =>
    `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${run(text)}<w:bookmarkEnd w:id="${id}"/>`
  const s = await open(
    VIDEO_DIR,
    p(run('Price 100, discount ') + marked(1, 'Discount', '25') + run(' each')) +
      p(run('Quarter 4 of 2026 sold ') + marked(2, 'Q4', '40') + run(' units')) +
      p('Twice the discount: ') +
      p('One more than the quarter: ') +
      p('tail'),
    'formula.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    const formula = page.locator('.formula-dialog')
    const insert = async (at: string, text: string) => {
      await caretAtEnd(page, at)
      await clickMenu(app, ['Table', 'Formula…'])
      await formula.locator('input[name="formula"]').fill(text)
      await formula.locator('button.primary').click()
      await expect(formula).toBeHidden()
    }
    // the bookmark covers 25, not the paragraph's first number
    await insert('Twice the discount: ', '=Discount*2')
    await expect(page.locator('.doc-page p', { hasText: 'Twice the discount' })).toHaveText(
      'Twice the discount: 50',
    )
    // no table around it: Q4 is the bookmark
    await insert('One more than the quarter: ', '=Q4+1')
    await expect(page.locator('.doc-page p', { hasText: 'One more than' })).toHaveText(
      'One more than the quarter: 41',
    )
    await page.screenshot({ path: test.info().outputPath('formula-bookmarks.png') })
    await saveFromMenu(app)
    await expect
      .poll(() => savedBody(docPath), { timeout: 30_000 })
      .toMatch(/<w:instrText[^>]*> ?=Discount\*2 ?<\/w:instrText>[\s\S]*?<w:t[^>]*>50<\/w:t>/)
    expect(await savedBody(docPath)).toMatch(
      /<w:instrText[^>]*> ?=Q4\+1 ?<\/w:instrText>[\s\S]*?<w:t[^>]*>41<\/w:t>/,
    )
  } finally {
    await s.close()
  }
})
