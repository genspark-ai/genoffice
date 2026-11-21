/** Word acceptance tasks, inserted objects: each walked along Word's path and read back from the saved file. */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { para, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  part,
  pickFile,
  ribbon,
  ribbonTab,
  savedBody,
  savedParagraph,
  selectParagraph,
  type Session,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-objects'
// 120 x 80 px, one colour
const PICTURE =
  'iVBORw0KGgoAAAANSUhEUgAAAHgAAABQCAIAAABd+SbeAAAAy0lEQVR4nO3SQQ0AIAADMUANEpGILFRwr9bAksvmPnfw3wo2ELrj0RGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IPRoPXj4CCNvpCHkAAAAASUVORK5CYII='

const cell = (page: Page, text: string) =>
  page.locator('.doc-page td', { hasText: new RegExp(`^${text}$`) })

/** the rows of the first saved table, each a list of its cells' texts */
const savedRows = async (docPath: string) =>
  ((await savedBody(docPath)).match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) ?? []).map((row) =>
    (row.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []).map((tc) =>
      (tc.match(/<w:t(?: [^>]*)?>[^<]*/g) ?? []).map((t) => t.replace(/<[^>]*>/, '')).join(''),
    ),
  )

test('Insert > Table makes a 3 x 3 table that Tab walks and extends; Table Layout adds a row, deletes a column and merges cells; Table Design applies a style and shading', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, ['before', 'after'].map(para).join(''), 'table.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'before' })).toBeVisible({
      timeout: 30_000,
    })
    await caretAtEnd(page, 'before')

    await ribbonTab(page, 'Insert')
    await ribbon(page, 'Table').click()
    // the picker is ten cells wide: this one is the third of its third row
    await page.locator('.table-picker-grid .table-cell').nth(22).click()
    const rows = page.locator('.doc-page table tr')
    await expect(rows).toHaveCount(3)
    await expect(page.locator('.doc-page table td')).toHaveCount(9)

    // Tab moves on from cell to cell, and in the last cell adds a row
    await page.locator('.doc-page table td').first().click()
    const texts = ['a1', 'b1', 'c1', 'a2', 'b2', 'c2', 'a3', 'b3', 'c3']
    for (const text of texts) {
      await page.keyboard.type(text)
      if (text !== 'c3') await page.keyboard.press('Tab')
    }
    await expect(page.locator('.doc-page table td')).toHaveText(texts)
    await page.keyboard.press('Tab')
    await expect(rows).toHaveCount(4)
    await page.keyboard.type('a4')
    await saveFromMenu(app)
    await expect
      .poll(() => savedRows(docPath), { timeout: 30_000 })
      .toEqual([
        ['a1', 'b1', 'c1'],
        ['a2', 'b2', 'c2'],
        ['a3', 'b3', 'c3'],
        ['a4', '', ''],
      ])

    await ribbonTab(page, 'Table Layout')
    await cell(page, 'a3').click()
    await ribbon(page, 'Insert Below').click()
    await expect(rows).toHaveCount(5)
    await cell(page, 'c2').click()
    await ribbon(page, 'Delete').click()
    await page.locator('.layout-menu button', { hasText: 'Delete Columns' }).click()
    await expect(page.locator('.doc-page table tr').first().locator('td')).toHaveCount(2)
    // whole cells select with a click and a Shift+click; a drag is a text selection
    await expect(async () => {
      await cell(page, 'a1').click()
      await cell(page, 'b1').click({ modifiers: ['Shift'] })
      await expect(ribbon(page, 'Merge Cells')).toBeEnabled({ timeout: 2_000 })
    }).toPass({ timeout: 20_000 })
    await ribbon(page, 'Merge Cells').click()
    await expect(page.locator('.doc-page td[colspan="2"]')).toHaveCount(1)

    await ribbonTab(page, 'Table Design')
    await cell(page, 'b2').click()
    await page.locator('[data-ribbon-body] .table-style-card', { hasText: 'Blue Header' }).click()
    // the style's header shading shows before the save writes the style out
    await expect(cell(page, 'a1b1')).toHaveCSS('background-color', 'rgb(217, 234, 247)')
    await ribbon(page, 'Shading').click()
    await page.locator('.gcp-palette .gcp-swatch[title="Yellow"]').click()
    await expect(cell(page, 'b2')).toHaveCSS('background-color', 'rgb(255, 255, 0)')
    await page.screenshot({ path: test.info().outputPath('table.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const body = await savedBody(docPath)
          const shaded = (body.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []).filter((tc) =>
            /<w:shd\b[^>]*w:fill="FFFF00"/.test(tc),
          )
          return {
            rows: await savedRows(docPath),
            merged: /<w:gridSpan w:val="2"\/>/.test(body),
            columns: (
              /<w:tblGrid>[\s\S]*?<\/w:tblGrid>/.exec(body)?.[0].match(/<w:gridCol\b/g) ?? []
            ).length,
            // the header shading is the style's, not the cells' own
            headerFill:
              /<w:tblStyle w:val="BlueHeader"\/>/.test(body) &&
              /w:styleId="BlueHeader"[\s\S]*?w:type="firstRow"[\s\S]*?w:fill="D9EAF7"/.test(
                await part(docPath, 'word/styles.xml'),
              ),
            borderColour: /<w:top\b[^>]*w:color="5B9BD5"/.test(body),
            shaded: shaded.map((tc) => /<w:t(?: [^>]*)?>([^<]*)/.exec(tc)?.[1]),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({
        rows: [['a1b1'], ['a2', 'b2'], ['a3', 'b3'], ['', ''], ['a4', '']],
        merged: true,
        columns: 2,
        headerFill: true,
        borderColour: true,
        shaded: ['b2'],
      })
  } finally {
    await s.close()
  }
})

test('Insert > Table leaves the caret in the first cell, so typing fills the table from its start', async () => {
  const s = await open(VIDEO_DIR, ['before', 'after'].map(para).join(''), 'table-caret.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'before' })).toBeVisible({
      timeout: 30_000,
    })
    await caretAtEnd(page, 'before')
    await ribbonTab(page, 'Insert')
    await ribbon(page, 'Table').click()
    await page.locator('.table-picker-grid .table-cell').nth(22).click()
    await expect(page.locator('.doc-page table td')).toHaveCount(9)
    await page.keyboard.type('a1')
    await page.keyboard.press('Tab')
    await page.keyboard.type('b1')
    await saveFromMenu(app)
    await expect
      .poll(() => savedRows(docPath), { timeout: 30_000 })
      .toEqual([
        ['a1', 'b1', ''],
        ['', '', ''],
        ['', '', ''],
      ])
  } finally {
    await s.close()
  }
})

test('Table Design > Table Styles saves the picked style as the table’s style', async () => {
  const s = await open(VIDEO_DIR, ['before', 'after'].map(para).join(''), 'table-style.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'before' })).toBeVisible({
      timeout: 30_000,
    })
    await caretAtEnd(page, 'before')
    await ribbonTab(page, 'Insert')
    await ribbon(page, 'Table').click()
    await page.locator('.table-picker-grid .table-cell').nth(22).click()
    await page.locator('.doc-page table td').first().click()
    await ribbonTab(page, 'Table Design')
    await page.locator('[data-ribbon-body] .table-style-card', { hasText: 'Blue Header' }).click()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const styleId = /<w:tblStyle w:val="([^"]+)"\/>/.exec(await savedBody(docPath))?.[1]
          return (
            !!styleId &&
            new RegExp(`<w:style\\b[^>]*w:type="table"[^>]*w:styleId="${styleId}"`).test(
              await part(docPath, 'word/styles.xml'),
            )
          )
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } finally {
    await s.close()
  }
})

/** drags on the page from one point to another, both relative to the first page */
async function dragOnPage(page: Page, from: [number, number], to: [number, number]) {
  const box = (await page.locator('.doc-page').first().boundingBox())!
  await page.mouse.move(box.x + from[0], box.y + from[1])
  await page.mouse.down()
  await page.mouse.move(box.x + (from[0] + to[0]) / 2, box.y + (from[1] + to[1]) / 2, { steps: 5 })
  await page.mouse.move(box.x + to[0], box.y + to[1], { steps: 5 })
  await page.mouse.up()
}

/**
 * Insert > Pictures after the paragraph holding `text`, Picture Format > Wrap Text > Square, then
 * a drag 200 px right and 60 px down, and a save. Gives back how to read the saved anchor: an
 * alignment or an offset in px on each axis.
 */
async function squarePictureDragged(s: Session, text: string) {
  const { app, page, docPath, dir } = s
  const pictureFile = join(dir, 'picture.png')
  writeFileSync(pictureFile, Buffer.from(PICTURE, 'base64'))
  await caretAtEnd(page, text)
  await pickFile(app, pictureFile)
  await ribbonTab(page, 'Insert')
  await ribbon(page, 'Pictures').click()
  const picture = page.locator('.doc-page img[src^="data:image/png"]')
  await expect(picture).toHaveCount(1)
  await picture.click()
  await expect(page.locator('.ribbon-tab.active')).toHaveText('Picture Format')
  await page.locator('[data-ribbon-body] button[data-tip="Wrap Text"]').click()
  await page
    .locator('.layout-menu button', { hasText: /^Square/ })
    .first()
    .click()
  const anchor = async () => {
    const xml = /<wp:anchor\b[\s\S]*?<wp:wrapSquare\b/.exec(await savedBody(docPath))?.[0] ?? ''
    const along = (axis: string) => {
      const position =
        new RegExp(`<wp:position${axis}\\b[\\s\\S]*?</wp:position${axis}>`).exec(xml)?.[0] ?? ''
      const offset = /<wp:posOffset>(-?\d+)</.exec(position)?.[1]
      return offset ? Math.round(Number(offset) / 9525) : /<wp:align>(\w+)</.exec(position)?.[1]
    }
    return { wrapped: xml !== '', x: along('H'), y: along('V') }
  }
  await saveFromMenu(app)
  await expect.poll(anchor, { timeout: 30_000 }).toEqual({ wrapped: true, x: 'left', y: 0 })

  const at = (await picture.boundingBox())!
  const [x, y] = [at.x + at.width / 2, at.y + at.height / 2]
  await page.mouse.move(x, y)
  await page.mouse.down()
  await page.mouse.move(x + 80, y + 30, { steps: 5 })
  await page.mouse.move(x + 200, y + 60, { steps: 5 })
  await page.mouse.up()
  await expect
    .poll(async () => Math.abs((await picture.boundingBox())!.x - at.x - 200))
    .toBeLessThan(3)
  await saveFromMenu(app)
  return anchor
}

test('a square-wrapped picture dragged sideways is saved where it was dropped', async () => {
  const s = await open(VIDEO_DIR, ['first', 'second', 'third'].map(para).join(''), 'picture.docx')
  try {
    await expect(s.page.locator('.doc-page p', { hasText: 'first' })).toBeVisible({
      timeout: 30_000,
    })
    await expect
      .poll(await squarePictureDragged(s, 'first'), { timeout: 30_000 })
      .toEqual({ wrapped: true, x: 200, y: 60 })
  } finally {
    await s.close()
  }
})

test('Insert > Pictures places a picture that takes a wrapping and a drag; Insert > Shapes, Text Box and Chart add a rectangle, a text box and a column chart', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    ['first', 'second', 'third', 'fourth', 'fifth', 'sixth'].map(para).join(''),
    'objects.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'first' })).toBeVisible({ timeout: 30_000 })
    await expect.poll(await squarePictureDragged(s, 'first'), { timeout: 30_000 }).toMatchObject({
      wrapped: true,
      y: 60,
    })

    await ribbonTab(page, 'Insert')
    await ribbon(page, 'Shapes').click()
    await page.locator('[data-rb-panel] button[aria-label="Rectangle"]').click()
    await dragOnPage(page, [150, 600], [330, 700])
    const boxes = page.locator('.doc-page .doc-textbox')
    await expect(boxes).toHaveCount(1)
    await page.locator('.doc-page p', { hasText: 'sixth' }).click()

    await ribbonTab(page, 'Insert')
    await ribbon(page, 'Text Box').click()
    await page.locator('.rb-textbox-menu button', { hasText: /^Draw Text Box$/ }).click()
    await dragOnPage(page, [420, 600], [620, 680])
    await expect(boxes).toHaveCount(2)
    // the caret moves into the new box a moment after it is drawn
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.doc-textbox-editor')))
      .toBe(true)
    await page.keyboard.type('boxed words')
    await expect(boxes.nth(1)).toContainText('boxed words')
    await page.locator('.doc-page p', { hasText: 'sixth' }).click()

    await caretAtEnd(page, 'sixth')
    await ribbonTab(page, 'Insert')
    await ribbon(page, 'Chart').click()
    const chart = page.locator('.modal', { hasText: 'Insert Chart' })
    await expect(chart.locator('button', { hasText: /^Column$/ })).toHaveClass(/btn-primary/)
    await chart.locator('.modal-actions button.btn-primary').click()
    await expect(chart).toBeHidden()
    await page.screenshot({ path: test.info().outputPath('objects.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const body = await zip.file('word/document.xml')!.async('string')
          const charts = await Promise.all(
            Object.keys(zip.files)
              .filter((name) => /^word\/charts\/chart\d+\.xml$/.test(name))
              .map((name) => zip.file(name)!.async('string')),
          )
          const shapes = body.match(/<wps:wsp\b[\s\S]*?<\/wps:wsp>/g) ?? []
          return {
            rectangle: shapes.some(
              (shape) => /<a:prstGeom prst="rect"/.test(shape) && !shape.includes('boxed words'),
            ),
            textBox: shapes.some((shape) => /<wps:txbx>[\s\S]*boxed words/.test(shape)),
            chartInBody: /<c:chart\b[^>]*r:id=/.test(body),
            columnChart: charts.some((xml) => /<c:barDir val="col"\/>/.test(xml)),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ rectangle: true, textBox: true, chartInBody: true, columnChart: true })
  } finally {
    await s.close()
  }
})

test('Insert > Link puts a web address on the selected text; Insert > Advanced Symbol inserts a symbol and a special character', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, ['see the site', 'price'].map(para).join(''), 'link.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'price' })).toBeVisible({ timeout: 30_000 })

    await selectParagraph(page, 'see the site')
    await page.keyboard.press('ControlOrMeta+k')
    const link = page.locator('.modal', { hasText: 'Insert Hyperlink' })
    await expect(link.getByLabel('Text to display')).toHaveValue('see the site')
    await link.getByLabel('Address').fill('https://example.com/report')
    await link.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(link).toBeHidden()
    await expect(page.locator('.doc-page a.doc-link')).toHaveText('see the site')

    await caretAtEnd(page, 'price')
    await clickMenu(app, ['Insert', 'Advanced Symbol…'])
    const symbol = page.locator('.symbol-dialog')
    await symbol
      .locator('.symbol-grid .symbol-cell', { hasText: /^\u20ac$/ })
      .first()
      .click()
    await symbol.locator('button.primary').click()
    await symbol.locator('[role="tab"]', { hasText: 'Special Characters' }).click()
    await symbol
      .locator('.symbol-special-list button')
      .filter({ hasText: /Em Dash$/ })
      .click()
    await symbol.locator('button.primary').click()
    await page.keyboard.press('Escape')
    await expect(symbol).toBeHidden()
    await expect(page.locator('.doc-page p', { hasText: 'price' })).toHaveText('price\u20ac\u2014')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const linked = await savedParagraph(docPath, 'see the site')
          const id = /<w:hyperlink\b[^>]*r:id="([^"]+)"/.exec(linked)?.[1]
          const rels = await part(docPath, 'word/_rels/document.xml.rels')
          return {
            target: new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*>`)
              .exec(rels)?.[0]
              .match(/Target="([^"]+)"/)?.[1],
            external: new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*TargetMode="External"`).test(
              rels,
            ),
            text: (await savedBody(docPath)).includes('price\u20ac\u2014'),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ target: 'https://example.com/report', external: true, text: true })
  } finally {
    await s.close()
  }
})
