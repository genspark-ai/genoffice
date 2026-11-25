/**
 * What the table and picture commands still left out: the Insert / Delete
 * Cells hint says which options are refused and why, Delete Cells with a
 * shift is a tracked change, a cell shifted in downward moves back on reject
 * after the file is reopened, a bookmark made from a selection covers that
 * text for a formula and in the file, a tracked conversion carries the
 * bookmark to the table, and Change Picture works on a picture in a line of
 * text. Each is driven through the UI and read back from the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { rmSync } from 'node:fs'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  part,
  ribbon,
  ribbonTab,
  savedBody,
  selectParagraph,
  selectWord,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-table-picture-leftovers'

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const p = (content: string) => `<w:p>${content.startsWith('<') ? content : run(content)}</w:p>`
const tc = (text: string, tcPr = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>${p(text)}</w:tc>`
const BORDERS =
  '<w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join('') +
  '</w:tblBorders>'
const table = (rows: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid>` +
  `<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows.map((r) => `<w:tr>${r}</w:tr>`).join('')}</w:tbl>`

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

const resolveAll = async (page: Page, verb: 'Accept' | 'Reject') => {
  await ribbonTab(page, 'Review')
  await page
    .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: new RegExp(`^${verb}$`),
    })
    .click()
  await page
    .locator('[data-rb-panel] button', { hasText: new RegExp(`^${verb} All Changes$`) })
    .first()
    .click()
}

const HINT = 'The greyed-out options would move cells through a merged cell'

test('the cells dialogs say why an option is greyed out, and Delete Cells with a shift is a tracked change', async () => {
  test.setTimeout(240_000)
  const s = await open(
    VIDEO_DIR,
    p('Stock') +
      table([
        tc('A1') + tc('B1'),
        tc('A2') + tc('B2', '<w:vMerge w:val="restart"/>'),
        tc('A3') + tc('', '<w:vMerge/>'),
        tc('A4') + tc('B4'),
      ]) +
      p('tail'),
    'cells.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    const rows = async () => (await blocks(page))[1]
    const insert = page.locator('.insert-cells-dialog')
    const remove = page.locator('.delete-cells-dialog')

    // column A has no merged cell below A1: nothing is refused, nothing to explain
    await caretAtEnd(page, 'A1')
    await clickMenu(app, ['Table', 'Insert', 'Cells…'])
    await expect(insert.getByLabel('Shift cells down')).toBeEnabled()
    await expect(insert.getByLabel('Shift cells right')).toBeEnabled()
    await expect(insert.locator('.modal-desc')).toHaveCount(0)
    await page.keyboard.press('Escape')
    await expect(insert).toHaveCount(0)

    // below B1 the merged cell is in the way of a move down, not of a move right
    await caretAtEnd(page, 'B1')
    await clickMenu(app, ['Table', 'Insert', 'Cells…'])
    await expect(insert.getByLabel('Shift cells down')).toBeDisabled()
    await expect(insert.getByLabel('Shift cells right')).toBeEnabled()
    await expect(insert.locator('.modal-desc')).toHaveText(HINT)
    await page.screenshot({ path: test.info().outputPath('insert-cells-hint.png') })
    await page.keyboard.press('Escape')
    await expect(insert).toHaveCount(0)
    await clickMenu(app, ['Table', 'Delete', 'Cells…'])
    await expect(remove.getByLabel('Shift cells up')).toBeDisabled()
    await expect(remove.getByLabel('Shift cells left')).toBeEnabled()
    await expect(remove.locator('.modal-desc')).toHaveText(HINT)
    await page.keyboard.press('Escape')
    await expect(remove).toHaveCount(0)

    await trackChangesOn(page)
    const deleteCells = async (at: string, option: string) => {
      await caretAtEnd(page, at)
      await clickMenu(app, ['Table', 'Delete', 'Cells…'])
      await remove.getByLabel(option).check()
      await remove.locator('button.primary, button[type=submit]').first().click()
      await expect(remove).toHaveCount(0)
    }
    // left: the cell stays, struck, until the change is accepted
    await deleteCells('A1', 'Shift cells left')
    await expect(page.locator('.doc-page td.cell-rev-del')).toHaveCount(1)
    expect(await rows()).toEqual([['A1', 'B1'], ['A2', 'B2'], ['A3'], ['A4', 'B4']])
    await page.screenshot({ path: test.info().outputPath('delete-cells-tracked.png') })
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).toContain('<w:cellDel ')
    expect(await savedBody(docPath)).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>).)*<w:delText[^>]*>A1</)

    await resolveAll(page, 'Reject')
    await expect(page.locator('.doc-page td.cell-rev-del')).toHaveCount(0)
    expect(await rows()).toEqual([['A1', 'B1'], ['A2', 'B2'], ['A3'], ['A4', 'B4']])

    await deleteCells('A1', 'Shift cells left')
    await resolveAll(page, 'Accept')
    await expect.poll(rows).toEqual([['B1'], ['A2', 'B2'], ['A3'], ['A4', 'B4']])
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).not.toContain('<w:cellDel ')
    expect(await savedBody(docPath)).toContain('<w:gridAfter w:val="1"/>')

    // up: accepting moves the rest of the column up and leaves a blank cell at its end
    await deleteCells('A2', 'Shift cells up')
    await expect(page.locator('.doc-page td.cell-rev-del')).toHaveCount(1)
    await resolveAll(page, 'Accept')
    await expect.poll(rows).toEqual([['B1'], ['A3', 'B2'], ['A4'], ['', 'B4']])
    await page.screenshot({ path: test.info().outputPath('delete-cells-up-accepted.png') })
  } finally {
    await s.close()
  }
})

test('a cell shifted in downward moves back up when it is rejected after the file is reopened', async () => {
  test.setTimeout(300_000)
  const s = await open(
    VIDEO_DIR,
    p('Stock') + table([tc('A1') + tc('B1'), tc('A2') + tc('B2'), tc('A3') + tc('B3')]) + p('tail'),
    'shift-down.docx',
  )
  let launched = s.launched
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    await trackChangesOn(page)
    await caretAtEnd(page, 'A2')
    await clickMenu(app, ['Table', 'Insert', 'Cells…'])
    const dialog = page.locator('.insert-cells-dialog')
    await dialog.getByLabel('Shift cells down').check()
    await dialog.locator('button.primary, button[type=submit]').first().click()
    await expect(dialog).toHaveCount(0)
    const shifted = [
      ['A1', 'B1'],
      ['', 'B2'],
      ['A2', 'B3'],
      ['A3', ''],
    ]
    expect((await blocks(page))[1]).toEqual(shifted)
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).toContain('<w:cellIns ')
    expect((await savedBody(docPath)).match(/<w:tr[ >]/g)).toHaveLength(4)

    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const again = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(again.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    await expect.poll(async () => (await blocks(again))[1]).toEqual(shifted)
    await expect(again.locator('.doc-page td.cell-rev-ins')).toHaveCount(2)
    await resolveAll(again, 'Reject')
    await expect
      .poll(async () => (await blocks(again))[1])
      .toEqual([
        ['A1', 'B1'],
        ['A2', 'B2'],
        ['A3', 'B3'],
      ])
    await again.screenshot({ path: test.info().outputPath('shift-down-rejected-after-reopen.png') })
    await saveFromMenu(launched.app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).not.toContain('<w:cellIns ')
    expect((await savedBody(docPath)).match(/<w:tr[ >]/g)).toHaveLength(3)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(s.dir, { recursive: true, force: true })
  }
})

const addBookmark = async (app: ElectronApplication, page: Page, name: string) => {
  const dialog = page.locator('.bookmark-dialog')
  await clickMenu(app, ['Insert', 'Bookmark…'])
  await dialog.locator('input[name="bookmark-name"]').fill(name)
  await dialog.locator('button[name="add"]').click()
  await expect(dialog.locator('.bookmark-row', { hasText: name })).toHaveCount(1)
  await page.keyboard.press('Escape')
  await expect(dialog).toBeHidden()
}

test('a bookmark made from a selection covers that text: a formula reads it and the file marks its two ends', async () => {
  test.setTimeout(240_000)
  const s = await open(
    VIDEO_DIR,
    p('Price 100, discount 25 each') + p('Twice the discount: ') + p('tail'),
    'bookmark.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    await selectWord(page, '25')
    await addBookmark(app, page, 'Discount')

    const formula = page.locator('.formula-dialog')
    const twice = page.locator('.doc-page p', { hasText: 'Twice the discount' })
    await caretAtEnd(page, 'Twice the discount: ')
    await clickMenu(app, ['Table', 'Formula…'])
    await formula.locator('input[name="formula"]').fill('=Discount*2')
    await formula.locator('button.primary').click()
    await expect(formula).toBeHidden()
    // 25, not the paragraph's first number
    await expect(twice).toHaveText('Twice the discount: 50')

    // an edit beside the bookmark leaves its range alone
    await caretAtEnd(page, 'Price 100')
    await page.keyboard.type(' net')
    await selectParagraph(page, 'Twice the discount')
    await page.keyboard.press('F9')
    await expect(twice).toHaveText('Twice the discount: 50')
    await page.screenshot({ path: test.info().outputPath('bookmark-range-formula.png') })

    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).toContain('w:name="Discount"')
    const xml = await savedBody(docPath)
    expect(xml).toMatch(
      /discount <\/w:t><\/w:r><w:bookmarkStart w:id="(\d+)" w:name="Discount"\/><w:r>(?:(?!<\/w:r>).)*>25<\/w:t><\/w:r><w:bookmarkEnd w:id="\1"\/>/,
    )
    expect(xml.match(/w:name="Discount"/g)).toHaveLength(1)
  } finally {
    await s.close()
  }
})

test('a tracked Convert Text to Table carries the bookmark to the table, so accepting keeps it', async () => {
  test.setTimeout(240_000)
  const s = await open(
    VIDEO_DIR,
    p(
      '<w:bookmarkStart w:id="1" w:name="Total"/>' + run('Item,Qty') + '<w:bookmarkEnd w:id="1"/>',
    ) +
      p('Apple,3') +
      p('tail'),
    'convert-bookmark.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    await trackChangesOn(page)
    await selectParagraph(page, 'Item,Qty', 'Apple,3')
    await clickMenu(app, ['Table', 'Convert', 'Convert Text to Table…'])
    const dialog = page.locator('.convert-text-dialog')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).toContain('<w:tbl>')
    let xml = await savedBody(docPath)
    expect(xml.match(/w:name="Total"/g)).toHaveLength(1)
    expect(xml.indexOf('w:name="Total"')).toBeGreaterThan(xml.indexOf('<w:tbl>'))

    await resolveAll(page, 'Accept')
    await expect
      .poll(() => blocks(page))
      .toEqual([
        [
          ['Item', 'Qty'],
          ['Apple', '3'],
        ],
        'tail',
      ])
    await clickMenu(app, ['Insert', 'Bookmark…'])
    const bookmarks = page.locator('.bookmark-dialog')
    await expect(bookmarks.locator('.bookmark-row .bookmark-name')).toHaveText(['Total'])
    await expect(bookmarks.locator('.bookmark-row .bookmark-preview')).toHaveText(['Item'])
    await page.screenshot({ path: test.info().outputPath('converted-bookmark-kept.png') })
    await page.keyboard.press('Escape')
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).not.toMatch(/<w:(?:ins|del) /)
    xml = await savedBody(docPath)
    expect(xml).toMatch(/<w:tc>(?:(?!<\/w:tc>).)*w:name="Total"(?:(?!<\/w:tc>).)*>Item</)
  } finally {
    await s.close()
  }
})

const INLINE_PICTURE =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="762000" cy="381000"/>' +
  '<wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture 7"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="381000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

test('Change Picture works on a picture in a line of text: it stays in the line at its width', async () => {
  test.setTimeout(240_000)
  const s = await open(
    VIDEO_DIR,
    p(run('before ') + INLINE_PICTURE + run(' after')) + p('tail'),
    'inline-picture.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
    const picture = page.locator('.doc-page p', { hasText: 'before' }).locator('img').first()
    await expect(picture).toBeVisible()
    const box = async () => {
      const r = (await picture.boundingBox())!
      return [Math.round(r.width), Math.round(r.height)]
    }
    expect(await box()).toEqual([80, 40])
    const src = await picture.getAttribute('src')
    const media = async () =>
      (await part(docPath, 'word/_rels/document.xml.rels')).match(/media\//g)

    // a picture as tall as it is wide, handed to the page as the clipboard's content: the system
    // clipboard is shared with whatever else runs on the machine, and reading it is covered by
    // docs-change-picture-sources
    await page.evaluate(async () => {
      const canvas = document.createElement('canvas')
      canvas.width = canvas.height = 20
      const ctx = canvas.getContext('2d')!
      ctx.fillStyle = 'rgb(0, 120, 220)'
      ctx.fillRect(0, 0, 20, 20)
      const png = await new Promise<Blob>((resolve) => canvas.toBlob((blob) => resolve(blob!)))
      navigator.clipboard.read = async () => [new ClipboardItem({ 'image/png': png })]
    })
    const sources = async () => {
      await picture.click({ button: 'right' })
      const menu = page.locator('.ctx-menu')
      await menu.waitFor({ timeout: 15_000 })
      const item = menu.locator('.ctx-item', { hasText: 'Change Picture' })
      await expect(item).toBeEnabled()
      await item.hover()
      return menu.locator('.ctx-submenu .ctx-item')
    }
    let items = await sources()
    await expect(items.locator('.ctx-label')).toHaveText([
      'From a File…',
      'From Icons…',
      'From Clipboard…',
    ])
    await page.screenshot({ path: test.info().outputPath('inline-change-picture-menu.png') })
    await items.nth(2).click()
    await expect.poll(() => picture.getAttribute('src')).not.toBe(src)
    await expect.poll(box).toEqual([80, 80])
    await expect(page.locator('.doc-page p', { hasText: 'before' })).toHaveText('before  after')
    const pasted = await picture.getAttribute('src')

    // an icon, drawn from its vector
    items = await sources()
    await items.nth(1).click()
    const icons = page.locator('.icons-dialog')
    await expect(icons).toBeVisible()
    await icons.locator('.icons-dialog-cell').nth(2).click()
    await icons.locator('button.primary').click()
    await expect(icons).toBeHidden()
    await expect.poll(() => picture.getAttribute('src')).not.toBe(pasted)
    expect(await picture.getAttribute('src')).toMatch(/^data:image\/svg\+xml/)
    await expect.poll(box).toEqual([80, 80])
    await page.screenshot({ path: test.info().outputPath('inline-picture-changed.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => (await savedBody(docPath)).includes('svgBlip'), { timeout: 30_000 })
      .toBe(true)
    const xml = await savedBody(docPath)
    // still one drawing between the two runs of text, 80 px square
    expect(xml).toMatch(/before <\/w:t><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing><wp:inline/)
    expect(xml.match(/<w:drawing>/g)).toHaveLength(1)
    expect(xml).toContain('<wp:extent cx="762000" cy="762000"/>')
    expect(xml).not.toContain('aidocs-new-pic-')
    expect(xml).not.toContain('r:embed="rIdImg"')
    expect((await media())!.length).toBeGreaterThanOrEqual(2)
  } finally {
    await s.close()
  }
})
