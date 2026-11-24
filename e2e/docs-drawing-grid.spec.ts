/**
 * Format > Document > Document Grid > Drawing Grid: the grid a file carries
 * is shown and can be changed, View > Gridlines draws it, a dragged or drawn
 * shape snaps to it, and the grid is saved while the snap switch stays on
 * this computer.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith, para, saveFromMenu, shapeParagraph } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-drawing-grid'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const GRID =
  '<w:defaultTabStop w:val="720"/>' +
  '<w:drawingGridHorizontalSpacing w:val="360"/><w:drawingGridVerticalSpacing w:val="360"/>' +
  '<w:displayHorizontalDrawingGridEvery w:val="0"/><w:displayVerticalDrawingGridEvery w:val="2"/>' +
  '<w:characterSpacingControl w:val="doNotCompress"/>'

async function docxWithGrid(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(
    await docxWith(
      para('First') +
        // a preset shape is moved by its body; a plain rectangle is a text box and takes the click
        shapeParagraph(1, 40, 30, 120, 60).replace('prst="rect"', 'prst="ellipse"') +
        para('Second') +
        para('Third'),
    ),
  )
  zip.file('word/settings.xml', `<?xml version="1.0"?><w:settings ${W_NS}>${GRID}</w:settings>`)
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function clickMenu(app: ElectronApplication, path: string[]) {
  const click = () =>
    app.evaluate(({ Menu }, path) => {
      let items = Menu.getApplicationMenu()?.items
      let item: Electron.MenuItem | undefined
      for (const label of path) {
        item = items?.find((candidate) => candidate.label === label)
        items = item?.submenu?.items
      }
      if (!item?.enabled) return false
      item.click()
      return true
    }, path)
  await expect.poll(click, { timeout: 15_000 }).toBe(true)
}

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const settingsXml = (docPath: string) => part(docPath, 'word/settings.xml')

async function openDocs(docPath: string, userDataDir?: string, shapes = 1) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    ...(userDataDir ? { userDataDir } : {}),
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page .doc-textbox')).toHaveCount(shapes, { timeout: 30_000 })
  return { launched, page }
}

async function close(launched: LaunchedApp, name: string, keepUserData = false) {
  await closeAndSaveVideo(launched, name)
  if (!keepUserData) rmSync(launched.userDataDir, { recursive: true, force: true })
}

async function openDrawingGrid(app: ElectronApplication, page: Page) {
  await page.locator('.doc-page p', { hasText: 'Third' }).click()
  await clickMenu(app, ['Format', 'Document…'])
  const documentDialog = page.locator('.doc-dialog')
  await documentDialog.getByRole('tab', { name: 'Document Grid' }).click()
  await documentDialog.getByRole('button', { name: 'Drawing Grid…' }).click()
  const dialog = page.locator('.draw-grid-dialog')
  await expect(dialog).toBeVisible()
  return { dialog, documentDialog }
}

/** the shape's corner in grid steps from the margin corner of the page, and the gridlines drawn */
const gridState = (page: Page, stepPx: number) =>
  page.evaluate((stepPx) => {
    const paper = document.querySelector('.doc-page') as HTMLElement
    const zoomEl = paper.closest('.doc-zoom')
    const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
    const style = getComputedStyle(paper)
    const rect = paper.getBoundingClientRect()
    const box = document.querySelector('.doc-page .doc-textbox')!.getBoundingClientRect()
    const steps = (client: number, edge: number, padding: string) =>
      (client - edge - parseFloat(padding) * zoom) / (stepPx * zoom)
    return {
      x: steps(box.left, rect.left, style.paddingLeft),
      y: steps(box.top, rect.top, style.paddingTop),
      lines: style.backgroundImage === 'none' ? '' : style.backgroundSize,
      columns: style.getPropertyValue('--docs-grid-columns').trim(),
    }
  }, stepPx)

const offGrid = (steps: number) => Math.abs(steps - Math.round(steps))

async function dragShape(page: Page, dx: number, dy: number) {
  const at = (await page.locator('.doc-page .doc-textbox').first().boundingBox())!
  const x = at.x + at.width / 2
  const y = at.y + at.height / 2
  await page.mouse.move(x, y)
  await page.mouse.down()
  await page.mouse.move(x + dx / 2, y + dy / 2, { steps: 4 })
  await page.mouse.move(x + dx, y + dy, { steps: 4 })
  await page.mouse.up()
}

test('Drawing Grid shows the grid of the file, changes it, and Gridlines draws it', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-drawing-grid-')))
  const docPath = join(dir, 'grid.docx')
  writeFileSync(docPath, await docxWithGrid())
  const first = await openDocs(docPath)
  try {
    const { launched, page } = first
    const { dialog, documentDialog } = await openDrawingGrid(launched.app, page)
    await expect(dialog.locator('input[name="hSpacing"]')).toHaveValue('18')
    await expect(dialog.locator('input[name="vSpacing"]')).toHaveValue('18')
    await expect(dialog.locator('input[name="snapToGrid"]')).not.toBeChecked()
    await expect(dialog.locator('input[name="displayGridlines"]')).not.toBeChecked()
    await expect(dialog.locator('input[name="verticalLines"]')).not.toBeChecked()
    await expect(dialog.locator('input[name="vEvery"]')).toBeDisabled()
    await expect(dialog.locator('input[name="hEvery"]')).toHaveValue('2')
    await expect(dialog.locator('input[name="useMargins"]')).toBeChecked()
    await expect(dialog.locator('input[name="hOrigin"]')).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('drawing-grid.png') })

    // a spacing out of range cannot be applied
    await dialog.locator('input[name="hSpacing"]').fill('0')
    await expect(dialog.locator('button.primary')).toBeDisabled()
    await dialog.locator('input[name="hSpacing"]').fill('24')
    await dialog.locator('input[name="verticalLines"]').check()
    await dialog.locator('input[name="vEvery"]').fill('1')
    await dialog.locator('input[name="displayGridlines"]').check()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // its own OK applied it: the Document dialog is cancelled
    await documentDialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(documentDialog).toBeHidden()

    // a vertical line every 24 pt, a horizontal one every 2 x 18 pt
    await expect.poll(async () => (await gridState(page, 32)).lines).toBe('32px 48px, 32px 48px')
    expect((await gridState(page, 32)).columns).toBe('')
    await page.screenshot({ path: test.info().outputPath('gridlines.png') })
    // View > Gridlines is the same switch
    await clickMenu(launched.app, ['View', 'Gridlines'])
    await expect.poll(async () => (await gridState(page, 32)).lines).toBe('')
    await clickMenu(launched.app, ['View', 'Gridlines'])
    await expect.poll(async () => (await gridState(page, 32)).lines).toBe('32px 48px, 32px 48px')

    await saveFromMenu(launched.app)
    await expect
      .poll(() => settingsXml(docPath), { timeout: 20_000 })
      .toContain(
        '<w:defaultTabStop w:val="720"/><w:drawingGridHorizontalSpacing w:val="480"/>' +
          '<w:drawingGridVerticalSpacing w:val="360"/>' +
          '<w:displayHorizontalDrawingGridEvery w:val="1"/>' +
          '<w:displayVerticalDrawingGridEvery w:val="2"/><w:characterSpacingControl',
      )
  } finally {
    await close(first.launched, 'drawing-grid-set')
  }

  const second = await openDocs(docPath)
  try {
    const { launched, page } = second
    const { dialog } = await openDrawingGrid(launched.app, page)
    await expect(dialog.locator('input[name="hSpacing"]')).toHaveValue('24')
    await expect(dialog.locator('input[name="vSpacing"]')).toHaveValue('18')
    await expect(dialog.locator('input[name="verticalLines"]')).toBeChecked()
    await expect(dialog.locator('input[name="vEvery"]')).toHaveValue('1')
    // an OK that changes nothing leaves the document clean and the file as it is
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    const before = await settingsXml(docPath)
    expect(before).toContain('<w:drawingGridHorizontalSpacing w:val="480"/>')
  } finally {
    await close(second.launched, 'drawing-grid-reopen')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a dragged shape snaps to the drawing grid once Snap objects is on, and the switch survives a restart', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-drawing-grid-')))
  const docPath = join(dir, 'snap.docx')
  writeFileSync(docPath, await docxWithGrid())
  const first = await openDocs(docPath)
  let snapped = { x: 0, y: 0 }
  try {
    const { launched, page } = first
    // 18 pt = 24 px steps; the shape starts off the grid and a drag leaves it off
    await dragShape(page, 37, 21)
    await expect
      .poll(async () => {
        const at = await gridState(page, 24)
        return offGrid(at.x) > 0.02 || offGrid(at.y) > 0.02
      })
      .toBe(true)

    const { dialog, documentDialog } = await openDrawingGrid(launched.app, page)
    await dialog.locator('input[name="snapToGrid"]').check()
    await dialog.locator('button.primary').click()
    await documentDialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(documentDialog).toBeHidden()

    const start = await gridState(page, 24)
    // a click selects the shape where it is, off the grid: only a drag snaps
    const shape = (await page.locator('.doc-page .doc-textbox').first().boundingBox())!
    await page.mouse.move(shape.x + shape.width / 2, shape.y + shape.height / 2)
    await page.mouse.down()
    await page.mouse.move(shape.x + shape.width / 2 + 1, shape.y + shape.height / 2 + 1)
    await page.mouse.up()
    await page.waitForTimeout(300)
    expect(await gridState(page, 24)).toEqual(start)
    await dragShape(page, 41, 53)
    await expect
      .poll(async () => {
        const at = await gridState(page, 24)
        return Math.max(offGrid(at.x), offGrid(at.y))
      })
      .toBeLessThan(0.03)
    snapped = await gridState(page, 24)
    expect(Math.abs(snapped.x - start.x - 41 / 24)).toBeLessThanOrEqual(0.55)
    expect(Math.abs(snapped.y - start.y - 53 / 24)).toBeLessThanOrEqual(0.55)
    await page.screenshot({ path: test.info().outputPath('snapped.png') })

    // a drawn text box starts and ends on grid points
    const insertTab = page.locator('.ribbon-tab', { hasText: /^Insert$/ })
    if (!(await insertTab.getAttribute('class'))?.includes('active')) await insertTab.click()
    await page.locator('.rb-big', { hasText: 'Text Box' }).click()
    await page.locator('.rb-textbox-menu button', { hasText: 'Draw Text Box' }).click()
    const paper = page.locator('.doc-page[contenteditable="true"]')
    await expect(paper).toHaveClass(/doc-shape-drawing/)
    const paperBox = (await paper.boundingBox())!
    const x0 = paperBox.x + 133
    const y0 = paperBox.y + 451
    await page.mouse.move(x0, y0)
    await page.mouse.down()
    await page.mouse.move(x0 + 60, y0 + 30, { steps: 4 })
    await page.mouse.move(x0 + 107, y0 + 67, { steps: 4 })
    await page.mouse.up()
    await expect(page.locator('.doc-page .doc-textbox')).toHaveCount(2)
    const drawn = () =>
      page.evaluate(() => {
        const page = document.querySelector('.doc-page') as HTMLElement
        const style = getComputedStyle(page)
        const rect = page.getBoundingClientRect()
        const boxes = Array.from(document.querySelectorAll('.doc-page .doc-textbox'))
        const box = boxes.map((el) => el.getBoundingClientRect()).sort((a, b) => b.top - a.top)[0]
        return [
          (box.left - rect.left - parseFloat(style.paddingLeft)) / 24,
          (box.top - rect.top - parseFloat(style.paddingTop)) / 24,
          box.width / 24,
          box.height / 24,
        ]
      })
    await expect.poll(async () => Math.max(...(await drawn()).map(offGrid))).toBeLessThan(0.05)
    // 133 -> 37 px into the text area -> the 48 px line; 107 px wide -> to the 144 px line
    expect((await drawn()).map(Math.round)).toEqual([2, 15, 4, 3])
    await page.locator('.doc-page p', { hasText: 'First' }).click()

    await saveFromMenu(launched.app)
    // the save has landed once the drawn box is in the file
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).match(/<wp:anchor /g)?.length, {
        timeout: 20_000,
      })
      .toBe(2)
    // the snap switch is not part of the document
    expect(await settingsXml(docPath)).toContain(GRID)
  } finally {
    await close(first.launched, 'drawing-grid-snap', true)
  }

  const userDataDir = first.launched.userDataDir
  let second: Awaited<ReturnType<typeof openDocs>> | undefined
  try {
    second = await openDocs(docPath, userDataDir, 2)
    const { launched, page } = second
    // the saved position is the snapped one
    await expect
      .poll(async () => {
        const at = await gridState(page, 24)
        return Math.abs(at.x - snapped.x) + Math.abs(at.y - snapped.y)
      })
      .toBeLessThan(0.06)
    const { dialog, documentDialog } = await openDrawingGrid(launched.app, page)
    await expect(dialog.locator('input[name="snapToGrid"]')).toBeChecked()
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await documentDialog.getByRole('button', { name: 'Cancel' }).click()
    await dragShape(page, -50, 20)
    await expect
      .poll(async () => {
        const at = await gridState(page, 24)
        return Math.max(offGrid(at.x), offGrid(at.y))
      })
      .toBeLessThan(0.03)
    expect((await gridState(page, 24)).x).toBeLessThan(snapped.x - 1)
  } finally {
    if (second) await close(second.launched, 'drawing-grid-restart', true)
    rmSync(userDataDir, { recursive: true, force: true })
    rmSync(dir, { recursive: true, force: true })
  }
})
