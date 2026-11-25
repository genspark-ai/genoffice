import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('The hand tool pans by drag, and the pen thickness is chosen, saved and remembered', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-hand-'))
  const pdfPath = join(dir, 'pages.pdf')
  const doc = await PDFDocument.create()
  for (let i = 0; i < 4; i++) doc.addPage([400, 600])
  await writeFile(pdfPath, await doc.save({ useObjectStreams: false }))
  const savedInkWidth = async () => {
    const saved = await PDFDocument.load(await readFile(pdfPath))
    const annots = saved.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annots || annots.size() === 0) return null
    const bs = annots.lookup(0, PDFDict).lookupMaybe(PDFName.of('BS'), PDFDict)
    return bs?.lookupMaybe(PDFName.of('W'), PDFNumber)?.asNumber() ?? null
  }

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-pen-width-hand-tool',
    openFile: pdfPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://pdf/')
    const content = page.locator('.pdf-page-content').first()
    await expect(content.locator('canvas')).toBeVisible()
    const scroll = page.locator('.pdf-scroll')
    const scrollTop = () => scroll.evaluate((el) => el.scrollTop)
    const drag = async (dy: number) => {
      const box = (await scroll.boundingBox())!
      const x = box.x + box.width / 2
      const y = box.y + box.height / 2
      await page.mouse.move(x, y)
      await page.mouse.down()
      await page.mouse.move(x, y + dy / 2, { steps: 4 })
      await page.mouse.move(x, y + dy, { steps: 4 })
      await page.mouse.up()
    }

    await drag(-120)
    expect(await scrollTop()).toBe(0)

    await page.locator('.ribbon-tab', { hasText: /^View$/ }).click()
    await page.getByRole('button', { name: 'Hand Tool', exact: true }).click()
    await expect(scroll).toHaveClass(/pdf-hand/)
    await drag(-120)
    await expect.poll(scrollTop).toBeGreaterThan(100)

    // back to text selection from the page's right-click menu
    await scroll.click({ button: 'right' })
    await page.getByRole('menuitemradio', { name: 'Select Text' }).click()
    await expect(scroll).not.toHaveClass(/pdf-hand/)
    await scroll.evaluate((el) => (el.scrollTop = 0))

    await page.getByRole('button', { name: 'Annotate', exact: true }).click()
    await page.getByRole('button', { name: 'Thickness', exact: true }).click()
    await page.getByRole('menuitemradio', { name: '5 pt' }).click()
    await page.getByRole('button', { name: 'Draw', exact: true }).click()
    const box = (await content.boundingBox())!
    await page.mouse.move(box.x + box.width * 0.3, box.y + box.height * 0.3)
    await page.mouse.down()
    await page.mouse.move(box.x + box.width * 0.6, box.y + box.height * 0.35, { steps: 6 })
    await page.mouse.up()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(savedInkWidth).toBe(5)
    expect(await page.evaluate(() => localStorage.getItem('genoffice-pdf-draw-width'))).toBe('5')
  } finally {
    await closeAndSaveVideo(launched, 'pdf-pen-width-hand-tool')
  }
})
