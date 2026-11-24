import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFArray, PDFDocument, PDFName } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

const PAGE = { width: 400, height: 600 }
/** Horizontal strokes, as distances from the top of the page */
const STROKE_TOPS = [80, 160]

test('A freehand stroke saved in the file is deleted by click and by the eraser', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-saved-ink-'))
  const pdfPath = join(dir, 'ink.pdf')
  const doc = await PDFDocument.create()
  const first = doc.addPage([PAGE.width, PAGE.height])
  first.node.set(
    PDFName.of('Annots'),
    doc.context.obj(
      STROKE_TOPS.map((top) => {
        const y = PAGE.height - top
        return doc.context.register(
          doc.context.obj({
            Type: 'Annot',
            Subtype: 'Ink',
            Rect: [98, y - 2, 302, y + 2],
            InkList: [[100, y, 300, y]],
            C: [0.86, 0.22, 0.18],
            BS: { W: 2 },
            F: 4,
          }),
        )
      }),
    ),
  )
  await writeFile(pdfPath, await doc.save({ useObjectStreams: false }))
  const inkCount = async () => {
    const saved = await PDFDocument.load(await readFile(pdfPath))
    return saved.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)?.size() ?? 0
  }

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-delete-saved-drawing',
    openFile: pdfPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://pdf/')
    const content = page.locator('.pdf-page-content').first()
    await expect(content.locator('canvas')).toBeVisible()
    const clickStroke = async (top: number) => {
      const box = (await content.boundingBox())!
      await page.mouse.click(box.x + box.width / 2, box.y + (box.height * top) / PAGE.height)
    }

    // The saved annotations load a moment after the page paints: click until the popup answers
    await expect(async () => {
      await clickStroke(STROKE_TOPS[0]!)
      await expect(page.locator('.pdf-del-popup')).toBeVisible({ timeout: 1000 })
    }).toPass({ timeout: 15_000 })
    await page.locator('.pdf-del-popup-danger').click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(inkCount).toBe(1)

    await page.getByRole('button', { name: 'Annotate', exact: true }).click()
    await page.getByRole('button', { name: 'Eraser', exact: true }).click()
    await expect(async () => {
      await clickStroke(STROKE_TOPS[1]!)
      await page.keyboard.press('ControlOrMeta+s')
      await expect.poll(inkCount, { timeout: 2000 }).toBe(0)
    }).toPass({ timeout: 20_000 })
  } finally {
    await closeAndSaveVideo(launched, 'pdf-delete-saved-drawing')
  }
})
