import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

/** Single empty page at a photo-scan size (3024×4032 pt), like image-based PDFs */
function hugePagePdf(): Buffer {
  const objects = [
    '<</Type/Catalog/Pages 2 0 R>>',
    '<</Type/Pages/Kids[3 0 R]/Count 1>>',
    '<</Type/Page/Parent 2 0 R/MediaBox[0 0 3024 4032]>>',
  ]
  let body = '%PDF-1.4\n'
  const offsets: number[] = []
  objects.forEach((obj, i) => {
    offsets.push(body.length)
    body += `${i + 1} 0 obj\n${obj}\nendobj\n`
  })
  const xrefStart = body.length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) body += `${String(off).padStart(10, '0')} 00000 n \n`
  body += `trailer\n<</Size ${objects.length + 1}/Root 1 0 R>>\nstartxref\n${xrefStart}\n%%EOF\n`
  return Buffer.from(body, 'latin1')
}

test('image-sized PDF opens at true fit-to-width, not the old 50% zoom floor', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-fit-'))
  const pdfPath = join(dir, 'huge.pdf')
  await writeFile(pdfPath, hugePagePdf())

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-fit-zoom',
    openFile: pdfPath,
  })
  try {
    const editorPage = await waitForPageWithUrl(launched.app, 'pdf/out')
    await expect(editorPage.locator('.pdf-page').first()).toBeVisible()

    await expect
      .poll(async () => Number(await editorPage.locator('.zoom-slider').inputValue()))
      .toBeLessThan(50)

    const pageW = await editorPage
      .locator('.pdf-page')
      .first()
      .evaluate((el) => el.getBoundingClientRect().width)
    const viewW = await editorPage.locator('.pdf-scroll').evaluate((el) => el.clientWidth)
    expect(pageW).toBeLessThanOrEqual(viewW)
  } finally {
    await closeAndSaveVideo(launched, 'pdf-fit-zoom')
  }
})
