import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

/** One page whose line is drawn as two show-text runs ("...000003" + "761"),
 *  like generated bank statements that split a number mid-way. */
function twoRunTextPdf(): Buffer {
  const stream = 'BT /F1 12 Tf 72 700 Td (MDSUSER 000003) Tj (761) Tj ET'
  const objects = [
    '<</Type/Catalog/Pages 2 0 R>>',
    '<</Type/Pages/Kids[3 0 R]/Count 1>>',
    '<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]/Resources<</Font<</F1 4 0 R>>>>/Contents 5 0 R>>',
    '<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>',
    `<</Length ${stream.length}>>\nstream\n${stream}\nendstream`,
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

/**
 * Regression for "Shift+→ cannot extend the selection over the trailing
 * digits" (alpha feedback): the viewer's window keydown handler mapped
 * ArrowRight to a page flip and preventDefault-ed it without checking
 * shiftKey, so the browser's native selection extension never ran. With a
 * non-collapsed selection present, Shift+navigation must reach the browser.
 */
test('Shift+ArrowRight extends the text selection instead of flipping the page', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-select-'))
  const pdfPath = join(dir, 'tworuns.pdf')
  await writeFile(pdfPath, twoRunTextPdf())

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-shift-arrow-selection',
    openFile: pdfPath,
  })
  try {
    const editorPage = await waitForPageWithUrl(launched.app, 'pdf/out')
    await expect(editorPage.locator('.textLayer span').first()).toBeVisible({ timeout: 30_000 })

    // select up to "...000003", like a finished mouse drag that stopped short
    const before = await editorPage.evaluate(() => {
      const span = document.querySelector('.textLayer span')
      const textNode = span?.firstChild
      if (!textNode || textNode.nodeType !== Node.TEXT_NODE) return null
      const range = document.createRange()
      range.setStart(textNode, 0)
      range.setEnd(textNode, 'MDSUSER 000003'.length)
      const selection = window.getSelection()
      selection?.removeAllRanges()
      selection?.addRange(range)
      return selection?.toString() ?? null
    })
    expect(before).toBe('MDSUSER 000003')

    const scrollBefore = await editorPage
      .locator('.pdf-scroll')
      .evaluate((el) => ({ top: el.scrollTop, left: el.scrollLeft }))

    await editorPage.keyboard.press('Shift+ArrowRight')

    // the selection grew into the second text run...
    await expect
      .poll(async () => editorPage.evaluate(() => window.getSelection()?.toString() ?? ''))
      .toBe('MDSUSER 0000037')

    // ...and the viewer did not treat the key as navigation
    const scrollAfter = await editorPage
      .locator('.pdf-scroll')
      .evaluate((el) => ({ top: el.scrollTop, left: el.scrollLeft }))
    expect(scrollAfter).toEqual(scrollBefore)
  } finally {
    await closeAndSaveVideo(launched, 'pdf-shift-arrow-selection')
  }
})
