/**
 * Format > Text Effects through the menu and its dialog: the outline, shadow
 * and glow show on the selected text, are in the saved file as Word's own
 * markup, and show again when the file is reopened.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

test('text effects are set from the dialog, shown, saved and read back', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-text-fx-')))
  const docPath = join(dir, 'effects.docx')
  writeFileSync(docPath, await docxWith(para('Plain fancy words') + para('Last')))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-text-effects',
    openFile: docPath,
  })
  /** the styled piece of the first paragraph: its text and what draws it */
  const styled = (page: import('@playwright/test').Page) =>
    page.evaluate(() => {
      const span = document.querySelector('.doc-page p span[style*="text-shadow"]') as HTMLElement
      return span
        ? {
            text: span.textContent,
            stroke: span.style.getPropertyValue('-webkit-text-stroke'),
            shadows: span.style.textShadow.match(/rgba?\(/g)?.length ?? 0,
          }
        : null
    })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain fancy words' })).toBeVisible({
      timeout: 30_000,
    })
    // "fancy": the first paragraph's text starts at position 1
    await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from: 7, to: 12 }).run()
    })
    await app.evaluate(({ Menu }) => {
      const format = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Format')
      format?.submenu?.items.find((item) => item.label === 'Text Effects…')?.click()
    })
    const dialog = page.locator('.text-fx-dialog')
    const pick = (group: string, label: RegExp) =>
      dialog.locator(`[data-fx="${group}"] .text-fx-choice`, { hasText: label }).click()
    // nothing is on yet
    await expect(dialog.locator('.text-fx-choice.active')).toHaveText([
      'None',
      'None',
      'None',
      'None',
    ])
    await pick('outline', /^1\.5 pt$/)
    await dialog.locator('[data-fx="outline"] .text-fx-swatch[aria-label="#FF0000"]').click()
    await pick('shadow', /^Offset: Bottom Right$/)
    await pick('glow', /^8 pt$/)
    await page.screenshot({ path: test.info().outputPath('text-effects-dialog.png') })
    await dialog.locator('button.primary').click()

    // the outline is a stroke; the shadow and the two layers of the glow are one text-shadow
    await expect
      .poll(() => styled(page))
      .toEqual({ text: 'fancy', stroke: '1.5pt rgb(255, 0, 0)', shadows: 3 })
    await page.screenshot({ path: test.info().outputPath('text-effects-result.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
          return [
            /<w:document[^>]*xmlns:w14=/.test(xml),
            xml.includes('<w14:textOutline w14:w="19050"'),
            xml.includes('<w14:srgbClr w14:val="FF0000"/>'),
            xml.includes('w14:dir="2700000"'),
            xml.includes('<w14:glow w14:rad="101600">'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-text-effects')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-text-effects',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain fancy words' })).toBeVisible({
      timeout: 30_000,
    })
    await expect
      .poll(() => styled(page))
      .toEqual({ text: 'fancy', stroke: '1.5pt rgb(255, 0, 0)', shadows: 3 })
  } finally {
    await closeAndSaveVideo(launched, 'docs-text-effects')
    rmSync(dir, { recursive: true, force: true })
  }
})
