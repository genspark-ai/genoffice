import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * List completion: a numbered list gets "Set Numbering Value…" (right-click) →
 * restart at 5, "Adjust List Indents…" → number position / text indent /
 * follow-with space, and the Numbering gallery's Define New Number Format…
 * → upper roman. Saved numbering.xml carries startOverride, the rewritten
 * level (w:ind + w:suff) and the new abstractNum.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown> }
}

async function minimalDocx(text: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function part(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return (await zip.file(name)?.async('string')) ?? ''
}

test.describe('docs list completion', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-lists-')))
    docPath = join(dir, 'lists.docx')
    writeFileSync(docPath, await minimalDocx('start'))
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('set numbering value, adjust indents and define a number format reach numbering.xml', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-list-completion',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const editor = await waitForPageWithUrl(app, '://docs/')
      await editor.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      await expect(editor.locator('.doc-page')).toContainText('start')
      await editor.locator('.doc-page').click()
      await editor.keyboard.press('ControlOrMeta+a')
      await editor.keyboard.type('one')
      await expect(editor.locator('.doc-page')).toContainText('one')
      // the Numbering button turns the paragraph into a list; Enter continues it
      const numberingButton = editor.getByRole('button', { name: 'Numbering', exact: true }).first()
      await expect(numberingButton).toBeEnabled()
      await numberingButton.click()
      await expect(editor.locator('.doc-page .doc-li')).toHaveCount(1)
      await editor.locator('.doc-page .doc-li').first().click()
      await editor.keyboard.press('End')
      await editor.keyboard.press('Enter')
      await editor.keyboard.type('two')
      await editor.keyboard.press('Enter')
      await editor.keyboard.type('three')
      const items = editor.locator('.doc-page .doc-li')
      await expect(items).toHaveCount(3)
      await expect(items.nth(2)).toHaveAttribute('data-marker', '3.')

      // right-click the third item → Set Numbering Value… → start at 5
      const third = items.nth(2)
      const box = (await third.boundingBox())!
      const menu = editor.locator('.ctx-menu')
      // the menu opens through a renderer/main handshake; a click landing while
      // the window is still settling can be swallowed, so retry a couple of times
      const openMenu = async () => {
        for (let attempt = 0; attempt < 3; attempt++) {
          await editor.mouse.click(box.x + box.width / 2, box.y + box.height / 2, {
            button: 'right',
          })
          if (await menu.isVisible({ timeout: 3_000 }).catch(() => false)) return
          await editor.waitForTimeout(500)
        }
        await expect(menu).toBeVisible()
      }
      await openMenu()
      await menu.getByRole('button', { name: 'Set Numbering Value…' }).click()
      const valueDialog = editor.locator('.modal.list-value-dialog')
      await expect(valueDialog).toBeVisible()
      await valueDialog.getByLabel('Set value to').fill('5')
      await expect(valueDialog).toContainText('5.')
      await valueDialog.screenshot({ path: screenshotPath('docs-list-set-value') })
      await valueDialog.getByRole('button', { name: 'OK' }).click()
      await expect(items.nth(2)).toHaveAttribute('data-marker', '5.')

      // right-click again → Adjust List Indents… → text indent 2 cm, follow with space
      await openMenu()
      await menu.getByRole('button', { name: 'Adjust List Indents…' }).click()
      const indentsDialog = editor.locator('.modal.list-indents-dialog')
      await expect(indentsDialog).toBeVisible()
      await indentsDialog.getByLabel('Number position').fill('1 cm')
      await indentsDialog.getByLabel('Text indent').fill('2 cm')
      await indentsDialog.getByRole('button', { name: 'Follow number with' }).click()
      await editor.getByRole('option', { name: 'Space' }).click()
      await indentsDialog.screenshot({ path: screenshotPath('docs-list-adjust-indents') })
      await indentsDialog.getByRole('button', { name: 'OK' }).click()
      await expect(items.nth(2)).toHaveAttribute('data-suff', 'space')

      // Numbering gallery → Define New Number Format… → upper roman, on a new paragraph after the list
      await items.nth(2).click()
      await editor.keyboard.press('End')
      await editor.keyboard.press('Enter')
      await editor.keyboard.press('Enter')
      await editor.keyboard.type('roman')
      await editor.locator('.rb-split-wrap .rb-caret[aria-label="Numbering"]').click()
      await editor.getByRole('button', { name: 'Define New Number Format…' }).click()
      const numberDialog = editor.locator('.modal.list-number-dialog')
      await expect(numberDialog).toBeVisible()
      await numberDialog.getByRole('button', { name: 'Number style' }).click()
      await editor.getByRole('option', { name: 'I, II, III, ...', exact: true }).click()
      await numberDialog.screenshot({ path: screenshotPath('docs-list-define-number') })
      await numberDialog.getByRole('button', { name: 'OK' }).click()
      await expect(editor.locator('.doc-page .doc-li').last()).toHaveAttribute('data-marker', 'I.')

      expect(
        await editor.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!()),
      ).toBe(true)
      await expect
        .poll(() => part(docPath, 'word/numbering.xml'), { timeout: 15_000 })
        .toContain('<w:startOverride w:val="5"/>')
      const numbering = await part(docPath, 'word/numbering.xml')
      expect(numbering).toContain('<w:suff w:val="space"/>')
      expect(numbering).toContain('<w:ind w:left="1134" w:hanging="567"/>')
      expect(numbering).toContain('<w:numFmt w:val="upperRoman"/>')
      const body = await part(docPath, 'word/document.xml')
      expect(body.match(/<w:numPr>/g)?.length).toBe(4)
    } finally {
      await closeAndSaveVideo(launched, 'docs-list-completion')
    }
  })
})
