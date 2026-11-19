/**
 * A document that starts with a picture opens with a caret before the
 * picture, not with the picture selected: the first key typed leaves it alone.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
)

const emu = (px: number) => px * 9525

const picturePara = (w: number, h: number) =>
  `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="1" name="Picture 1"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

async function fixture(body: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  zip.file('word/media/image1.png', PNG)
  return zip.generateAsync({ type: 'nodebuffer' })
}

type EditorWindow = {
  __aidocs?: {
    editor?: {
      state: { selection: { toJSON(): { type: string } }; doc: { textContent: string } }
    }
  }
}

const selectionType = (page: Page) =>
  page.evaluate(
    () => (window as unknown as EditorWindow).__aidocs?.editor?.state.selection.toJSON().type,
  )

const editorFocused = (page: Page) =>
  page.evaluate(() => Boolean(document.activeElement?.closest('.doc-page')))

for (const [name, body] of [
  ['a picture alone', picturePara(200, 150)],
  ['a picture before text', picturePara(200, 150) + para('After')],
] as const) {
  test(`${name}: the document opens with a caret before the picture`, async () => {
    test.setTimeout(120_000)
    const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-leading-picture-')))
    const docPath = join(dir, 'picture.docx')
    writeFileSync(docPath, await fixture(body))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-open-leading-picture',
      openFile: docPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      const img = page.locator('.doc-page[contenteditable="true"] img')
      await expect(img).toHaveCount(1, { timeout: 30_000 })
      await expect.poll(() => editorFocused(page), { timeout: 15_000 }).toBe(true)
      expect(await selectionType(page)).not.toBe('node')

      await page.keyboard.press('Backspace')
      await expect(img).toHaveCount(1)

      // typing goes into a paragraph of its own, ahead of the picture
      await page.keyboard.type('x')
      await expect(img).toHaveCount(1)
      await expect(page.locator('.doc-page[contenteditable="true"]').first()).toContainText('x')
    } finally {
      await closeAndSaveVideo(launched, 'docs-open-leading-picture')
      rmSync(dir, { recursive: true, force: true })
    }
  })
}
