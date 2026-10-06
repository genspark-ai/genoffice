import { expect, test, type Locator, type Page } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * Withholding a span from the model, driven through the real UI: right-click a
 * selection, name it, and the words stay in the document and in the file while
 * the model's view is the placeholder. Screenshots are taken at each step so the
 * result can be looked at, not just asserted.
 */

const SECRET = '13800138000'
const LABEL = '客户电话'
const POLL = { timeout: 20_000 }

/** the tiptap instance both editors hang off their root element */
type EditorHandle = {
  chain: () => {
    focus: () => { setTextSelection: (r: { from: number; to: number }) => { run: () => boolean } }
  }
  state: { doc: { descendants: (fn: (n: unknown, pos: number) => void) => void } }
}

async function selectText(
  root: Locator,
  needle: string,
): Promise<{ from: number; to: number } | null> {
  return root.evaluate((el, n) => {
    const ed = (el as unknown as { editor: EditorHandle }).editor
    let hit: { from: number; to: number } | null = null
    ed.state.doc.descendants((node, pos) => {
      if (hit) return
      const textNode = node as { isText?: boolean; text?: string }
      if (!textNode.isText || !textNode.text) return
      const at = textNode.text.indexOf(n as string)
      if (at !== -1) hit = { from: pos + at, to: pos + at + (n as string).length }
    })
    if (hit) ed.chain().focus().setTextSelection(hit).run()
    return hit
  }, needle)
}

/**
 * Where `needle` sits on screen. Walking the DOM and building a Range is
 * deliberate: `window.getSelection()` is not reliable here, and a right-click
 * outside the live selection collapses it in Chromium — which is how the docs
 * run first picked up the word under the pointer instead of the marked one.
 */
async function pointOnWord(page: Page, needle: string): Promise<{ x: number; y: number }> {
  return page.evaluate((w) => {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT)
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const idx = node.textContent?.indexOf(w as string) ?? -1
      if (idx === -1) continue
      const range = document.createRange()
      range.setStart(node, idx)
      range.setEnd(node, idx + (w as string).length)
      const box = range.getBoundingClientRect()
      return { x: box.left + box.width / 2, y: box.top + box.height / 2 }
    }
    throw new Error(`word not found on screen: ${w}`)
  }, needle)
}

test.describe('withholding a span from the model', () => {
  test('markdown: name a selection, and the marker survives the save', async () => {
    test.setTimeout(180_000)
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-redact-md-'))
    const mdPath = join(dir, 'call.md')
    await writeFile(
      mdPath,
      `# Customer call\n\nCall ${SECRET} now to confirm the order.\n\nOur office opens at 09:00.\n`,
    )

    const launched = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'redact-md',
      openFile: mdPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://markdown/')
      const editor = page.locator('.doc-editor')
      await expect(page.locator('.doc-editor h1')).toHaveText('Customer call')

      // the feature ships behind a ribbon switch, off by default
      const toggle = page.locator(
        '.rb-btn[data-tip="Right-click a selection to hide it from the model."]',
      )
      await expect(toggle).toHaveCount(1)
      await page.screenshot({ path: screenshotPath('md-1-feature-off') })
      await toggle.click()
      await expect(toggle).toHaveClass(/active/)

      expect(await selectText(editor, SECRET)).not.toBeNull()
      await page.screenshot({ path: screenshotPath('md-2-selected') })

      const at = await pointOnWord(page, SECRET)
      await page.mouse.click(at.x, at.y, { button: 'right' })

      const menuItem = page.locator('.redact-menu-item')
      await expect(menuItem).toHaveText('Hide the selection from AI', { timeout: 15_000 })
      await page.screenshot({ path: screenshotPath('md-3-context-menu') })
      await menuItem.click()

      const dialog = page.locator('.redact-dialog')
      await expect(dialog).toBeVisible()
      // the seed is the selected text, and the preview shows the marker as typed
      await expect(dialog.locator('.redact-dialog-input')).toHaveValue(SECRET.slice(0, 24))
      await page.screenshot({ path: screenshotPath('md-4-dialog') })

      const input = dialog.locator('.redact-dialog-input')
      await input.fill(LABEL)
      await expect(dialog.locator('.redact-dialog-preview')).toHaveText(`{{${LABEL}}}`)
      await page.screenshot({ path: screenshotPath('md-5-dialog-preview') })

      await dialog.locator('.redact-dialog-btn.primary').click()
      await expect(dialog).toHaveCount(0)
      await expect(editor.locator('.redact-span')).toHaveCount(1)
      await expect(editor.locator('.redact-span')).toHaveAttribute('data-label', LABEL)
      // the real words are still in the document — the span is a mark over them
      await expect(editor).toContainText(SECRET)
      await page.screenshot({ path: screenshotPath('md-6-marked') })

      await page.keyboard.press('Meta+s')
      await expect
        .poll(async () => (await readFile(mdPath, 'utf8')).includes('data-redaction'), POLL)
        .toBe(true)
      const saved = await readFile(mdPath, 'utf8')
      expect(saved).toContain(SECRET)
      expect(saved).toContain(`data-label="${LABEL}"`)
      await page.screenshot({ path: screenshotPath('md-7-saved') })
    } finally {
      await closeAndSaveVideo(launched, 'redact-md')
    }
  })

  test('docs: name a selection, and the run keeps its words in the .docx', async () => {
    test.setTimeout(180_000)
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-redact-docs-'))
    const docxPath = join(dir, 'call.docx')
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
    await writeFile(
      docxPath,
      await buildDocx({
        bodyXml:
          `<w:p><w:r><w:t>Customer call</w:t></w:r></w:p>` +
          p(`Call ${SECRET} now to confirm the order.`) +
          p('Our office opens at 09:00.'),
      }),
    )
    expect(W).toContain('wordprocessingml')

    const launched = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'redact-docs',
      openFile: docxPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      const editor = page.locator('.doc-page[contenteditable="true"]')
      await expect(editor).toContainText('Customer call', { timeout: 30_000 })

      expect(
        await page.evaluate((n) => {
          const ed = (window as unknown as { __aidocs: { editor: EditorHandle } }).__aidocs.editor
          let hit: { from: number; to: number } | null = null
          ed.state.doc.descendants((node, pos) => {
            if (hit) return
            const textNode = node as { isText?: boolean; text?: string }
            if (!textNode.isText || !textNode.text) return
            const at = textNode.text.indexOf(n as string)
            if (at !== -1) hit = { from: pos + at, to: pos + at + (n as string).length }
          })
          if (hit) ed.chain().focus().setTextSelection(hit).run()
          return hit
        }, SECRET),
      ).not.toBeNull()
      await page.screenshot({ path: screenshotPath('docs-1-selected') })

      const at = await pointOnWord(page, SECRET)
      await page.mouse.click(at.x, at.y, { button: 'right' })

      const menuItem = page.locator('.ctx-menu .ctx-item', {
        hasText: 'Hide the selection from AI',
      })
      await expect(menuItem).toHaveCount(1, { timeout: 15_000 })
      await page.screenshot({ path: screenshotPath('docs-2-context-menu') })
      await menuItem.click()

      const dialog = page.locator('.modal', { hasText: 'Hide the selection from AI' })
      await expect(dialog).toBeVisible()
      const input = dialog.locator('input').first()
      await expect(input).toHaveValue(SECRET.slice(0, 24))
      await page.screenshot({ path: screenshotPath('docs-3-dialog') })

      await input.fill(LABEL)
      await expect(dialog).toContainText(`{{${LABEL}}}`)
      await page.screenshot({ path: screenshotPath('docs-4-dialog-preview') })

      await dialog.locator('.modal-actions button').last().click()
      await expect(dialog).toHaveCount(0)
      await expect(editor.locator('.redact-span')).toHaveCount(1)
      // the words themselves are untouched: a border, not a deletion
      await expect(editor).toContainText(SECRET)
      await page.screenshot({ path: screenshotPath('docs-5-marked') })

      await page.keyboard.press('Meta+s')
      // a .docx is a zip: document.xml is deflated, so the run properties can
      // only be read after inflating the part, not by scanning the raw bytes
      const readDocumentXml = async (): Promise<string> => {
        const zip = await JSZip.loadAsync(await readFile(docxPath))
        const part = zip.file('word/document.xml')
        if (!part) throw new Error('word/document.xml missing from the saved package')
        return part.async('string')
      }
      await expect
        .poll(async () => (await readDocumentXml()).includes('go:redact'), POLL)
        .toBe(true)

      const xml = await readDocumentXml()
      // the label rides in the file, and Word gets a real character border
      expect(xml).toContain(`w:label="${LABEL}"`)
      expect(xml).toContain('w:bdr')
      // and the words are still there — a mark over them, never a deletion
      expect(xml).toContain(SECRET)
      expect(xml).not.toContain('{{')
      await page.screenshot({ path: screenshotPath('docs-6-saved') })
    } finally {
      await closeAndSaveVideo(launched, 'redact-docs')
    }
  })
})
