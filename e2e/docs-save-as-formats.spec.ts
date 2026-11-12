import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Save As offers PDF, HTML and Markdown next to Word. Picking one writes an
 * exported copy through the existing export; the open document keeps its own
 * .docx path and is not rewritten.
 */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

async function oneLineDocx(): Promise<Buffer> {
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Quarterly summary</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('Save As writes a PDF, HTML or Markdown copy and leaves the document on its .docx', async () => {
  test.setTimeout(180_000)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-save-as-formats-'))
  const docPath = join(dir, 'report.docx')
  await writeFile(docPath, await oneLineDocx())
  const original = await readFile(docPath)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-save-as-formats',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Quarterly summary' })).toBeVisible()

    const saveAsTo = async (target: string) => {
      await app.evaluate(({ dialog, shell }, out) => {
        dialog.showSaveDialog = (async (...args: unknown[]) => {
          const options = args[args.length - 1] as { filters?: { extensions: string[] }[] }
          ;(globalThis as Record<string, unknown>).__saveAsFilters = options.filters
          return { canceled: false, filePath: out }
        }) as unknown as typeof dialog.showSaveDialog
        shell.showItemInFolder = () => {}
      }, target)
      await app.evaluate(({ webContents }) => {
        webContents
          .getAllWebContents()
          .find((w) => w.getURL().includes('://docs/'))
          ?.send('menu:command', 'save-as')
      })
      await expect
        .poll(async () => (await stat(target).catch(() => null))?.size ?? 0, {
          timeout: 60_000,
        })
        .toBeGreaterThan(0)
    }

    const md = join(dir, 'copy.md')
    await saveAsTo(md)
    expect(await readFile(md, 'utf8')).toContain('Quarterly summary')
    const filters = await app.evaluate(
      () => (globalThis as Record<string, unknown>).__saveAsFilters as { extensions: string[] }[],
    )
    expect(filters.map((filter) => filter.extensions[0])).toEqual(['docx', 'pdf', 'html', 'md'])

    const html = join(dir, 'copy.html')
    await saveAsTo(html)
    expect(await readFile(html, 'utf8')).toContain('Quarterly summary')

    const pdf = join(dir, 'copy.pdf')
    await saveAsTo(pdf)
    expect((await readFile(pdf)).subarray(0, 5).toString()).toBe('%PDF-')

    // the document did not move to a copy and was not rewritten
    expect(await readFile(docPath)).toEqual(original)
    const shellPage = await waitForPageWithUrl(app, 'shell/out')
    await expect(shellPage.locator('.tab-bar .tab-item', { hasText: 'report.docx' })).toHaveCount(1)
  } finally {
    await closeAndSaveVideo(launched, 'docs-save-as-formats')
  }
})
