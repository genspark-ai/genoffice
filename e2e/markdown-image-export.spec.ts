import { test, expect } from '@playwright/test'
import { mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PNG } from 'pngjs'
import { PDFDocument } from 'pdf-lib'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

async function exportMenu(app: import('@playwright/test').ElectronApplication, label: string) {
  await app.evaluate(({ Menu, BrowserWindow }, text) => {
    const find = (items: Electron.MenuItem[]): Electron.MenuItem | undefined => {
      for (const item of items) {
        if (item.label === text) return item
        const child = item.submenu && find(item.submenu.items)
        if (child) return child
      }
    }
    const menu = find(Menu.getApplicationMenu()!.items)
    if (!menu) throw new Error(`Missing File menu item: ${text}`)
    menu.click(undefined as never, BrowserWindow.getFocusedWindow()!, undefined as never)
  }, label)
}

test('Markdown File menu exports paginated PNG files with formula and local image', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'markdown-image-export-'))
  const source = join(dir, 'report.md')
  const image = new PNG({ width: 80, height: 50 })
  for (let p = 0; p < image.data.length; p += 4) {
    image.data[p] = 230
    image.data[p + 1] = 20
    image.data[p + 2] = 20
    image.data[p + 3] = 255
  }
  await writeFile(join(dir, 'red.png'), PNG.sync.write(image))
  await writeFile(
    source,
    '# Export test\n\n$E=mc^2$\n\n![red](red.png)\n\n' +
      Array.from(
        { length: 65 },
        (_, i) => `Paragraph ${i + 1}: pagination must preserve every line.`,
      ).join('\n\n'),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'markdown-image-export',
    openFile: source,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://markdown/')
    await expect(page.locator('.doc-editor h1')).toHaveText('Export test')
    await expect(page.locator('.doc-editor .katex')).toHaveCount(1)
    await page.waitForFunction(() =>
      [
        ...document.querySelectorAll<HTMLImageElement>(
          '.doc-editor img[src]:not(.ProseMirror-separator)',
        ),
      ].every((img) => img.complete && img.naturalWidth > 0),
    )
    await launched.app.evaluate(({ dialog, shell }, dir) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [dir] })
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: `${dir}/reference.pdf` })
      shell.showItemInFolder = () => {}
    }, dir)
    await exportMenu(launched.app, 'Export as Images…')
    await expect(page.locator('.status-export')).toContainText(
      /Exported \d+ images|Image export failed/,
      { timeout: 45_000 },
    )
    await expect(page.locator('.status-export')).toContainText(/Exported \d+ images/)
    const folders = (await readdir(dir)).filter((name) => name.startsWith('report-images-'))
    expect(folders).toHaveLength(1)
    const output = join(dir, folders[0])
    const files = (await readdir(output)).sort()
    expect(files.length).toBeGreaterThan(1)
    for (const file of files) {
      const png = PNG.sync.read(await readFile(join(output, file)))
      expect(png.width).toBeGreaterThan(1500)
      expect(png.height).toBeGreaterThan(2200)
      // Every page has painted content, not just the correct PNG header.
      let ink = 0
      for (let i = 0; i < png.data.length; i += 4) if (png.data[i] < 200) ink++
      expect(ink).toBeGreaterThan(100)
    }
    const firstBytes = await readFile(join(output, files[0]))
    await test.info().attach('exported-page-01', { body: firstBytes, contentType: 'image/png' })
    const first = PNG.sync.read(firstBytes)
    let red = 0
    for (let i = 0; i < first.data.length; i += 4)
      if (first.data[i] > 180 && first.data[i + 1] < 70 && first.data[i + 2] < 70) red++
    expect(red).toBeGreaterThan(500)

    // Both outputs must use the same print pagination.
    await exportMenu(launched.app, 'Export as PDF…')
    await expect.poll(async () => (await readdir(dir)).includes('reference.pdf')).toBe(true)
    const pdfBytes = await readFile(join(dir, 'reference.pdf'))
    expect(pdfBytes.toString('latin1')).toContain('KaTeX')
    const pdf = await PDFDocument.load(pdfBytes)
    expect(files.length).toBe(pdf.getPageCount())

    await launched.app.evaluate(({ dialog }) => {
      dialog.showOpenDialog = async () => ({ canceled: true, filePaths: [] })
    })
    // PDF export may open another tab; select the Markdown editor again.
    const shellPage = await waitForPageWithUrl(launched.app, 'shell/out')
    await shellPage.locator('.tab-bar .tab-item', { hasText: 'report.md' }).click()
    await exportMenu(launched.app, 'Export as Images…')
    await expect(page.locator('.status-export')).toHaveCount(0)
    expect((await readdir(dir)).filter((name) => name.startsWith('report-images-'))).toEqual(
      folders,
    )
  } finally {
    await closeAndSaveVideo(launched, 'markdown-image-export')
    await rm(dir, { recursive: true, force: true })
  }
})
