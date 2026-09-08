import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { PNG } from 'pngjs'
import { afterEach, describe, expect, it } from 'vitest'

import { exportSlidesPdf, type PdfExportWindow } from '../src/main/pdf-export'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function outputPath(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-slides-pdf-test-'))
  roots.push(root)
  return join(root, 'export.pdf')
}

function singlePixelPngBase64(): string {
  const png = new PNG({ width: 1, height: 1 })
  png.data.set([0x12, 0x34, 0x56, 0xff])
  return PNG.sync.write(png).toString('base64')
}

function deterministicNoisePngBase64(): string {
  const png = new PNG({ width: 800, height: 800 })
  let state = 0x12345678
  const nextByte = (): number => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    return state >>> 24
  }
  for (let i = 0; i < png.data.length; i += 4) {
    png.data[i] = nextByte()
    png.data[i + 1] = nextByte()
    png.data[i + 2] = nextByte()
    png.data[i + 3] = 255
  }
  return PNG.sync.write(png).toString('base64')
}

class TestPdfWindow implements PdfExportWindow {
  loadedPath: string | null = null
  loadedHtml = ''
  destroyed = false
  tempDirectoryExistedAtDestroy: boolean | null = null
  shouldFailLoad = false
  shouldFailPrint = false
  executedScript: string | null = null
  executedWithUserGesture: boolean | undefined
  printOptions: Electron.PrintToPDFOptions | null = null

  async loadFile(path: string): Promise<void> {
    this.loadedPath = path
    this.loadedHtml = await readFile(path, 'utf8')
    if (this.shouldFailLoad) throw new Error('load failed')
  }

  webContents = {
    executeJavaScript: async (script: string, userGesture?: boolean): Promise<void> => {
      this.executedScript = script
      this.executedWithUserGesture = userGesture
    },
    printToPDF: async (options: Electron.PrintToPDFOptions): Promise<Buffer> => {
      this.printOptions = options
      if (this.shouldFailPrint) throw new Error('print failed')
      return Buffer.from('PDF')
    },
  }

  destroy(): void {
    if (this.loadedPath) this.tempDirectoryExistedAtDestroy = existsSync(dirname(this.loadedPath))
    this.destroyed = true
  }
}

describe('slides PDF export', () => {
  it('loads a temporary HTML file containing all PNGs, writes the PDF, then removes the directory', async () => {
    const win = new TestPdfWindow()
    const filePath = await outputPath()
    const opened: string[] = []
    const chromiumDataUrlLimit = 2 * 1024 * 1024
    const firstPng = singlePixelPngBase64()
    const secondPng = deterministicNoisePngBase64()
    const decodedSecondPng = PNG.sync.read(Buffer.from(secondPng, 'base64'))

    expect(decodedSecondPng.width).toBeGreaterThan(0)
    expect(decodedSecondPng.height).toBeGreaterThan(0)

    const result = await exportSlidesPdf({
      pngsBase64: [firstPng, secondPng],
      widthPx: 1600,
      heightPx: 900,
      filePath,
      createWindow: () => win,
      openExportedPdf: (path) => opened.push(path),
    })

    expect(result).toEqual({ ok: true, path: filePath })
    expect(basename(win.loadedPath!)).toBe('slides.html')
    expect(basename(dirname(win.loadedPath!))).toMatch(/^genoffice-slides-pdf-/)
    expect(win.loadedHtml).toContain(`data:image/png;base64,${firstPng}`)
    expect(win.loadedHtml.length).toBeGreaterThan(chromiumDataUrlLimit)
    expect(win.loadedHtml.endsWith(`${secondPng}"></div></body></html>`)).toBe(true)
    expect(win.loadedHtml).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
    expect(win.executedScript).toContain('document.fonts.ready')
    expect(win.executedWithUserGesture).toBe(true)
    expect(win.printOptions).toEqual({
      landscape: false,
      printBackground: true,
      pageSize: { width: 13.333, height: 7.5 },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      preferCSSPageSize: false,
    })
    await expect(readFile(filePath)).resolves.toEqual(Buffer.from('PDF'))
    expect(opened).toEqual([filePath])
    expect(existsSync(dirname(win.loadedPath!))).toBe(false)
    expect(win.destroyed).toBe(true)
    expect(win.tempDirectoryExistedAtDestroy).toBe(true)
  })

  it('removes the temporary directory and destroys the window when loading fails', async () => {
    const win = new TestPdfWindow()
    win.shouldFailLoad = true

    const result = await exportSlidesPdf({
      pngsBase64: ['png'],
      widthPx: 4,
      heightPx: 3,
      filePath: await outputPath(),
      createWindow: () => win,
      openExportedPdf: () => {},
    })

    expect(result).toEqual({ ok: false, error: 'Error: load failed' })
    expect(win.loadedPath).not.toBeNull()
    expect(existsSync(dirname(win.loadedPath!))).toBe(false)
    expect(win.destroyed).toBe(true)
  })

  it('removes the temporary directory and destroys the window when PDF printing fails', async () => {
    const win = new TestPdfWindow()
    win.shouldFailPrint = true

    const result = await exportSlidesPdf({
      pngsBase64: ['png'],
      widthPx: 4,
      heightPx: 3,
      filePath: await outputPath(),
      createWindow: () => win,
      openExportedPdf: () => {},
    })

    expect(result).toEqual({ ok: false, error: 'Error: print failed' })
    expect(win.loadedPath).not.toBeNull()
    expect(existsSync(dirname(win.loadedPath!))).toBe(false)
    expect(win.destroyed).toBe(true)
  })
})
