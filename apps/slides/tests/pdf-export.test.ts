import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { PNG } from 'pngjs'
import { afterEach, describe, expect, it } from 'vitest'

import {
  buildPdfExportHtml,
  exportPageWidthIn,
  exportSlidesPdf,
  type PdfExportWindow,
} from '../src/main/pdf-export'

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
      pages: [{ png: firstPng }, { png: secondPng }],
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
    // vector pages carry their bitmaps as SVG <image>, outside document.images
    expect(win.executedScript).toContain("querySelectorAll('svg image')")
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
      pages: [{ png: 'png' }],
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
      pages: [{ png: 'png' }],
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

  it('overlays hyperlink rects as anchors so printToPDF emits link annotations', async () => {
    const win = new TestPdfWindow()
    const result = await exportSlidesPdf({
      pages: [{ png: 'png1' }, { png: 'png2' }],
      widthPx: 1600,
      heightPx: 900,
      filePath: await outputPath(),
      links: [
        [
          { x: 0.1, y: 0.2, w: 0.25, h: 0.05, href: 'https://x.test/?a=1&b="2"' },
          { x: 0.5, y: 0.5, w: 0.1, h: 0.1, href: '#pg2' },
          // corrupt geometry must not produce NaN% markup
          { x: Number.NaN, y: 0, w: 1, h: 1, href: 'https://dropped.test/' },
        ],
        [],
      ],
      createWindow: () => win,
      openExportedPdf: () => {},
    })

    expect(result.ok).toBe(true)
    // pages carry ids so "#pgN" anchors become in-document PDF destinations
    expect(win.loadedHtml).toContain('<div class="page" id="pg1">')
    expect(win.loadedHtml).toContain('<div class="page" id="pg2">')
    expect(win.loadedHtml).toContain(
      '<a href="https://x.test/?a=1&amp;b=&quot;2&quot;" style="left:10%;top:20%;width:25%;height:5%"></a>',
    )
    expect(win.loadedHtml).toContain(
      '<a href="#pg2" style="left:50%;top:50%;width:10%;height:10%"></a>',
    )
    expect(win.loadedHtml).not.toContain('dropped.test')
    expect(win.loadedHtml).not.toContain('NaN')
  })

  it('inlines SVG pages as-is and scopes the font CSS to the head', () => {
    const svg =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540"><text x="1" y="2">Hi</text></svg>'
    const html = buildPdfExportHtml(
      [{ svg }, { png: 'png' }],
      13.333,
      7.5,
      undefined,
      "@font-face{font-family:'Carlito';src:url(data:font/ttf;base64,AAAA)}</style><script>1</script>",
    )
    expect(html).toContain(`<div class="page" id="pg1">${svg}</div>`)
    expect(html).toContain('<div class="page" id="pg2"><img src="data:image/png;base64,png"></div>')
    expect(html).toContain("@font-face{font-family:'Carlito'")
    // a stray </style> in the CSS cannot break out of the style block
    expect(html).not.toContain('</style><script>')
    expect(html.match(/<\/style>/g)).toHaveLength(1)
    expect(html).toContain('.page img, .page > svg { display: block; width: 100%; height: 100%; }')
  })

  it('emits no anchors when the export carries no links', () => {
    const html = buildPdfExportHtml([{ png: 'png' }], 13.333, 7.5)
    expect(html).toContain('<div class="page" id="pg1">')
    expect(html).not.toContain('<a ')
  })

  it('keeps real capture dimensions exact and falls back for degenerate ones', async () => {
    expect(exportPageWidthIn(1600, 900)).toBe(13.333)
    expect(exportPageWidthIn(4, 3)).toBe(10)
    // zeroed/negative/non-finite captures fall back to 16:9 instead of 0/NaN/Infinity pages
    for (const [w, h] of [
      [0, 900],
      [1600, 0],
      [-1600, 900],
      [Number.NaN, 900],
      [1600, Number.POSITIVE_INFINITY],
    ] as const) {
      expect(exportPageWidthIn(w, h)).toBe(13.333)
    }
    // absurd ratios clamp instead of emitting poster-sized pages
    expect(exportPageWidthIn(100000, 10)).toBe(37.5)
    expect(exportPageWidthIn(10, 100000)).toBe(1.5)

    const win = new TestPdfWindow()
    const result = await exportSlidesPdf({
      pages: [{ png: 'png' }],
      widthPx: 0,
      heightPx: 0,
      filePath: await outputPath(),
      createWindow: () => win,
      openExportedPdf: () => {},
    })
    expect(result.ok).toBe(true)
    expect(win.loadedHtml).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
    expect(win.printOptions).toMatchObject({ pageSize: { width: 13.333, height: 7.5 } })
  })
})
