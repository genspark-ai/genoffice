/** Fourth fidelity-fuzz axis, second half: the docx -> editor text ->
 *  printable-HTML sanitizer path.
 *
 *  The docs suite's roundtrip-fuzz.test.ts proves byte fidelity inside DOCX;
 *  this one proves the text survives the format boundary on the way to the
 *  printable page — the helper the AI create-document command uses to turn
 *  model-written HTML into a printable document. It is NOT the path a .docx
 *  takes when the user prints or exports to PDF: that path renders the
 *  pagination preview's sheets in Chromium, which needs Electron.
 *
 *  The Chromium rendering is stubbed: loadFile receives the real sanitized HTML
 *  the helper writes, and printToPDF lays its visible text out with pdf-lib.
 *  The read-back is the independent reader (pdf.js), not the printer's own
 *  input.
 *
 *  The corruption half of the original draft duplicated
 *  roundtrip-hostile.test.ts (on the default seeds every refusal comes out of
 *  parseDocx, which that suite already covers), so it is gone: this axis is the
 *  printable-HTML fidelity, not the parse.
 *
 *  Corruption is seeded. CI runs PRINT_FUZZ_SEEDS; a failing seed reproduces
 *  exactly via buildFuzzDocx(seed). */
import { readFileSync } from 'node:fs'
import { Editor } from '@tiptap/core'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  buildPrintableHtml,
  printHtmlToPdf,
  type PrintWindow,
} from '../../../packages/electron-utils/src/print-html-pdf'
import { buildFuzzDocx } from './helpers/fuzz-docx'

const SEEDS = Math.max(1, Number(process.env.PRINT_FUZZ_SEEDS ?? 8))
const LOAD_TIMEOUT_MS = 8000
const anchorText = (seed: number): string => `Anchor seed ${seed}`

/** Collect the text the editor would render — the same DOM the print path sees. */
async function printedTextOf(docx: Uint8Array): Promise<string[]> {
  const parsed = await parseDocx(docx)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  try {
    const out: string[] = []
    const walk = (node: unknown): void => {
      if (node === null || typeof node !== 'object') return
      const n = node as { text?: unknown; content?: unknown }
      if (typeof n.text === 'string' && n.text.trim() !== '') out.push(n.text)
      if (Array.isArray(n.content)) for (const child of n.content) walk(child)
    }
    walk(editor.getJSON())
    return out
  } finally {
    editor.destroy()
  }
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Visible text of the sanitized printable page, decoded back from entities. */
function bodyTextOf(html: string): string[] {
  const body = /<body[^>]*>([\s\S]*?)<\/body>/i.exec(html)?.[1] ?? ''
  return body
    .split(/<\/p>|<\/h[1-6]>|<br\s*\/?>/i)
    .map((chunk) =>
      chunk
        .replace(/<[^>]*>/g, '')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&')
        .replace(/\s+/g, ' ')
        .trim(),
    )
    .filter((line) => line.length > 0)
}

/** Stand-in for Chromium: one printed line per paragraph, Helvetica/WinAnsi.
 *  Non-WinAnsi glyphs (CJK) are blanks in the PDF but remain in the HTML, so
 *  the HTML-layer assertion below is what covers them. */
async function renderHtmlToPdf(html: string): Promise<Buffer> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const lines = bodyTextOf(html).map((line) => line.replace(/[^\x20-\x7e]/g, ' '))
  let page = doc.addPage([595, 842])
  let y = 790
  for (const line of lines) {
    if (y < 40) {
      page = doc.addPage([595, 842])
      y = 790
    }
    page.drawText(line.slice(0, 200), { x: 40, y, size: 9, font })
    y -= 14
  }
  return Buffer.from(await doc.save())
}

async function roundTrip(
  docx: Uint8Array,
  title: string,
): Promise<{ html: string; text: string; source: string[] }> {
  const paragraphs = await printedTextOf(docx)
  const body = paragraphs.map((text) => `<p>${escapeHtml(text)}</p>`).join('\n')
  const html = buildPrintableHtml(title, body)
  let loaded = ''
  const win: PrintWindow = {
    loadFile: async (filePath: string) => {
      loaded = readFileSync(filePath, 'utf8')
    },
    destroy: () => undefined,
    webContents: {
      printToPDF: async () => renderHtmlToPdf(loaded),
    },
  }
  const pdf = await printHtmlToPdf(html, () => win)
  const task = getDocument({
    data: new Uint8Array(pdf).slice(),
    verbosity: 0,
  })
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    const doc = await Promise.race([
      task.promise,
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => reject(new Error('pdf.js read-back timed out')), LOAD_TIMEOUT_MS)
      }),
    ])
    const chunks: string[] = []
    for (let i = 1; i <= Math.min(doc.numPages, 64); i++) {
      const page = await doc.getPage(i)
      const content = await page.getTextContent()
      chunks.push(...content.items.map((item) => ('str' in item ? item.str : '')))
      await page.cleanup()
    }
    // The source paragraphs travel with the result: the HTML is compared with
    // what the editor produced, never with itself.
    return { html: loaded, text: chunks.join(' ').replace(/\s+/g, ' '), source: paragraphs }
  } finally {
    if (timer !== undefined) clearTimeout(timer)
    void task.destroy().catch(() => undefined)
  }
}

const normalise = (text: string): string => text.replace(/\s+/g, ' ').trim()

describe('seeded docx -> editor text -> printable HTML fuzz', () => {
  it('a healthy document keeps its anchor and visible text through the sanitizer', async () => {
    for (let seed = 1; seed <= SEEDS; seed++) {
      const { html, text, source } = await roundTrip(await buildFuzzDocx(seed), `seed ${seed}`)
      expect(text, `seed ${seed} anchor in PDF`).toContain(anchorText(seed))
      // Guard against an empty generator turning the fidelity checks into no-ops.
      expect(source.length, `seed ${seed} source paragraphs`).toBeGreaterThan(3)
      // The sanitizer keeps every paragraph's text, including CJK literals. The
      // comparison is with the editor's own text: comparing the printed HTML
      // with the HTML it was decoded from could never fail.
      for (const paragraph of source.slice(0, 20)) {
        const probe = paragraph.slice(0, 24)
        if (probe.length < 4) continue
        expect(html, `seed ${seed} sanitized HTML`).toContain(escapeHtml(probe))
      }
      // ASCII prose is what a WinAnsi PDF font can carry; assert a sample of it
      // survived the PDF boundary as well.
      const ascii = source
        .map((line) => line.replace(/[^\x20-\x7e]/g, ' '))
        .map(normalise)
        .filter((line) => /^[A-Za-z0-9 ,.:_-]{12,}$/.test(line))
      expect(ascii.length, `seed ${seed} ASCII prose to read back`).toBeGreaterThan(0)
      for (const line of ascii.slice(0, 5)) {
        expect(text, `seed ${seed} printed line`).toContain(line.slice(0, 24))
      }
    }
  }, 60000)
})
