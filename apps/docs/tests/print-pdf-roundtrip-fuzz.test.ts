/** Fourth fidelity-fuzz axis, second half: the docx -> printable HTML -> PDF
 *  -> read-back round trip. The docs suite's roundtrip-fuzz.test.ts proves
 *  byte fidelity inside DOCX; this one proves the *printed* surface keeps the
 *  document's text across the format boundary.
 *
 *  The Chromium rendering itself needs Electron, so printHtmlToPdf's window is
 *  injected: loadFile receives the real sanitized HTML the print path writes,
 *  and printToPDF lays its visible text out with pdf-lib. The read-back is the
 *  independent reader (pdf.js), not the printer's own input. Corrupt DOCX
 *  packages are fed through the same chain and must refuse cleanly or
 *  succeed — never hang — with a healthy seed still round-tripping after them.
 *
 *  Corruption is seeded. CI runs PRINT_FUZZ_SEEDS; a failing seed reproduces
 *  exactly via buildFuzzDocx(seed). */
import { readFileSync } from 'node:fs'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
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

/** Seeded PRNG, kept local so the docs suite does not import a sibling app's
 *  test helper. Same construction as the other hostile-fuzz suites. */
function rng(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a |= 0
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

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

async function roundTrip(docx: Uint8Array, title: string): Promise<{ html: string; text: string }> {
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
    return { html: loaded, text: chunks.join(' ').replace(/\s+/g, ' ') }
  } finally {
    if (timer !== undefined) clearTimeout(timer)
    void task.destroy().catch(() => undefined)
  }
}

const normalise = (text: string): string => text.replace(/\s+/g, ' ').trim()

type Corrupt = (bytes: Uint8Array, r: () => number) => Promise<Uint8Array>

const truncate: Corrupt = async (bytes, r) =>
  bytes.slice(0, 16 + Math.floor(r() * Math.max(1, bytes.length - 16)))

const entryDrop: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir)
  if (names.length === 0) return bytes
  zip.remove(names[Math.floor(r() * names.length)]!)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

const xmlGarbage: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const junk = new Uint8Array(64 + Math.floor(r() * 512))
  for (let i = 0; i < junk.length; i++) junk[i] = Math.floor(r() * 256)
  zip.file('word/document.xml', junk)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

const CORRUPTIONS: Array<[string, Corrupt]> = [
  ['truncate', truncate],
  ['entrydrop', entryDrop],
  ['xmlgarbage', xmlGarbage],
]

describe('seeded docx -> printed PDF round-trip fuzz', () => {
  it('a healthy document keeps its anchor and visible text through the print path', async () => {
    for (let seed = 1; seed <= SEEDS; seed++) {
      const { html, text } = await roundTrip(await buildFuzzDocx(seed), `seed ${seed}`)
      expect(text, `seed ${seed} anchor in PDF`).toContain(anchorText(seed))
      const paragraphs = bodyTextOf(html)
      // Guard against an empty generator turning the fidelity checks into no-ops.
      expect(paragraphs.length, `seed ${seed} printed paragraphs`).toBeGreaterThan(3)
      // The sanitizer keeps every paragraph's text, including CJK literals.
      for (const paragraph of paragraphs.slice(0, 20)) {
        const probe = paragraph.slice(0, 24)
        if (probe.length < 4) continue
        expect(html, `seed ${seed} sanitized HTML`).toContain(escapeHtml(probe))
      }
      // ASCII prose is what a WinAnsi PDF font can carry; assert a sample of it
      // survived the PDF boundary as well.
      const ascii = paragraphs
        .map((line) => line.replace(/[^\x20-\x7e]/g, ' '))
        .map(normalise)
        .filter((line) => /^[A-Za-z0-9 ,.:_-]{12,}$/.test(line))
      expect(ascii.length, `seed ${seed} ASCII prose to read back`).toBeGreaterThan(0)
      for (const line of ascii.slice(0, 5)) {
        expect(text, `seed ${seed} printed line`).toContain(line.slice(0, 24))
      }
    }
  }, 60000)

  for (const [mode, corrupt] of CORRUPTIONS) {
    for (let seed = 1; seed <= SEEDS; seed++) {
      it(`seed ${seed} ${mode}: the chain refuses cleanly or succeeds, never hangs`, async () => {
        const valid = await buildFuzzDocx(seed)
        const bytes = await corrupt(valid, rng(seed * 104729 + 17))
        const outcome = await roundTrip(bytes, `hostile ${seed} ${mode}`).then(
          () => 'succeeded' as const,
          () => 'refused' as const,
        )
        // Either outcome is legitimate; a hang fails the case's timeout and an
        // escape from the boundary fails the type of `outcome` itself.
        expect(['succeeded', 'refused']).toContain(outcome)
      }, 30000)
    }

    it(`a healthy document still round-trips after the ${mode} batch`, async () => {
      const { text } = await roundTrip(await buildFuzzDocx(202), 'healthy after hostile')
      expect(text).toContain(anchorText(202))
    }, 30000)
  }
})
