import { test, expect } from '@playwright/test'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Unequal-column sections (w:equalWidth="0" with explicit w:col children — the
 * shape CNKI/PDF-to-Word converters emit) collapsing to one-or-two-character
 * columns beside a blank page, with the page count changing between zoom
 * levels: the report in genoffice#1952 (screenshots at 100% vs 110%). The
 * fixture mirrors that document's structural skeleton — continuous sections
 * with custom column lists interleaved with single-column nextPage sections —
 * and the sweep pins, at several zoom levels across the whole document, that
 * no paragraph ever renders as a collapsed sliver and the total page count
 * stays constant (pagination must not depend on zoom).
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const SENTENCE =
  '本文采用小波阈值消噪，采取阈值以下全部置零、阈值以上保留原始值和阈值的差值的消噪策略。'

function cjkParagraphs(count: number, chars: number): string {
  const body = SENTENCE.repeat(Math.ceil(chars / SENTENCE.length)).slice(0, chars)
  return Array.from(
    { length: count },
    (_, i) => `<w:p><w:r><w:t>第${i + 1}段 ${body}</w:t></w:r></w:p>`,
  ).join('')
}

/** The docs URL can load with no editor mounted (the renderer birth race
 *  tracked upstream): a reload re-runs the bootstrap and recovers. */
async function awaitDocsEditor(page: import('@playwright/test').Page): Promise<void> {
  const editor = page.locator('.ProseMirror').first()
  try {
    await editor.waitFor({ timeout: 45_000 })
    return
  } catch {
    await page.reload({ waitUntil: 'domcontentloaded' })
    await editor.waitFor({ timeout: 90_000 })
  }
}

async function unequalColsDocx(): Promise<Buffer> {
  const sectPr = (inner: string): string =>
    `<w:p><w:pPr><w:sectPr>${inner}</w:sectPr></w:pPr></w:p>`
  const PAGE =
    '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
  const SINGLE = `<w:sectPr>${PAGE}</w:sectPr>`
  // the CNKI shape: unequal explicit columns on a continuous section
  const TWO_COL =
    '<w:sectPr><w:type w:val="continuous"/><w:cols w:num="2" w:equalWidth="0"><w:col w:w="4905" w:space="176"/><w:col w:w="4987"/></w:cols>' +
    PAGE +
    '</w:sectPr>'
  const xml =
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>` +
    cjkParagraphs(10, 400) +
    sectPr(SINGLE) +
    cjkParagraphs(14, 400) +
    sectPr(TWO_COL) +
    cjkParagraphs(10, 400) +
    `<w:sectPr>${PAGE}</w:sectPr>` +
    '</w:body></w:document>'
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file('word/document.xml', xml)
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('unequal-column sections never collapse and pagination ignores zoom', async () => {
  test.setTimeout(300_000)
  const artifacts = join(__dirname, 'artifacts')
  mkdirSync(artifacts, { recursive: true })
  const docPath = join(artifacts, 'unequal-cols-fixture.docx')
  writeFileSync(docPath, await unequalColsDocx())
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: docPath,
    videoDir: 'docs-unequal-cols-zoom-sweep',
  })
  const { app } = launched
  try {
    // the docs editor lives in its own window/view: address THAT page,
    // not launchShell's first window (which stays on Home)
    const page = await waitForPageWithUrl(app, '://docs/')
    await awaitDocsEditor(page)
    await page.waitForTimeout(2500)

    // collapse signature, normalized by zoom: a paragraph rendered as a
    // sliver (~1-2 characters per line) while holding a full sentence
    const scan = () =>
      page.evaluate(() => {
        const zoomFactor =
          parseFloat(
            getComputedStyle(document.querySelector<HTMLElement>('.doc-zoom')!).zoom ?? '1',
          ) || 1
        const hits: Array<{ text: string; w: number; top: number }> = []
        for (const p of document.querySelectorAll('.doc-page p')) {
          const r = p.getBoundingClientRect()
          const text = (p.textContent ?? '').replace(/\s+/g, '')
          if (
            text.length >= 6 &&
            r.width / zoomFactor < 48 &&
            r.width / zoomFactor >= -1 &&
            r.height / zoomFactor > 40
          ) {
            hits.push({
              text: text.slice(0, 24),
              w: Math.round(r.width / zoomFactor),
              top: Math.round(r.top + scrollY),
            })
          }
        }
        return { hits, zoomFactor }
      })

    const fullScan = async () => {
      const total = await page.evaluate(() => {
        const el = document.querySelector('.editor-scroll')
        return el ? el.scrollHeight : document.documentElement.scrollHeight
      })
      const step = await page.evaluate(() => Math.floor(innerHeight * 0.85))
      for (let y = 0; y < total; y += step) {
        await page.evaluate((top) => {
          const el = document.querySelector('.editor-scroll')
          if (el) el.scrollTo(0, top)
          else scrollTo(0, top)
        }, y)
        await page.waitForTimeout(600)
        const { hits } = await scan()
        if (hits.length > 0) return hits
      }
      return null
    }

    const setZoom = async (z: number) => {
      await page.evaluate((value) => {
        const slider = document.querySelector<HTMLInputElement>('.zoom-slider')!
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
        setter.call(slider, String(value))
        slider.dispatchEvent(new Event('input', { bubbles: true }))
        slider.dispatchEvent(new Event('change', { bubbles: true }))
      }, z)
      await page.waitForTimeout(1400)
    }

    const totals = new Set<number>()
    let collapse: unknown = null
    for (const z of [100, 90, 80, 110, 150]) {
      await setZoom(z)
      const status = await page.evaluate(
        () => document.querySelector('.status-page')?.textContent ?? '',
      )
      const total = Number(/(\d+)\s*$/.exec(status)?.[1] ?? 0)
      expect(total, `status bar at ${z}%`).toBeGreaterThan(0)
      totals.add(total)
      collapse = await fullScan()
      if (collapse) {
        console.log(`COLLAPSE at ${z}%`, JSON.stringify(collapse))
        await page.screenshot({ path: `artifacts/unequal-cols-collapse-z${z}.png` })
        break
      }
    }
    expect(collapse, 'no column collapse at any zoom').toBeNull()
    expect([...totals], 'total page count identical at every zoom').toHaveLength(1)
  } finally {
    await closeAndSaveVideo(launched, 'docs-unequal-cols-zoom-sweep')
  }
})

test('a corrupt w:cols count flows as one full-width column, not character slivers', async () => {
  test.setTimeout(300_000)
  const corrupt =
    '<?xml version="1.0" encoding="UTF-8"?><w:document ' +
    W +
    '><w:body>' +
    cjkParagraphs(16, 400) +
    '<w:sectPr><w:type w:val="continuous"/><w:cols w:num="200" w:space="720"/>' +
    '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
    '</w:body></w:document>'
  const artifacts = join(__dirname, 'artifacts')
  mkdirSync(artifacts, { recursive: true })
  const docPath = join(artifacts, 'corrupt-cols-fixture.docx')
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file('word/document.xml', corrupt)
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: docPath,
    videoDir: 'docs-unequal-cols-zoom-sweep',
  })
  const { app } = launched
  try {
    // the docs editor lives in its own window/view: address THAT page,
    // not launchShell's first window (which stays on Home)
    const page = await waitForPageWithUrl(app, '://docs/')
    await awaitDocsEditor(page)
    await page.waitForTimeout(2500)
    const hits = await page.evaluate(() => {
      const zoomFactor =
        parseFloat(
          getComputedStyle(document.querySelector<HTMLElement>('.doc-zoom')!).zoom ?? '1',
        ) || 1
      const out: Array<{ text: string; w: number }> = []
      for (const p of document.querySelectorAll('.doc-page p')) {
        const r = p.getBoundingClientRect()
        const text = (p.textContent ?? '').replace(/\s+/g, '')
        if (
          text.length >= 6 &&
          (r.width / zoomFactor < 48 || r.width / zoomFactor < 0) &&
          r.height / zoomFactor > 40
        ) {
          out.push({ text: text.slice(0, 24), w: Math.round(r.width / zoomFactor) })
        }
      }
      return out.slice(0, 8)
    })
    console.log('CORRUPT_COLS', JSON.stringify(hits))
    expect(hits, 'collapsed paragraphs under a corrupt cols count').toHaveLength(0)
  } finally {
    await closeAndSaveVideo(launched, 'docs-unequal-cols-zoom-sweep')
  }
})
