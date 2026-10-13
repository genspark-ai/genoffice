import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * A Word-style table of contents must survive the PDF export as *clickable*
 * links. Chromium turns `href="#name"` into a /Link annotation resolved through
 * the catalog's named destinations, so this checks the bytes that come out:
 * every TOC name has a /Dests entry, and every entry lands on the page that
 * actually prints its target heading.
 *
 * The unit tests can only assert the DOM. The pagination preview prints one
 * full-document clone per page, hides the editor canvas ahead of them, and
 * React rewrites every clone's innerHTML from a cached page HTML string — so the
 * interesting failures (a duplicated id, a destination lost to the chunked
 * merge) are invisible to jsdom and only show up in the exported file.
 */

interface AidocsWindow {
  __exportPdf?: () => Promise<boolean>
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

/** chapter headings, each carrying a Word internal bookmark */
const CHAPTERS = ['Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot']
const tocName = (i: number) => `_Toc90000${i}`

/**
 * A TOC document: the contents page hyperlinks to Word internal `_Toc…` names,
 * and each chapter starts on its own page with a matching bookmarkStart. The
 * parser files `_Toc…` under hiddenBookmarks — reading only the user-facing
 * bookmark attr leaves such a document with no anchors at all.
 */
async function tocDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  const toc = CHAPTERS.map(
    (name, i) =>
      `<w:p><w:hyperlink w:anchor="${tocName(i)}"><w:r><w:t>${name}</w:t></w:r></w:hyperlink></w:p>`,
  ).join('')
  const chapters = CHAPTERS.map(
    (name, i) =>
      `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>` +
      `<w:bookmarkStart w:id="${100 + i}" w:name="${tocName(i)}"/>` +
      `<w:r><w:t>${name} chapter body</w:t></w:r>` +
      `<w:bookmarkEnd w:id="${100 + i}"/></w:p>` +
      Array.from({ length: 6 }, () => `<w:p><w:r><w:t>filler line</w:t></w:r></w:p>`).join(''),
  ).join('')
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
      '<w:p><w:r><w:t>Contents</w:t></w:r></w:p>' +
      toc +
      chapters +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** named destinations in the PDF -> name -> 1-based page number */
async function namedDests(path: string): Promise<Map<string, number>> {
  const { PDFDocument, PDFName } = await import('pdf-lib')
  const pdf = await PDFDocument.load(readFileSync(path))
  const raw = pdf.context.lookup(pdf.catalog.get(PDFName.of('Dests')))
  const out = new Map<string, number>()
  if (!raw || typeof (raw as { entries?: unknown }).entries !== 'function') return out
  for (const [key, value] of (raw as never as Map<string, unknown>).entries()) {
    const arr = pdf.context.lookup(value)
    if (!arr || typeof (arr as { size?: unknown }).size !== 'function') continue
    const target = String((arr as never as { get(i: number): unknown }).get(0))
    const idx = pdf.getPages().findIndex((p) => p.ref.toString() === target)
    if (idx >= 0) out.set(String(key).replace(/^\//, ''), idx + 1)
  }
  return out
}

/** every /Link annotation: the destination name it points at */
async function linkNames(path: string): Promise<string[]> {
  const { PDFDocument, PDFName } = await import('pdf-lib')
  const pdf = await PDFDocument.load(readFileSync(path))
  const names: string[] = []
  pdf.getPages().forEach((p) => {
    const annots = pdf.context.lookup(p.node.get(PDFName.of('Annots')))
    const list =
      annots && typeof (annots as { asArray?: unknown }).asArray === 'function'
        ? (annots as never as { asArray(): unknown[] }).asArray()
        : []
    for (const ref of list) {
      const a = pdf.context.lookup(ref)
      if (a.get(PDFName.of('Subtype'))?.toString() === '/Link') {
        names.push(String(a.get(PDFName.of('Dest'))).replace(/^\//, ''))
      }
    }
  })
  return names
}

/** the 1-based page each chapter's body text actually prints on */
async function chapterTextPages(path: string): Promise<number[]> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  const pages: string[] = []
  for (let i = 1; i <= doc.numPages; i++) {
    const content = await (await doc.getPage(i)).getTextContent()
    pages.push(
      content.items
        .map((item) => ('str' in item ? item.str : ''))
        .join('')
        .replace(/\s+/g, ' '),
    )
  }
  return CHAPTERS.map((c) => pages.findIndex((t) => t.includes(`${c} chapter body`)) + 1)
}

/** wait until the pagination preview has laid out at least `count` sheets */
async function waitForPages(page: Page, count: number): Promise<void> {
  await expect
    .poll(() => page.locator('.pv-page').count(), { timeout: 60_000 })
    .toBeGreaterThanOrEqual(count)
}

test.describe('docs TOC links in the exported PDF', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-toc-')))
    docPath = join(dir, 'toc.docx')
    writeFileSync(docPath, await tocDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('every TOC entry resolves to the page that prints its chapter', async () => {
    test.setTimeout(240_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-bookmark-links',
      openFile: docPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__exportPdf),
        undefined,
        { timeout: 30_000 },
      )
      const viewTab = page.locator('.ribbon-tab', { hasText: /^View$/ })
      if (!/\bactive\b/.test((await viewTab.getAttribute('class')) ?? '')) await viewTab.click()
      await page.getByRole('button', { name: 'Page Preview', exact: true }).click()
      await waitForPages(page, CHAPTERS.length + 1)

      const outPath = join(dir, 'toc.pdf')
      await launched.app.evaluate(({ dialog: d, shell }, out) => {
        d.showSaveDialog = async () => ({ canceled: false, filePath: out })
        shell.showItemInFolder = () => {}
      }, outPath)
      expect(await page.evaluate(() => (window as unknown as AidocsWindow).__exportPdf!())).toBe(
        true,
      )

      const dests = await namedDests(outPath)
      const links = await linkNames(outPath)
      const printed = await chapterTextPages(outPath)

      // every Word internal TOC name made it into the catalog, and every entry
      // on the contents page became a link annotation that points at one
      expect([...dests.keys()].sort()).toEqual(CHAPTERS.map((_, i) => tocName(i)).sort())
      expect(links.filter((n) => n.startsWith('_Toc')).sort()).toEqual(
        CHAPTERS.map((_, i) => tocName(i)).sort(),
      )

      // the destinations are not merely present: each one lands on the page its
      // chapter actually prints on, in document order, after the contents page
      expect(CHAPTERS.map((_, i) => dests.get(tocName(i)))).toEqual(printed)
      expect(new Set(printed).size).toBe(CHAPTERS.length)
      expect(Math.min(...printed)).toBeGreaterThan(1)
    } finally {
      await closeAndSaveVideo(launched, 'docs-bookmark-links')
    }
  })
})
