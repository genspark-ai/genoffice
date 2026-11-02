import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { test, expect, type Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Pagination invariants in the real editor, on documents that exercise tables,
 * columns, sections, floats and tab stops. Every safety net the pass chain owns
 * (`__pageDebug`) must stay silent, the model-side BlockPlan must agree with
 * the DOM scrape (I9), the three slicings of one document (canvas, preview,
 * exported PDF) must agree on the page count, and a forced re-measure of the
 * untouched document must reproduce the same pages and block geometry.
 *
 * A document known to break an invariant is listed in KNOWN_FAILURES with the
 * hazard it tracks (X1..X10 in the pagination loop review) and runs as
 * `test.fixme`, so the entry is retired by the fix instead of hiding it.
 */

const CORPUS = resolve(__dirname, '../apps/docs/tests/pagination-corpus/docx')
/** synthetic float-table / anchored-picture / table-width documents (geometry/build.py) */
const GEOMETRY = resolve(__dirname, '../apps/docs/tests/pagination-corpus/geometry')
const TRUTHS = resolve(__dirname, '../packages/docx-engine/tests/fixtures/truths')

const DOCS: Array<{ name: string; file: string }> = [
  { name: 'long-table', file: join(CORPUS, '05-long-table.docx') },
  { name: 'tblheader-repeat', file: join(CORPUS, '23-tblheader-repeat.docx') },
  { name: 'oversized-table', file: join(CORPUS, '19-oversized-table.docx') },
  { name: 'mixed-content-tables', file: join(CORPUS, '15-mixed-content.docx') },
  { name: 'two-columns', file: join(CORPUS, '22-two-columns.docx') },
  { name: 'continuous-section', file: join(CORPUS, '21-continuous-section.docx') },
  { name: 'multi-section', file: join(CORPUS, '08-multi-section-paper.docx') },
  { name: 'float-anchor', file: join(TRUTHS, 'render-anchor-own-line-band/probe.docx') },
  { name: 'tab-stops', file: join(TRUTHS, 'tab-stop-no-min-advance/probe.docx') },
  {
    name: 'bottom-float-section-break',
    file: join(TRUTHS, 'render-bottom-float-section-break/probe.docx'),
  },
  { name: 'geo-float-tables-body', file: join(GEOMETRY, 'geo-float-tables-body.docx') },
  { name: 'geo-float-tables-nested-hf', file: join(GEOMETRY, 'geo-float-tables-nested-hf.docx') },
  { name: 'geo-anchored-pictures', file: join(GEOMETRY, 'geo-anchored-pictures.docx') },
  { name: 'geo-table-overwide-centred', file: join(GEOMETRY, 'geo-table-overwide-centred.docx') },
  { name: 'geo-table-pct-nonuniform', file: join(GEOMETRY, 'geo-table-pct-nonuniform.docx') },
  { name: 'ruled-header', file: join(TRUTHS, 'render-hf-side-wrap-flow/probe.docx') },
  { name: 'header-anchors', file: join(TRUTHS, 'render-hf-anchor-reserve/probe.docx') },
]

/** doc name -> hazard id and note */
const KNOWN_FAILURES: Record<string, string> = {}

interface DebugSlice {
  start: number
  end: number
  section: number
}

interface DebugBlock {
  top: number
  height: number
}

interface PageDebug {
  slices: DebugSlice[]
  blocks: DebugBlock[]
  followUps: number
  retriggerLog: string[]
  loopStops: number
  gapBoundaryMismatch: number
  /** I11: PageCut facts vs the gap builder's DOM re-derivation (lead, capacity, hf variant) */
  pageCutMismatch: number
  pageCutDiag: string[]
  tabFreezes: number
  sliceIterationsCapped: boolean
  sliceRunaway: boolean
  rowSplitOverflows: number
  gaps: number
  /** I9: model-side BlockPlan vs the DOM scrape, one entry per disagreeing fact */
  planDiff: Array<{ cls: string; block: number; plan: unknown; scrape: unknown }>
  /** float channel writes (--tblp-dy, doc-table-float-flow, --page-float-dy) per pass of the chain */
  layoutHealth: { passes: FlipCounts[]; flipsAfterFirst: FlipCounts }
}

interface FlipCounts {
  tblpDy: number
  floatFlow: number
  pageFloatDy: number
  carry: number
}

interface DebugWindow {
  __pageDebug?: PageDebug
  __previewDebug?: {
    slices: DebugSlice[]
    shadowed: boolean
    previewCutMismatch: number
    previewCutDiag: string[]
  }
  /** set before opening the preview: it re-slices its snapshot and compares (previewCutMismatch) */
  __previewShadow?: boolean
  __openPagePreview?: () => void
  __exportPdf?: () => Promise<boolean>
}

const readDebug = (page: Page) =>
  page.evaluate(() => {
    const dbg = (window as unknown as DebugWindow).__pageDebug!
    return {
      slices: dbg.slices.map((s) => ({ start: s.start, end: s.end, section: s.section })),
      blocks: dbg.blocks.map((b) => ({ top: b.top, height: b.height })),
      followUps: dbg.followUps,
      retriggerLog: dbg.retriggerLog,
      loopStops: dbg.loopStops,
      gapBoundaryMismatch: dbg.gapBoundaryMismatch,
      pageCutMismatch: dbg.pageCutMismatch,
      pageCutDiag: dbg.pageCutDiag,
      tabFreezes: dbg.tabFreezes,
      sliceIterationsCapped: dbg.sliceIterationsCapped,
      sliceRunaway: dbg.sliceRunaway,
      rowSplitOverflows: dbg.rowSplitOverflows,
      // a multi-cell row cut gets a zero-height .page-gap-cut marker instead of a gap band
      gaps: document.querySelectorAll(
        '.doc-page .page-gap:not(.page-gap-carry), .doc-page .page-gap-cut:not(.page-cut-overlay)',
      ).length,
      planDiff: dbg.planDiff,
      layoutHealth: dbg.layoutHealth,
    } satisfies PageDebug
  })

/** wait until pagination stops mutating: page count, last-page bottom and block geometry stable */
async function settle(page: Page): Promise<PageDebug> {
  await page.waitForSelector('.doc-page', { timeout: 30_000 })
  await page.evaluate(() => document.fonts.ready.then(() => undefined))
  const deadline = Date.now() + 40_000
  let prev = ''
  let stableSince = Date.now()
  while (Date.now() < deadline) {
    const cur = await page.evaluate(() => {
      const last = document.querySelectorAll('.doc-page')[0]
      const dbg = (window as unknown as DebugWindow).__pageDebug
      const geo = dbg?.blocks
        .map((b) => `${Math.round(b.top * 4)}/${Math.round(b.height * 4)}`)
        .join(',')
      return `${dbg?.slices.length ?? 0}:${last ? Math.round(last.getBoundingClientRect().bottom) : 0}:${geo ?? ''}`
    })
    if (cur !== prev) {
      prev = cur
      stableSince = Date.now()
    } else if (Date.now() - stableSince > 2_000) {
      return readDebug(page)
    }
    await page.waitForTimeout(200)
  }
  throw new Error(`doc layout never settled (last: ${prev.slice(0, 80)})`)
}

const sliceKey = (slices: DebugSlice[]) =>
  slices.map((s) => `${s.start.toFixed(1)}-${s.end.toFixed(1)}@${s.section}`).join(' ')

async function pdfPageCount(path: string): Promise<number> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  return doc.numPages
}

test.describe('pagination invariants (rendered)', () => {
  for (const d of DOCS) {
    const known = KNOWN_FAILURES[d.name]
    const run = known ? test.fixme : test
    run(`${d.name}${known ? ` (known: ${known})` : ''}`, async () => {
      test.setTimeout(240_000)
      const outDir = mkdtempSync(join(tmpdir(), 'genoffice-pagination-invariants-'))
      const launched = await launchShell({
        onboardingSeen: true,
        videoDir: `docs-pagination-invariants-${d.name}`,
        openFile: d.file,
      })
      try {
        const page = await waitForPageWithUrl(launched.app, '://docs/')
        // tall enough for width-fit to settle at 100%
        await launched.app.evaluate(({ BrowserWindow }) => {
          BrowserWindow.getAllWindows()[0]?.setBounds({ x: 0, y: 0, width: 1360, height: 1500 })
        })
        const dbg = await settle(page)
        const chain = dbg.retriggerLog.join(' > ')

        // (a) every safety net stayed silent
        expect(dbg.loopStops, `follow-up chain refused as a loop: ${chain}`).toBe(0)
        expect(
          dbg.retriggerLog.some((r) => /!(cap|repeat)$/.test(r)),
          chain,
        ).toBe(false)
        expect(dbg.tabFreezes, 'tab-stop measurer froze').toBe(0)
        expect(dbg.sliceIterationsCapped, 'slice fixed point hit its pass bound').toBe(false)
        expect(dbg.sliceRunaway, 'column-turn limit fired').toBe(false)
        expect(dbg.rowSplitOverflows, 'row split past its fragment budget').toBe(0)
        expect(dbg.gapBoundaryMismatch, 'page gaps built must equal page boundaries').toBe(0)
        expect(
          dbg.pageCutMismatch,
          `I11: PageCut agrees with the gap builder: ${dbg.pageCutDiag.join(' | ')}`,
        ).toBe(0)
        expect(dbg.gaps, 'one gap widget per page boundary').toBe(dbg.slices.length - 1)
        expect(dbg.planDiff, 'I9: BlockPlan agrees with the measureBlocks scrape').toEqual([])
        expect(
          dbg.layoutHealth.flipsAfterFirst,
          `float channels rewritten after the first pass: ${JSON.stringify(dbg.layoutHealth.passes)}`,
        ).toEqual({ tblpDy: 0, floatFlow: 0, pageFloatDy: 0, carry: 0 })

        // (c) a forced full re-measure of the untouched document changes nothing
        await page.evaluate(() => document.fonts.dispatchEvent(new Event('loadingdone')))
        await page.waitForTimeout(400)
        const again = await settle(page)
        expect(sliceKey(again.slices), 'slices after a forced re-measure').toBe(
          sliceKey(dbg.slices),
        )
        expect(again.blocks.length).toBe(dbg.blocks.length)
        for (let i = 0; i < dbg.blocks.length; i++) {
          expect(
            Math.abs(again.blocks[i].top - dbg.blocks[i].top) +
              Math.abs(again.blocks[i].height - dbg.blocks[i].height),
            `block ${i} geometry after a forced re-measure`,
          ).toBeLessThan(0.5)
        }
        expect(again.loopStops, `re-measure chain: ${again.retriggerLog.join(' > ')}`).toBe(0)
        expect(
          again.followUps,
          `re-measure chain: ${again.retriggerLog.join(' > ')}`,
        ).toBeLessThanOrEqual(1)
        for (const [i, p] of again.layoutHealth.passes.entries())
          expect(p, `re-measure pass ${i} rewrote a float channel`).toEqual({
            tblpDy: 0,
            floatFlow: 0,
            pageFloatDy: 0,
            carry: 0,
          })

        // (b) canvas == preview == exported PDF: the sheets print the canvas
        // PageCut list; the preview's shadow re-slicing of its own snapshot must
        // agree with it (previewCutMismatch), and the PDF page count with both
        await page.evaluate(() => {
          const w = window as unknown as DebugWindow
          w.__previewShadow = true
          w.__openPagePreview!()
        })
        await expect
          .poll(() => page.locator('.pv-page').count(), { timeout: 30_000 })
          .toBe(dbg.slices.length)
        await page.waitForTimeout(1_000)
        const previewDbg = await page.evaluate(
          () => (window as unknown as DebugWindow).__previewDebug,
        )
        const preview = previewDbg?.slices ?? []
        expect(await page.locator('.pv-page').count(), 'preview sheets').toBe(dbg.slices.length)
        expect(preview.length, 'preview slicing').toBe(dbg.slices.length)
        expect(previewDbg?.shadowed, 'preview shadow ran').toBe(true)
        expect(
          previewDbg?.previewCutMismatch,
          `preview shadow slicing vs canvas cuts: ${previewDbg?.previewCutDiag.join(' | ')}`,
        ).toBe(0)
        for (let i = 0; i < dbg.slices.length; i++) {
          expect(
            Math.abs(preview[i].start - dbg.slices[i].start) +
              Math.abs(preview[i].end - dbg.slices[i].end),
            `preview page ${i + 1} window vs canvas`,
          ).toBeLessThan(0.5)
          expect(preview[i].section, `preview page ${i + 1} section`).toBe(dbg.slices[i].section)
        }
        // (d) one strip geometry: the canvas gap strip and the preview strip of a
        // page are the same height (same DOM, same width, same floating drawings)
        const strips = await page.evaluate(() => {
          const h = (el: Element | null) => (el ? el.getBoundingClientRect().height : null)
          const out: Array<{ page: number; kind: string; canvas: number; preview: number | null }> =
            []
          for (const el of document.querySelectorAll<HTMLElement>('.doc-page .page-gap-hf')) {
            const idx = Number(el.dataset.hfPageIdx)
            const kind = el.dataset.hfKind!
            const pv = document.querySelectorAll('.pv-page')[idx]
            out.push({
              page: idx + 1,
              kind,
              canvas: h(el)!,
              preview: h(pv?.querySelector(`.page-hf-${kind}`) ?? null),
            })
          }
          return out
        })
        for (const s of strips) {
          if (s.preview === null) continue
          expect(
            Math.abs(s.canvas - s.preview),
            `${s.kind} strip height on page ${s.page}: canvas ${s.canvas} vs preview ${s.preview}`,
          ).toBeLessThan(1)
        }

        const pdfPath = join(outDir, `${d.name}.pdf`)
        await launched.app.evaluate(({ dialog, shell }, out) => {
          dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
          shell.showItemInFolder = () => {}
        }, pdfPath)
        expect(
          await page.evaluate(() => (window as unknown as DebugWindow).__exportPdf!()),
          'PDF export',
        ).toBe(true)
        expect(await pdfPageCount(pdfPath), 'exported PDF pages').toBe(dbg.slices.length)
      } finally {
        await closeAndSaveVideo(launched, `docs-pagination-invariants-${d.name}`)
      }
    })
  }
})
