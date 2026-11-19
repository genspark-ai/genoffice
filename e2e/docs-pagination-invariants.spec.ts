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
  {
    name: 'geo-float-table-split-boxes',
    file: join(GEOMETRY, 'geo-float-table-split-boxes.docx'),
  },
  {
    name: 'geo-footnoted-table-tail-seam',
    file: join(GEOMETRY, 'geo-footnoted-table-tail-seam.docx'),
  },
  { name: 'geo-toc-narrow-column', file: join(GEOMETRY, 'geo-toc-narrow-column.docx') },
  { name: 'ruled-header', file: join(TRUTHS, 'render-hf-side-wrap-flow/probe.docx') },
  { name: 'header-anchors', file: join(TRUTHS, 'render-hf-anchor-reserve/probe.docx') },
]

/** doc name -> hazard id and note */
const KNOWN_FAILURES: Record<string, string> = {
  // the page-2 cut is a one-line tie: pass 1 splits the row at 1728.0, the fonts-loadingdone re-measure
  // moves it whole to page 3 (1694.2) and the document grows 2.0 px (2538.1 -> 2540.1). Snapping the
  // row-fit comparison to the LayoutUnit grid did not change it (CI run 42779841935), so this is not
  // float drift: the re-measure reads the split-state DOM (row splitExtra / translated cell children)
  // and decides the row differently. Needs an Electron session to pin.
  'geo-table-pct-nonuniform':
    'row-split decision differs between first pass and split-state re-measure',
}

interface DebugSlice {
  start: number
  end: number
  section: number
}

interface DebugBlock {
  top: number
  height: number
}

interface DebugFloatSplit {
  blockTop: number
  dyPx: number
  cutYs: number[]
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
  /** one `start:end:section:lead:...:flags` record per page (pagination-cuts pageCutSig) */
  pageCutSig: string
  pageCutDiag: string[]
  /** T4.5: page readers (float shifts, cell boxes, frames, status bar, thumbnails)
   *  derive the same page from the cuts as from the gap widgets */
  pageReaderMismatch: number
  pageReaderDiag: string[]
  pageReaderShadowed: boolean
  floatSplits: DebugFloatSplit[]
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
  /** set before a pass: every page reader also walks the gap widgets and compares (pageReaderMismatch) */
  __pageReaderShadow?: boolean
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
      blocks: dbg.blocks.map((b) => ({
        top: b.top,
        height: b.height,
      })),
      followUps: dbg.followUps,
      retriggerLog: dbg.retriggerLog,
      loopStops: dbg.loopStops,
      gapBoundaryMismatch: dbg.gapBoundaryMismatch,
      pageCutMismatch: dbg.pageCutMismatch,
      pageCutSig: dbg.pageCutSig,
      pageCutDiag: dbg.pageCutDiag,
      pageReaderMismatch: dbg.pageReaderMismatch,
      pageReaderDiag: dbg.pageReaderDiag,
      pageReaderShadowed: dbg.pageReaderShadowed,
      floatSplits: dbg.floatSplits,
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

        // (c) a forced full re-measure of the untouched document changes nothing;
        // run it with the page-reader shadow on: every reader must derive the same
        // page from the cuts as from the gap widgets
        await page.evaluate(() => {
          ;(window as unknown as DebugWindow).__pageReaderShadow = true
          document.fonts.dispatchEvent(new Event('loadingdone'))
        })
        await page.waitForTimeout(400)
        const again = await settle(page)
        expect(again.pageReaderShadowed, 'page-reader shadow ran').toBe(true)
        expect(
          again.pageReaderMismatch,
          `page readers vs cuts: ${again.pageReaderDiag.join(' | ')}`,
        ).toBe(0)
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

        // a page opening on a footnoted table's bottom edge is a tail seam (its
        // bottom border stays on the table's page, no rule atop the next one)
        if (d.name === 'geo-footnoted-table-tail-seam') {
          const cuts = dbg.pageCutSig.split('|')
          expect(cuts.length, dbg.pageCutSig).toBe(2)
          expect(cuts[1].split(':').pop(), dbg.pageCutSig).toBe('T')
        }

        // TOC leaders are sized against the title as wrapped at the section's
        // width: every leader + page unit closes its title's last line
        if (d.name === 'geo-toc-narrow-column') {
          const drops = await page.evaluate(() =>
            [...document.querySelectorAll<HTMLElement>('.ProseMirror .doc-toc-line')].map((el) => {
              const title = el.querySelector('.doc-toc-title')!.getBoundingClientRect()
              const tail = el.querySelector('.doc-toc-tail')!.getBoundingClientRect()
              return `${el.textContent?.slice(0, 12)}:${tail.top >= title.bottom - 1 ? 'dropped' : 'inline'}`
            }),
          )
          expect(drops).toHaveLength(4)
          expect(drops.filter((d) => d.endsWith(':dropped'))).toEqual([])
        }

        // (b) H7: a split floating table's later pages own the boxes on them. The
        // page-positioned cell box in the second portion sits on that portion's
        // paper (page turn = any non-carry gap widget, the in-table ones included)
        // and the slicer agrees; the anchor paragraph's page-relative-X picture
        // keeps its paper offset on the page the carry moved it to
        if (d.name === 'geo-float-table-split-boxes') {
          const own = await page.evaluate(() => {
            const pm = document.querySelector<HTMLElement>('.doc-page')!
            const zoom =
              parseFloat(getComputedStyle(pm.closest<HTMLElement>('.doc-zoom')!).zoom || '1') || 1
            const turns = [...document.querySelectorAll<HTMLElement>('.doc-page .page-gap')]
              .filter((g) => !g.classList.contains('page-gap-carry'))
              .map((g) => ({
                top: g.getBoundingClientRect().top,
                bottom: g.getBoundingClientRect().bottom,
                mt: parseFloat(g.style.getPropertyValue('--gap-mt')) || 0,
              }))
              .sort((a, b) => a.top - b.top)
            const pageOf = (y: number) => turns.filter((t) => t.top < y).length
            const paperTopOf = (pageIdx: number) =>
              pageIdx === 0
                ? pm.getBoundingClientRect().top
                : turns[pageIdx - 1].bottom - turns[pageIdx - 1].mt * zoom
            const box = pm.querySelector<HTMLElement>(
              '.doc-cell-boxes-page > .doc-textbox[data-cell-page]',
            )
            const tr = box?.closest('tr')
            const table = tr?.closest('table')
            const rows = table
              ? [...table.querySelectorAll(':scope > tbody > tr')].filter(
                  (r) =>
                    !r.classList.contains('page-gap') &&
                    !r.classList.contains('page-repeat-header'),
                )
              : []
            const floatGaps = table
              ? table.querySelectorAll(':scope > tbody > tr.page-gap').length
              : 0
            const img = pm.querySelector<HTMLElement>('img[data-page-rel-x]')
            const r = (el: Element | null | undefined) => el?.getBoundingClientRect()
            return {
              zoom,
              floatGaps,
              rowIndex: tr ? rows.indexOf(tr) : -1,
              rowsAbovePx: tr
                ? rows
                    .slice(0, rows.indexOf(tr))
                    .reduce((s, r) => s + r.getBoundingClientRect().height, 0) / zoom
                : NaN,
              rowPage: tr ? pageOf(r(tr)!.top + 1) : -1,
              pinPage: box ? Number(box.dataset.pinPage) : -1,
              boxTop: r(box)?.top ?? NaN,
              boxPaperTop: tr ? paperTopOf(pageOf(r(tr)!.top + 1)) : NaN,
              boxOffY: box ? parseFloat(box.dataset.cellPageY ?? '') : NaN,
              imgPage: img ? pageOf(r(img)!.top + 1) : -1,
              imgLeft: r(img)?.left ?? NaN,
              pmLeft: pm.getBoundingClientRect().left,
              imgOffX: img ? parseFloat(img.dataset.pageRelXPx ?? '') : NaN,
            }
          })
          expect(own.floatGaps, 'the floating table splits across pages').toBeGreaterThan(0)
          expect(own.rowPage, 'the box host row sits on a later page').toBeGreaterThan(0)
          expect(own.pinPage, 'cell box pinned to its row page').toBe(own.rowPage)
          expect(
            Math.abs(own.boxTop - (own.boxPaperTop + own.boxOffY * own.zoom)),
            'cell box at its page-relative Y on that paper',
          ).toBeLessThan(1.5)
          // the slicer's page of the host row: its cut list is the gap list
          const split = dbg.floatSplits[0]
          expect(split, 'slicer split the floating table').toBeTruthy()
          const rowTop = split.blockTop + split.dyPx + own.rowsAbovePx
          const slicerPage = dbg.slices.filter((s, i) => i > 0 && s.start <= rowTop + 0.5).length
          expect(own.pinPage, 'cell box page equals the slicer page of its row').toBe(slicerPage)
          expect(own.imgPage, 'anchor paragraph picture lands on a later page').toBeGreaterThan(0)
          expect(
            Math.abs(own.imgLeft - (own.pmLeft + own.imgOffX * own.zoom)),
            'page-relative X picture measured from the paper edge',
          ).toBeLessThan(1.5)
        }

        // (e) canvas == preview == exported PDF: the sheets print the canvas
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
