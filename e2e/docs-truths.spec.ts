import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { test, expect, type Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Executable Word truths, rendered: every fixture under
 * packages/docx-engine/tests/fixtures/truths with class "render" opens in the real
 * Docs editor and its truth.json cases are checked against the pagination debug
 * snapshot (`__pageDebug`: page slices, gapless block geometry, split-row heights)
 * and against DOM line boxes. Every fixture also asserts the pass invariants:
 * gaps built == page boundaries, follow-up passes settle, and no re-trigger reason
 * repeats back to back (the ping-pong signature of a measurement that reads its
 * own decoration). The pure-formula half lives in apps/docs/tests/truths.test.ts.
 */

const TRUTHS_DIR = resolve(__dirname, '../packages/docx-engine/tests/fixtures/truths')
const PT_PER_PX = 72 / 96

interface TruthCase {
  kind: string
  note?: string
  tol?: number
  block?: number
  page?: number
  control?: number
  from?: number
  to?: number
  nth?: number
  expect: unknown
  status?: 'failing'
}

interface Truth {
  id: string
  title: string
  section: string
  class: 'pure' | 'render'
  docx?: string
  status?: 'failing'
  cases: TruthCase[]
  dir: string
}

interface DebugBlock {
  top: number
  height: number
  docxIndex: number | null
}

interface PageDebug {
  slices: unknown[]
  followUps: number
  retrigger: string[]
  retriggerLog: string[]
  loopStops?: number
  gapBoundaryMismatch?: number
  blocks: DebugBlock[]
  tableRows: Array<{ top: number; rows: Array<{ h: number }> }>
}

interface TruthWindow {
  __pageDebug?: PageDebug
}

const truths: Truth[] = readdirSync(TRUTHS_DIR, { withFileTypes: true })
  .filter((d) => d.isDirectory() && existsSync(join(TRUTHS_DIR, d.name, 'truth.json')))
  .map((d) => ({
    ...(JSON.parse(readFileSync(join(TRUTHS_DIR, d.name, 'truth.json'), 'utf8')) as Truth),
    dir: join(TRUTHS_DIR, d.name),
  }))
  .filter((t) => t.class === 'render')
  .sort((a, b) => a.id.localeCompare(b.id))

/** wait until pagination stops mutating: page count, last-page bottom and block geometry stable */
async function settle(page: Page): Promise<PageDebug> {
  await page.waitForSelector('.doc-page', { timeout: 30_000 })
  await page.evaluate(() => document.fonts.ready.then(() => undefined))
  const deadline = Date.now() + 40_000
  let prev = ''
  let stableSince = Date.now()
  while (Date.now() < deadline) {
    const cur = await page.evaluate(() => {
      const pages = document.querySelectorAll('.doc-page')
      const last = pages[pages.length - 1]
      const dbg = (window as unknown as TruthWindow).__pageDebug
      const geo = dbg?.blocks
        .map((b) => `${Math.round(b.top * 4)}/${Math.round(b.height * 4)}`)
        .join(',')
      return `${pages.length}:${last ? Math.round(last.getBoundingClientRect().bottom) : 0}:${geo ?? ''}`
    })
    if (cur !== prev) {
      prev = cur
      stableSince = Date.now()
    } else if (Date.now() - stableSince > 2_000) {
      return page.evaluate(() => {
        const dbg = (window as unknown as TruthWindow).__pageDebug!
        return {
          slices: dbg.slices,
          followUps: dbg.followUps,
          retrigger: dbg.retrigger,
          retriggerLog: dbg.retriggerLog,
          loopStops: dbg.loopStops,
          gapBoundaryMismatch: dbg.gapBoundaryMismatch,
          blocks: dbg.blocks,
          tableRows: dbg.tableRows,
        }
      })
    }
    await page.waitForTimeout(200)
  }
  throw new Error(`doc layout never settled (last: ${prev.slice(0, 80)})`)
}

function blockAt(dbg: PageDebug, docxIndex: number): DebugBlock {
  const b = dbg.blocks.find((x) => x.docxIndex === docxIndex)
  if (!b) throw new Error(`no block with docxIndex ${docxIndex}`)
  return b
}

/** 1-based page a block starts on (block tops and slices share the gapless flow) */
function pageOfBlock(dbg: PageDebug, docxIndex: number): number {
  const b = blockAt(dbg, docxIndex)
  const slices = dbg.slices as Array<{ start: number; end: number }>
  const i = slices.findIndex((s) => b.top >= s.start - 0.5 && b.top < s.end)
  if (i < 0) throw new Error(`block ${docxIndex} at ${b.top} lies on no page`)
  return i + 1
}

/** the block's DOM node: blocks carry data-idx = docxIndex (passthrough chips and
 *  gap widgets sit between them, so child ordinals do not map to blocks) */
const BLOCK_SEL = (docxIndex: number) => `.editor-scroll .ProseMirror > [data-idx="${docxIndex}"]`

/** line box height (px, zoom 1) of a block element that must hold one line */
async function singleLineBoxPx(page: Page, docxIndex: number): Promise<number> {
  return page.evaluate((sel) => {
    const el = document.querySelector<HTMLElement>(sel)
    if (!el) throw new Error(`no block element ${sel}`)
    const range = document.createRange()
    range.selectNodeContents(el)
    // inline rects are glyph boxes; the block's content box is the line box once there is one line
    const tops = new Set(Array.from(range.getClientRects()).map((r) => Math.round(r.top)))
    if (tops.size !== 1) throw new Error(`${sel} lays out ${tops.size} lines, expected one`)
    const cs = getComputedStyle(el)
    const zoomHost = el.closest('.doc-zoom') as HTMLElement | null
    const zoom = zoomHost ? parseFloat(getComputedStyle(zoomHost).zoom || '1') || 1 : 1
    const inner =
      el.getBoundingClientRect().height -
      (parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom)) -
      (parseFloat(cs.borderTopWidth) + parseFloat(cs.borderBottomWidth))
    return inner / zoom
  }, BLOCK_SEL(docxIndex))
}

/** top (px, zoom 1) of the first glyph box of a block: Word's baseline distances compare as glyph-top distances */
async function firstTextTopPx(page: Page, docxIndex: number): Promise<number> {
  return page.evaluate((sel) => {
    const pm = document.querySelector('.editor-scroll .ProseMirror')
    const el = document.querySelector<HTMLElement>(sel)
    if (!pm || !el) throw new Error(`no block element ${sel}`)
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
    const first = walker.nextNode()
    if (!first) throw new Error(`${sel} has no text`)
    const range = document.createRange()
    range.selectNodeContents(first)
    const rect = range.getClientRects()[0]
    const zoomHost = el.closest('.doc-zoom') as HTMLElement | null
    const zoom = zoomHost ? parseFloat(getComputedStyle(zoomHost).zoom || '1') || 1 : 1
    const origin = pm.getBoundingClientRect().top
    return (rect.top - origin) / zoom
  }, BLOCK_SEL(docxIndex))
}

/** top (pt) of the first glyph box on the 1-based page, measured from that page's top;
 *  page frames are rebuilt from the gap widgets exactly like pageFramesFromGaps. `body`
 *  reads the flow text, `header` the page's header strip (the first page's edge area,
 *  later pages' gap strips) */
async function pageTextTopPt(
  page: Page,
  pageNo: number,
  part: 'body' | 'header' = 'body',
): Promise<number> {
  return page.evaluate(
    ({ n, part }) => {
      const pm = document.querySelector('.editor-scroll .ProseMirror')
      const wrap = pm?.closest('.page-wrap') as HTMLElement | null
      if (!pm || !wrap) throw new Error('no editor page wrap')
      const zoomHost = pm.closest('.doc-zoom') as HTMLElement | null
      const zoom = zoomHost ? parseFloat(getComputedStyle(zoomHost).zoom || '1') || 1 : 1
      const wr = wrap.getBoundingClientRect()
      const gaps = Array.from(
        wrap.querySelectorAll<HTMLElement>('.page-gap:not(.page-gap-carry), .page-gap-cut'),
      )
        .map((g) => {
          const r = g.getBoundingClientRect()
          const cs = getComputedStyle(g)
          return {
            top: (r.top - wr.top) / zoom,
            height: r.height / zoom,
            mb: parseFloat(cs.getPropertyValue('--gap-mb')) || 0,
            mt: parseFloat(cs.getPropertyValue('--gap-mt')) || 0,
          }
        })
        .sort((a, b) => a.top - b.top)
      const frames: Array<{ top: number; bottom: number }> = []
      let top = 0
      for (const g of gaps) {
        frames.push({ top, bottom: g.top + g.mb })
        top = g.top + g.height - g.mt
      }
      frames.push({ top, bottom: wr.height / zoom })
      const f = frames[n - 1]
      if (!f) throw new Error(`no page frame ${n} (${frames.length} pages)`)
      // header/footer strips of later pages live inside the gap widgets: body text only
      const inBody = (node: Node) =>
        !node.parentElement?.closest('.page-gap, .page-gap-cut, .page-float-host')
      const inHeader = (node: Node) => {
        const strip = node.parentElement?.closest<HTMLElement>('.page-hf-header, .page-gap-hf')
        if (!strip) return false
        return strip.classList.contains('page-gap-hf')
          ? strip.dataset.hfKind === 'header' && strip.dataset.hfPageIdx === String(n - 1)
          : n === 1
      }
      const root = part === 'header' ? wrap : pm
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
        acceptNode: (node) =>
          node.textContent &&
          node.textContent.trim() !== '' &&
          (part === 'header' ? inHeader(node) : inBody(node))
            ? NodeFilter.FILTER_ACCEPT
            : NodeFilter.FILTER_REJECT,
      })
      let best = Infinity
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        const range = document.createRange()
        range.selectNodeContents(t)
        for (const r of Array.from(range.getClientRects())) {
          const y = (r.top - wr.top) / zoom
          if (r.height === 0 || y < f.top - 0.5 || y >= f.bottom) continue
          if (y < best) best = y
        }
      }
      if (best === Infinity) throw new Error(`no ${part} text on page ${n}`)
      return ((best - f.top) * 72) / 96
    },
    { n: pageNo, part },
  )
}

/** top (pt) of the first nested table of a top-level table block, measured from the
 *  content-box top of the cell that holds it (Word positions a floated nested table
 *  against its cell, not the page) */
async function nestedTableTopPt(page: Page, docxIndex: number, nth: number): Promise<number> {
  return page.evaluate(
    ({ sel, nth }) => {
      const nested = document.querySelectorAll<HTMLElement>(`${sel} td table`)[nth]
      if (!nested) throw new Error(`no nested table ${nth} in ${sel}`)
      const td = nested.closest('td')!
      const cs = getComputedStyle(td)
      const contentTop =
        td.getBoundingClientRect().top + parseFloat(cs.borderTopWidth) + parseFloat(cs.paddingTop)
      const zoomHost = td.closest('.doc-zoom') as HTMLElement | null
      const zoom = zoomHost ? parseFloat(getComputedStyle(zoomHost).zoom || '1') || 1 : 1
      return (((nested.getBoundingClientRect().top - contentTop) / zoom) * 72) / 96
    },
    { sel: BLOCK_SEL(docxIndex), nth },
  )
}

/** height (pt) of the nth equation box of a block: the inline-math span wraps the
 *  MathML, so its border box is the equation's line extents */
async function mathBoxPt(page: Page, docxIndex: number, nth: number): Promise<number> {
  return page.evaluate(
    ({ sel, nth }) => {
      const el = document.querySelectorAll<HTMLElement>(`${sel} span[data-inline-math]`)[nth]
      if (!el) throw new Error(`no equation ${nth} in ${sel}`)
      const zoomHost = el.closest('.doc-zoom') as HTMLElement | null
      const zoom = zoomHost ? parseFloat(getComputedStyle(zoomHost).zoom || '1') || 1 : 1
      return ((el.getBoundingClientRect().height / zoom) * 72) / 96
    },
    { sel: BLOCK_SEL(docxIndex), nth },
  )
}

test.describe('executable Word truths (rendered)', () => {
  test.skip(
    process.platform !== 'darwin',
    'line geometry is probed against the macOS font set; run on macOS',
  )

  for (const t of truths) {
    test(`${t.id}: ${t.title}`, async () => {
      const launched = await launchShell({
        onboardingSeen: true,
        videoDir: `docs-truths-${t.id}`,
        openFile: join(t.dir, t.docx ?? 'probe.docx'),
      })
      try {
        const editorPage = await waitForPageWithUrl(launched.app, '://docs/')
        // tall enough for width-fit to settle at 100%: page geometry below is in CSS px
        await launched.app.evaluate(({ BrowserWindow }) => {
          const win = BrowserWindow.getAllWindows()[0]
          win?.setBounds({ x: 0, y: 0, width: 1360, height: 1500 })
        })
        const dbg = await settle(editorPage)
        // a fixture whose rendering is a known backlog item still proves the layout settles
        if (t.status === 'failing') return

        // pass idempotence: a full re-measure of the settled document (the
        // font-load path) must reproduce the same pages and block geometry;
        // a difference means a decoration of the previous pass leaked into
        // this one's measurement
        await editorPage.evaluate(() => document.fonts.dispatchEvent(new Event('loadingdone')))
        await editorPage.waitForTimeout(400)
        const again = await settle(editorPage)
        expect(again.slices.length, 'page count after a forced re-measure').toBe(dbg.slices.length)
        expect(again.blocks.length, 'block count after a forced re-measure').toBe(dbg.blocks.length)
        for (let i = 0; i < dbg.blocks.length; i++) {
          expect(
            Math.abs(again.blocks[i].top - dbg.blocks[i].top) +
              Math.abs(again.blocks[i].height - dbg.blocks[i].height),
            `block ${i} geometry after a forced re-measure`,
          ).toBeLessThan(0.5)
        }

        // A3 invariants, for free on every fixture: gaps == boundaries; the
        // self-scheduled pass chain converged (followUps counts its depth: one
        // follow-up that settled reads 1) well under the cap, with no reason
        // firing twice in a row (the ping-pong signature of a measurement that
        // reads its own decoration)
        expect(dbg.gapBoundaryMismatch ?? 0, 'page gaps built must equal page boundaries').toBe(0)
        expect(
          dbg.loopStops ?? 0,
          `follow-up chain refused as a loop: ${dbg.retriggerLog.join(' > ')}`,
        ).toBe(0)
        expect(
          dbg.followUps,
          `follow-up chain ${dbg.retriggerLog.join(' > ')}`,
        ).toBeLessThanOrEqual(1)
        for (let i = 1; i < dbg.retriggerLog.length; i++) {
          expect(dbg.retriggerLog[i], `re-trigger chain ${dbg.retriggerLog.join(' > ')}`).not.toBe(
            dbg.retriggerLog[i - 1],
          )
        }

        for (const c of t.cases) {
          if (c.status === 'failing') continue
          const label = `${c.kind}${c.note ? ` - ${c.note}` : ''}`
          switch (c.kind) {
            case 'pages':
              expect.soft(dbg.slices.length, label).toBe(c.expect)
              break
            case 'blockPage':
              expect.soft(pageOfBlock(dbg, c.block!), label).toBe(c.expect)
              break
            case 'pageTextTopPt': {
              const got = await pageTextTopPt(editorPage, c.page!)
              expect
                .soft(Math.abs(got - (c.expect as number)), `${label}: ${got.toFixed(2)}pt`)
                .toBeLessThanOrEqual(c.tol ?? 0.6)
              break
            }
            case 'headerTextTopPt': {
              const got = await pageTextTopPt(editorPage, c.page!, 'header')
              expect
                .soft(Math.abs(got - (c.expect as number)), `${label}: got ${got}`)
                .toBeLessThan(c.tol ?? 0.6)
              break
            }
            case 'pageTextTopDeltaPt': {
              // against a control page: independent of the face's ascent convention
              const got =
                (await pageTextTopPt(editorPage, c.page!)) -
                (await pageTextTopPt(editorPage, c.control!))
              expect
                .soft(Math.abs(got - (c.expect as number)), `${label}: ${got.toFixed(2)}pt`)
                .toBeLessThanOrEqual(c.tol ?? 0.6)
              break
            }
            case 'blockTopDeltaPt': {
              const a = blockAt(dbg, c.from!)
              const b = blockAt(dbg, c.to!)
              expect
                .soft(Math.abs((b.top - a.top) * PT_PER_PX - (c.expect as number)), label)
                .toBeLessThanOrEqual(c.tol ?? 0.2)
              break
            }
            case 'textTopDeltaPt': {
              // same-page paragraphs: glyph-top distance == Word baseline distance (same face and size)
              for (const i of [c.from!, c.to!]) blockAt(dbg, i) // both are model blocks
              const delta =
                ((await firstTextTopPx(editorPage, c.to!)) -
                  (await firstTextTopPx(editorPage, c.from!))) *
                PT_PER_PX
              expect
                .soft(Math.abs(delta - (c.expect as number)), `${label}: ${delta.toFixed(2)}pt`)
                .toBeLessThanOrEqual(c.tol ?? 0.2)
              break
            }
            case 'blockHeightPt': {
              const b = blockAt(dbg, c.block!)
              expect
                .soft(Math.abs(b.height * PT_PER_PX - (c.expect as number)), label)
                .toBeLessThanOrEqual(c.tol ?? 0.2)
              break
            }
            case 'tableHeightPt': {
              // Chromium lands each row on a 0.5px step, so a 4-row table carries up to 2px of noise
              const b = blockAt(dbg, c.block!)
              expect
                .soft(Math.abs(b.height * PT_PER_PX - (c.expect as number)), label)
                .toBeLessThanOrEqual(c.tol ?? 2)
              break
            }
            case 'rowHeightsPt': {
              const b = blockAt(dbg, c.block!)
              const rows = dbg.tableRows.find((r) => Math.abs(r.top - b.top) < 0.5)
              expect.soft(rows, `${label}: table rows at top ${b.top}`).toBeTruthy()
              if (!rows) break
              const got = rows.rows.map((r) => r.h * PT_PER_PX)
              const want = c.expect as number[]
              expect.soft(got.length, label).toBe(want.length)
              for (let i = 0; i < want.length; i++) {
                expect
                  .soft(Math.abs(got[i] - want[i]), `${label} row ${i}`)
                  .toBeLessThanOrEqual(c.tol ?? 0.6)
              }
              break
            }
            case 'gridLineBox': {
              // invariant A3.4: a one-line grid paragraph's DOM line box is the model's line height
              const px = await singleLineBoxPx(editorPage, c.block!)
              const wantPx = (c.expect as number) / PT_PER_PX
              expect
                .soft(Math.abs(px - wantPx), `${label}: line box ${px}px vs ${wantPx}px`)
                .toBeLessThanOrEqual(0.02)
              const b = blockAt(dbg, c.block!)
              expect
                .soft(Math.abs(b.height - wantPx), `${label}: block height`)
                .toBeLessThanOrEqual(0.02)
              break
            }
            case 'nestedTableTopPt': {
              const got = await nestedTableTopPt(editorPage, c.block!, c.nth ?? 0)
              expect
                .soft(Math.abs(got - (c.expect as number)), `${label}: ${got.toFixed(2)}pt`)
                .toBeLessThanOrEqual(c.tol ?? 0.6)
              break
            }
            case 'mathLinePt': {
              const got = await mathBoxPt(editorPage, c.block!, c.nth ?? 0)
              expect
                .soft(Math.abs(got - (c.expect as number)), `${label}: ${got.toFixed(2)}pt`)
                .toBeLessThanOrEqual(c.tol ?? 0.6)
              break
            }
            default:
              throw new Error(`unknown render case kind "${c.kind}" in ${t.id}`)
          }
        }
      } finally {
        await closeAndSaveVideo(launched, `docs-truths-${t.id}`)
      }
    })
  }
})
