/**
 * Route-B prototype probe for the large-document lag (#2781): opens a
 * table-heavy document (~25 pages of bordered 30x6 tables) and pins the
 * parking mechanism in real Chromium — the sticky cost gate arms, most
 * blocks park, the caret reveals a parked block and typing lands, and
 * scrolling reveals parked content. Numbers print for the research note;
 * only mechanism facts are asserted.
 *
 * A plain-text document of the same page count must NOT park (the wake-up
 * tax is a net loss there); that side is covered by the gate unit tests
 * (parkingArmDecided).
 *
 * Local only: this directory is in the main config's testIgnore, so the CI
 * sweep never collects it — opening this document wedges Playwright's
 * Electron CDP attach (electron.launch never resolves — observed locally
 * and on CI runners alike, twice, each burning the full 15-minute budget),
 * and two full launches exceed the CI test budget even where the launch
 * succeeds. The mechanism's CI coverage lives in the gate unit tests.
 *
 * Run locally (build first: `npm run build:all`):
 *   npx playwright test --config e2e/local/playwright.local.config.ts
 *
 * Parking is opt-in (Preferences ▸ layout parking, layout-parking-pref);
 * the probe exercises the development override window.__layoutParkingOn so
 * it measures the mechanism, not the settings UI. window.__layoutParkingOff
 * forces parking back off for an A/B leg.
 */
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from '../helpers'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const TABLES = 40
const ROWS = 30
/** blocks in the probe document: per table an intro and a closing paragraph
 *  plus ROWS table rows (page.evaluate callbacks cannot see these Node
 *  constants — indexes derived from them are inlined at the use sites) */
const BLOCKS = TABLES * 2 + TABLES * ROWS
const MID_BLOCK = Math.floor(BLOCKS / 2)

const cell = (t: string) =>
  `<w:tc><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`

async function bigDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  const row = (r: number, t: number) =>
    `<w:tr>${[0, 1, 2, 3, 4, 5].map((c) => cell(`R${r}C${c} value ${t * 7 + c}`)).join('')}</w:tr>`
  let body = ''
  for (let t = 0; t < TABLES; t++) {
    body += p(`Table ${t + 1} of ${TABLES} — quarterly figures.`)
    const rows = Array.from({ length: ROWS }, (_, r) => row(r + t * ROWS, r)).join('')
    body +=
      `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="1500"/>${'<w:gridCol w:w="1500"/>'.repeat(5)}</w:tblGrid>${rows}</w:tbl>`
    body += p(`Closing remarks for table ${t + 1} before the next section begins.`)
  }
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}` +
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
      `</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

interface TypingStats {
  median: number
  p95: number
  max: number
  longFrames: number
  samples: number
}

function stats(gaps: number[]): TypingStats {
  const sorted = [...gaps].sort((a, b) => a - b)
  const at = (q: number) => sorted[Math.min(sorted.length - 1, Math.floor(q * sorted.length))]!
  return {
    median: Math.round(at(0.5)),
    p95: Math.round(at(0.95)),
    max: Math.round(sorted[sorted.length - 1]!),
    longFrames: gaps.filter((g) => g > 250).length,
    samples: gaps.length,
  }
}

/** records requestAnimationFrame gaps while the Node side types */
/** main-thread responsiveness while typing: the Node side measures the round
    trip of a trivial evaluate around each keystroke — the queuing delay the
    typist feels. No rAF dependency (an occluded window throttles frames). */
async function typingRtts(page: Page, text: string, delayMs: number): Promise<number[]> {
  const rtts: number[] = []
  for (const ch of text) {
    await page.keyboard.type(ch, { delay: delayMs })
    const t0 = Date.now()
    await page.evaluate(() => 1)
    rtts.push(Date.now() - t0)
  }
  return rtts
}

async function runOnce(parking: boolean, docPath: string): Promise<void> {
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: docPath,
    videoDir: 'large-doc-parking',
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('.doc-page p').first().waitFor({ timeout: 60_000 })
    // parking is opt-in and off by default: the probe rides the development
    // override (set before the first pass ends, so the sticky gate arms with
    // parking wanted); the off leg just never sets it
    if (parking)
      await page.evaluate(
        () => ((window as never as Record<string, unknown>).__layoutParkingOn = true),
      )
    // the first passes run with nothing parked; parking engages at pass end
    await page.waitForFunction(
      () => {
        const dbg = (window as never as Record<string, Record<string, unknown>>).__pageDebug
        return Boolean(dbg && (dbg.passGen as number) >= 1 && dbg.remeasureMs)
      },
      { timeout: 180_000 },
    )
    if (parking)
      // the flag may have landed after the pass that armed the gate: wait
      // until a dispatch actually parked the offscreen set
      await page.waitForFunction(
        () => document.querySelectorAll('.ProseMirror .cv-parked').length > 0,
        { timeout: 120_000 },
      )

    const before = await page.evaluate(() => {
      const dbg = (window as never as Record<string, Record<string, unknown>>).__pageDebug
      return {
        parked: document.querySelectorAll('.ProseMirror .cv-parked').length,
        blocks: document.querySelectorAll('.ProseMirror > *').length,
        remeasureMs: dbg?.remeasureMs as number,
        pages: document.querySelectorAll('.page-gap').length + 1,
      }
    })

    // click into the middle of the document: a parked region when parking is
    // on — the caret must reveal it (content-visibility relevance). The
    // paragraph is revealed by the scroll first, so the click's stability
    // wait does not race the reveal, and its timeout stays bounded.
    const mid = page.locator('.doc-page p').nth(MID_BLOCK)
    await mid.evaluate((el) => el.scrollIntoView({ block: 'center' }))
    await mid.click({ timeout: 30_000 })
    const midState = await page.evaluate((idx) => {
      const el = document.querySelectorAll('.doc-page p')[idx] as HTMLElement
      return {
        contentVisibility: getComputedStyle(el).contentVisibility,
        parked: document.querySelectorAll('.ProseMirror .cv-parked').length,
      }
    }, MID_BLOCK)

    // typing: main-thread RTT sampled per keystroke, inserted where the
    // caret sits (a parked region when parking is on)
    const gaps = await typingRtts(page, 'The quick brown fox jumps. ', 30)

    // let the debounce pass fire and finish
    await page.waitForTimeout(2_500)
    const after = await page.evaluate(() => {
      const dbg = (window as never as Record<string, Record<string, unknown>>).__pageDebug
      return {
        parked: document.querySelectorAll('.ProseMirror .cv-parked').length,
        remeasureMs: dbg?.remeasureMs as number,
      }
    })
    // the typed text must be in the document (caret reveal worked); the
    // mid paragraph index is MID_BLOCK (1280 blocks, floor(/2)) inlined —
    // page callbacks cannot see Node constants
    const landed = await page.evaluate(() => {
      const el = document.querySelectorAll('.doc-page p')[640] as HTMLElement
      return el.textContent ?? ''
    })
    expect(landed).toContain('quick brown fox')

    // scroll into parked territory: content reveals on arrival. The check
    // reads the paragraph the scroll targeted (60 from the end, passed as an
    // argument — page callbacks cannot see Node constants)
    const farIdx = await page.evaluate(() => {
      const els = document.querySelectorAll('.doc-page p')
      const far = els[els.length - 60] as HTMLElement
      far.scrollIntoView()
      return els.length - 60
    })
    await page.waitForTimeout(600)
    const farVisible = await page.evaluate((idx) => {
      const far = document.querySelectorAll('.doc-page p')[idx] as HTMLElement
      return getComputedStyle(far).contentVisibility
    }, farIdx)
    expect(farVisible).not.toBe('auto')

    console.log(
      [
        `\n=== parking ${parking ? 'ON' : 'OFF'} ===`,
        `blocks ${before.blocks}, parked after open: ${before.parked}`,
        `pages ~${before.pages}, last pass (open settle): ${before.remeasureMs?.toFixed(0)} ms`,
        `mid click: contentVisibility=${midState.contentVisibility}, parked=${midState.parked}`,
        `typing 400 frames: ${JSON.stringify(stats(gaps))}`,
        `after typing: parked=${after.parked}, last pass ${after.remeasureMs?.toFixed(0)} ms`,
      ].join('\n'),
    )

    // mechanism assertions: parking engages, and the caret's paragraph is
    // not parked while the caret is inside it
    if (parking) {
      expect(before.parked).toBeGreaterThan(before.blocks * 0.8)
      expect(midState.contentVisibility).not.toBe('auto')
    }
  } finally {
    await app.close()
  }
}

// Not collected by the CI sweep (e2e/local is in the main config's
// testIgnore): the reason lives in the file header — the launch wedges CDP
// attach in CI and two full launches exceed the CI budget.

test('large document: layout parking engages, reveals on caret and scroll', async () => {
  test.setTimeout(900_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-large-doc-')))
  const docPath = join(dir, 'large.docx')
  writeFileSync(docPath, await bigDocx())
  try {
    // the parking-off leg prints the A/B numbers for the research note and
    // carries no assertions — parking is off by default, so it simply never
    // sets the development override
    if (!process.env.CI) await runOnce(false, docPath)
    await runOnce(true, docPath)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
})
