/**
 * Route-B prototype probe for the large-document lag (#2781): a ~100-page
 * document opens twice, once with layout parking (content-visibility on
 * offscreen blocks) and once without, and the probe measures where the
 * continuous-DOM cost actually lives in real Chromium — typing frame gaps,
 * the measurement pass duration, and the parked-block count.
 *
 * The assertions pin the mechanism, not the latencies (machine-dependent):
 * parking engages on a large document, the caret reveals a parked block and
 * typing lands, scrolling reveals parked content, and the measured numbers
 * are printed for the research note. Run locally, not in CI:
 *   npx playwright test e2e/large-doc-parking.spec.ts
 */
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from './helpers'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const PARAGRAPHS = 5000
const MID = 2500

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

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
  const body = Array.from({ length: PARAGRAPHS }, (_, i) =>
    p(
      i % 50 === 0
        ? `Paragraph ${i} opens a block of copy that runs a little longer so the page count lands near a hundred pages.`
        : `Paragraph ${i} body copy.`,
    ),
  ).join('')
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
    if (!parking)
      await page.evaluate(
        () => ((window as never as Record<string, unknown>).__layoutParkingOff = true),
      )
    // the first passes run with nothing parked; parking engages at pass end
    await page.waitForFunction(
      () => {
        const dbg = (window as never as Record<string, Record<string, unknown>>).__pageDebug
        return Boolean(dbg && (dbg.passGen as number) >= 1 && dbg.remeasureMs)
      },
      { timeout: 180_000 },
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
    // on — the caret must reveal it (content-visibility relevance)
    await page.locator('.doc-page p').nth(MID).click()
    const midState = await page.evaluate((idx) => {
      const el = document.querySelectorAll('.doc-page p')[idx] as HTMLElement
      return {
        contentVisibility: getComputedStyle(el).contentVisibility,
        parked: document.querySelectorAll('.ProseMirror .cv-parked').length,
      }
    }, MID)

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
    // the typed text must be in the document (caret reveal worked)
    const landed = await page.evaluate(() => {
      const el = document.querySelectorAll('.doc-page p')[MID] as HTMLElement
      return el.textContent ?? ''
    })
    expect(landed).toContain('quick brown fox')

    // scroll into parked territory: content reveals on arrival
    await page.evaluate(() => {
      const els = document.querySelectorAll('.doc-page p')
      const far = els[PARAGRAPHS - 60] as HTMLElement
      far.scrollIntoView()
    })
    await page.waitForTimeout(600)
    const farVisible = await page.evaluate(() => {
      const far = document.querySelectorAll('.doc-page p')[PARAGRAPHS - 60] as HTMLElement
      return getComputedStyle(far).contentVisibility
    })
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

test('large document: layout parking engages, reveals on caret and scroll', async () => {
  test.setTimeout(900_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-large-doc-')))
  const docPath = join(dir, 'large.docx')
  writeFileSync(docPath, await bigDocx())
  try {
    await runOnce(false, docPath)
    await runOnce(true, docPath)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
})
