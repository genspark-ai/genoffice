/**
 * A justified paragraph straddling a page boundary must not render its
 * page-last line stretched sparse (r177): the inline page-gap
 * widget forces the line break at the paginator's cut position, and when
 * justify-shrink decorations later move the natural wrap points without a doc
 * change, a stale cut leaves a short line that text-align:justify blows up
 * with huge word gaps until the user edits at the boundary.
 *
 * The fixture is compatibilityMode-15 justified French text (space-shrink
 * active). Each round edits the FIRST paragraph — shifting every downstream
 * wrap point and forcing shrink re-decisions — then asserts that on every
 * page-crossing line the largest word gap stays within a few natural space
 * widths.
 */
import { test, expect } from '@playwright/test'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const FIXTURE = join(__dirname, 'assets/justify-pagegap-fr.docx')

/** per inline page gap: word count and max word-gap of the line right above
 *  the widget, as a multiple of the natural space advance */
async function sparseReport(editor: Page): Promise<Array<{ words: number; ratio: number }>> {
  return editor.evaluate(() => {
    const out: Array<{ words: number; ratio: number }> = []
    const canvas = document.createElement('canvas').getContext('2d')!
    for (const gap of Array.from(document.querySelectorAll('.page-gap-inline'))) {
      const para = gap.closest('p') ?? gap.parentElement
      if (!para || gap.closest('table')) continue
      const cs = getComputedStyle(para)
      canvas.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`
      const spaceW = canvas.measureText(' ').width
      const gapTop = gap.getBoundingClientRect().top
      const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT)
      const boxes: Array<{ l: number; r: number; t: number; b: number }> = []
      let tn: Node | null
      while ((tn = walker.nextNode())) {
        if (!(gap.compareDocumentPosition(tn) & Node.DOCUMENT_POSITION_PRECEDING)) continue
        const text = tn.textContent ?? ''
        const re = /[^ ]+/g
        let m: RegExpExecArray | null
        while ((m = re.exec(text))) {
          const r = document.createRange()
          r.setStart(tn, m.index)
          r.setEnd(tn, m.index + m[0].length)
          for (const rect of Array.from(r.getClientRects())) {
            if (rect.width > 0.01)
              boxes.push({ l: rect.left, r: rect.right, t: rect.top, b: rect.bottom })
          }
        }
      }
      const above = boxes.filter((b) => b.b <= gapTop + 1)
      if (above.length === 0) continue
      const lowest = above.reduce((a, b) => (b.b > a.b ? b : a))
      const line = above
        .filter((b) => b.t < lowest.b - 1 && b.b > lowest.t + 1)
        .sort((a, b) => a.l - b.l)
      let maxGap = 0
      for (let i = 1; i < line.length; i++) maxGap = Math.max(maxGap, line[i].l - line[i - 1].r)
      const zoom =
        para instanceof HTMLElement ? para.getBoundingClientRect().width / para.offsetWidth || 1 : 1
      out.push({ words: line.length, ratio: maxGap / zoom / spaceW })
    }
    return out
  })
}

const wait = (ms: number) => new Promise((r) => setTimeout(r, ms))

test('page-crossing justified lines keep natural word gaps across reflows', async () => {
  test.setTimeout(150_000)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'justify-pagegap',
    openFile: FIXTURE,
  })
  const { app } = launched
  try {
    const editor = await waitForPageWithUrl(app, '://docs/')
    await editor.locator('.doc-page .page-gap-inline').first().waitFor({ timeout: 30000 })
    await wait(2500) // pagination + shrink settle

    const baseline = await sparseReport(editor)
    expect(baseline.length).toBeGreaterThan(0) // fixture must actually cross pages mid-paragraph
    for (const g of baseline) expect(g.ratio).toBeLessThan(3)

    // grow the first paragraph by a full line+ each round: every downstream
    // block moves, the page boundary lands in new lines/paragraphs and the
    // shrink extension re-decides around each fresh cut (cumulative growth,
    // no deletes — per-keystroke Delete rounds blew the CI time budget)
    const sentence =
      'Zut alors vraiment ce paragraphe grandit encore un peu plus pour pousser toutes les lignes vers le bas maintenant '
    for (let round = 1; round <= 6; round++) {
      await editor.evaluate(() =>
        document.querySelector('.ProseMirror p')?.scrollIntoView({ block: 'center' }),
      )
      await editor
        .locator('.ProseMirror p')
        .first()
        .click({ position: { x: 8, y: 8 } })
      await editor.keyboard.press('Home')
      await editor.keyboard.type(sentence, { delay: 2 })
      await wait(1500) // > pagination debounce + shrink RAF rounds
      const report = await sparseReport(editor)
      expect(report.length, `round ${round}: mid-paragraph cuts vanished`).toBeGreaterThan(0)
      for (const g of report) {
        expect(g.ratio, `round ${round}: sparse page-last line (${g.words} words)`).toBeLessThan(3)
      }
    }
  } finally {
    await closeAndSaveVideo(launched, 'justify-pagegap')
  }
})
