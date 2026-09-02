import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Regression net for the table page-gap flicker (#1088): a fixed-layout table
 * WITHOUT a <colgroup> (the shape AI inserts and HTML pastes used to produce)
 * straddling a page boundary must keep a stable layout. Pre-fix, the in-table
 * page-gap row's td[colspan=1000] widened the column grid, collapsing every
 * cell to ~1px and flipping the page count in an endless ~3Hz remeasure loop.
 * Render-plane oscillation is invisible to file-based checks (the saved docx
 * is fine), so it gets its own E2E probe: sample layout at the flicker period
 * and demand a single stable state.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown }
  __pageDebug?: { slices?: unknown[] }
}

test.describe('docs table page-gap stability', () => {
  test('colgroup-less table straddling a page boundary keeps pages and column widths stable', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-table-gap-flicker' })
    const { app, page } = launched
    try {
      await expect(page.locator('.quick-card').first()).toContainText('AI Docs')
      await page.locator('.quick-card').first().click()
      const editorPage = await waitForPageWithUrl(app, 'docs/out')
      await editorPage.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      // the new-tab flow replaces the editor document once the blank template
      // loads (a raced insert gets wiped): wait until the doc reference holds
      // still across consecutive polls before touching it
      await editorPage.waitForFunction(
        () => {
          const w = window as unknown as AidocsWindow & {
            __eeStableDoc?: unknown
            __eeStableN?: number
          }
          const doc = (w.__aidocs?.editor as { state?: { doc?: unknown } } | undefined)?.state?.doc
          if (!doc) return false
          if (w.__eeStableDoc === doc) w.__eeStableN = (w.__eeStableN ?? 0) + 1
          else {
            w.__eeStableDoc = doc
            w.__eeStableN = 0
          }
          return (w.__eeStableN ?? 0) >= 3
        },
        undefined,
        { timeout: 30_000, polling: 400 },
      )

      // lead-in paragraphs, then a 40-row table with NO colWidthsPct: it renders
      // without a colgroup and is tall enough that pagination must cut inside it
      await editorPage.evaluate(() => {
        const editor = (window as unknown as AidocsWindow).__aidocs!.editor as {
          state: { doc: { content: { size: number } } }
          commands: { insertContentAt: (pos: number, content: unknown) => boolean }
        }
        const para = (text: string) => ({ type: 'docParagraph', content: [{ type: 'text', text }] })
        const cell = (text: string) => ({ type: 'docTableCell', content: [para(text)] })
        const nodes: unknown[] = []
        for (let i = 1; i <= 12; i++) {
          nodes.push(para(`Lead-in paragraph ${i} pushing the table onto a page boundary.`))
        }
        nodes.push({
          type: 'docTable',
          content: Array.from({ length: 40 }, (_, r) => ({
            type: 'docTableRow',
            content: [cell(`Row ${r + 1} item description`), cell(`Spec clause ${r + 1}.X`)],
          })),
        })
        editor.commands.insertContentAt(editor.state.doc.content.size, nodes)
      })
      // let pagination settle before sampling
      await editorPage.waitForTimeout(3500)

      // sample layout at the pre-fix flicker period (~300ms remeasure debounce)
      const samples: Array<{ pages: number; cellW: number; gapColspans: number[] }> = []
      for (let i = 0; i < 14; i++) {
        samples.push(
          await editorPage.evaluate(() => {
            const table = document.querySelector('.ProseMirror table.doc-table')
            const firstCell = table?.querySelector('tr:not(.page-gap):not(.page-repeat-header) td')
            const gapCells = Array.from(
              document.querySelectorAll<HTMLTableCellElement>(
                '.ProseMirror tr.page-gap-table > td',
              ),
            )
            return {
              pages: (window as unknown as AidocsWindow).__pageDebug?.slices?.length ?? -1,
              cellW: firstCell ? Math.round(firstCell.getBoundingClientRect().width) : -1,
              gapColspans: gapCells.map((c) => c.colSpan),
            }
          }),
        )
        await editorPage.waitForTimeout(300)
      }

      const pages = new Set(samples.map((s) => s.pages))
      const widths = new Set(samples.map((s) => s.cellW))
      const gapSpans = samples.flatMap((s) => s.gapColspans)
      expect(pages.size, `page count oscillates: ${[...pages].join(',')}`).toBe(1)
      expect(widths.size, `cell width oscillates: ${[...widths].join(',')}px`).toBe(1)
      // collapsed columns read ~1px; a healthy two-column grid is hundreds wide
      expect(samples[0].cellW).toBeGreaterThan(100)
      expect(
        gapSpans.length,
        'no in-table page-gap row: the table never straddled a page boundary',
      ).toBeGreaterThan(0)
      for (const span of gapSpans) {
        expect(span, 'gap colspan must equal the real column grid').toBe(2)
      }
    } finally {
      await closeAndSaveVideo(launched, 'docs-table-gap-flicker')
    }
  })
})
