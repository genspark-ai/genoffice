import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * The renderer-facing tab API behind the tear-off gesture: a tab leaves the
 * strip for a window that follows the pointer, and comes back as a tab when
 * the pointer returns — the live WebContentsView is reparented both ways, so
 * the document is never reloaded. Driven through the API because Playwright
 * cannot hold a pointer drag across two OS windows.
 */

interface TabsApiOnWindow {
  aiOfficeTabs: {
    list(): Promise<Array<{ id: string; kind: string; title: string; active: boolean }>>
    tearOff(id: string, screenX: number, screenY: number): Promise<boolean>
    dragTornWindow(screenX: number, screenY: number): void
    dockTornWindow(index: number): Promise<void>
    endTornDrag(): Promise<void>
    detach(id: string): Promise<void>
  }
}

test.describe('tab tear-off and dock', () => {
  test('a torn-off markdown tab docks back without a reload; Open in New Window covers it too', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-tear-'))
    const mdPath = join(dir, 'notes.md')
    await writeFile(mdPath, '# Tear me off\n\nA paragraph.\n')
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'tab-tear-off-dock',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      // with a file on argv the first window Playwright sees may be the editor
      // view, not the shell renderer that hosts the tab strip
      const page = await waitForPageWithUrl(app, 'shell/out')
      const mdPage = await waitForPageWithUrl(app, '://markdown/')
      await expect(mdPage.locator('body')).toBeVisible()
      // survives only while the renderer is never reloaded
      await mdPage.evaluate(() => {
        ;(window as unknown as { __tearMarker: number }).__tearMarker = 42
      })

      const tabs = await page.evaluate(() =>
        (window as unknown as TabsApiOnWindow).aiOfficeTabs.list(),
      )
      const mdTab = tabs.find((t) => t.kind === 'markdown')
      expect(mdTab, 'the markdown file opened in a tab').toBeTruthy()

      const windowsNow = () =>
        app.evaluate(({ BrowserWindow }) =>
          BrowserWindow.getAllWindows()
            .filter((w) => !w.isDestroyed())
            .map((w) => ({ title: w.getTitle(), position: w.getPosition() })),
        )

      // ── tear off: the tab leaves the strip, a window appears under the pointer
      const torn = await page.evaluate(
        (id) => (window as unknown as TabsApiOnWindow).aiOfficeTabs.tearOff(id, 640, 480),
        mdTab!.id,
      )
      expect(torn).toBe(true)
      await expect(page.locator('.tab-bar .tab-item:not(.tab-home)')).toHaveCount(0)
      const afterTear = await windowsNow()
      expect(afterTear).toHaveLength(2)
      const tornWindow = afterTear.find((w) => w.title === 'notes.md')
      expect(tornWindow, 'the detached window carries the file name').toBeTruthy()

      // ── the held pointer steers the window
      await page.evaluate(() =>
        (window as unknown as TabsApiOnWindow).aiOfficeTabs.dragTornWindow(700, 540),
      )
      await expect
        .poll(async () => (await windowsNow()).find((w) => w.title === 'notes.md')?.position)
        .toEqual([tornWindow!.position[0] + 60, tornWindow!.position[1] + 60])

      // ── back over the strip: the window folds into a tab again, same document
      await page.evaluate(() =>
        (window as unknown as TabsApiOnWindow).aiOfficeTabs.dockTornWindow(1),
      )
      const docked = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(docked).toHaveCount(1)
      await expect(docked).toHaveClass(/active/)
      await expect(docked).toContainText('notes.md')
      await expect.poll(windowsNow).toHaveLength(1)
      expect(
        await mdPage.evaluate(() => (window as unknown as { __tearMarker: number }).__tearMarker),
      ).toBe(42)

      // ── the context-menu path ("Open in New Window") reaches markdown as well
      const redocked = await page.evaluate(() =>
        (window as unknown as TabsApiOnWindow).aiOfficeTabs.list(),
      )
      const again = redocked.find((t) => t.kind === 'markdown')!
      await page.evaluate(
        (id) => (window as unknown as TabsApiOnWindow).aiOfficeTabs.detach(id),
        again.id,
      )
      await expect(page.locator('.tab-bar .tab-item:not(.tab-home)')).toHaveCount(0)
      await expect.poll(windowsNow).toHaveLength(2)
      expect(
        await mdPage.evaluate(() => (window as unknown as { __tearMarker: number }).__tearMarker),
      ).toBe(42)

      // a clean detached window closes without a prompt
      await app.evaluate(({ BrowserWindow }) => {
        for (const w of BrowserWindow.getAllWindows()) if (w.getTitle() === 'notes.md') w.close()
      })
      await expect.poll(windowsNow).toHaveLength(1)
    } finally {
      await closeAndSaveVideo(launched, 'tab-tear-off-dock')
    }
  })
})
