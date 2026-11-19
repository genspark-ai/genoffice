/**
 * Tools > Track Changes > Highlight Changes: "Highlight changes in printed
 * document" is its own tick. Cleared, the print sheet shows and prints the
 * clean pages while the screen keeps its markup, and the screen's view comes
 * back when the sheet closes.
 */
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

test('a print without highlighted changes leaves the markup on screen', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-highlight-changes-print',
  })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('.ProseMirror').first().click()
    await page.keyboard.type('tracked text')
    const clickMenu = (top: string, path: string[]) =>
      expect
        .poll(
          () =>
            app.evaluate(
              ({ Menu }, { top, path }) => {
                let item = Menu.getApplicationMenu()?.items.find((i) => i.label === top)
                for (const label of path)
                  item = item?.submenu?.items.find((i) => i.label.startsWith(label))
                if (!item?.enabled) return false
                item.click()
                return true
              },
              { top, path },
            ),
          { timeout: 15_000 },
        )
        .toBe(true)
    const dialog = page.locator('.hlc-dialog')
    const printBox = dialog.locator('input[name="highlight-print"]')
    const root = page.locator('.app')

    await clickMenu('Tools', ['Track Changes', 'Highlight Changes'])
    await expect(dialog).toBeVisible()
    await expect(printBox).toBeChecked()
    await printBox.uncheck()
    await page.screenshot({ path: test.info().outputPath('highlight-changes-print.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // the screen keeps All Markup
    await expect(root).not.toHaveClass(/rev-display-/)

    await clickMenu('File', ['Print'])
    const sheet = page.locator('.print-dialog')
    await expect(sheet).toBeVisible({ timeout: 30_000 })
    await expect(root).toHaveClass(/rev-display-none/)
    await sheet.locator('button', { hasText: 'Cancel' }).click()
    await expect(sheet).toBeHidden()
    await expect(root).not.toHaveClass(/rev-display-/)

    // the tick is kept for the next time the dialog opens
    await clickMenu('Tools', ['Track Changes', 'Highlight Changes'])
    await expect(printBox).not.toBeChecked()
  } finally {
    await closeAndSaveVideo(launched, 'docs-highlight-changes-print')
  }
})
