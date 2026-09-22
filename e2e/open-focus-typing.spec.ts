import { test, expect } from '@playwright/test'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { ElectronApplication, Page } from 'playwright-core'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const DOCX = resolve(__dirname, 'assets/justify-pagegap-fr.docx')
const XLSX = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

/**
 * A freshly opened document must own the keyboard. The regression had
 * two layers — the shell never handed webContents focus to a tab activated
 * from the Home list (the click sits on the chrome webContents), and the docs
 * editor never focused its body even in a focused view. Both are exercised by
 * the real flow: land on Home, open a file exactly like the Home list does,
 * then type without a single click into the document.
 */

/** the Home-list click that precedes an open leaves keyboard focus on chrome */
async function openFromHome(app: ElectronApplication, home: Page, file: string): Promise<void> {
  await app.evaluate(({ BrowserWindow }) => {
    const win = BrowserWindow.getAllWindows()[0]
    win.focus()
    win.webContents.focus()
  })
  await home.evaluate(
    (p) =>
      (window as unknown as { aiOffice: { openPath(p: string): Promise<void> } }).aiOffice.openPath(
        p,
      ),
    file,
  )
}

/** the fix's observable end state: the document's editable surface holds focus */
async function waitForEditableFocus(page: Page): Promise<void> {
  await page.waitForFunction(() => document.activeElement?.isContentEditable === true, null, {
    timeout: 30_000,
  })
}

/** Playwright's keyboard bypasses Electron-level focus, so the shell layer is
 *  asserted directly: the opened document's webContents must hold the window's
 *  keyboard focus (on unfixed code it stays on the Home/chrome webContents). */
async function expectViewFocused(app: ElectronApplication, urlPart: string): Promise<void> {
  await expect
    .poll(() =>
      app.evaluate(
        ({ webContents }, part) =>
          webContents
            .getAllWebContents()
            .filter((wc) => wc.isFocused())
            .map((wc) => wc.getURL())
            .some((u) => u.includes(part)),
        urlPart,
      ),
    )
    .toBe(true)
}

test('docs: typing works immediately after opening a file from Home', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-openfocus-e2e-'))
  const docx = join(scratch, 'open-focus.docx')
  await copyFile(DOCX, docx)

  const launched = await launchShell({ onboardingSeen: true, videoDir: 'open-focus-docs' })
  try {
    await openFromHome(launched.app, launched.page, docx)
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    await docs.waitForSelector('.ProseMirror', { timeout: 30_000 })
    await expectViewFocused(launched.app, '://docs/')
    await waitForEditableFocus(docs)

    await docs.keyboard.type('ROW226FOCUS')
    await expect
      .poll(() => docs.evaluate(() => document.querySelector('.ProseMirror')?.textContent ?? ''))
      .toContain('ROW226FOCUS')
  } finally {
    await closeAndSaveVideo(launched, 'open-focus-docs')
  }
})

test('sheets: typing into the active cell works immediately after opening from Home', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-openfocus-e2e-'))
  const xlsx = join(scratch, 'open-focus.xlsx')
  await copyFile(XLSX, xlsx)

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'open-focus-sheets',
    // the real production open path adopts the pre-mounted spare sheets view
    // (helpers disable it by default) — the focus bug only shows through it
    env: { GENOFFICE_DEBUG_HOOKS: '1', GENOFFICE_NO_SPARE_VIEW: '' },
  })
  try {
    // let the spare view mount (scheduled 1.5s after the shell finishes loading)
    await launched.page.waitForTimeout(2_500)
    await openFromHome(launched.app, launched.page, xlsx)
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(
      () =>
        (window as unknown as { __genofficeDebug?: { univerAPI?: unknown } }).__genofficeDebug
          ?.univerAPI,
      null,
      { timeout: 60_000 },
    )
    await expectViewFocused(launched.app, '://sheets/')
    await waitForEditableFocus(sheets)

    await sheets.keyboard.type('4242')
    await sheets.keyboard.press('Enter')
    await expect
      .poll(() =>
        sheets.evaluate(() => {
          const api = (
            window as unknown as {
              __genofficeDebug: {
                univerAPI: {
                  getActiveWorkbook(): {
                    getActiveSheet(): { getRange(a: string): { getValue(): unknown } }
                  }
                }
              }
            }
          ).__genofficeDebug.univerAPI
          return api.getActiveWorkbook()?.getActiveSheet()?.getRange('A1')?.getValue() ?? null
        }),
      )
      .toBe(4242)
  } finally {
    await closeAndSaveVideo(launched, 'open-focus-sheets')
  }
})
