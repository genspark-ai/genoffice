/**
 * One figure, on its own: Word's Keyboard Shortcuts sheet.
 * `npx tsx tools/gen-help-shortcut-figure.ts [lang …]`
 *
 * ## Why this is not in gen-help-screenshots.ts
 *
 * That script photographs the shell: Home, Settings, the tab strip, and a
 * ribbon it opens from inside the app. To reach Word it would have to open a
 * document, and both ways of doing that break it:
 *
 * - **Clicking the file in Home's list** puts the shortcut sheet behind a
 *   dependency on that one row's hit-testing. The row is `draggable` and its
 *   left-hand cell stops click propagation, so the click has to land on the
 *   name cell — a detail that fails silently into a 30-second timeout.
 * - **Opening the document on argv** works, but it makes a Word window exist
 *   from the first frame, and the shared script picks its window by position
 *   (`app.windows()[0]`). That may now be Word's, and the Home-only steps that
 *   follow go looking for a sidebar this window does not have.
 *
 * Handing the document over on argv and finding the shell window by URL fixes
 * the second problem, at the price of the first script carrying a launch path,
 * a window-finder and a menu-trigger it otherwise has no use for. This file does
 * one thing instead, and the scratch environment below is copied from the other
 * one rather than reinvented — a figure captured against a different HOME would
 * be a figure of the developer's session, which is the thing the shared script
 * exists to prevent.
 *
 * ## Two things that have to be got right
 *
 * - **`apps/docs/out` must exist.** The app serves each module's built renderer
 *   over its own scheme (installRendererProtocol in main/index.ts), and a
 *   missing file answers 404 — so the Word window comes up with an empty body
 *   and no error anywhere. Build it: `cd apps/docs && npx electron-vite build`.
 * - **`⌘/` cannot be typed.** It is a main-process menu accelerator, consumed
 *   before the renderer sees the key; the menu item's click is what sends the
 *   `shortcuts` command that opens this dialog.
 *
 * ## Pass one language at a time
 *
 * `npx tsx tools/gen-help-shortcut-figure.ts zh en` writes the CJK figure
 * and then hangs: the second launch never gets its own window. Each run does
 * use a fresh userData, so this is not the single-instance lock the argv comment
 * above is about — whatever it is, the honest workaround is one language per
 * invocation, which is how both committed figures were produced.
 */
import { _electron as electron, type CDPSession, type Page } from '@playwright/test'
import { createRequire } from 'node:module'
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { buildBlankDocx } from '../packages/docx-engine/src/blank'

const REPO = resolve(__dirname, '..')
const SHELL_DIR = join(REPO, 'apps/shell')
const OUT_DIR = join(REPO, 'apps/shell/src/renderer/src/i18n/help/topics/img')

const LANGS = process.argv.slice(2).filter((a) => !a.startsWith('-'))
const TARGET_LANGS = LANGS.length ? LANGS : ['zh', 'en']

async function capture(page: Page, cdp: CDPSession): Promise<Buffer> {
  const { data } = await cdp.send('Page.captureScreenshot', {
    format: 'png',
    fromSurface: true,
  })
  return Buffer.from(data, 'base64')
}

/** 1360x850 CSS at 2x, matching the committed figures. */
async function frame(page: Page, width = 1360): Promise<CDPSession> {
  await page.setViewportSize({ width, height: 850 })
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Emulation.setDeviceMetricsOverride', {
    width,
    height: 850,
    deviceScaleFactor: 2,
    mobile: false,
  })
  await page.waitForTimeout(600)
  return cdp
}

async function main(): Promise<void> {
  mkdirSync(OUT_DIR, { recursive: true })
  for (const lang of TARGET_LANGS) {
    const root = mkdtempSync(join(tmpdir(), 'help-shortcuts-'))
    const home = join(root, 'home')
    const userData = join(root, 'userdata')
    const samples = join(root, 'help-samples')
    for (const dir of [home, userData, samples]) mkdirSync(dir, { recursive: true })
    writeFileSync(join(samples, 'doc.docx'), await buildBlankDocx())

    const require = createRequire(join(SHELL_DIR, 'package.json'))
    const { ELECTRON_RUN_AS_NODE: _drop, ...hostEnv } = process.env
    const args: string[] = []
    if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
    args.push(SHELL_DIR, join(samples, 'doc.docx'))

    const app = await electron.launch({
      executablePath: require('electron') as unknown as string,
      args,
      env: {
        ...hostEnv,
        HOME: home,
        GENOFFICE_USER_DATA: userData,
        GENOFFICE_AUTH_DIR: join(home, '.genoffice'),
        GENOFFICE_NO_SPARE_VIEW: '1',
        GENOFFICE_LANG: lang,
      },
    })

    try {
      let docsPage: Page | undefined
      const deadline = Date.now() + 60_000
      while (!docsPage && Date.now() < deadline) {
        for (const w of app.windows()) {
          const url = await w.evaluate(() => window.location.href).catch(() => '')
          if (url.includes('/docs/') && (await w.locator('.doc-page').count())) docsPage = w
        }
        if (!docsPage) await new Promise((r) => setTimeout(r, 250))
      }
      if (!docsPage) {
        throw new Error(
          'the Word editor never rendered. If apps/docs/out is missing the renderer ' +
            'protocol answers 404 and the window comes up blank with no error — build it ' +
            'with "npx electron-vite build" in apps/docs.',
        )
      }

      const cdp = await frame(docsPage)
      await docsPage.waitForSelector('.ribbon', { timeout: 30_000 })
      const sheet = await app.evaluate(({ Menu }) => {
        const help = Menu.getApplicationMenu()?.items.find((i) => i.role === 'help')
        const item = help?.submenu?.items.find((i) => (i.accelerator ?? '').includes('/'))
        if (!item?.click) throw new Error('the Help menu has no shortcut item')
        item.click()
        return item.label ?? ''
      })
      await docsPage.waitForSelector('.modal-shortcuts', { timeout: 15_000 })
      // the dialog lists every binding, so it is scrolled to the top before the
      // shot rather than left wherever it happened to mount
      await docsPage.evaluate(() => {
        document.querySelector('.modal-shortcuts')?.scrollTo(0, 0)
      })
      await docsPage.waitForTimeout(400)

      const name = `word-shortcuts.${lang}.png`
      writeFileSync(join(OUT_DIR, name), await capture(docsPage, cdp))
      process.stdout.write(`wrote ${name} (sheet "${sheet}")\n`)
    } finally {
      await app.close().catch(() => {})
    }
  }
}

main().catch((err) => {
  console.error(err)
  process.exitCode = 1
})
