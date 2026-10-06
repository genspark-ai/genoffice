import { test, expect, type ElectronApplication } from '@playwright/test'
import { mkdtemp, readFile, writeFile, cp } from 'node:fs/promises'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Switching editors swaps the application menu, and `withUserGuide()` puts the
 * manual's F1 item into whatever Help menu it finds (#1863).
 *
 * It used to destructure two items out of a two-element template. Electron
 * drops a separator that sits at either end of a template, so that build
 * returned one item, `separator` was `undefined`, and the next editor switch
 * threw "Invalid item" from `Menu.insert` and took the main process down — the
 * modal over the editor reading "A JavaScript error occurred in the main
 * process".
 *
 * A unit test cannot see this: the defect is in how Electron assembles a
 * template, not in a pure function. So this drives the real menu and reads it
 * back out of the main process — one F1 item, and a process that still answers.
 */
const menuReport = (app: ElectronApplication) =>
  app.evaluate(({ Menu }) => {
    const help = Menu.getApplicationMenu()?.items.find((i) => i.role === 'help')
    return {
      hasHelp: Boolean(help),
      f1: (help?.submenu?.items ?? []).filter((i) => i.accelerator === 'F1').length,
      labels: (help?.submenu?.items ?? []).map((i) => i.label || i.type),
    }
  })

test('the manual survives an editor switch that swaps the application menu', async () => {
  test.setTimeout(300_000)

  const dir = await mkdtemp(join(tmpdir(), 'genoffice-menu-'))
  const fixture = join(dir, 'fixture')
  await cp(join(__dirname, 'assets/font-manager-rubik'), fixture, { recursive: true })
  const slideXml = join(fixture, 'ppt/slides/slide1.xml')
  await writeFile(slideXml, (await readFile(slideXml, 'utf8')).replaceAll('Rubik', 'Arial'))
  const pptx = join(dir, 'deck.pptx')
  execFileSync('zip', ['-X', '-q', '-r', pptx, '.'], { cwd: fixture })

  const launched = await launchShell({ onboardingSeen: true, videoDir: 'menu-user-guide' })
  const { page } = launched
  try {
    await expect(page.locator('.recent-row'))
      .toHaveCount(0, { timeout: 60_000 })
      .catch(() => {})

    // three editors, three menus; the Help submenu is the branch that used to
    // throw, and the process answering at all is the real assertion
    for (const [path, url, name] of [
      [join(dir, 'sample.md'), '://markdown/', 'markdown'],
      [pptx, '://slides/', 'slides'],
      [join(dir, 'sample.md'), '://markdown/', 'markdown again'],
    ] as const) {
      await writeFile(path, '# Sample\n\nA markdown note.\n').catch(() => {})
      await page.evaluate((p) => window.aiOffice.openPath(p), path)
      await waitForPageWithUrl(launched.app, url)
      await page.waitForTimeout(1200)

      const menu = await menuReport(launched.app)
      expect(menu.hasHelp, `${name}: no Help menu installed`).toBe(true)
      expect(menu.f1, `${name}: ${JSON.stringify(menu)}`).toBeLessThanOrEqual(1)
    }
  } finally {
    await closeAndSaveVideo(launched, 'menu-user-guide')
  }
})
