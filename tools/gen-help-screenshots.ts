/**
 * Regenerate the manual's screenshots: `npx tsx tools/gen-help-screenshots.ts`
 *
 * ## Why a script and not a hand capture
 *
 * The figures were first shot from a developer's own session, which put their
 * account name in the greeting and their e-mail address in Settings → Account.
 * A committed screenshot is forever, so the fix is not "remember to crop" but a
 * capture that *cannot* pick up the developer's account: a scratch HOME, a
 * scratch userData, and a scratch samples folder, all discarded afterwards.
 *
 * ## Why HOME and not just GENOFFICE_USER_DATA
 *
 * The account name is not in userData at all. `gskApiKey()` reads
 * `~/.genspark-tool-cli/config.json` through `homedir()`, and GENOFFICE_USER_DATA
 * says nothing about that — so a run that only overrides userData still finds
 * the developer's key and greets them by name (verified: that is what a local
 * e2e run does today). HOME is the one lever that covers both it and the
 * `~/.genoffice/bin` shim. `GENOFFICE_AUTH_DIR` is not enough on its own: it
 * redirects GenOffice's own key and leaves the shared gsk CLI config in place.
 *
 * With no key, `gskLoginInfo()` returns null, the app reads as signed out, and
 * the greeting, the account row and the Genspark Projects entry all disappear —
 * which is also the state every CI run and every user without a Genspark account
 * is in. Nothing of the developer's own session is read or written.
 */
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'
import { createRequire } from 'node:module'
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { buildBlankDocx } from '../packages/docx-engine/src/blank'
import { createBlankPptx } from '../packages/pptx-engine/src/blank'
import { buildSheetsFixture } from '../apps/sheets/tests/fixture-builder'

const REPO = resolve(__dirname, '..')
const SHELL_DIR = join(REPO, 'apps/shell')
const OUT_DIR = join(REPO, 'apps/shell/src/renderer/src/i18n/help/topics/img')

/** The UI languages the manual ships figures for; see the registry's language set. */
const LANGS = process.argv.slice(2).filter((a) => !a.startsWith('-'))
const TARGET_LANGS = LANGS.length ? LANGS : ['zh', 'en']

/** A one-page PDF, written by hand so the capture needs no PDF library. */
function minimalPdf(): Buffer {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    null, // stream, filled below
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  const stream = `BT /F1 18 Tf 72 700 Td (GenOffice Manual Sample) Tj ET\nBT /F1 11 Tf 72 660 Td (This paragraph is here so you can try text selection.) Tj ET\nBT /F1 11 Tf 72 620 Td (SECRET ONE needs redaction.) Tj ET\nBT /F1 11 Tf 72 580 Td (More filler text on the sample page.) Tj ET`
  objects[3] = `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`
  let pdf = '%PDF-1.4\n'
  const offsets: number[] = []
  objects.forEach((body, i) => {
    offsets.push(pdf.length)
    pdf += `${i + 1} 0 obj\n${body}\nendobj\n`
  })
  const xref = pdf.length
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Buffer.from(pdf, 'latin1')
}

/** The scratch profile: never the developer's own HOME or userData. */
interface Scratch {
  home: string
  userData: string
  samples: string
  dispose: () => void
}

async function makeScratch(): Promise<Scratch> {
  const root = mkdtempSync(join(tmpdir(), 'genoffice-help-shots-'))
  const home = join(root, 'home')
  const userData = join(root, 'userdata')
  const samples = join(root, 'help-samples')
  // Fixed and platform-shaped, not a per-run temp path: a value that changes
  // per run or per machine makes the committed figure churn, and a path
  // carrying the developer's home directory puts their macOS username in a PNG
  // forever. os.tmpdir() is not enough — on macOS it is
  // /var/folders/<xx>/<per-user-hash>/T, so it differs per machine. This has to
  // be absolute and creatable, which is all resolveDefaultSaveDir asks of a
  // configured dir.
  const saveDir = process.platform === 'win32' ? 'C:\\GenOffice' : '/tmp/GenOffice'
  for (const dir of [home, userData, samples, saveDir]) mkdirSync(dir, { recursive: true })

  // the file list the Home figure shows; types only — the figures are about
  // the surface around them, not the documents
  writeFileSync(join(samples, 'doc.docx'), await buildBlankDocx())
  writeFileSync(join(samples, 'deck.pptx'), await createBlankPptx())
  writeFileSync(join(samples, 'book.xlsx'), await buildSheetsFixture())
  writeFileSync(join(samples, 'sample.pdf'), minimalPdf())
  writeFileSync(join(samples, 'notes.md'), '# Sample\n\nA markdown note.\n')

  // folderRoots is a plain list of absolute paths, so the sidebar tree can be
  // seeded without driving the native folder picker.
  //
  // defaultSaveDir needs seeding too, and for a different reason: the Settings
  // figure shows it, and the fallback is the real user's ~/Documents — so a
  // capture without this line puts the developer's home directory and macOS
  // username into a committed PNG. The path has to be absolute and writable
  // (resolveDefaultSaveDir ignores a configured one that is not).
  writeFileSync(
    join(userData, 'app-settings.json'),
    JSON.stringify(
      {
        onboardingSeen: true,
        folderRoots: [samples],
        defaultSaveDir: saveDir,
      },
      null,
      2,
    ),
  )

  return { home, userData, samples, dispose: () => rmSync(root, { recursive: true, force: true }) }
}

async function launch(scratch: Scratch, lang: string): Promise<ElectronApplication> {
  const require = createRequire(join(SHELL_DIR, 'package.json'))
  // ELECTRON_RUN_AS_NODE (set by CI hosts) would boot Electron as plain Node
  const { ELECTRON_RUN_AS_NODE: _drop, ...hostEnv } = process.env
  const args: string[] = []
  if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
  args.push(SHELL_DIR)
  return electron.launch({
    executablePath: require('electron') as unknown as string,
    args,
    env: {
      ...hostEnv,
      HOME: scratch.home,
      GENOFFICE_USER_DATA: scratch.userData,
      GENOFFICE_AUTH_DIR: join(scratch.home, '.genoffice'),
      GENOFFICE_NO_SPARE_VIEW: '1',
      GENOFFICE_LANG: lang,
    },
  })
}

/**
 * Size the capture to match the committed figures: a 1360x850 CSS viewport at
 * 2x, so the PNG is 2720x1700. The device scale has to be set through CDP —
 * setViewportSize alone is CSS pixels, and a half-resolution figure goes soft
 * on every retina display the manual is read on.
 */
async function frame(page: Page): Promise<void> {
  await page.setViewportSize({ width: 1360, height: 850 })
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1360,
    height: 850,
    deviceScaleFactor: 2,
    mobile: false,
  })
  await page.waitForTimeout(600)
}

/** Dismiss the star prompt: it is a promo overlay, not part of the screen the
 *  figure is documenting, and it sits over the file list. */
async function dismissStarPrompt(page: Page): Promise<void> {
  const close = page.locator('.star-prompt button, .star-prompt [aria-label]').last()
  if (await close.count()) await close.click().catch(() => {})
  await page.keyboard.press('Escape').catch(() => {})
  await page.waitForTimeout(300)
}

async function shoot(page: Page, name: string, lang: string): Promise<void> {
  const file = join(OUT_DIR, `${name}.${lang}.png`)
  await page.screenshot({ path: file })
  process.stdout.write(`wrote ${name}.${lang}.png\n`)
}

async function main(): Promise<void> {
  mkdirSync(OUT_DIR, { recursive: true })
  for (const lang of TARGET_LANGS) {
    const scratch = await makeScratch()
    let app: ElectronApplication | null = null
    try {
      app = await launch(scratch, lang)
      const page = app.windows()[0] ?? (await app.firstWindow())
      await frame(page)

      await dismissStarPrompt(page)

      // The folder view is the figure: it carries everything the landing view
      // has (the quick-start cards) plus the folders, the file table and the
      // search box, and it needs no recents — which a fresh profile has none
      // of, since a file has to be opened for one to be recorded.
      //
      // `.tree-item` is the sidebar, `.recent-row` the table: both render a
      // "more actions" trigger, so an unscoped match opens the folder's menu
      // instead of a file's.
      // click the name, not the row: the <li> also carries the expand chevron
      // and the row actions, and its centre can land on either
      const treeFolder = page
        .locator('.tree-item')
        .filter({ hasText: 'help-samples' })
        .locator('.tree-name')
      await treeFolder.first().click()
      const rows = page.locator('.recent-row')
      try {
        await rows.first().waitFor({ state: 'visible', timeout: 15_000 })
      } catch {
        const seen = await page.evaluate(() =>
          [...document.querySelectorAll('[class*="row"],[class*="item"]')]
            .map((e) => `${(e.className || '').toString().slice(0, 40)}`)
            .filter((c, i, a) => a.indexOf(c) === i)
            .slice(0, 25)
            .join(' | '),
        )
        throw new Error(`folder view did not list files; rows on screen: ${seen}`)
      }

      // 1. the Home screen, signed out, listing the sample files
      await shoot(page, 'home-screen', lang)

      // 2. the same list with a file's own menu open — the surface the
      //    file-operations topic is about
      // the row's own trigger, by class: its aria-label is localized, and the
      // sidebar tree carries a trigger with the same wording
      // a file row specifically: the first row is the sub-folder, and its menu
      // is a folder menu (rename/move/delete *folder*)
      const fileRow = rows.filter({ hasText: 'doc.docx' }).first()
      await fileRow.locator('.recent-actions .more-btn').click()
      await page.locator('.ctx-menu, .row-menu, [role="menu"]').first().waitFor({ timeout: 10_000 })
      await shoot(page, 'file-ops', lang)
      await page.keyboard.press('Escape')
      await page.waitForTimeout(300)

      // 3. Settings. There is no gear in the top-right: the modal opens from the
      //    account row at the bottom-left (Home.tsx's AccountEntry says so), and
      //    its aria-label is localized, so match the class rather than the text.
      //    The figure is the General pane because that is where the article's
      //    own subjects live — language and theme; Account is where the e-mail
      //    used to leak from.
      await page.locator('.account-btn').click()
      const nav = page.locator('.set-nav-item').filter({ hasText: /General|通用|一般/ })
      if (await nav.count()) await nav.first().click()
      await page.waitForTimeout(500)
      await shoot(page, 'settings-integrations', lang)
    } finally {
      await app?.close().catch(() => {})
      scratch.dispose()
    }
  }
}

void main()
