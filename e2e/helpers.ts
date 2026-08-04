/**
 * Shared launcher for Electron E2E tests.
 *
 * Each test boots the built shell (`apps/shell/out`) against a scratch
 * userData dir (via GENOFFICE_USER_DATA) so runs never touch real settings
 * and never collide with a running install's single-instance lock.
 * Build first: `npm run build:all`.
 */
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'

export const SHELL_DIR = resolve(__dirname, '../apps/shell')
export const ARTIFACTS_DIR = resolve(__dirname, 'artifacts')

const SHELL_MAIN = join(SHELL_DIR, 'out/main/index.js')

interface LaunchOptions {
  /** reuse a previous scratch dir to simulate a second launch */
  userDataDir?: string
  /** UI language override (GENOFFICE_LANG); defaults to English for stable assertions */
  lang?: string
  /** pre-seed app-settings.json with onboardingSeen=true to start at the home screen */
  onboardingSeen?: boolean
  /** subdir of e2e/artifacts to store this launch's video in */
  videoDir: string
  /** absolute document path passed as argv, opened in an editor tab on launch */
  openFile?: string
}

export interface LaunchedApp {
  app: ElectronApplication
  page: Page
  userDataDir: string
}

export async function launchShell(options: LaunchOptions): Promise<LaunchedApp> {
  if (!existsSync(SHELL_MAIN)) {
    throw new Error(`Missing build output at ${SHELL_MAIN} — run \`npm run build:all\` first`)
  }
  const userDataDir = options.userDataDir ?? (await mkdtemp(join(tmpdir(), 'genoffice-e2e-')))
  if (options.onboardingSeen) {
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true }),
    )
  }
  const require = createRequire(join(SHELL_DIR, 'package.json'))
  const executablePath = require('electron') as unknown as string
  // ELECTRON_RUN_AS_NODE (set by VS Code/CI hosts) would boot Electron as
  // plain Node with no windows — strip it so the app always starts as an app
  const { ELECTRON_RUN_AS_NODE: _electronRunAsNode, ...hostEnv } = process.env
  // Linux CI runners restrict unprivileged user namespaces (no usable SUID
  // sandbox) and run under xvfb without GPU — without these the window opens
  // but the renderer never loads. The suite drives trusted local builds only.
  // Switches go before the app path so Chromium is guaranteed to consume them
  // and they never leak into the argv the app parses for documents to open.
  const args: string[] = []
  if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
  args.push(SHELL_DIR)
  if (options.openFile) args.push(options.openFile)
  const app = await electron.launch({
    executablePath,
    args,
    env: {
      ...hostEnv,
      GENOFFICE_USER_DATA: userDataDir,
      GENOFFICE_LANG: options.lang ?? 'en',
      ...(process.platform === 'linux' ? { ELECTRON_DISABLE_SANDBOX: '1' } : {}),
    },
    // Playwright's Electron screencast wedges the page CDP session on Linux
    // (page.url() stays empty, no lifecycle events, evaluate hangs) — record
    // only where it works
    recordVideo:
      process.platform === 'linux'
        ? undefined
        : {
            dir: join(ARTIFACTS_DIR, options.videoDir),
            size: { width: 1280, height: 800 },
          },
  })
  const page = await app.firstWindow()
  await waitForDocumentReady(app, page)
  return { app, page, userDataDir }
}

/**
 * Playwright can attach to the Electron window mid-navigation and miss the
 * load lifecycle events entirely (Linux timing) — waitForLoadState then hangs
 * on a page that is actually loaded. Polling through evaluate uses the live
 * CDP session instead of the missed events.
 */
async function waitForDocumentReady(
  app: ElectronApplication,
  page: Page,
  timeoutMs = 30_000,
): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    // the pre-navigation about:blank document also reports readyState complete
    const ready = await page
      .evaluate(() =>
        document.readyState !== 'loading' && window.location.href !== 'about:blank'
          ? document.readyState
          : null,
      )
      .catch(() => null)
    if (ready) return
    await new Promise((r) => setTimeout(r, 100))
  }
  // process list tells renderer-spawn failures apart from slow loads
  const diag = await app
    .evaluate(({ app: electronApp, BrowserWindow }) => ({
      processes: electronApp.getAppMetrics().map((m) => m.type),
      contents: BrowserWindow.getAllWindows().map((w) => w.webContents.getURL()),
    }))
    .catch((e) => String(e))
  throw new Error(`Shell window never loaded (url: ${page.url()}, diag: ${JSON.stringify(diag)})`)
}

/**
 * Close the app and return the recorded video path for the given page.
 *
 * Open editor tabs trigger a native Save/Don't Save/Cancel dialog on close,
 * which would block app.close() forever — stub the dialog to answer
 * "Don't Save" (button index 1) so shutdown stays unattended. If close still
 * hangs, kill the process after 20s so the suite never wedges.
 */
export async function closeAndSaveVideo(
  launched: LaunchedApp,
  name: string,
): Promise<string | undefined> {
  const video = launched.page.video()
  await launched.app
    .evaluate(({ dialog }) => {
      dialog.showMessageBox = (async () => ({
        response: 1,
        checkboxChecked: false,
      })) as typeof dialog.showMessageBox
    })
    .catch(() => {})
  let killTimer: NodeJS.Timeout | undefined
  await Promise.race([
    launched.app.close(),
    new Promise<void>((resolvePromise) => {
      killTimer = setTimeout(() => {
        launched.app.process().kill()
        resolvePromise()
      }, 20_000)
    }),
  ])
  if (killTimer) clearTimeout(killTimer)
  if (!video) return undefined
  const target = join(ARTIFACTS_DIR, 'videos', `${name}.webm`)
  try {
    await video.saveAs(target)
    return target
  } catch {
    return undefined
  }
}

export function screenshotPath(name: string): string {
  return join(ARTIFACTS_DIR, 'screenshots', `${name}.png`)
}

/**
 * Wait for a page whose URL contains `urlPart` (e.g. an editor WebContentsView).
 * Checks windows that already exist before listening, so it never races the
 * view being created between launch and the first waitForEvent call.
 */
export async function waitForPageWithUrl(
  app: ElectronApplication,
  urlPart: string,
  timeoutMs = 30_000,
): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      if (candidate.url().includes(urlPart)) return candidate
      // page.url() stays empty when attach raced navigation; ask the document
      const href = await candidate.evaluate(() => window.location.href).catch(() => '')
      if (href.includes(urlPart)) return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error(`No window with URL containing "${urlPart}"`)
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}
