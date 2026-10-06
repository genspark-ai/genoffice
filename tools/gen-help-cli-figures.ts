/**
 * Regenerate the three manual figures that are not screens: `npx tsx tools/gen-help-cli-figures.ts`
 *
 * ## Why these three are separate from gen-help-screenshots.ts
 *
 * The other fourteen figures are the running app. `install`, `cli` and `mcp` are
 * not screens, so they cannot be captured the way the app is: there is no window
 * that lists the installer files, prints `genoffice --help`, or registers an MCP
 * server. Composing those by hand would put invented text in a committed PNG, so
 * every line in these three figures is stdout captured from the real binary.
 *
 * ## Why the CLI is run as a bundle and not through tsx
 *
 * `npx tsx packages/cli/src/index.ts --help` does not work and cannot be made to:
 * `packages/pptx-ops/src/op-docs.ts` imports its op guides with Vite's `?raw`
 * suffix, which is not a Node loader feature, so the run dies on
 * `ERR_UNKNOWN_FILE_EXTENSION ".md"`. The supported build is the esbuild bundle
 * (`npm run build -w @genoffice/cli`), which resolves `?raw` and is what the
 * packaged app actually ships. So the figure shows what the shipped command
 * prints, not what a source run would.
 *
 * ## Why mcp is captured with `--help` and not `mcp install all`
 *
 * Both were run. `mcp install all` and `mcp list` print the absolute path of the
 * resolved `genoffice` binary and the absolute path of every agent's config file
 * under `$HOME`, so a capture of either puts this machine's build path and
 * account name into a committed PNG and churns on every run. `mcp --help` is the
 * same real command printing its real usage, carries no path at all, and covers
 * what the article documents: the install/uninstall/list syntax, `--http`,
 * `--host`, `--token`, `--compact-schemas`, `--dir` and `--force`.
 *
 * ## Why the figure looks like the app and not like a terminal theme
 *
 * Green-on-black would be a picture of something this product does not ship. The
 * colors below are the app's own `[data-theme='dark']` block from
 * `packages/ui/src/tokens.css`, and the type is the same stack as the app's
 * `.set-code` block in `apps/shell/src/renderer/src/settings.css`
 * (ui-monospace, 12px). The card is `--surface` on `--canvas` with a
 * `--border` edge, which is how the app draws a panel. Same width and scale as
 * the other figures, 1360 CSS px at 2x; the height follows the captured output,
 * because a fixed 850px box leaves `mcp --help` a third empty.
 *
 * `install` is the exception: it is a real screen, so it is captured from the
 * running app on the same scratch profile the app generator uses, and nothing
 * of the developer's own session is read. No command run here writes outside
 * the scratch directories, and every one of them is deleted afterwards.
 */
import {
  _electron as electron,
  chromium,
  type ElectronApplication,
  type Page,
} from '@playwright/test'
import { spawnSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const REPO = resolve(__dirname, '..')
const SHELL_DIR = join(REPO, 'apps/shell')
const OUT_DIR = join(REPO, 'apps/shell/src/renderer/src/i18n/help/topics/img')
const CLI_BUNDLE = join(REPO, 'packages/cli/dist/genoffice.cjs')

/** Only these three names; the other figures in OUT_DIR belong to the other script. */
type Figure = 'cli' | 'mcp' | 'install'

/**
 * Run the real CLI and take its stdout verbatim.
 *
 * A non-zero exit is not a failure here: `mcp install all` on a machine with no
 * coding agent is a real result the article describes, so the exit code is
 * reported rather than turned into a thrown error that would silently invite
 * hand-written output.
 */
function capture(args: string[], env: NodeJS.ProcessEnv = {}): string {
  const run = spawnSync(process.execPath, [CLI_BUNDLE, ...args], {
    encoding: 'utf8',
    env: { ...process.env, ...env },
    maxBuffer: 32 * 1024 * 1024,
  })
  if (run.error) throw run.error
  const out = `${run.stdout ?? ''}${run.stderr ?? ''}`.replace(/\s+$/, '')
  if (!out) throw new Error(`genoffice ${args.join(' ')} printed nothing (exit ${run.status})`)
  return out
}

/** The shipped bundle, built on demand so the figure can be regenerated from a clean tree. */
function ensureCliBundle(): void {
  if (existsSync(CLI_BUNDLE)) return
  process.stdout.write('building the CLI bundle...\n')
  const build = spawnSync('npm', ['run', 'build', '-w', '@genoffice/cli'], {
    cwd: REPO,
    encoding: 'utf8',
    stdio: 'inherit',
  })
  if (build.status !== 0 || !existsSync(CLI_BUNDLE)) {
    throw new Error('could not build packages/cli; run npm run build -w @genoffice/cli by hand')
  }
}

/** The captured text goes into the page as text, never as markup. */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/**
 * One captured command, drawn as a panel in the app's own dark theme.
 *
 * `command` is the line the reader would type and `output` is what came back;
 * they are kept apart so the figure reads as a session rather than as prose.
 */
function figureHtml(title: string, blocks: { command: string; output: string }[]): string {
  const body = blocks
    .map(
      (b) =>
        `<div class="cmd">$ <span class="name">${escapeHtml(b.command)}</span></div>` +
        `<pre class="out">${escapeHtml(b.output)}</pre>`,
    )
    .join('\n')
  return `<!doctype html>
<html data-theme="dark"><head><meta charset="utf-8"><title>${escapeHtml(title)}</title>
<style>
  /* App dark theme, copied from packages/ui/src/tokens.css [data-theme='dark']. */
  :root {
    --canvas: #2b2b2b; --surface: #1e1e1e; --surface-subtle: #2a2a2a;
    --border: #3a3a3a; --text-primary: #e4e4e4; --text-secondary: #a0a0a0;
    --text-tertiary: #8a8a8a; --brand: #4a9eff; --radius-12: 12px;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    width: 1360px; background: var(--canvas);
    font-family: 'Arial', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: var(--text-primary); padding: 40px; display: flex; flex-direction: column;
  }
  /* The app's own panel: --surface on --canvas with a --border edge. */
  .card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius-12); overflow: hidden;
    display: flex; flex-direction: column;
  }
  .head {
    display: flex; align-items: center; gap: 10px; padding: 14px 18px;
    border-bottom: 1px solid var(--border); background: var(--surface-subtle);
    font-size: 13px; color: var(--text-secondary);
  }
  .head .app { color: var(--text-primary); font-weight: 600; }
  .head .sep { color: var(--text-tertiary); }
  .term {
    /* The same stack the app uses for .set-code in settings.css, a step down
       from 12px so the whole of a long --help fits the figure without cropping. */
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-size: 11px; line-height: 1.45; padding: 14px 18px; overflow: hidden;
  }
  .cmd { color: var(--brand); margin-bottom: 10px; }
  .cmd .name { color: var(--text-primary); }
  .out {
    margin: 0 0 14px; white-space: pre-wrap; word-break: break-word;
    font: inherit; color: var(--text-primary);
  }
  .out:last-child { margin-bottom: 0; }
</style></head>
<body>
  <div class="card">
    <div class="head"><span class="app">GenOffice</span><span class="sep">—</span><span>${escapeHtml(title)}</span></div>
    <div class="term">
${body}
    </div>
  </div>
</body></html>
`
}

/**
 * Same width and scale as the other figures (1360 CSS px at 2x) so they sit
 * consistently in the article, but the height follows the captured output: a
 * fixed 850px box leaves `mcp --help` a third empty, which reads as a mistake.
 */
async function shootHtml(html: string, name: string): Promise<void> {
  const browser = await chromium.launch()
  try {
    const page = await browser.newPage({
      viewport: { width: 1360, height: 850 },
      deviceScaleFactor: 2,
    })
    await page.setContent(html, { waitUntil: 'load' })
    // A long unbreakable run that spills sideways is cropped, and cropping real
    // output is a lie about what the command printed, so it is a build failure.
    const spill = await page.evaluate(() => {
      const el = document.querySelector('.term') as HTMLElement | null
      return el ? el.scrollWidth - el.clientWidth : 0
    })
    if (spill > 0) {
      throw new Error(
        `${name}: the captured output overflows the figure by ${spill}px sideways; the type scale has to come down rather than clip it`,
      )
    }
    // Shrink the viewport onto the content: a screenshot is at least one
    // viewport tall, so a short capture would otherwise keep 850px of canvas.
    const height = await page.evaluate(
      () => document.documentElement.getBoundingClientRect().height,
    )
    await page.setViewportSize({ width: 1360, height: Math.ceil(height) })
    await page.screenshot({ path: join(OUT_DIR, `${name}.en.png`) })
    process.stdout.write(`wrote ${name}.en.png (1360x${Math.ceil(height)} @2x)\n`)
  } finally {
    await browser.close()
  }
}

/**
 * Dismiss the star prompt: it is a promo overlay, not part of the screen the
 * figure documents, and it lands on top of the pane on some runs. Same
 * dismissal the app generator does, so the two scripts agree on what a capture
 * is allowed to include.
 */
async function dismissStarPrompt(page: Page): Promise<void> {
  const close = page.locator('.star-prompt button, .star-prompt [aria-label]').last()
  if (await close.count()) await close.click().catch(() => {})
  await page.keyboard.press('Escape').catch(() => {})
  await page.waitForTimeout(300)
}

/**
 * The install figure: Settings > About, the one place in the app that reports
 * the build you are on and replaces it. Same scratch profile as the app
 * generator, so the figure cannot pick up a developer's account or home path.
 *
 * One row of this pane is not reproducible: "Open Source" appends the live
 * stargazer count, which the main process fetches from api.github.com. It is
 * real and correct in the figure, but it moves whenever the repo is starred, so
 * a regeneration that shows a different number is the count changing and not a
 * capture going wrong. Three ways of pinning it were tried and all failed
 * without touching product code, which is why it is documented instead of
 * hidden: contextBridge freezes window.aiOffice so the renderer cannot be made
 * to ignore it, a Chromium --host-resolver-rules switch does not reach the main
 * process's own fetch, and Electron strips NODE_OPTIONS=--require, so a preload
 * shim never runs. The app has no offline flag for it either. If this ever
 * needs to be stable, the fix belongs in the product (a capture flag next to
 * the fetch), not in this script.
 */
async function shootInstall(): Promise<void> {
  const require = createRequire(join(SHELL_DIR, 'package.json'))
  const { ELECTRON_RUN_AS_NODE: _drop, ...hostEnv } = process.env
  if (!existsSync(join(SHELL_DIR, 'out/main/index.js'))) {
    throw new Error(
      'apps/shell is not built; run `npm run build:all` first so the About pane can be captured',
    )
  }
  const root = mkdtempSync(join(tmpdir(), 'genoffice-install-figs-'))
  const home = join(root, 'home')
  const userData = join(root, 'userdata')
  for (const dir of [home, userData]) mkdirSync(dir, { recursive: true })
  // Onboarding has to be marked seen: its welcome overlay sits over the window
  // and swallows the click that opens Settings. The app generator seeds the same
  // flag, and a fresh profile is otherwise in exactly this state.
  writeFileSync(
    join(userData, 'app-settings.json'),
    JSON.stringify({ onboardingSeen: true }, null, 2),
  )
  let app: ElectronApplication | null = null
  try {
    app = await electron.launch({
      executablePath: require('electron') as unknown as string,
      args: [SHELL_DIR],
      env: {
        ...hostEnv,
        HOME: home,
        GENOFFICE_USER_DATA: userData,
        GENOFFICE_AUTH_DIR: join(home, '.genoffice'),
        GENOFFICE_NO_SPARE_VIEW: '1',
        GENOFFICE_LANG: 'en',
      },
    })
    const page: Page = app.windows()[0] ?? (await app.firstWindow())
    await page.setViewportSize({ width: 1360, height: 850 })
    const cdp = await page.context().newCDPSession(page)
    await cdp.send('Emulation.setDeviceMetricsOverride', {
      width: 1360,
      height: 850,
      deviceScaleFactor: 2,
      mobile: false,
    })
    await page.waitForTimeout(2500)
    await dismissStarPrompt(page)
    await page.locator('.account-btn').click()
    // the nav label is localized, so match the class and the English word
    const about = page.locator('.set-nav-item').filter({ hasText: /About/ }).first()
    await about.waitFor({ state: 'visible', timeout: 20_000 })
    await about.click()
    await page.waitForTimeout(1200)
    await page.screenshot({ path: join(OUT_DIR, 'install.en.png') })
    process.stdout.write('wrote install.en.png\n')
  } finally {
    await app?.close().catch(() => {})
    rmSync(root, { recursive: true, force: true })
  }
}

async function main(): Promise<void> {
  mkdirSync(OUT_DIR, { recursive: true })
  const want = new Set(process.argv.slice(2).filter((a) => !a.startsWith('-')))
  const take = (name: Figure): boolean => want.size === 0 || want.has(name)

  if (take('cli')) {
    ensureCliBundle()
    await shootHtml(
      figureHtml('genoffice --help', [
        { command: 'genoffice --help', output: capture(['--help']) },
      ]),
      'cli',
    )
  }

  if (take('mcp')) {
    ensureCliBundle()
    await shootHtml(
      figureHtml('genoffice mcp --help', [
        { command: 'genoffice mcp --help', output: capture(['mcp', '--help']) },
      ]),
      'mcp',
    )
  }

  if (take('install')) await shootInstall()
}

void main()
