/**
 * BrowserDriver over playwright-core driving a locally installed Chrome. Used
 * by the tests and the CLI; the app uses the Electron driver instead. Not
 * re-exported from the package index so app bundles never pull in Playwright.
 */
import { existsSync } from 'node:fs'
import { chromium } from 'playwright-core'
import type { Browser, BrowserContext, Page } from 'playwright-core'
import type { BrowserDriver, ScreenshotOptions, Viewport } from '../driver'

const CHROME_CANDIDATES = [
  process.env.CHROME_PATH,
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  '/usr/bin/google-chrome',
  '/usr/bin/google-chrome-stable',
  '/usr/bin/chromium',
  '/usr/bin/chromium-browser',
].filter((p): p is string => Boolean(p))

export function findChrome(): string {
  const found = CHROME_CANDIDATES.find((p) => existsSync(p))
  if (!found) throw new Error('Chrome/Chromium not found. Set the CHROME_PATH env variable.')
  return found
}

export function launchChrome(): Promise<Browser> {
  return chromium.launch({
    executablePath: findChrome(),
    headless: true,
    args: ['--no-sandbox', '--disable-dev-shm-usage', '--force-device-scale-factor=2'],
  })
}

export class PlaywrightDriver implements BrowserDriver {
  private constructor(
    private readonly context: BrowserContext,
    private readonly page: Page,
  ) {}

  /** One isolated context per conversion; the browser is shared by the caller. */
  static async create(browser: Browser, viewport: Viewport): Promise<PlaywrightDriver> {
    const context = await browser.newContext({
      viewport: { width: viewport.width, height: viewport.height },
      deviceScaleFactor: viewport.deviceScaleFactor,
    })
    return new PlaywrightDriver(context, await context.newPage())
  }

  async setViewport(viewport: Viewport): Promise<void> {
    // deviceScaleFactor is fixed per context (set in create); the converter
    // only ever changes width/height afterwards.
    await this.page.setViewportSize({ width: viewport.width, height: viewport.height })
  }

  async goto(url: string, { timeoutMs }: { timeoutMs: number }): Promise<void> {
    try {
      await this.page.goto(url, { waitUntil: 'networkidle', timeout: timeoutMs })
    } catch {
      // network never went idle: convert whatever has rendered (matches the CLI's behavior)
    }
  }

  evaluate<T>(fn: string | ((arg?: any) => T | Promise<T>), arg?: unknown): Promise<T> {
    if (typeof fn === 'string') return this.page.evaluate(fn) as Promise<T>
    return this.page.evaluate(fn as (arg: unknown) => T, arg)
  }

  async addScript(source: string): Promise<void> {
    await this.page.addScriptTag({ content: source })
  }

  async addInitScript(source: string): Promise<void> {
    await this.page.addInitScript(source)
  }

  async screenshot(options: ScreenshotOptions): Promise<Uint8Array> {
    // fullPage makes `clip` a document-coordinate rect captured beyond the
    // viewport (CDP captureBeyondViewport); without it Playwright trims the
    // clip to the viewport and below-the-fold shots come back empty.
    return this.page.screenshot({
      type: 'png',
      fullPage: Boolean(options.clip),
      clip: options.clip,
      omitBackground: options.omitBackground,
    })
  }

  async screenshotElement(
    selector: string,
    options: Pick<ScreenshotOptions, 'omitBackground'>,
  ): Promise<Uint8Array | null> {
    const el = await this.page.$(selector)
    if (!el) return null
    // Not elementHandle.screenshot(): that waits for the element to be visible
    // and stable, but the converter also shoots boxes it has just hidden
    // (broken-image placeholders). Clip the element's page box instead.
    // Origin clamped like the converter's own clips: a box overhanging the
    // document's top/left edge is captured from the edge instead.
    const box = await el.evaluate((node) => {
      const r = node.getBoundingClientRect()
      return {
        x: Math.max(0, Math.round(r.left + window.scrollX)),
        y: Math.max(0, Math.round(r.top + window.scrollY)),
        width: Math.round(r.width),
        height: Math.round(r.height),
      }
    })
    if (box.width <= 0 || box.height <= 0) return null
    return this.screenshot({ clip: box, omitBackground: options.omitBackground })
  }

  async elementExists(selector: string): Promise<boolean> {
    return (await this.page.$(selector)) !== null
  }

  async close(): Promise<void> {
    await this.context.close()
  }
}
