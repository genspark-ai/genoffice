/**
 * html2docx BrowserDriver over a hidden BrowserWindow. The converter drives
 * the page through webContents.executeJavaScript and the Chrome DevTools
 * Protocol (device metrics, network idle, screenshots), so the export sees
 * the document exactly as the preview does: scripts on, assets through
 * html-asset://, no preload and no Node.
 */
import { BrowserWindow } from 'electron'
import type { Debugger, WebContents } from 'electron'
import type { BrowserDriver, ScreenshotOptions, Viewport } from '../../../../packages/html2docx/src'

const NETWORK_QUIET_MS = 500
const POLL_MS = 100

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

export class ElectronBrowserDriver implements BrowserDriver {
  private readonly inflight = new Set<string>()

  private constructor(
    private readonly win: BrowserWindow,
    private readonly wc: WebContents,
    private readonly cdp: Debugger,
  ) {}

  static async create(viewport: Viewport): Promise<ElectronBrowserDriver> {
    const win = new BrowserWindow({
      show: false,
      width: viewport.width,
      height: viewport.height,
      webPreferences: {
        sandbox: true,
        contextIsolation: true,
        nodeIntegration: false,
        backgroundThrottling: false,
      },
    })
    const wc = win.webContents
    wc.setWindowOpenHandler(() => ({ action: 'deny' }))
    // nobody can answer a beforeunload prompt in a hidden window
    wc.on('will-prevent-unload', (event) => event.preventDefault())
    // a fresh window has no renderer yet and CDP commands would wait for one
    await wc.loadURL('about:blank')
    const cdp = wc.debugger
    cdp.attach('1.3')
    const driver = new ElectronBrowserDriver(win, wc, cdp)
    cdp.on('message', (_event, method, params) => driver.onCdpMessage(method, params))
    await cdp.sendCommand('Page.enable')
    await cdp.sendCommand('Network.enable')
    await driver.setViewport(viewport)
    return driver
  }

  private onCdpMessage(method: string, params: any): void {
    if (method === 'Network.requestWillBeSent') this.inflight.add(params.requestId)
    else if (method === 'Network.loadingFinished' || method === 'Network.loadingFailed') {
      this.inflight.delete(params.requestId)
    } else if (method === 'Page.javascriptDialogOpening') {
      // alert/confirm/prompt would block the renderer forever; dismiss like Playwright does
      void this.cdp.sendCommand('Page.handleJavaScriptDialog', { accept: false }).catch(() => {})
    }
  }

  async setViewport(viewport: Viewport): Promise<void> {
    this.win.setContentSize(viewport.width, viewport.height)
    await this.cdp.sendCommand('Emulation.setDeviceMetricsOverride', {
      width: viewport.width,
      height: viewport.height,
      deviceScaleFactor: viewport.deviceScaleFactor,
      mobile: false,
    })
  }

  async goto(url: string, { timeoutMs }: { timeoutMs: number }): Promise<void> {
    const deadline = Date.now() + timeoutMs
    this.inflight.clear()
    // loadURL settles on load or failure; a stalled subresource can hold it past
    // the deadline, so stop the load and convert what has rendered (CLI behavior)
    let timer: NodeJS.Timeout | undefined
    const timedOut = new Promise<'timeout'>((resolve) => {
      timer = setTimeout(() => resolve('timeout'), timeoutMs)
    })
    // a main-frame load error (or the ERR_ABORTED our stop() causes after a
    // timeout) means: convert whatever rendered, like the CLI does
    const load = this.wc.loadURL(url).then(
      () => 'loaded' as const,
      () => 'failed' as const,
    )
    try {
      const result = await Promise.race([load, timedOut])
      if (result === 'timeout' && !this.wc.isDestroyed()) this.wc.stop()
    } finally {
      clearTimeout(timer)
    }
    // networkidle: no request in flight for NETWORK_QUIET_MS; give up at the deadline
    let quietSince: number | null = null
    while (Date.now() < deadline) {
      if (this.inflight.size === 0) {
        quietSince ??= Date.now()
        if (Date.now() - quietSince >= NETWORK_QUIET_MS) return
      } else {
        quietSince = null
      }
      await sleep(POLL_MS)
    }
  }

  evaluate<T>(fn: string | ((arg?: any) => T | Promise<T>), arg?: unknown): Promise<T> {
    const expression =
      typeof fn === 'string'
        ? fn
        : `(${fn.toString()})(${arg === undefined ? '' : JSON.stringify(arg)})`
    return this.wc.executeJavaScript(expression, true) as Promise<T>
  }

  async addScript(source: string): Promise<void> {
    await this.wc.executeJavaScript(source, true)
  }

  async addInitScript(source: string): Promise<void> {
    await this.cdp.sendCommand('Page.addScriptToEvaluateOnNewDocument', { source })
  }

  async screenshot(options: ScreenshotOptions): Promise<Uint8Array> {
    if (options.omitBackground) {
      await this.cdp.sendCommand('Emulation.setDefaultBackgroundColorOverride', {
        color: { r: 0, g: 0, b: 0, a: 0 },
      })
    }
    try {
      const { data } = (await this.cdp.sendCommand('Page.captureScreenshot', {
        format: 'png',
        captureBeyondViewport: true,
        ...(options.clip ? { clip: { ...options.clip, scale: 1 } } : {}),
      })) as { data: string }
      return Buffer.from(data, 'base64')
    } finally {
      if (options.omitBackground) {
        await this.cdp.sendCommand('Emulation.setDefaultBackgroundColorOverride', {})
      }
    }
  }

  async screenshotElement(
    selector: string,
    options: Pick<ScreenshotOptions, 'omitBackground'>,
  ): Promise<Uint8Array | null> {
    const box = await this.evaluate<{ x: number; y: number; width: number; height: number } | null>(
      (sel: string) => {
        const node = document.querySelector(sel)
        if (!node) return null
        const r = node.getBoundingClientRect()
        return {
          x: Math.max(0, Math.round(r.left + window.scrollX)),
          y: Math.max(0, Math.round(r.top + window.scrollY)),
          width: Math.round(r.width),
          height: Math.round(r.height),
        }
      },
      selector,
    )
    if (!box || box.width <= 0 || box.height <= 0) return null
    return this.screenshot({ clip: box, omitBackground: options.omitBackground })
  }

  elementExists(selector: string): Promise<boolean> {
    return this.evaluate<boolean>((sel: string) => document.querySelector(sel) !== null, selector)
  }

  async close(): Promise<void> {
    try {
      this.cdp.detach()
    } catch {
      // already detached with the page
    }
    if (!this.win.isDestroyed()) this.win.destroy()
  }
}
