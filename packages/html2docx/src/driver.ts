/**
 * The browser the converter renders in. The extraction layer runs inside the
 * page (DOM + computed style + element screenshots), so the host supplies a
 * page-like driver: Playwright for tests and the CLI, Electron's webContents
 * in the app. Implementations share the contract test in tests/.
 */
export interface Viewport {
  width: number
  height: number
  deviceScaleFactor: number
}

export interface ClipRect {
  x: number
  y: number
  width: number
  height: number
}

export interface ScreenshotOptions {
  /** Page-coordinate clip; captures beyond the viewport. */
  clip?: ClipRect
  /** Transparent instead of the page background (isolated floats). */
  omitBackground?: boolean
}

export interface BrowserDriver {
  setViewport(viewport: Viewport): Promise<void>
  /** Resolves once the network is idle; a timeout resolves too (the page is used as-is). */
  goto(url: string, options: { timeoutMs: number }): Promise<void>
  /** `fn` is serialized with toString() and gets one JSON-serializable argument; a string is evaluated as an expression. */
  evaluate<T>(fn: string | ((arg?: any) => T | Promise<T>), arg?: unknown): Promise<T>
  /** Inject a classic script (the in-page extractor helpers). */
  addScript(source: string): Promise<void>
  /** Runs before any page script on every navigation (debug tracing only). */
  addInitScript(source: string): Promise<void>
  screenshot(options: ScreenshotOptions): Promise<Uint8Array>
  /** Screenshot of the first element matching `selector`'s box (even when hidden or
   *  zero-opacity — never wait for visibility); null when absent or empty. */
  screenshotElement(
    selector: string,
    options: Pick<ScreenshotOptions, 'omitBackground'>,
  ): Promise<Uint8Array | null>
  elementExists(selector: string): Promise<boolean>
  close(): Promise<void>
}
