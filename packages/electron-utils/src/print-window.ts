import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/** Throwaway window supplied by the caller so each app keeps its own webPreferences. */
export interface PrintDialogWindow {
  loadFile(path: string): Promise<void>
  webContents: {
    executeJavaScript(script: string, userGesture?: boolean): Promise<unknown>
    print(
      options: { silent: boolean; printBackground: boolean },
      callback: (success: boolean, failureReason: string) => void,
    ): void
  }
  show(): void
  focus(): void
  isDestroyed(): boolean
  destroy(): void
}

export type PrintDialogOutcome =
  | { ok: true }
  /** the user closed the system dialog: their own outcome, not a failure */
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** Fonts and bitmaps decoded before print(), or pages print with fallback fonts and missing images. */
export const PRINT_READY_SCRIPT =
  'Promise.all([document.fonts.ready, ...Array.from(document.images).map((i) => i.decode().catch(() => {}))])'

// a user-authored script that never yields must not wedge printing for the session
const PRINT_READY_TIMEOUT_MS = 5_000

export interface PrintDocumentOptions {
  /** final document HTML, assets already resolved by the caller */
  html: string
  /** needs scripting enabled: the readiness probe runs through executeJavaScript */
  window: PrintDialogWindow
  fileName?: string
  /** temp-dir prefix, so a leaked dir is attributable to the calling app */
  dirPrefix?: string
  platform?: NodeJS.Platform
}

/**
 * Load a print document into the window and wait for its readiness probe
 * (fonts and images decoded). The caller owns the window and must rm the
 * returned temp dir afterwards. Shared by the system-dialog print and the
 * silent printToPDF routes, which need the same loaded document.
 */
export async function loadPrintHtml(
  win: PrintDialogWindow,
  html: string,
  fileName = 'print.html',
  dirPrefix = 'genoffice-print-',
): Promise<{ tempDir: string }> {
  // a temp file, not a data: URL: long documents truncate the URL
  const tempDir = await mkdtemp(join(tmpdir(), dirPrefix))
  try {
    const htmlPath = join(tempDir, fileName)
    await writeFile(htmlPath, html, 'utf8')
    await win.loadFile(htmlPath)
    let timer: ReturnType<typeof setTimeout> | undefined
    await Promise.race([
      win.webContents.executeJavaScript(PRINT_READY_SCRIPT, true),
      new Promise<void>((resolve) => {
        timer = setTimeout(resolve, PRINT_READY_TIMEOUT_MS)
      }),
    ]).catch((err: unknown) => {
      console.error('[print] readiness probe failed:', err)
    })
    clearTimeout(timer)
  } catch (err) {
    // a failed load leaves nothing for the caller to clean: they only ever
    // see a tempDir when the document is loaded and ready to print
    await rm(tempDir, { recursive: true, force: true })
    throw err
  }
  return { tempDir }
}

/**
 * Print a document through the system print dialog in a throwaway window.
 * The window outlives the dialog (destroying it as print() opens closes the
 * sheet) and is destroyed on every path.
 */
export async function printHtmlDocument({
  html,
  window: win,
  fileName = 'print.html',
  dirPrefix = 'genoffice-print-',
  platform = process.platform,
}: PrintDocumentOptions): Promise<PrintDialogOutcome> {
  let tempDir: string | null = null
  try {
    tempDir = (await loadPrintHtml(win, html, fileName, dirPrefix)).tempDir
    // Windows attaches the native dialog to the printed window, which must be visible
    if (platform === 'win32') {
      win.show()
      win.focus()
    }
    const result = await new Promise<{ success: boolean; failureReason: string }>((resolve) => {
      win.webContents.print({ silent: false, printBackground: true }, (success, failureReason) =>
        resolve({ success, failureReason }),
      )
    })
    if (!result.success) {
      if (result.failureReason === 'Print job canceled') return { ok: true, canceled: true }
      return { ok: false, error: result.failureReason || 'the print job failed' }
    }
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  } finally {
    try {
      if (!win.isDestroyed()) win.destroy()
    } finally {
      if (tempDir) await rm(tempDir, { recursive: true, force: true })
    }
  }
}
