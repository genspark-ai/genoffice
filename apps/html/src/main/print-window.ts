import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/**
 * Print a document through the system print dialog in a throwaway window.
 *
 * These are the same four platform rules
 * `apps/slides/src/main/print-window.ts` encodes, kept deliberately in step
 * with it rather than extracted: the obvious shared home is
 * `packages/electron-utils` (both apps already depend on it, and its
 * print-html-pdf.ts is the sibling of this), but the two apps can only be
 * verified together, and this worktree's `node_modules` is a symlink into
 * another checkout, so a cross-app move here could not be typechecked or
 * tested. The duplication is one file and one exported constant; the
 * `PRINT_READY_SCRIPT` below is byte-identical to the slides one. Worth
 * extracting as a follow-up that can run both suites.
 *
 * The rules are not obvious, so they are spelled out rather than left to the
 * caller:
 *
 * 1. Windows attaches the native dialog to the printed window, so that window
 *    must be visible or the dialog never appears. Every other platform shows a
 *    sheet, so the window stays hidden there.
 * 2. Fonts and images must be decoded before `print()`, or the output misses
 *    webfonts and bitmaps. This needs `executeJavaScript`, which Chromium
 *    rejects outright when the window is created with `javascript: false`, so
 *    the caller must leave scripting on. See `html-main.ts` for why the print
 *    window differs from the PDF export window on that point.
 * 3. `print()`'s callback reports both a `success` flag and a `failureReason`;
 *    a reason is not by itself an error. "Print job canceled" is the user
 *    closing the dialog, which is their own outcome and must stay silent.
 * 4. The window has to outlive the dialog: destroying it when `print()` is
 *    called closes the sheet as it opens, so teardown waits for the callback.
 *
 * The window is supplied by the caller so each app keeps its own
 * `webPreferences` and parent-window handling.
 */
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

/** Fonts and every bitmap decoded before printing, or pages print with fallback
 * fonts and missing images. Mirrors the slides print path exactly. */
export const PRINT_READY_SCRIPT =
  'Promise.all([document.fonts.ready, ...Array.from(document.images).map((i) => i.decode().catch(() => {}))])'

/**
 * A user-authored document can keep `executeJavaScript` pending forever (a
 * script that never yields — the same hazard the HTML docx export guards
 * against). Combined with the renderer's in-flight print guard, a promise that
 * never settles would wedge printing for the rest of the session, which is
 * worse than printing without the extra wait. Bounded, then print anyway.
 */
const PRINT_READY_TIMEOUT_MS = 5_000

export interface PrintDocumentOptions {
  /** Final document HTML, already rewritten for asset resolution by the caller. */
  html: string
  /** Throwaway window; the caller keeps ownership of its creation options. */
  window: PrintDialogWindow
  /** File name written inside the temp dir. */
  fileName?: string
  /** Temp-dir prefix, so a leaked dir is attributable to the calling app. */
  dirPrefix?: string
  platform?: NodeJS.Platform
}

export async function printHtmlDocument({
  html,
  window: win,
  fileName = 'print.html',
  dirPrefix = 'genoffice-print-',
  platform = process.platform,
}: PrintDocumentOptions): Promise<PrintDialogOutcome> {
  let tempDir: string | null = null
  try {
    // Through a temp file, not a data: URL: long documents truncate the URL.
    tempDir = await mkdtemp(join(tmpdir(), dirPrefix))
    const htmlPath = join(tempDir, fileName)
    await writeFile(htmlPath, html, 'utf8')
    await win.loadFile(htmlPath)
    // Best effort. A window that cannot run the probe (or a frame that died
    // under us) still prints; it just may miss webfonts, which is the state
    // this app shipped before the wait existed.
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
    // Windows attaches the native dialog to the printed window, which must be
    // visible. Slides does the same; without it no dialog appears at all.
    if (platform === 'win32') {
      win.show()
      win.focus()
    }
    // Rule 4: the dialog is modal to this window, so it must outlive the call.
    const result = await new Promise<{ success: boolean; failureReason: string }>((resolve) => {
      win.webContents.print({ silent: false, printBackground: true }, (success, failureReason) =>
        resolve({ success, failureReason }),
      )
    })
    if (!result.success) {
      // Rule 3: a cancelled dialog is the user's own outcome, not an error.
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
