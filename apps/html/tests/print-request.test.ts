import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { HTML_CHANNELS } from '../src/shared/ipc'

// The main and preload modules import electron, and App.tsx pulls in the editor,
// preview and i18n, so none of the three can be loaded in this jsdom suite. The
// print wiring is therefore pinned by reading the sources: the defect being
// guarded is a missing registration, which is exactly what this checks.
const source = (relative: string) =>
  readFile(fileURLToPath(new URL(relative, import.meta.url)), 'utf8')

const APP = '../src/renderer/App.tsx'
const PRELOAD = '../src/preload/index.ts'
const MAIN = '../src/main/html-main.ts'

describe('HTML print wiring', () => {
  it('declares a print request channel and a print call distinct from it', () => {
    // Same channel for both would make the renderer's reply land on the menu's
    // own event and never reach the handler.
    expect(HTML_CHANNELS.printRequest).toBe('html:print-request')
    expect(HTML_CHANNELS.printHtml).toBe('html:print-html')
  })

  it('subscribes the renderer to the print request and disposes the subscription', async () => {
    const app = await source(APP)
    expect(app).toContain('window.htmlApi.onPrintRequest(')
    // Leaving the listener attached across an unmount/remount would print the
    // document twice from one keystroke.
    const dispose = app.match(/const offPrint = window\.htmlApi\.onPrintRequest\(/)
    expect(dispose, 'the subscription result must be kept for cleanup').not.toBeNull()
    expect(app).toMatch(/offPrint\(\)/)
  })

  it('exposes the print call over preload and handles it in main', async () => {
    const [preload, main] = await Promise.all([source(PRELOAD), source(MAIN)])
    expect(preload).toContain('onPrintRequest:')
    expect(preload).toMatch(/printHtml:.*HTML_CHANNELS\.printHtml/)
    expect(main).toMatch(/ipcMain\.handle\(\s*HTML_CHANNELS\.printHtml,/)
  })

  it('prints the document the renderer hands over, not the stale saved text', async () => {
    // The print request carries the live buffer: printing what was last written
    // to disk would silently drop unsaved edits.
    const app = await source(APP)
    expect(app).toMatch(/runPrint[\s\S]{0,900}serializeDocText\(/)
    expect(app).toMatch(/printHtml\(\{/)
  })

  it('gates the print and releases the gate on every exit path', async () => {
    // The gate itself, and its release, are proven in print-guard.test.ts; this
    // pins that App.tsx actually routes the keystroke through it.
    const app = await source(APP)
    expect(app).toMatch(/const printingRef = useRef\(false\)/)
    expect(app).toMatch(/runGuardedPrint\(\s*printingRef,/)
  })

  it('surfaces a failed print instead of swallowing ok:false', async () => {
    // The old handler ended in .catch(() => {}), so a printer that refused the
    // job left the user with a document that silently never printed.
    const app = await source(APP)
    expect(app).toMatch(/setNotice\(t\('printFailed', \{ error \}\)\)/)
    const handler = app.slice(app.indexOf('const runPrint = useCallback'))
    expect(handler).not.toContain('.catch(() => {})')
  })

  it('prints from a window that can run the fonts/images readiness probe', async () => {
    // Chromium rejects executeJavaScript when a window is created with
    // javascript: false, which would make the readiness wait impossible.
    const main = await source(MAIN)
    const printFn = main.slice(main.indexOf('function printHtml('))
    expect(printFn).toMatch(/new BrowserWindow\(\{\s*show: false/)
    expect(printFn).not.toMatch(/javascript: false/)
    expect(printFn).toContain('printHtmlDocument({')
  })

  it('hands the print outcome back to the renderer, cancel included', async () => {
    // The handler must not re-wrap the outcome into a bare { ok: true }: that
    // is what made every failure look like a success to the renderer.
    const main = await source(MAIN)
    const handler = main.slice(main.indexOf('HTML_CHANNELS.printHtml,'))
    expect(handler).toMatch(/return printHtml\(request\.html, savePathByWc\.get\(e\.sender\.id\)\)/)
  })
})
