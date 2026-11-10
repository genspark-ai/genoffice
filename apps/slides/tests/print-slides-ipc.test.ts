/**
 * The print handler must read its silent-job options from `PrintSlidesOp.job`:
 * the dialog nests the printer choice there, so `printJobOptions(op)` saw no
 * `deviceName`, reported "system dialog" for every job and the picker, copies,
 * duplex and driver options never took effect (genoffice#2378 review).
 */
import { tmpdir } from 'node:os'
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

const mocks = vi.hoisted(() => ({
  handlers: new Map<string, (...args: never[]) => unknown>(),
  windows: [] as Array<Record<string, any>>,
  printSlidesHtml: vi.fn(async (_html: string, _win: unknown) => ({ ok: true })),
  printPdfViaCups: vi.fn(async (_pdf: Buffer, _job: Record<string, unknown>) => ({ ok: true })),
}))

vi.mock('electron', () => {
  class FakeBrowserWindow {
    readonly webContents = {
      executeJavaScript: async (): Promise<void> => {},
      print: vi.fn(
        (options: Record<string, unknown>, callback: (success: boolean, reason: string) => void) =>
          callback(true, ''),
      ),
      printToPDF: vi.fn(async () => Buffer.from('%PDF-1.4 print')),
    }
    loadedPaths: string[] = []
    destroyed = false
    constructor(_options: Record<string, unknown>) {
      mocks.windows.push(this as never)
    }
    async loadFile(path: string): Promise<void> {
      this.loadedPaths.push(path)
    }
    show(): void {}
    focus(): void {}
    isDestroyed(): boolean {
      return this.destroyed
    }
    destroy(): void {
      this.destroyed = true
    }
    static fromWebContents(): null {
      return null
    }
    static getFocusedWindow(): null {
      return null
    }
  }
  return {
    app: {
      getPath: () => tmpdir(),
      getName: () => 'slides-test',
      getVersion: () => '0',
      whenReady: () => Promise.resolve(),
    },
    dialog: {},
    clipboard: { writeBuffer: () => {} },
    ipcMain: {
      handle: (channel: string, fn: (...args: never[]) => unknown) =>
        mocks.handlers.set(channel, fn),
      on: () => {},
      removeHandler: () => {},
    },
    BrowserWindow: FakeBrowserWindow,
    webContents: { getAllWebContents: () => [], fromId: () => null },
    Menu: { buildFromTemplate: () => ({}), setApplicationMenu: () => {} },
    session: {
      defaultSession: { protocol: { handle: () => {} }, setDisplayMediaRequestHandler: () => {} },
    },
    nativeImage: {},
    shell: { openPath: () => Promise.resolve('') },
    desktopCapturer: { getSources: async () => [] },
    WebContentsView: class {},
  }
})

// The shaping chain loads a harfbuzz wasm that cannot resolve in the unit-test env;
// the print handler never measures text.
vi.mock('../src/main/shaped-metrics', () => ({
  shapedMetricsReady: async () => {},
  refineComplexWidths: async () => {},
  complexScriptOf: () => null,
  initShapedMetrics: () => {},
  gtMeasure: (t: string) => ({ width: t.length * 8 }),
  shapedFamily: () => null,
  shapedMeasure: () => ({ width: 0 }),
}))

vi.mock('../src/main/print-window', () => ({ printSlidesHtml: mocks.printSlidesHtml }))
// macOS silent jobs spool through lp; the handler-level test only asserts the
// route and the resolved job options, never touches CUPS.
vi.mock('@genoffice/electron-utils/print-cups', () => ({
  printPdfViaCups: mocks.printPdfViaCups,
  printToPdfOptionsFor: (request: Record<string, unknown>) => request,
}))

import { registerSlidesIpc } from '../src/main/slides-main'

function call(channel: string, ...args: unknown[]): unknown {
  return (mocks.handlers.get(channel) as (...args: unknown[]) => unknown)(
    { sender: { id: 1 } },
    ...args,
  )
}

const OP = {
  pngsBase64: [Buffer.from('slide-png').toString('base64')],
  widthPx: 1280,
  heightPx: 720,
}

beforeAll(() => registerSlidesIpc())
beforeEach(() => {
  mocks.windows.length = 0
  mocks.printSlidesHtml.mockClear()
  mocks.printPdfViaCups.mockClear()
})

describe('slides:print job options', () => {
  it('prints silently through the printer nested under op.job', async () => {
    const result = (await call('slides:print', {
      ...OP,
      job: {
        deviceName: 'Office-Printer',
        copies: 2,
        collate: false,
        duplex: 'longEdge',
        driverOptions: [{ name: 'MediaType', value: 'stationery' }],
      },
    })) as { ok: boolean }

    expect(result).toEqual({ ok: true })
    // a picked printer must not fall through to the system-dialog route
    expect(mocks.printSlidesHtml).not.toHaveBeenCalled()
    const win = mocks.windows[0]!
    if (process.platform === 'darwin') {
      // macOS routes the silent job through CUPS with printToPDF output
      expect(win.webContents.printToPDF).toHaveBeenCalledTimes(1)
      expect(win.webContents.print).not.toHaveBeenCalled()
      expect(mocks.printPdfViaCups).toHaveBeenCalledTimes(1)
      const [bytes, options] = mocks.printPdfViaCups.mock.calls[0]!
      expect(Buffer.isBuffer(bytes)).toBe(true)
      expect(options).toMatchObject({
        silent: true,
        deviceName: 'Office-Printer',
        copies: 2,
        collate: false,
        duplexMode: 'longEdge',
        driverOptions: [{ name: 'MediaType', value: 'stationery' }],
      })
    } else {
      expect(win.webContents.print).toHaveBeenCalledTimes(1)
      expect(win.webContents.printToPDF).not.toHaveBeenCalled()
      expect(mocks.printPdfViaCups).not.toHaveBeenCalled()
      const [options] = win.webContents.print.mock.calls[0] as [Record<string, unknown>]
      expect(options).toMatchObject({
        silent: true,
        deviceName: 'Office-Printer',
        copies: 2,
        collate: false,
        duplexMode: 'longEdge',
        printBackground: true,
      })
    }
  })

  it('treats the OS-default pick (deviceName "") as a silent job too', async () => {
    const result = (await call('slides:print', { ...OP, job: { deviceName: '' } })) as {
      ok: boolean
    }
    expect(result).toEqual({ ok: true })
    expect(mocks.printSlidesHtml).not.toHaveBeenCalled()
    if (process.platform === 'darwin') {
      expect(mocks.printPdfViaCups).toHaveBeenCalledTimes(1)
      expect(mocks.printPdfViaCups.mock.calls[0]![1]).toMatchObject({ silent: true })
    } else {
      const [options] = mocks.windows[0]!.webContents.print.mock.calls[0] as [
        Record<string, unknown>,
      ]
      expect(options).toMatchObject({ silent: true })
    }
  })

  it('keeps the system dialog when no job was sent', async () => {
    const result = (await call('slides:print', OP)) as { ok: boolean }
    expect(result).toEqual({ ok: true })
    expect(mocks.printSlidesHtml).toHaveBeenCalledTimes(1)
    const win = mocks.windows[0]!
    const [html, printedWindow] = mocks.printSlidesHtml.mock.calls[0]!
    expect(html).toEqual(expect.any(String))
    expect(printedWindow).toBe(win)
    expect(win.webContents.print).not.toHaveBeenCalled()
    expect(win.webContents.printToPDF).not.toHaveBeenCalled()
    expect(mocks.printPdfViaCups).not.toHaveBeenCalled()
  })

  it('hands the handout layout and thumbnail order to the page assembly', async () => {
    const result = (await call('slides:print', {
      ...OP,
      layout: 'handout6',
      orientation: 'portrait',
      order: 'vertical',
    })) as { ok: boolean }
    expect(result).toEqual({ ok: true })
    const [html] = mocks.printSlidesHtml.mock.calls[0]! as unknown as [string]
    expect(html).toContain('grid-auto-flow: column;')
    expect(html).toContain('@page { size: 8.27in 11.69in; margin: 0; }')
  })
})
