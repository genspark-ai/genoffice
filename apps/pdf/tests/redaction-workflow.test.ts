import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { PDF_CHANNELS, type RedactionInput, type SavePdfResult } from '../src/shared/ipc'
import { readPdfText } from '../src/main/read-text'

type Handler = (event: { sender: FakeContents }, ...args: any[]) => any
const handlers = new Map<string, Handler>()
const pick = vi.fn()
let lastContents: FakeContents
let nextId = 0
class FakeContents {
  id = ++nextId
  listeners = new Map<string, () => void>()
  onSaveAs: (path: string) => Promise<void> = async () => {}
  once = (event: string, fn: () => void) => this.listeners.set(event, fn)
  on = this.once
  setWindowOpenHandler = vi.fn()
  loadURL = vi.fn()
  loadFile = vi.fn()
  isDestroyed = () => false
  send = (channel: string, path: string) => {
    if (channel === PDF_CHANNELS.saveAsRequest) void this.onSaveAs(path)
  }
}
vi.mock('electron', () => ({
  app: { on: vi.fn(), whenReady: () => new Promise(() => {}) },
  dialog: { showSaveDialog: (...args: unknown[]) => pick(...args) },
  shell: {},
  BrowserWindow: { fromWebContents: () => null },
  WebContentsView: class {
    webContents = (lastContents = new FakeContents())
  },
  ipcMain: {
    handle: (channel: string, fn: Handler) => handlers.set(channel, fn),
    on: (channel: string, fn: Handler) => handlers.set(channel, fn),
    removeHandler: vi.fn(),
  },
}))
import { createPdfView } from '../src/main/pdf-main'

const dirs: string[] = []
afterEach(async () => {
  pick.mockReset()
  await Promise.all(dirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true })))
})
async function setup() {
  const dir = await mkdtemp(join(tmpdir(), 'redaction-workflow-'))
  dirs.push(dir)
  const source = join(dir, 'original.pdf')
  const copy = join(dir, 'redacted.pdf')
  const pdf = await PDFDocument.create()
  const page = pdf.addPage([300, 200])
  const font = await pdf.embedFont(StandardFonts.Helvetica)
  page.drawText('SECRET-ONE', { x: 20, y: 130, size: 18, font })
  page.drawText('SECRET-TWO', { x: 20, y: 60, size: 18, font })
  page.drawText('PUBLIC', { x: 20, y: 20, size: 18, font })
  const original = await pdf.save()
  await writeFile(source, original)
  createPdfView(source)
  const wc = lastContents
  const call = (channel: string, ...args: unknown[]) =>
    handlers.get(channel)!({ sender: wc }, ...args)
  let regions: RedactionInput[] = [{ pageIndex: 0, rect: [18, 125, 160, 155] }]
  let result: SavePdfResult | undefined
  wc.onSaveAs = async (targetPath) => {
    result = await call(PDF_CHANNELS.save, {
      path: call(PDF_CHANNELS.consumePending),
      targetPath,
      markups: [],
      drawings: [],
      formValues: [],
      stamps: [],
      redactions: regions,
    })
    call(PDF_CHANNELS.saveAsResult, result!.ok)
  }
  pick.mockResolvedValue({ canceled: false, filePath: copy })
  return {
    source,
    copy,
    original,
    call,
    wc,
    result: () => result,
    setRegions: (next: RedactionInput[]) => {
      regions = next
    },
  }
}

describe('redaction working document', () => {
  it('switches to the saved copy and accumulates a second redaction without another dialog', async () => {
    const s = await setup()
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.source)).toBe(true)
    expect(s.result()).toMatchObject({ ok: true })
    expect(s.call(PDF_CHANNELS.consumePending)).toBe(s.copy)
    expect(new Uint8Array(await s.call(PDF_CHANNELS.readFile, s.copy))).toEqual(
      new Uint8Array(await readFile(s.copy)),
    )
    await expect(s.call(PDF_CHANNELS.readFile, s.source)).rejects.toThrow('not granted')
    s.wc.listeners.get('did-start-loading')?.()
    expect(s.call(PDF_CHANNELS.consumePending)).toBe(s.copy)
    s.setRegions([{ pageIndex: 0, rect: [18, 55, 160, 85] }])
    pick.mockRejectedValue(new Error('A second Save As dialog must not open'))
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.copy)).toBe(true)
    const text = (await readPdfText(new Uint8Array(await readFile(s.copy)))).pages[0]!.text
    expect(text).toContain('PUBLIC')
    expect(text).not.toContain('SECRET-ONE')
    expect(text).not.toContain('SECRET-TWO')
    expect(new Uint8Array(await readFile(s.source))).toEqual(s.original)
  })

  it('keeps the original active when the dialog is cancelled or saving fails', async () => {
    const s = await setup()
    pick.mockResolvedValueOnce({ canceled: true })
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.source)).toBe(false)
    expect(s.call(PDF_CHANNELS.consumePending)).toBe(s.source)
    s.setRegions([{ pageIndex: 9, rect: [18, 55, 160, 85] }])
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.source)).toBe(false)
    expect(s.call(PDF_CHANNELS.consumePending)).toBe(s.source)
    expect(new Uint8Array(await readFile(s.source))).toEqual(s.original)
    await expect(s.call(PDF_CHANNELS.readFile, s.copy)).rejects.toThrow('not granted')
  })

  it('preserves the working copy on a failed later apply and allows retry', async () => {
    const s = await setup()
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.source)).toBe(true)
    const firstCopy = await readFile(s.copy)
    s.setRegions([{ pageIndex: 9, rect: [18, 55, 160, 85] }])
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.copy)).toBe(false)
    expect(s.call(PDF_CHANNELS.consumePending)).toBe(s.copy)
    expect(await readFile(s.copy)).toEqual(firstCopy)
    s.setRegions([{ pageIndex: 0, rect: [18, 55, 160, 85] }])
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.copy)).toBe(true)
    const text = (await readPdfText(new Uint8Array(await readFile(s.copy)))).pages[0]!.text
    expect(text).toContain('PUBLIC')
    expect(text).not.toContain('SECRET')
  })

  it('still refuses to overwrite the original when first applying redactions', async () => {
    const s = await setup()
    pick.mockResolvedValue({ canceled: false, filePath: s.source })
    expect(await s.call(PDF_CHANNELS.requestRedactionCopy, s.source)).toBe(false)
    expect(new Uint8Array(await readFile(s.source))).toEqual(s.original)
  })
})
