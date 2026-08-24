import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

/**
 * pdf-main open-path bookkeeping: the renderer calls consume-pending on every
 * mount — including remounts caused by View > Reload (Cmd+R). The path must
 * survive re-consumption or a reload strands the tab on "No file to open".
 */

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()

interface FakeWebContents {
  id: number
  once: ReturnType<typeof vi.fn>
  on: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  loadFile: ReturnType<typeof vi.fn>
  listeners: Map<string, () => void>
}

let nextWcId = 1
let lastWebContents: FakeWebContents

function makeFakeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  const wc: FakeWebContents = {
    id: nextWcId++,
    listeners,
    once: vi.fn((event: string, handler: () => void) => {
      listeners.set(event, handler)
    }),
    on: vi.fn((event: string, handler: () => void) => {
      listeners.set(event, handler)
    }),
    setWindowOpenHandler: vi.fn(),
    loadURL: vi.fn(),
    loadFile: vi.fn(),
  }
  lastWebContents = wc
  return wc
}

vi.mock('electron', () => ({
  app: { on: vi.fn(), whenReady: vi.fn(() => new Promise(() => {})) },
  dialog: {},
  shell: {},
  BrowserWindow: class {},
  WebContentsView: class {
    webContents = makeFakeWebContents()
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    on: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    removeHandler: vi.fn(),
  },
}))

import { PDF_CHANNELS } from '../src/shared/ipc'
import { configurePdfRuntime, createPdfView, pdfIsDirty } from '../src/main/pdf-main'
import type { CreateDocumentRequest, CreateDocumentResult } from '../src/shared/ipc'

function makePdfFile(): string {
  const dir = mkdtempSync(join(tmpdir(), 'pdf-open-path-'))
  const path = join(dir, 'doc.pdf')
  writeFileSync(path, '%PDF-1.4\n%%EOF\n')
  return path
}

const consume = (wcId: number) =>
  handlers.get(PDF_CHANNELS.consumePending)?.({ sender: { id: wcId } })

describe('pdf open-path lifecycle', () => {
  it('consume-pending returns the path again after a reload (second consume)', () => {
    const path = makePdfFile()
    createPdfView(path)
    const wcId = lastWebContents.id

    expect(consume(wcId)).toBe(path)
    // View > Reload remounts the renderer, which consumes again
    expect(consume(wcId)).toBe(path)
  })

  it('drops the path when the view is destroyed', () => {
    const path = makePdfFile()
    createPdfView(path)
    const wc = lastWebContents

    expect(consume(wc.id)).toBe(path)
    wc.listeners.get('destroyed')?.()
    expect(consume(wc.id)).toBeNull()
  })

  it('clears dirty on did-start-loading so a reload does not keep the close-save prompt', () => {
    const path = makePdfFile()
    createPdfView(path)
    const wc = lastWebContents
    handlers.get(PDF_CHANNELS.dirtyChanged)?.({ sender: { id: wc.id } }, true)
    expect(pdfIsDirty(wc.id)).toBe(true)
    wc.listeners.get('did-start-loading')?.()
    expect(pdfIsDirty(wc.id)).toBe(false)
  })
})

describe('pdf create-document IPC', () => {
  const request: CreateDocumentRequest = {
    type: 'pdf',
    title: 'Summary',
    content: '<h1>Summary</h1><p>Body</p>',
  }

  it('allows a registered pathless PDF view and forwards a normalized request', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentResult> => ({
      ok: true,
      path: '/tmp/Summary.pdf',
    }))
    configurePdfRuntime({ preloadPath: '', createDocument })
    createPdfView()
    const result = await handlers.get(PDF_CHANNELS.createDocument)?.(
      { sender: { id: lastWebContents.id } },
      request,
    )

    expect(result).toEqual({ ok: true, path: '/tmp/Summary.pdf' })
    expect(createDocument).toHaveBeenCalledWith(request)
  })

  it('rejects unregistered or destroyed senders', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentResult> => ({ ok: true }))
    configurePdfRuntime({ preloadPath: '', createDocument })
    createPdfView(makePdfFile())
    const handler = handlers.get(PDF_CHANNELS.createDocument)

    await expect(handler?.({ sender: { id: 999_999 } }, request)).resolves.toMatchObject({
      ok: false,
      error: expect.stringContaining('not a registered PDF view'),
    })

    const wc = lastWebContents
    wc.listeners.get('destroyed')?.()
    await expect(handler?.({ sender: { id: wc.id } }, request)).resolves.toMatchObject({
      ok: false,
      error: expect.stringContaining('not a registered PDF view'),
    })
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('rejects malformed and oversized requests before calling the host', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentResult> => ({ ok: true }))
    configurePdfRuntime({ preloadPath: '', createDocument })
    createPdfView(makePdfFile())
    const handler = handlers.get(PDF_CHANNELS.createDocument)
    const event = { sender: { id: lastWebContents.id } }

    for (const bad of [
      { ...request, type: 'xlsx' },
      { ...request, title: ' ' },
      { ...request, content: ' ' },
      { ...request, content: 'x'.repeat(2_000_001) },
    ]) {
      await expect(handler?.(event, bad)).resolves.toMatchObject({
        ok: false,
        error: expect.stringContaining('invalid create-document request'),
      })
    }
    expect(createDocument).not.toHaveBeenCalled()
  })
})
