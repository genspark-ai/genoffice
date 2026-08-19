import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument } from 'pdf-lib'
import { afterEach, describe, expect, it, vi } from 'vitest'

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()

interface FakeWebContents {
  id: number
  once: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
  listeners: Map<string, () => void>
}

let nextWcId = 1
let lastWebContents: FakeWebContents
const operationHook = vi.hoisted(() => ({
  beforeAtomicWrite: null as (() => void) | null,
}))

function makeFakeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  const webContents = {
    id: nextWcId++,
    listeners,
    once: vi.fn((event: string, handler: () => void) => {
      listeners.set(event, handler)
    }),
    setWindowOpenHandler: vi.fn(),
  }
  lastWebContents = webContents
  return webContents
}

vi.mock('../src/main/save-pdf', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/main/save-pdf')>()
  return {
    ...actual,
    insertBlankPageBytes: async (
      ...args: Parameters<typeof actual.insertBlankPageBytes>
    ): Promise<Uint8Array> => {
      const bytes = await actual.insertBlankPageBytes(...args)
      operationHook.beforeAtomicWrite?.()
      return bytes
    },
  }
})

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

import { createPdfView } from '../src/main/pdf-main'
import { PDF_CHANNELS } from '../src/shared/ipc'
import type { InsertBlankPageResult } from '../src/shared/ipc'

let dir: string | null = null

afterEach(() => {
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = null
  operationHook.beforeAtomicWrite = null
})

describe('in-place PDF page operation writes', () => {
  it('removes the .gensave temp when an insert-page rename fails', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-page-write-'))
    const path = join(dir, 'doc.pdf')
    const doc = await PDFDocument.create()
    doc.addPage([100, 200])
    writeFileSync(path, await doc.save({ useObjectStreams: false }))
    const original = readFileSync(path)
    createPdfView(path)

    // Simulate another process replacing the target with a directory after the
    // operation reads it, so the helper's final file-over-directory rename fails.
    const preserved = join(dir, 'preserved.pdf')
    operationHook.beforeAtomicWrite = () => {
      renameSync(path, preserved)
      mkdirSync(path)
      writeFileSync(join(path, 'occupied'), 'x')
    }
    const result = (await handlers.get(PDF_CHANNELS.insertBlankPage)!(
      { sender: { id: lastWebContents.id } },
      { path, afterPageIndex: 0 },
    )) as InsertBlankPageResult

    expect(result.ok).toBe(false)
    expect(readFileSync(preserved)).toEqual(original)
    expect(readdirSync(dir).filter((name) => name.includes('.gensave-'))).toEqual([])
  })
})
