import { existsSync, mkdtempSync, readFileSync, rmSync, truncateSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument } from 'pdf-lib'
import { afterEach, describe, expect, it, vi } from 'vitest'

/**
 * mergePdf size cap: the selection is stat'ed and summed before anything is
 * read (an oversized merge is refused without reading it first), and the
 * picked files are then read one at a time — Promise.all used to hold every
 * picked file in memory at once on the main process, where an OOM takes the
 * whole app down, not one tab.
 */

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()

interface FakeWebContents {
  id: number
  once: ReturnType<typeof vi.fn>
  on: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
}

let nextWcId = 1
let lastWebContents: FakeWebContents

function makeFakeWebContents(): FakeWebContents {
  const webContents: FakeWebContents = {
    id: nextWcId++,
    once: vi.fn(),
    on: vi.fn(),
    setWindowOpenHandler: vi.fn(),
    loadURL: vi.fn(),
  }
  lastWebContents = webContents
  return webContents
}

const dialogMock = vi.hoisted(() => ({ showOpenDialog: vi.fn() }))

vi.mock('electron', () => ({
  app: {
    on: vi.fn(),
    whenReady: vi.fn(() => new Promise(() => {})),
    getPath: vi.fn((name: string) => join(tmpdir(), `pdf-merge-test-${name}`)),
  },
  dialog: dialogMock,
  shell: { showItemInFolder: vi.fn(), openExternal: vi.fn() },
  BrowserWindow: class {
    static fromWebContents = vi.fn(() => null)
    static getFocusedWindow = vi.fn(() => null)
  },
  WebContentsView: class {
    webContents = makeFakeWebContents()
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    on: vi.fn(),
    removeHandler: vi.fn(),
  },
}))

import { MAX_MERGE_TOTAL_BYTES, readMergeInputs } from '../src/main/pdf-main'
import type { MergeInputIo } from '../src/main/pdf-main'
import { createPdfView } from '../src/main/pdf-main'
import { PDF_CHANNELS } from '../src/shared/ipc'
import type { MergePdfRequest, MergePdfResult } from '../src/shared/ipc'

/** 1.1 GiB — over the cap, and far enough past it that MB rounding can't hide it */
const OVER_CAP_BYTES = 1_181_076_280
const SIX_HUNDRED_MIB = 629_145_600

let dir: string | null = null

/** Real one-page PDF (pdf-lib, same as the other page-operation tests) */
async function writePdf(path: string): Promise<void> {
  const doc = await PDFDocument.create()
  doc.addPage([100, 200])
  writeFileSync(path, await doc.save({ useObjectStreams: false }))
}

/** Small file whose stat size reports `bytes` — sparse, so no real disk is used */
function writeSparsePdf(path: string, bytes: number): void {
  writeFileSync(path, '%PDF-1.4\n%%EOF\n')
  truncateSync(path, bytes)
}

const invokeMerge = (request: MergePdfRequest): Promise<MergePdfResult> =>
  handlers.get(PDF_CHANNELS.mergePdf)!(
    { sender: { id: lastWebContents.id } },
    request,
  ) as Promise<MergePdfResult>

/** In-memory io that records every read in call order and can gate one read */
function makeFakeIo(sizes: Record<string, number>): MergeInputIo & {
  readCalls: string[]
  gate(path: string): Promise<void>
  openGate(): void
} {
  const readCalls: string[] = []
  let blockedPath = ''
  let onReadStart: (() => void) | null = null
  let openGate: () => void = () => {}
  return {
    readCalls,
    // Resolves when the read of `path` starts; the read itself stays in
    // flight until openGate() so the test can inspect the in-flight state.
    gate(path: string) {
      blockedPath = path
      return new Promise<void>((resolve) => {
        onReadStart = resolve
      })
    },
    openGate: () => openGate(),
    stat: async (path: string) => ({ size: sizes[path] ?? 0 }),
    readFile: async (path: string) => {
      readCalls.push(path)
      if (path === blockedPath) {
        onReadStart?.()
        await new Promise<void>((resolve) => {
          openGate = resolve
        })
      }
      return Buffer.from(`bytes-of-${path}`)
    },
  }
}

afterEach(() => {
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = null
  dialogMock.showOpenDialog.mockReset()
})

describe('readMergeInputs', () => {
  const base = '/tmp/base.pdf'
  const first = '/tmp/first.pdf'
  const second = '/tmp/second.pdf'

  it('refuses an over-cap selection before reading anything', async () => {
    const io = makeFakeIo({ [base]: 100, [first]: OVER_CAP_BYTES })

    await expect(readMergeInputs(base, [first], io)).rejects.toThrow(
      `pdf: merge too large — the selected files total ${Math.round((100 + OVER_CAP_BYTES) / (1024 * 1024))} MB, ` +
        `over the ${Math.round(MAX_MERGE_TOTAL_BYTES / (1024 * 1024))} MB merge limit`,
    )
    // The refusal must not cost a single read — stat only.
    expect(io.readCalls).toEqual([])
  })

  it('refuses when individually-fine files sum over the cap', async () => {
    const io = makeFakeIo({ [base]: 100, [first]: SIX_HUNDRED_MIB, [second]: SIX_HUNDRED_MIB })

    // 600 + 600 MiB plus the base crosses the 1024 MiB cap only as a sum
    await expect(readMergeInputs(base, [first, second], io)).rejects.toThrow(
      /over the 1024 MB merge limit/,
    )
    expect(io.readCalls).toEqual([])
  })

  it('reads the picked files one at a time, then the base document', async () => {
    const io = makeFakeIo({ [base]: 100, [first]: 200, [second]: 300 })
    const firstReadStarted = io.gate(first)

    const pending = readMergeInputs(base, [first, second], io)
    await firstReadStarted
    // While the first picked read is in flight, nothing else may have started
    expect(io.readCalls).toEqual([first])

    io.openGate()
    const { base: baseBytes, others } = await pending

    expect(io.readCalls).toEqual([first, second, base])
    expect(others.map((bytes) => Buffer.from(bytes).toString())).toEqual([
      `bytes-of-${first}`,
      `bytes-of-${second}`,
    ])
    expect(Buffer.from(baseBytes).toString()).toBe(`bytes-of-${base}`)
  })
})

describe('mergePdf handler', () => {
  it('refuses an over-cap dialog pick through the real stat path', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-merge-cap-'))
    const base = join(dir, 'base.pdf')
    const huge = join(dir, 'huge.pdf')
    await writePdf(base)
    writeSparsePdf(huge, OVER_CAP_BYTES)
    createPdfView(base)
    dialogMock.showOpenDialog.mockResolvedValue({ canceled: false, filePaths: [huge] })

    const result = await invokeMerge({ path: base, suggestedName: 'merged.pdf' })

    expect(result.ok).toBe(false)
    if (!result.ok)
      expect(result.error).toMatch(/pdf: merge too large — .* over the 1024 MB merge limit/)
  })

  it('merges an under-cap selection and writes the output', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-merge-ok-'))
    const base = join(dir, 'base.pdf')
    const first = join(dir, 'first.pdf')
    const second = join(dir, 'second.pdf')
    await writePdf(base)
    await writePdf(first)
    await writePdf(second)
    createPdfView(base)
    dialogMock.showOpenDialog.mockResolvedValue({
      canceled: false,
      filePaths: [first, second],
    })

    const result = await invokeMerge({ path: base, suggestedName: 'merged.pdf' })

    expect(result).toMatchObject({ ok: true, appendedCount: 2 })
    if (result.ok && 'savedPath' in result) {
      expect(existsSync(result.savedPath)).toBe(true)
      expect(readFileSync(result.savedPath).subarray(0, 5).toString()).toBe('%PDF-')
    }
  })
})
