import {
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

/**
 * pdf:auto-rename — content-derived naming for shell-created blank PDFs
 * (pdf's analog of sheets' autoRenameWorkbook). Renames only paths that are
 * both granted to the calling view AND still carrying the shell's untitled
 * name; user-chosen names must never move.
 */

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()

interface FakeWebContents {
  id: number
  once: ReturnType<typeof vi.fn>
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

import type { PdfAutoRenameResult, SavePdfResult } from '../src/shared/ipc'
import { PDF_CHANNELS } from '../src/shared/ipc'
import {
  createPdfView,
  markPdfUntitledPath,
  movePdfFileNoClobber,
  setPdfRenamedHook,
} from '../src/main/pdf-main'

const tempDirs = new Set<string>()

function makePdfFile(name = 'Untitled PDF.pdf'): string {
  const dir = mkdtempSync(join(tmpdir(), 'pdf-auto-rename-'))
  tempDirs.add(dir)
  const path = join(dir, name)
  writeFileSync(path, '%PDF-1.4\n%%EOF\n')
  return path
}

const rename = (wcId: number, path: string, base: unknown) =>
  handlers.get(PDF_CHANNELS.autoRename)?.(
    { sender: { id: wcId } },
    path,
    base,
  ) as PdfAutoRenameResult

const isUntitled = (wcId: number, path: string) =>
  handlers.get(PDF_CHANNELS.isUntitled)?.({ sender: { id: wcId } }, path) as boolean

const readGranted = (wcId: number, path: string) =>
  handlers.get(PDF_CHANNELS.readFile)?.({ sender: { id: wcId } }, path) as Promise<ArrayBuffer>

const saveGranted = (wcId: number, path: string) =>
  handlers.get(PDF_CHANNELS.save)?.({ sender: { id: wcId } }, { path }) as Promise<SavePdfResult>

afterEach(() => {
  vi.restoreAllMocks()
  setPdfRenamedHook(() => {})
  for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true })
  tempDirs.clear()
})

describe('pdf auto-rename', () => {
  it('renames an untitled-marked file, fires the shell hook, and consumes the flag', () => {
    const path = makePdfFile()
    markPdfUntitledPath(path)
    createPdfView(path)
    const wcId = lastWebContents.id
    const hook = vi.fn()
    setPdfRenamedHook(hook)

    // the untitled flag is queryable (gates the renderer's after-AI-run silent save)
    expect(isUntitled(wcId, path)).toBe(true)

    const result = rename(wcId, path, 'Rental Agreement')
    expect(result.renamed).toBe(true)
    expect(basename(result.path!)).toBe('Rental Agreement.pdf')
    expect(existsSync(result.path!)).toBe(true)
    expect(existsSync(path)).toBe(false)
    expect(hook).toHaveBeenCalledWith(expect.objectContaining({ id: wcId }), path, result.path)

    // the untitled flag is consumed: a second proposal must not move the file again
    expect(isUntitled(wcId, path)).toBe(false)
    expect(rename(wcId, result.path!, 'Other Name').renamed).toBe(false)
    // the view keeps working on the new path (readFile grant follows the rename)
    expect(handlers.get(PDF_CHANNELS.consumePending)?.({ sender: { id: wcId } })).toBe(result.path)
  })

  it('sanitizes illegal filename characters and caps the length', () => {
    const path = makePdfFile()
    markPdfUntitledPath(path)
    createPdfView(path)

    const result = rename(lastWebContents.id, path, '  Q3: "Plan" <draft>?  ')
    expect(result.renamed).toBe(true)
    expect(basename(result.path!)).toBe('Q3 Plan draft.pdf')
  })

  it('retries an occupied candidate without altering the winning file bytes', () => {
    const path = makePdfFile()
    const occupiedPath = join(path, '..', 'Report.pdf')
    const sentinel = Buffer.from([0, 255, 17, 99, 42])
    writeFileSync(occupiedPath, sentinel)
    markPdfUntitledPath(path)
    createPdfView(path)

    const result = rename(lastWebContents.id, path, 'Report')
    expect(result.renamed).toBe(true)
    expect(basename(result.path!)).toBe('Report-2.pdf')
    expect(readFileSync(occupiedPath)).toEqual(sentinel)
  })

  it('revokes the old path even when it is recreated and keeps the new path reloadable', async () => {
    const path = makePdfFile()
    markPdfUntitledPath(path)
    createPdfView(path)
    const wcId = lastWebContents.id

    const result = rename(wcId, path, 'Reload Target')
    expect(result.renamed).toBe(true)
    const target = result.path!
    const recreated = Buffer.from('not the renamed document')
    writeFileSync(path, recreated)

    await expect(readGranted(wcId, path)).rejects.toThrow('path not granted')
    await expect(saveGranted(wcId, path)).resolves.toEqual({
      ok: false,
      error: 'pdf: path not granted to this view',
    })
    expect(rename(wcId, path, 'Hijacked').renamed).toBe(false)
    expect(readFileSync(path)).toEqual(recreated)

    expect(handlers.get(PDF_CHANNELS.consumePending)?.({ sender: { id: wcId } })).toBe(target)
    const bytes = Buffer.from(new Uint8Array(await readGranted(wcId, target)))
    expect(bytes.toString()).toBe('%PDF-1.4\n%%EOF\n')
  })

  it('falls back to an exclusive copy when hard links are unavailable', () => {
    const source = makePdfFile()
    const target = join(source, '..', 'Copy Fallback.pdf')
    const original = readFileSync(source)

    const result = movePdfFileNoClobber(source, target, {
      link: () => {
        throw Object.assign(new Error('hard links unavailable'), { code: 'EPERM' })
      },
    })

    expect(result).toBe('moved')
    expect(existsSync(source)).toBe(false)
    expect(readFileSync(target)).toEqual(original)
  })

  it('rolls back the reserved target when removing the source fails', () => {
    const source = makePdfFile()
    const target = join(source, '..', 'Rollback Target.pdf')
    const original = readFileSync(source)

    const result = movePdfFileNoClobber(source, target, {
      unlink: (path) => {
        if (path === source) {
          throw Object.assign(new Error('simulated unlink failure'), { code: 'EACCES' })
        }
        unlinkSync(path)
      },
    })

    expect(result).toBe('failed')
    expect(readFileSync(source)).toEqual(original)
    expect(existsSync(target)).toBe(false)
  })

  it('restores the source and fails if the reserved target is replaced mid-move', () => {
    const source = makePdfFile()
    const target = join(source, '..', 'Raced Target.pdf')
    const original = readFileSync(source)
    const replacement = Buffer.from('concurrent replacement')
    let targetIdentityReads = 0

    const result = movePdfFileNoClobber(source, target, {
      identity: (path) => {
        if (path === target && ++targetIdentityReads === 2) {
          unlinkSync(target)
          writeFileSync(target, replacement)
        }
        const stats = statSync(path, { bigint: true })
        return `${stats.dev}:${stats.ino}`
      },
    })

    expect(result).toBe('failed')
    expect(readFileSync(source)).toEqual(original)
    expect(readFileSync(target)).toEqual(replacement)
  })

  it('keeps the old grant and untitled state after a pre-move failure', async () => {
    const path = makePdfFile()
    const bytes = readFileSync(path)
    markPdfUntitledPath(path)
    createPdfView(path)
    const wcId = lastWebContents.id
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})

    unlinkSync(path)
    expect(rename(wcId, path, 'Will Retry').renamed).toBe(false)
    writeFileSync(path, bytes)

    expect(isUntitled(wcId, path)).toBe(true)
    expect(Buffer.from(new Uint8Array(await readGranted(wcId, path)))).toEqual(bytes)
    expect(rename(wcId, path, 'Will Retry').renamed).toBe(true)
    warn.mockRestore()
  })

  it('commits renderer bookkeeping even when the shell rename hook throws', async () => {
    const path = makePdfFile()
    markPdfUntitledPath(path)
    createPdfView(path)
    const wcId = lastWebContents.id
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    setPdfRenamedHook(() => {
      throw new Error('hook failed')
    })

    const result = rename(wcId, path, 'Hook Safe')
    expect(result.renamed).toBe(true)
    expect(handlers.get(PDF_CHANNELS.consumePending)?.({ sender: { id: wcId } })).toBe(result.path)
    expect(
      Buffer.from(new Uint8Array(await readGranted(wcId, result.path!))).byteLength,
    ).toBeGreaterThan(0)
    expect(warn).toHaveBeenCalledWith('[pdf] auto-rename hook failed:', expect.any(Error))
    warn.mockRestore()
  })

  it('never touches files that are not marked untitled or not granted to the view', () => {
    const opened = makePdfFile('user-named.pdf')
    createPdfView(opened)
    const wcId = lastWebContents.id
    // granted but not untitled (a regular file the user opened)
    expect(isUntitled(wcId, opened)).toBe(false)
    expect(rename(wcId, opened, 'New Name').renamed).toBe(false)
    expect(existsSync(opened)).toBe(true)

    // untitled but not granted to this view
    const foreign = makePdfFile()
    markPdfUntitledPath(foreign)
    expect(rename(wcId, foreign, 'New Name').renamed).toBe(false)
    expect(existsSync(foreign)).toBe(true)
  })

  it('rejects proposals that sanitize to nothing', () => {
    const path = makePdfFile()
    markPdfUntitledPath(path)
    createPdfView(path)
    expect(rename(lastWebContents.id, path, '  ...  ').renamed).toBe(false)
    expect(rename(lastWebContents.id, path, 42).renamed).toBe(false)
    expect(existsSync(path)).toBe(true)
  })
})
