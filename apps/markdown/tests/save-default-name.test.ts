import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

/**
 * The main-process half of model file naming: the name the renderer proposes
 * becomes the Save As dialog's default.
 *
 * The renderer half is covered in `file-naming.test.ts`; this is the seam
 * between them — the request field the renderer fills and the dialog the
 * reader sees. A name that never reached the dialog would leave the whole
 * feature a no-op, and nothing else would notice.
 */

type IpcHandler = (event: { sender: FakeWebContents }, ...args: unknown[]) => unknown

interface FakeWebContents {
  id: number
  listeners: Map<string, () => void>
  isDestroyed: ReturnType<typeof vi.fn>
  loadFile: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  once: ReturnType<typeof vi.fn>
  send: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
}

const handlers = new Map<string, IpcHandler>()
const showSaveDialogWithMemory = vi.fn()
const webContents: FakeWebContents[] = []
const temporaryDirectories: string[] = []
let nextWebContentsId = 1

function makeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  const contents: FakeWebContents = {
    id: nextWebContentsId++,
    listeners,
    isDestroyed: vi.fn(() => false),
    loadFile: vi.fn(),
    loadURL: vi.fn(),
    once: vi.fn((event: string, listener: () => void) => listeners.set(event, listener)),
    send: vi.fn(),
    setWindowOpenHandler: vi.fn(),
  }
  webContents.push(contents)
  return contents
}

vi.mock('electron', () => ({
  app: {
    getPath: vi.fn(() => tmpdir()),
    on: vi.fn(),
    quit: vi.fn(),
    whenReady: vi.fn(() => new Promise(() => {})),
  },
  BrowserWindow: class {
    static fromWebContents() {
      return null
    }
    static getFocusedWindow() {
      return null
    }
  },
  dialog: {
    showMessageBox: vi.fn(),
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => handlers.set(channel, handler)),
    on: vi.fn((channel: string, handler: IpcHandler) => handlers.set(channel, handler)),
    removeHandler: vi.fn(),
  },
  net: { fetch: vi.fn() },
  protocol: { handle: vi.fn() },
  shell: { openExternal: vi.fn() },
  WebContentsView: class {
    webContents = makeWebContents()
  },
}))

vi.mock('@genoffice/electron-utils', () => ({
  configuredDefaultSaveDir: vi.fn(() => tmpdir()),
  contextMenuLabels: vi.fn(() => ({})),
  installContextMenu: vi.fn(),
  installNavigationGuard: vi.fn(),
  rendererUrl: vi.fn(() => 'genoffice-app://markdown/index.html'),
  safeExternalUrl: vi.fn(() => null),
  showOpenDialogWithMemory: vi.fn(),
  showSaveDialogWithMemory: (...args: unknown[]) => showSaveDialogWithMemory(...args),
}))

import { createMarkdownView } from '../src/main/markdown-main'
import { MARKDOWN_CHANNELS } from '../src/shared/ipc'

/** The dialog's default path, from the last save dialog that was shown. */
function lastDefaultPath(): string {
  const call = showSaveDialogWithMemory.mock.calls.at(-1)
  expect(call, 'the save dialog should have been shown').toBeTruthy()
  const options = call?.[2] as { defaultPath?: string } | undefined
  return options?.defaultPath ?? ''
}

afterEach(async () => {
  vi.restoreAllMocks()
  showSaveDialogWithMemory.mockReset()
  showSaveDialogWithMemory.mockResolvedValue({ canceled: true })
  for (const contents of webContents.splice(0)) contents.listeners.get('destroyed')?.()
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('the name a save offers reaches the dialog', () => {
  it('prefills the dialog with the name the renderer proposed', async () => {
    const view = createMarkdownView(null)
    const contents = view.webContents as unknown as FakeWebContents
    await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      { text: '# notes', imageSources: [], mode: 'saveAs', defaultName: 'Quarterly notes' },
    )
    // The reader still chooses where the file goes: the name is a proposal in
    // the dialog, not a path the save picks on its own.
    expect(showSaveDialogWithMemory).toHaveBeenCalled()
    expect(basename(lastDefaultPath())).toBe('Quarterly notes.md')
  })

  it('leaves the dialog on its own default when nothing was proposed', async () => {
    const view = createMarkdownView(null)
    const contents = view.webContents as unknown as FakeWebContents
    await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      { text: '# notes', imageSources: [], mode: 'saveAs' },
    )
    expect(showSaveDialogWithMemory).toHaveBeenCalled()
    const path = lastDefaultPath()
    expect(path.endsWith('.md')).toBe(true)
    expect(path).not.toContain('Quarterly notes')
  })

  it('cleans a proposed name of the characters a path cannot carry', async () => {
    // The name comes from outside the app, so the same filter runs on it: a
    // stem carrying a separator would otherwise become a path somewhere else.
    const view = createMarkdownView(null)
    const contents = view.webContents as unknown as FakeWebContents
    await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      { text: '# notes', imageSources: [], mode: 'saveAs', defaultName: 'a/b:c' },
    )
    expect(basename(lastDefaultPath())).toBe('a_b_c.md')
  })

  it('writes the file where the reader said, not where the name pointed', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'markdown-naming-'))
    temporaryDirectories.push(directory)
    const view = createMarkdownView(null)
    const contents = view.webContents as unknown as FakeWebContents
    const target = join(directory, 'markdown-naming-chosen.md')
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: target })
    const result = await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      { text: '# notes', imageSources: [], mode: 'saveAs', defaultName: 'Quarterly notes' },
    )
    expect(result).toMatchObject({ ok: true, path: target })
  })
})
