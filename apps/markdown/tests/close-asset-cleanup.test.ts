import { existsSync } from 'node:fs'
import { chmod, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

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
const showMessageBox = vi.fn()
const showSaveDialogWithMemory = vi.fn()
const webContents: FakeWebContents[] = []
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
    showMessageBox: (...args: unknown[]) => showMessageBox(...args),
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
  safeExternalUrl: vi.fn(() => null),
  showOpenDialogWithMemory: vi.fn(),
  showSaveDialogWithMemory: (...args: unknown[]) => showSaveDialogWithMemory(...args),
}))

import {
  readOwnedAssetManifest,
  reconcileOwnedAssets,
  writeImageIntoOwnedAssets,
} from '../src/main/asset-lifecycle'
import { createMarkdownView, requestMarkdownClose } from '../src/main/markdown-main'
import { MARKDOWN_CHANNELS } from '../src/shared/ipc'

const temporaryDirectories: string[] = []

async function createDocument(): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'markdown-close-assets-'))
  temporaryDirectories.push(directory)
  const documentPath = join(directory, 'note.md')
  await writeFile(documentPath, '# Saved document')
  return documentPath
}

function markDirty(contents: FakeWebContents): void {
  handlers.get(MARKDOWN_CHANNELS.dirtyChanged)?.({ sender: contents }, true)
}

afterEach(async () => {
  vi.restoreAllMocks()
  showMessageBox.mockReset()
  showSaveDialogWithMemory.mockReset()
  for (const contents of webContents.splice(0)) contents.listeners.get('destroyed')?.()
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('requestMarkdownClose asset cleanup', () => {
  it("deletes pending owned assets on Don't Save while preserving saved assets", async () => {
    const documentPath = await createDocument()
    const saved = await writeImageIntoOwnedAssets(
      documentPath,
      'saved.png',
      Buffer.from('saved image'),
    )
    await reconcileOwnedAssets(documentPath, [saved])
    const pending = await writeImageIntoOwnedAssets(
      documentPath,
      'pending.png',
      Buffer.from('pending image'),
    )
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    markDirty(contents)
    showMessageBox.mockResolvedValue({ response: 1 })

    await expect(requestMarkdownClose(contents as never)).resolves.toBe(true)

    expect(existsSync(join(dirname(documentPath), saved))).toBe(true)
    expect(existsSync(join(dirname(documentPath), pending))).toBe(false)
    expect(await readOwnedAssetManifest(documentPath)).toMatchObject({
      files: [{ name: 'saved.png' }],
      pending: {},
    })
  })

  it('preserves pending owned assets when close is canceled', async () => {
    const documentPath = await createDocument()
    const pending = await writeImageIntoOwnedAssets(
      documentPath,
      'pending.png',
      Buffer.from('pending image'),
    )
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    markDirty(contents)
    showMessageBox.mockResolvedValue({ response: 2 })

    await expect(requestMarkdownClose(contents as never)).resolves.toBe(false)

    expect(existsSync(join(dirname(documentPath), pending))).toBe(true)
    expect(await readOwnedAssetManifest(documentPath)).toMatchObject({
      files: [{ name: 'pending.png' }],
      pending: { 'note.md': ['pending.png'] },
    })
  })

  it('preserves referenced assets when the user successfully saves before closing', async () => {
    const documentPath = await createDocument()
    const pending = await writeImageIntoOwnedAssets(
      documentPath,
      'pending.png',
      Buffer.from('pending image'),
    )
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    markDirty(contents)
    showMessageBox.mockResolvedValue({ response: 0 })

    const closing = requestMarkdownClose(contents as never)
    await vi.waitFor(() =>
      expect(contents.send).toHaveBeenCalledWith(MARKDOWN_CHANNELS.closeSaveRequest),
    )
    const saveResult = await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      {
        text: `![pending](${pending})`,
        imageSources: [pending],
        mode: 'save',
      },
    )
    handlers.get(MARKDOWN_CHANNELS.closeSaveResult)?.({ sender: contents }, true)

    await expect(closing).resolves.toBe(true)
    expect(saveResult).toMatchObject({ ok: true, path: documentPath })
    expect(existsSync(join(dirname(documentPath), pending))).toBe(true)
    expect(await readOwnedAssetManifest(documentPath)).toMatchObject({
      files: [{ name: 'pending.png' }],
      documents: { 'note.md': ['pending.png'] },
      pending: {},
    })
  })
})

describe('Save As source pending ownership', () => {
  it('clears the old document key for same-directory Save As without deleting new references', async () => {
    const sourcePath = await createDocument()
    const saved = await writeImageIntoOwnedAssets(
      sourcePath,
      'saved.png',
      Buffer.from('saved image'),
    )
    await reconcileOwnedAssets(sourcePath, [saved])
    await writeFile(sourcePath, `![saved](${saved})`)
    const pending = await writeImageIntoOwnedAssets(
      sourcePath,
      'pending.png',
      Buffer.from('pending image'),
    )
    const targetPath = join(dirname(sourcePath), 'copy.md')
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })
    const view = createMarkdownView(sourcePath)
    const contents = view.webContents as unknown as FakeWebContents

    const saveResult = await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      {
        text: `![saved](${saved})\n![pending](${pending})`,
        imageSources: [saved, pending],
        mode: 'saveAs',
      },
    )

    expect(saveResult).toMatchObject({ ok: true, path: targetPath })
    expect(existsSync(join(dirname(sourcePath), pending))).toBe(true)
    expect(await readFile(sourcePath, 'utf8')).toBe(`![saved](${saved})`)
    expect(await readOwnedAssetManifest(targetPath)).toMatchObject({
      files: [{ name: 'pending.png' }, { name: 'saved.png' }],
      documents: {
        'copy.md': ['pending.png', 'saved.png'],
        'note.md': ['saved.png'],
      },
      pending: {},
    })
  })

  it('removes cross-directory source orphans while preserving old saved references', async () => {
    const sourcePath = await createDocument()
    const saved = await writeImageIntoOwnedAssets(
      sourcePath,
      'saved.png',
      Buffer.from('saved image'),
    )
    await reconcileOwnedAssets(sourcePath, [saved])
    await writeFile(sourcePath, `![saved](${saved})`)
    const pending = await writeImageIntoOwnedAssets(
      sourcePath,
      'pending.png',
      Buffer.from('pending image'),
    )
    const targetDirectory = await mkdtemp(join(tmpdir(), 'markdown-save-as-target-'))
    temporaryDirectories.push(targetDirectory)
    const targetPath = join(targetDirectory, 'copy.md')
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })
    const view = createMarkdownView(sourcePath)
    const contents = view.webContents as unknown as FakeWebContents

    const saveResult = await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      {
        text: `![saved](${saved})\n![pending](${pending})`,
        imageSources: [saved, pending],
        mode: 'saveAs',
      },
    )

    expect(saveResult).toMatchObject({
      ok: true,
      path: targetPath,
      imageRewrites: [
        { from: saved, to: 'assets/saved.png' },
        { from: pending, to: 'assets/pending.png' },
      ],
    })
    expect(existsSync(join(dirname(sourcePath), pending))).toBe(false)
    expect(await readFile(join(dirname(sourcePath), saved), 'utf8')).toBe('saved image')
    expect(await readFile(sourcePath, 'utf8')).toBe(`![saved](${saved})`)
    expect(await readOwnedAssetManifest(sourcePath)).toMatchObject({
      files: [{ name: 'saved.png' }],
      documents: { 'note.md': ['saved.png'] },
      pending: {},
    })
    expect(await readOwnedAssetManifest(targetPath)).toMatchObject({
      files: [{ name: 'pending.png' }, { name: 'saved.png' }],
      documents: { 'copy.md': ['pending.png', 'saved.png'] },
      pending: {},
    })
  })

  it.skipIf(process.platform === 'win32')(
    'resolves source pending ownership even when destination reconciliation reports an error',
    async () => {
      const sourcePath = await createDocument()
      const keep = await writeImageIntoOwnedAssets(
        sourcePath,
        'keep.png',
        Buffer.from('keep image'),
      )
      const orphan = await writeImageIntoOwnedAssets(
        sourcePath,
        'orphan.png',
        Buffer.from('orphan image'),
      )
      // Simulate a source file that already references one still-pending asset.
      await writeFile(sourcePath, `![keep](${keep})`)

      const targetDirectory = await mkdtemp(join(tmpdir(), 'markdown-save-as-error-target-'))
      temporaryDirectories.push(targetDirectory)
      const targetPath = join(targetDirectory, 'copy.md')
      const blocked = await writeImageIntoOwnedAssets(
        targetPath,
        'blocked.png',
        Buffer.from('unrelated target asset'),
      )
      const blockedPath = join(targetDirectory, blocked)
      await chmod(blockedPath, 0)
      showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })
      const warning = vi.spyOn(console, 'warn').mockImplementation(() => {})
      const view = createMarkdownView(sourcePath)
      const contents = view.webContents as unknown as FakeWebContents

      let saveResult: unknown
      try {
        saveResult = await handlers.get(MARKDOWN_CHANNELS.save)?.(
          { sender: contents },
          {
            text: `![keep](${keep})\n![orphan](${orphan})`,
            imageSources: [keep, orphan],
            mode: 'saveAs',
          },
        )
      } finally {
        await chmod(blockedPath, 0o600)
      }

      expect(saveResult).toMatchObject({ ok: true, path: targetPath })
      expect(warning).toHaveBeenCalledWith(
        '[markdown] asset reconciliation incomplete:',
        expect.arrayContaining([expect.stringContaining('blocked.png')]),
      )
      expect(existsSync(join(dirname(sourcePath), keep))).toBe(true)
      expect(existsSync(join(dirname(sourcePath), orphan))).toBe(false)
      expect(await readOwnedAssetManifest(sourcePath)).toMatchObject({
        files: [{ name: 'keep.png' }],
        pending: {},
      })
      expect(existsSync(join(targetDirectory, 'assets', 'keep.png'))).toBe(true)
      expect(existsSync(join(targetDirectory, 'assets', 'orphan.png'))).toBe(true)
    },
  )

  it('preserves source pending ownership and rolls back copied assets when Save As fails', async () => {
    const sourcePath = await createDocument()
    const pending = await writeImageIntoOwnedAssets(
      sourcePath,
      'pending.png',
      Buffer.from('pending image'),
    )
    const targetDirectory = await mkdtemp(join(tmpdir(), 'markdown-save-as-failure-'))
    temporaryDirectories.push(targetDirectory)
    const targetPath = join(targetDirectory, 'copy.md')
    // A directory at the picked file path makes the final atomic rename fail
    // after destination assets have already been prepared.
    await mkdir(targetPath)
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })
    const view = createMarkdownView(sourcePath)
    const contents = view.webContents as unknown as FakeWebContents

    const saveResult = await handlers.get(MARKDOWN_CHANNELS.save)?.(
      { sender: contents },
      {
        text: `![pending](${pending})`,
        imageSources: [pending],
        mode: 'saveAs',
      },
    )

    expect(saveResult).toMatchObject({ ok: false })
    expect(existsSync(join(dirname(sourcePath), pending))).toBe(true)
    expect(await readOwnedAssetManifest(sourcePath)).toMatchObject({
      files: [{ name: 'pending.png' }],
      pending: { 'note.md': ['pending.png'] },
    })
    expect(existsSync(join(targetDirectory, 'assets', 'pending.png'))).toBe(false)
  })
})
