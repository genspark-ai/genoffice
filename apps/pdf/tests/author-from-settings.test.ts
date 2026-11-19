// @vitest-environment node
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFString } from 'pdf-lib'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { systemAuthorName } from '@genoffice/electron-utils'
import type { PdfApi, SavePdfResult } from '../src/shared/ipc'

type Handler = (event: { sender: FakeContents }, ...args: any[]) => any
type Listener = (event: unknown, ...args: unknown[]) => void
const { handlers, rendererListeners, exposed } = vi.hoisted(() => ({
  handlers: new Map<string, Handler>(),
  rendererListeners: new Map<string, Set<Listener>>(),
  exposed: new Map<string, unknown>(),
}))
let userData = ''
let lastContents: FakeContents
let nextId = 0
class FakeContents {
  id = ++nextId
  once = vi.fn()
  on = vi.fn()
  setWindowOpenHandler = vi.fn()
  loadURL = vi.fn()
  loadFile = vi.fn()
  isDestroyed = () => false
  send = vi.fn()
}
vi.mock('electron', () => ({
  app: {
    on: vi.fn(),
    whenReady: () => new Promise(() => {}),
    getPath: () => userData,
  },
  dialog: {},
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
  contextBridge: {
    exposeInMainWorld: (key: string, value: unknown) => exposed.set(key, value),
  },
  ipcRenderer: {
    invoke: async (channel: string, ...args: unknown[]) =>
      handlers.get(channel)!({ sender: lastContents }, ...args),
    send: vi.fn(),
    on: (channel: string, fn: Listener) => {
      const set = rendererListeners.get(channel) ?? new Set<Listener>()
      set.add(fn)
      rendererListeners.set(channel, set)
    },
    removeListener: (channel: string, fn: Listener) => rendererListeners.get(channel)?.delete(fn),
  },
}))
vi.mock('@genoffice/electron-utils/drop-open', () => ({ installDropOpenBridge: () => {} }))
import { createPdfView } from '../src/main/pdf-main'
import '../src/preload/index'

const api = () => exposed.get('pdfApi') as PdfApi
const shellBroadcast = (channel: string, ...args: unknown[]) => {
  for (const fn of rendererListeners.get(channel) ?? []) fn({}, ...args)
}

const dirs: string[] = []
afterEach(async () => {
  await Promise.all(dirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true })))
})

async function openWithSettings(settings?: Record<string, unknown>): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'pdf-author-'))
  dirs.push(dir)
  userData = dir
  if (settings) await writeFile(join(dir, 'app-settings.json'), JSON.stringify(settings))
  const source = join(dir, 'doc.pdf')
  const pdf = await PDFDocument.create()
  pdf.addPage([612, 792])
  await writeFile(source, await pdf.save())
  createPdfView(source)
  return source
}

async function savedAuthors(path: string): Promise<Map<string, unknown[]>> {
  const saved = await PDFDocument.load(new Uint8Array(await readFile(path)), {
    updateMetadata: false,
  })
  const { context } = saved
  const annots = context.lookup(saved.getPages()[0]!.node.get(PDFName.of('Annots')), PDFArray)
  const bySubtype = new Map<string, unknown[]>()
  for (let index = 0; index < annots.size(); index += 1) {
    const annot = context.lookup(annots.get(index), PDFDict)
    const subtype = context.lookup(annot.get(PDFName.of('Subtype')), PDFName).decodeText()
    const t = annot.get(PDFName.of('T'))
    const author = t instanceof PDFHexString || t instanceof PDFString ? t.decodeText() : t
    bySubtype.set(subtype, [...(bySubtype.get(subtype) ?? []), author])
  }
  return bySubtype
}

describe('the PDF annotation author', () => {
  it('is the user name set in Settings → General, as Docs, Sheets and Slides use it', async () => {
    await openWithSettings({ authorName: '  Ada   Lovelace ' })
    expect(await api().getUsername()).toBe('Ada Lovelace')
  })

  it('falls back to the computer account name when the setting is unset or blank', async () => {
    await openWithSettings()
    expect(await api().getUsername()).toBe(systemAuthorName())
    await openWithSettings({ authorName: '   ' })
    expect(await api().getUsername()).toBe(systemAuthorName())
  })

  it('follows a rename made in Settings while the document stays open', async () => {
    await openWithSettings({ authorName: 'Ada Lovelace' })
    const seen: string[] = []
    const off = api().onUsernameChanged((name) => seen.push(name))
    shellBroadcast('app:author-name-changed', 'Grace Hopper')
    shellBroadcast('app:author-name-changed', 42)
    off()
    shellBroadcast('app:author-name-changed', 'Someone Else')
    expect(seen).toEqual(['Grace Hopper'])
  })

  it('is written on the annotations a save adds, while AI Assistant notes keep their own', async () => {
    const source = await openWithSettings({ authorName: 'Ada Lovelace' })
    const author = await api().getUsername()
    const result: SavePdfResult = await api().save({
      path: source,
      markups: [
        {
          pageIndex: 0,
          type: 'highlight',
          color: [1, 1, 0],
          quads: [[100, 220, 200, 220, 100, 200, 200, 200]],
        },
      ],
      annotationAuthor: author,
      drawings: [
        { kind: 'rect', pageIndex: 0, color: [1, 0, 0], width: 1, rect: [10, 10, 60, 60] },
        {
          kind: 'note',
          pageIndex: 0,
          color: [1, 1, 0],
          at: [50, 700],
          contents: 'mine',
          author,
        },
        {
          kind: 'note',
          pageIndex: 0,
          color: [1, 1, 0],
          at: [80, 700],
          contents: 'theirs',
          author: 'AI Assistant',
        },
      ],
      formValues: [],
      stamps: [],
    })
    expect(result).toMatchObject({ ok: true })
    expect(Object.fromEntries(await savedAuthors(source))).toEqual({
      Highlight: ['Ada Lovelace'],
      Square: ['Ada Lovelace'],
      Text: ['Ada Lovelace', 'AI Assistant'],
    })
  })
})
