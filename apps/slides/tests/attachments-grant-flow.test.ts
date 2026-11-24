/**
 * The grant flow the genoffice#2758 review called out: a renderer-named files:add
 * must NOT extend what its view may read (the ledger only grows through
 * dialog picks and the preload bridge's drop/paste records), a dropped
 * path reads, and files beside the open deck read. The bridge+shell pair
 * is simulated by seeding the shared ledger the way the shell's
 * user-file-paths-seen handler does after a drop.
 */
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { mkdirSync, writeFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'

const handlers = new Map<string, (...args: unknown[]) => unknown>()
const deckSessions = vi.hoisted(() => new Map<number, { path?: string }>())

vi.mock('electron', () => ({
  app: { getPath: () => tmpdir() },
  dialog: {},
  ipcMain: {
    handle: (channel: string, fn: (...args: unknown[]) => unknown) => handlers.set(channel, fn),
  },
  BrowserWindow: class {},
  webContents: {},
  protocol: { registerSchemesAsPrivileged: vi.fn() },
}))
vi.mock('@genoffice/file-parse', () => ({
  parseFileToText: vi.fn(async (filePath: string) => ({
    ok: true,
    kind: 'text',
    text: `body of ${filePath}`,
  })),
}))
vi.mock('../src/main/session-state', () => ({
  dialogParent: vi.fn(() => null),
  sessions: deckSessions,
}))
vi.mock('@genoffice/electron-utils', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@genoffice/electron-utils')>()
  return { ...actual, showOpenDialogWithMemory: vi.fn() }
})

import { registerAttachmentIpc } from '../src/main/attachments-ipc'
import { sharedPathGrants } from '@genoffice/electron-utils'

type ReadResult = { ok: boolean; text?: string; error?: string }

function handler(channel: string): (...args: unknown[]) => unknown {
  const fn = handlers.get(channel)
  if (!fn) throw new Error(`${channel} not registered`)
  return fn
}

describe('attachment grants: the renderer cannot mint them', () => {
  it('files:add lists an attachment but leaves the ledger alone, so its reads stay refused', async () => {
    handlers.clear()
    registerAttachmentIpc()
    const wc = 7
    const neverDropped = join(tmpdir(), `go-never-dropped-${process.pid}.txt`)
    writeFileSync(neverDropped, 'secret')
    try {
      const add = handler('slides:files-add')({ sender: { id: wc } }, [neverDropped]) as {
        accepted: unknown[]
      }
      expect(add.accepted).toHaveLength(1)
      // the renderer named it; nothing recorded it, so reads stay refused
      const read = (await handler('slides:files-read')(
        { sender: { id: wc } },
        neverDropped,
        0,
        100,
      )) as ReadResult
      expect(read.ok).toBe(false)
      expect(read.error).toContain('not granted')
    } finally {
      sharedPathGrants().release(wc)
    }
  })

  it('a dropped path — recorded the way the bridge records drops — reads', async () => {
    handlers.clear()
    registerAttachmentIpc()
    const wc = 8
    const dropped = join(tmpdir(), `go-dropped-${process.pid}.txt`)
    writeFileSync(dropped, 'attachment body')
    try {
      // what the shell's user-file-paths-seen handler does after the
      // preload bridge reports an OS drop
      sharedPathGrants().grant(wc, [dropped])
      const read = (await handler('slides:files-read')(
        { sender: { id: wc } },
        dropped,
        0,
        100,
      )) as ReadResult
      expect(read.ok).toBe(true)
      expect(read.text).toContain('body of')
    } finally {
      writeFileSync(dropped, '')
    }
  })

  it('files beside the open deck read without any grant', async () => {
    handlers.clear()
    registerAttachmentIpc()
    const wc = 9
    const dir = join(tmpdir(), `go-deck-dir-${process.pid}`)
    mkdirSync(dir, { recursive: true })
    writeFileSync(join(dir, 'deck.pptx'), 'deck')
    const beside = join(dir, 'notes.txt')
    writeFileSync(beside, 'notes')
    deckSessions.set(wc, { path: join(dir, 'deck.pptx') })
    try {
      const read = (await handler('slides:files-read')(
        { sender: { id: wc } },
        beside,
        0,
        100,
      )) as ReadResult
      expect(read.ok).toBe(true)
      expect(read.text).toContain('notes')
    } finally {
      deckSessions.delete(wc)
    }
  })
})
