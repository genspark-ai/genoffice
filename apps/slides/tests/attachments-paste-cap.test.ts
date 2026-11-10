import { tmpdir } from 'node:os'
import { describe, expect, it, vi } from 'vitest'

const handlers = new Map<string, (...args: never[]) => unknown>()

vi.mock('electron', () => ({
  app: { getPath: () => tmpdir() },
  dialog: {},
  ipcMain: {
    handle: (channel: string, fn: (...args: never[]) => unknown) => handlers.set(channel, fn),
  },
  BrowserWindow: class {},
  webContents: {},
}))

vi.mock('@genoffice/file-parse', () => ({
  parseFileToText: vi.fn(async () => ({ ok: false, error: 'mocked' })),
}))
vi.mock('@genoffice/cli/agent-skills', () => ({}))
vi.mock('@genoffice/cli/install', () => ({}))
vi.mock('@genoffice/electron-utils', () => ({ showOpenDialogWithMemory: vi.fn() }))
// session-state pulls the font/shaping chain (harfbuzz wasm) which cannot
// load in the unit-test env; the paste handler under test never reaches it.
vi.mock('../src/main/session-state', () => ({ dialogParent: vi.fn(() => null) }))

import { registerAttachmentIpc } from '../src/main/attachments-ipc'

function pastedImageHandler() {
  handlers.clear()
  registerAttachmentIpc()
  const fn = handlers.get('slides:files-add-pasted-image')
  if (!fn) throw new Error('pasted-image handler not registered')
  return fn as (
    e: unknown,
    data: unknown,
    ext: unknown,
  ) => { accepted: unknown[]; rejected: string[] }
}

describe('pasted-image byte cap', () => {
  it('rejects oversized pastes before touching the temp disk', () => {
    const fn = pastedImageHandler()
    const big = fn({}, new Uint8Array(20 * 1024 * 1024), 'png')
    expect(big.accepted).toEqual([])
    expect(big.rejected).toHaveLength(1)
  })

  it('still accepts normal screenshots and rejects non-images', () => {
    const fn = pastedImageHandler()
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
      'base64',
    )
    const ok = fn({}, png, 'png')
    expect(ok.accepted).toHaveLength(1)
    expect(ok.rejected).toEqual([])
    const bad = fn({}, new Uint8Array([1, 2, 3]), 'txt')
    expect(bad.accepted).toEqual([])
    expect(bad.rejected).toHaveLength(1)
  })
})
