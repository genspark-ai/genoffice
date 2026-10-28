import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

vi.mock('electron', () => {
  const native = {
    showSaveDialog: async (options: unknown) => nativeSave(options),
    showOpenDialog: async (options: unknown) => nativeOpen(options),
    showMessageBox: async (options: unknown) => nativeMessage(options),
  }
  return { dialog: native, session: { defaultSession: {} } }
})

const nativeSave = vi.fn(async () => ({ canceled: true }))
const nativeOpen = vi.fn(async () => ({ canceled: true, filePaths: [] }))
const nativeMessage = vi.fn(async () => ({ response: 0 }))

import {
  beginWebDispatch,
  endWebDispatch,
  installWebDialogPatch,
  listWebDialogDir,
  resolveWebDialog,
  validateWebDialogPath,
  webDialogStartDirs,
} from '../src/web/web-dialogs'
import { WebTabContents } from '../src/web/web-contents'

function tabWithCapturedDialogs(): {
  tab: WebTabContents
  requests: { id: number; options: unknown }[]
} {
  const tab = new WebTabContents('/docs/')
  const requests: { id: number; options: unknown }[] = []
  tab.deliver = (channel, args) => {
    expect(channel).toBe('web:dialog')
    requests.push(args[0] as { id: number; options: unknown })
  }
  return { tab, requests }
}

function uiRoot(): string {
  return mkdtempSync(join(tmpdir(), 'webdlg-'))
}

describe('web dialogs', () => {
  beforeEach(() => {
    nativeSave.mockClear()
    nativeOpen.mockClear()
    nativeMessage.mockClear()
    installWebDialogPatch()
  })

  let dir = ''
  afterEach(() => {
    if (dir) rmSync(dir, { recursive: true, force: true })
    dir = ''
    endWebDispatch()
  })

  it('routes dialogs to the tab during a web dispatch', async () => {
    const { tab, requests } = tabWithCapturedDialogs()
    beginWebDispatch(tab)
    const dialog = (await import('electron')).dialog as unknown as {
      showSaveDialog: (options: unknown) => Promise<unknown>
    }
    const result = dialog.showSaveDialog({ title: 'Export', defaultPath: '/tmp/x.pdf' })
    await new Promise((r) => setTimeout(r, 20))
    expect(requests).toHaveLength(1)
    expect(requests[0].options).toMatchObject({ kind: 'save', title: 'Export' })
    resolveWebDialog(requests[0].id, { canceled: false, filePath: '/tmp/x.pdf' })
    await expect(result).resolves.toEqual({ canceled: false, filePath: '/tmp/x.pdf' })
    expect(nativeSave).not.toHaveBeenCalled()
  })

  it('keeps native dialogs outside a web dispatch', async () => {
    const dialog = (await import('electron')).dialog as unknown as {
      showMessageBox: (options: unknown) => Promise<unknown>
    }
    await dialog.showMessageBox({ message: 'hi', buttons: ['OK'] })
    expect(nativeMessage).toHaveBeenCalledTimes(1)
  })

  it('maps open dialogs and honors a canceled response', async () => {
    const { tab, requests } = tabWithCapturedDialogs()
    beginWebDispatch(tab)
    const dialog = (await import('electron')).dialog as unknown as {
      showOpenDialog: (options: unknown) => Promise<unknown>
    }
    const result = dialog.showOpenDialog({ title: 'Pick' })
    await new Promise((r) => setTimeout(r, 20))
    expect(requests).toHaveLength(1)
    endWebDispatch()
    resolveWebDialog(requests[0].id, { canceled: true })
    await expect(result).resolves.toEqual({ canceled: true, filePaths: [] })
  })

  it('lists a directory for the overlay: dirs first, filtered, hidden skipped', async () => {
    dir = uiRoot()
    mkdirSync(join(dir, 'sub'))
    writeFileSync(join(dir, 'b.png'), 'x')
    writeFileSync(join(dir, 'a.txt'), 'x')
    writeFileSync(join(dir, '.hidden.png'), 'x')
    const entries = await listWebDialogDir(dir, ['png'])
    expect(entries).toEqual([
      { name: 'sub', dir: true },
      { name: 'b.png', dir: false },
    ])
  })

  it('validates dialog targets against the filesystem', async () => {
    dir = uiRoot()
    const file = join(dir, 'f.pdf')
    writeFileSync(file, 'x')
    expect(validateWebDialogPath(file, 'open')).toBe(file)
    expect(validateWebDialogPath(join(dir, 'nope.pdf'), 'open')).toBeNull()
    expect(validateWebDialogPath(join(dir, 'new.pdf'), 'save')).toBe(join(dir, 'new.pdf'))
    expect(validateWebDialogPath(join(dir, 'gone', 'new.pdf'), 'save')).toBeNull()
  })

  it('starts the overlay at the dialog default path', () => {
    dir = uiRoot()
    const file = join(dir, 'report.docx')
    writeFileSync(file, 'x')
    expect(webDialogStartDirs(file)).toEqual([dir])
    const fallback = webDialogStartDirs()
    expect(fallback).toHaveLength(1)
    // '/' on POSIX, the drive root on Windows
    expect(fallback[0]).toMatch(/^(\/|[A-Za-z]:\\)$/)
  })
})
