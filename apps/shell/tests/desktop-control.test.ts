import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The desktop control domain against real temp stores: reads come straight
 * from the desktop-ipc/todos-ipc files, writes land through their atomic
 * paths, and a live Home view hears about a new wallpaper through the
 * notifyDesktopPrefsChanged hook.
 */

vi.mock('electron', () => ({
  app: { getPath: () => '', isPackaged: false },
  ipcMain: { handle: vi.fn() },
  dialog: {},
  BrowserWindow: { fromWebContents: () => null },
}))

import {
  DesktopControlError,
  desktopControlHandler,
  runDesktopAction,
  type DesktopControlDeps,
} from '../src/main/desktop-control'
import { desktopPrefsPath } from '../src/main/desktop-ipc'

let userData: string
let wallpapers: string
let custom: string
let notified: unknown[]

const WINDOWS = [
  { id: 'home', kind: 'home', title: 'GenOffice', active: true },
  { id: 'tab-2', kind: 'docs', title: 'report.docx', active: false, filePath: '/tmp/report.docx' },
]

const deps = (): DesktopControlDeps => ({
  userDataDir: () => userData,
  wallpapersDir: () => wallpapers,
  customWallpapersDir: () => custom,
  listWindows: () => WINDOWS,
  notifyDesktopPrefsChanged: (prefs) => notified.push(prefs),
})

beforeEach(() => {
  userData = mkdtempSync(join(tmpdir(), 'genoffice-desktop-control-'))
  wallpapers = join(userData, 'builtin-wallpapers')
  custom = join(userData, 'wallpapers')
  mkdirSync(wallpapers, { recursive: true })
  writeFileSync(
    join(wallpapers, 'wallpapers.json'),
    JSON.stringify([
      { id: 'teal-deep', nameKey: 'wallpaperTealDeep', filename: 'teal-deep.png' },
      { id: 'warm-dusk', nameKey: 'wallpaperWarmDusk', filename: 'warm-dusk.png' },
    ]),
  )
  notified = []
})

afterEach(() => {
  rmSync(userData, { recursive: true, force: true })
})

const readPersistedWallpaper = () =>
  (JSON.parse(readFileSync(desktopPrefsPath(userData), 'utf8')) as { wallpaper: string }).wallpaper

describe('get-state', () => {
  it('answers the prefs, layout summary, todos and windows from the live stores', async () => {
    writeFileSync(
      join(userData, 'desktop-layout.json'),
      JSON.stringify({ positions: { 'app:docs': { x: 12, y: 12 } } }),
    )
    writeFileSync(
      join(userData, 'todos.json'),
      JSON.stringify([{ id: 't1', text: 'sign the contract', done: false }]),
    )
    const state = await runDesktopAction(deps(), { cmd: 'desktop', action: 'get-state' })
    expect((state.prefs as { wallpaper: string }).wallpaper).toBe('default')
    expect(state.layout).toMatchObject({ positioned: 1 })
    expect(state.todos).toEqual([{ id: 't1', text: 'sign the contract', done: false }])
    expect(state.windows).toEqual(WINDOWS)
  })
})

describe('set-wallpaper', () => {
  it('accepts a builtin by id, with or without the builtin: prefix', async () => {
    for (const input of ['teal-deep', 'builtin:teal-deep']) {
      const result = await runDesktopAction(deps(), {
        cmd: 'desktop',
        action: 'set-wallpaper',
        wallpaper: input,
      })
      expect(result).toEqual({ wallpaper: 'builtin:teal-deep' })
      expect(readPersistedWallpaper()).toBe('builtin:teal-deep')
    }
    // a mounted Home view hears about the change
    expect(notified).toHaveLength(2)
  })

  it('restores the default aurora', async () => {
    const result = await runDesktopAction(deps(), {
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: 'default',
    })
    expect(result).toEqual({ wallpaper: 'default' })
    expect(readPersistedWallpaper()).toBe('default')
  })

  it('rejects an unknown name with the available ids attached', async () => {
    const failure = await desktopControlHandler(deps())({
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: 'nope',
    })
    expect(failure).toMatchObject({
      ok: false,
      error: { reason: 'invalid_argument', detail: { available: ['teal-deep', 'warm-dusk'] } },
    })
  })

  it('accepts an already-imported custom image by its plain name', async () => {
    mkdirSync(custom, { recursive: true })
    writeFileSync(join(custom, '1730000000-photo.png'), Buffer.from([0x89, 0x50, 0x4e, 0x47]))
    const result = await runDesktopAction(deps(), {
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: 'custom:1730000000-photo.png',
    })
    expect(result).toEqual({ wallpaper: 'custom:1730000000-photo.png' })
    expect(readPersistedWallpaper()).toBe('custom:1730000000-photo.png')
    expect(notified).toHaveLength(1)
  })

  it('refuses path-like custom names before touching the disk, aligned with the read side', async () => {
    for (const input of [
      'custom:../x.png',
      'custom:..\\x.png',
      'custom:sub/x.png',
      'custom:/etc/passwd',
      'custom:',
    ]) {
      const failure = await desktopControlHandler(deps())({
        cmd: 'desktop',
        action: 'set-wallpaper',
        wallpaper: input,
      })
      expect(failure).toMatchObject({ ok: false, error: { reason: 'invalid_argument' } })
    }
    // nothing was written: no prefs file exists at all, never poisoned with an
    // id the read side would reject
    expect(existsSync(desktopPrefsPath(userData))).toBe(false)
    expect(notified).toEqual([])
  })

  it('refuses a custom name that is not in the custom dir', async () => {
    mkdirSync(custom, { recursive: true })
    writeFileSync(join(custom, '1730000000-photo.png'), 'x')
    const failure = await desktopControlHandler(deps())({
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: 'custom:1730000000-other.png',
    })
    expect(failure).toMatchObject({
      ok: false,
      error: {
        reason: 'invalid_argument',
        detail: { suggestion: 'pass an absolute image path to import one' },
      },
    })
    expect(existsSync(desktopPrefsPath(userData))).toBe(false)
  })

  it('imports an absolute image path into the custom dir, like the picker does', async () => {
    const source = join(userData, 'my photo.png')
    writeFileSync(source, Buffer.from([0x89, 0x50, 0x4e, 0x47]))
    const result = await runDesktopAction(deps(), {
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: source,
    })
    const id = String(result.wallpaper)
    expect(id).toMatch(/^custom:\d+-my photo\.png$/)
    // the copy landed in the custom dir and the prefs point at it
    expect(readPersistedWallpaper()).toBe(id)
    expect(readFileSync(join(custom, id.slice('custom:'.length)))).toEqual(
      Buffer.from([0x89, 0x50, 0x4e, 0x47]),
    )
  })

  it('refuses a missing file, a directory and a non-image', async () => {
    const missing = await desktopControlHandler(deps())({
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: join(userData, 'gone.png'),
    })
    expect(missing).toMatchObject({ ok: false, error: { reason: 'file_not_found' } })

    const notImage = join(userData, 'notes.txt')
    writeFileSync(notImage, 'hello')
    const badExt = await desktopControlHandler(deps())({
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: notImage,
    })
    expect(badExt).toMatchObject({ ok: false, error: { reason: 'invalid_argument' } })

    const dir = await desktopControlHandler(deps())({
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: userData,
    })
    expect(dir).toMatchObject({ ok: false, error: { reason: 'invalid_argument' } })
  })

  it('keeps the desktop prefs otherwise untouched', async () => {
    writeFileSync(
      desktopPrefsPath(userData),
      JSON.stringify({
        wallpaper: 'builtin:teal-deep',
        carouselMinutes: 30,
        iconSize: 'lg',
        sortMode: 'name',
        wallpaperFit: 'tile',
        labelColor: '#aabbcc',
      }),
    )
    await runDesktopAction(deps(), {
      cmd: 'desktop',
      action: 'set-wallpaper',
      wallpaper: 'warm-dusk',
    })
    expect(JSON.parse(readFileSync(desktopPrefsPath(userData), 'utf8'))).toMatchObject({
      wallpaper: 'builtin:warm-dusk',
      carouselMinutes: 30,
      iconSize: 'lg',
      sortMode: 'name',
      wallpaperFit: 'tile',
      labelColor: '#aabbcc',
    })
  })
})

describe('todo add / list', () => {
  it('prepends like the widget and reads back through the same file', async () => {
    writeFileSync(
      join(userData, 'todos.json'),
      JSON.stringify([{ id: 'old', text: 'existing', done: true }]),
    )
    const added = (await runDesktopAction(deps(), {
      cmd: 'desktop',
      action: 'todo-add',
      text: '  buy milk  ',
    })) as { added: string; todos: Array<{ id: string; text: string; done: boolean }> }
    expect(added.added).toBe('buy milk')
    expect(added.todos).toHaveLength(2)
    expect(added.todos[0]).toMatchObject({ text: 'buy milk', done: false })
    expect(added.todos[1]).toMatchObject({ id: 'old' })

    const listed = await runDesktopAction(deps(), { cmd: 'desktop', action: 'todo-list' })
    expect(listed.todos).toEqual(added.todos)
  })

  it('refuses empty text', async () => {
    await expect(
      runDesktopAction(deps(), { cmd: 'desktop', action: 'todo-add', text: '   ' }),
    ).rejects.toThrow(DesktopControlError)
  })
})

describe('windows-list', () => {
  it('passes the live enumeration through', async () => {
    const result = await runDesktopAction(deps(), { cmd: 'desktop', action: 'windows-list' })
    expect(result.windows).toEqual(WINDOWS)
  })
})
