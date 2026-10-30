import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The Home desktop's two channels against a real temp dir: the handlers are
 * the only writers of `<userData>/desktop-layout.json`, so this pins what that
 * file is allowed to look like — roundtrip, corrupt-reads-empty, shape
 * refusal, entry/coordinate clamps, and that a save never leaves the file
 * half-written.
 */

const mockedRename = vi.hoisted(() => vi.fn())
/** the real rename, captured at mock time for the retry test's second attempt */
const realRename = vi.hoisted(() => ({ fn: undefined as typeof import('node:fs/promises').rename }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  realRename.fn = actual.rename
  return { ...actual, rename: mockedRename }
})

const handles = new Map<string, (...args: never[]) => unknown>()
let userData = ''

vi.mock('electron', () => ({
  app: { getPath: () => userData },
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      handles.set(channel, handler),
  },
}))

import { HOME_CHANNELS } from '../src/shared/home-api'
import {
  DEFAULT_DESKTOP_PREFS,
  customWallpaperName,
  desktopLayoutPath,
  desktopPrefsPath,
  registerDesktopIpc,
} from '../src/main/desktop-ipc'

const invoke = (channel: string, ...args: unknown[]) =>
  handles.get(channel)!({ sender: { id: 7 } } as never, ...args)

const readLayout = () => invoke(HOME_CHANNELS.desktopLayout) as Promise<unknown>
const saveLayout = (raw: unknown) => invoke(HOME_CHANNELS.desktopLayoutSave, raw) as Promise<void>
const readPrefs = () => invoke(HOME_CHANNELS.desktopPrefs) as Promise<unknown>
const savePrefs = (raw: unknown) => invoke(HOME_CHANNELS.desktopPrefsSave, raw) as Promise<void>
const wallpaperImage = (id: unknown) => invoke(HOME_CHANNELS.wallpaperImage, id) as Promise<string>
const pickCustom = () => invoke(HOME_CHANNELS.wallpaperPickCustom) as Promise<string>

const at = (x: number, y: number) => ({ x, y })

/** a whole, well-formed prefs object the read/save tests write around */
const VALID_PREFS = {
  wallpaper: 'builtin:teal-deep',
  carouselMinutes: 30,
  iconSize: 'sm',
  sortMode: 'kind',
  wallpaperFit: 'contain',
  labelColor: '#333333',
}

/** the file's parsed content, as the next launch would read it */
const readFile = (): unknown => JSON.parse(readFileSync(desktopLayoutPath(userData), 'utf8'))

/** scratch files a failed or finished atomic write might leave behind */
const leftovers = (): string[] =>
  readdirSync(userData).filter((name) => name.includes('.tmp') || name.startsWith('.desktop'))

beforeEach(() => {
  userData = mkdtempSync(join(tmpdir(), 'genoffice-desktop-'))
  // the write is the real scratch-then-rename unless a test overrides it
  mockedRename
    .mockReset()
    .mockImplementation((from: string, to: string) => realRename.fn!(from, to))
  registerDesktopIpc()
})

afterEach(() => {
  mockedRename.mockReset()
  rmSync(userData, { recursive: true, force: true })
})

describe('home:desktop-layout', () => {
  it('reads back what a save wrote, ids and coordinates intact', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12), 'file:/tmp/a.docx': at(108, 116) } })
    await expect(readLayout()).resolves.toEqual({
      positions: { 'app:docs': at(12, 12), 'file:/tmp/a.docx': at(108, 116) },
    })
  })

  it('answers an empty layout when the file does not exist yet', async () => {
    await expect(readLayout()).resolves.toEqual({ positions: {} })
  })

  it('answers an empty layout for a corrupt or shapeless file', async () => {
    for (const garbage of [
      '{oops',
      'null',
      '"nope"',
      '42',
      '[]',
      '{"items":{}}',
      '{"positions":[]}',
      '{"positions":"map"}',
      '{"positions":{"a":"cell"}}',
    ]) {
      writeFileSync(desktopLayoutPath(userData), garbage, 'utf8')
      await expect(readLayout(), `${garbage} should read empty`).resolves.toEqual({
        positions: {},
      })
    }
  })

  it('drops entries that lost their shape, keeping the rest of the map', async () => {
    writeFileSync(
      desktopLayoutPath(userData),
      JSON.stringify({
        positions: {
          'app:docs': at(12, 12),
          bad: { x: 'no', y: 12 },
          worse: 'not an object',
          'file:/tmp/a.docx': at(108, 116),
        },
      }),
      'utf8',
    )
    await expect(readLayout()).resolves.toEqual({
      positions: { 'app:docs': at(12, 12), 'file:/tmp/a.docx': at(108, 116) },
    })
  })

  it('normalizes a hand-edited file to the store invariant: caps and clamps', async () => {
    const positions: Record<string, { x: number; y: number }> = {}
    for (let i = 0; i < 600; i += 1) positions[`id${i}`] = at(-5 + i, 20000 - i)
    writeFileSync(desktopLayoutPath(userData), JSON.stringify({ positions }), 'utf8')
    const layout = (await readLayout()) as { positions: Record<string, { x: number; y: number }> }
    expect(Object.keys(layout.positions)).toHaveLength(500)
    expect(layout.positions.id0).toEqual(at(0, 10000))
    expect(layout.positions.id499).toEqual(at(494, 10000))
  })
})

describe('home:desktop-layout-save', () => {
  it('writes a pretty-printed { positions } object to <userData>/desktop-layout.json', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12) } })
    expect(readFile()).toEqual({ positions: { 'app:docs': at(12, 12) } })
    expect(readFileSync(desktopLayoutPath(userData), 'utf8')).toContain('\n  ')
  })

  it('refuses a payload that is not a positions map, leaving the file untouched', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12) } })
    const before = readFileSync(desktopLayoutPath(userData), 'utf8')

    for (const bad of [
      undefined,
      null,
      'a layout',
      42,
      [at(1, 1)],
      { positions: null },
      { positions: 'map' },
      { positions: [at(1, 1)] },
      { positions: { a: 'cell' } },
      { positions: { a: { x: 1 } } },
      { positions: { a: { x: '1', y: 2 } } },
      { positions: { a: { x: Number.NaN, y: 2 } } },
      { positions: { a: { x: 1, y: Number.POSITIVE_INFINITY } } },
    ]) {
      await expect(saveLayout(bad), `${JSON.stringify(bad)} should refuse`).rejects.toThrow()
    }
    expect(readFileSync(desktopLayoutPath(userData), 'utf8')).toBe(before)
    expect(leftovers()).toEqual([])
  })

  it('clamps coordinates into 0..10000 instead of refusing', async () => {
    await saveLayout({ positions: { far: at(-40, 40000), near: at(12, 12) } })
    await expect(readLayout()).resolves.toEqual({
      positions: { far: at(0, 10000), near: at(12, 12) },
    })
  })

  it('caps the map at 500 entries instead of refusing', async () => {
    const positions: Record<string, { x: number; y: number }> = {}
    for (let i = 0; i < 502; i += 1) positions[`id${i}`] = at(i, i)
    await saveLayout({ positions })
    const layout = (await readLayout()) as { positions: Record<string, unknown> }
    expect(Object.keys(layout.positions)).toHaveLength(500)
    expect(Object.keys(layout.positions)).toContain('id499')
    expect(Object.keys(layout.positions)).not.toContain('id500')
  })

  it('stores exactly the contract shape, dropping anything extra an entry carries', async () => {
    await saveLayout({ positions: { a: { x: 1, y: 2, note: 'foreign field' } } })
    expect(readFile()).toEqual({ positions: { a: at(1, 2) } })
  })

  it('lets an empty map empty the file, not delete it', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12) } })
    await saveLayout({ positions: {} })
    await expect(readLayout()).resolves.toEqual({ positions: {} })
  })

  it('keeps the last of two saves when they land in order', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12) } })
    await saveLayout({ positions: { 'app:docs': at(108, 12), 'app:sheets': at(12, 12) } })
    await expect(readLayout()).resolves.toEqual({
      positions: { 'app:docs': at(108, 12), 'app:sheets': at(12, 12) },
    })
  })
})

describe('the file is never half-written', () => {
  it('publishes through a scratch file and leaves none behind on success', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12) } })
    expect(readdirSync(userData)).toEqual(['desktop-layout.json'])
    expect(readFile()).toEqual({ positions: { 'app:docs': at(12, 12) } })
  })

  it('keeps the previous layout intact, and no scratch file behind, when the publish fails', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12) } })
    const before = readFileSync(desktopLayoutPath(userData), 'utf8')
    mockedRename.mockRejectedValue(Object.assign(new Error('busy'), { code: 'EPERM' }))
    try {
      await expect(saveLayout({ positions: { 'app:docs': at(108, 12) } })).rejects.toMatchObject({
        code: 'EPERM',
      })
      expect(readFileSync(desktopLayoutPath(userData), 'utf8')).toBe(before)
      expect(readdirSync(userData)).toEqual(['desktop-layout.json'])
    } finally {
      mockedRename.mockReset()
    }
  })

  it('is read the same after a write that retried a locked rename', async () => {
    await saveLayout({ positions: { 'app:docs': at(12, 12) } })
    // one EPERM (AV / cloud-sync lock), then the real rename
    mockedRename
      .mockRejectedValueOnce(Object.assign(new Error('busy'), { code: 'EPERM' }))
      .mockImplementationOnce((from: unknown, to: unknown) =>
        realRename.fn!(from as string, to as string),
      )
    try {
      await saveLayout({ positions: { 'app:docs': at(108, 12) } })
      await expect(readLayout()).resolves.toEqual({ positions: { 'app:docs': at(108, 12) } })
      expect(readdirSync(userData)).toEqual(['desktop-layout.json'])
    } finally {
      mockedRename.mockReset()
    }
  })
})

describe('home:desktop-prefs', () => {
  it('answers the defaults when the file does not exist yet', async () => {
    await expect(readPrefs()).resolves.toEqual(DEFAULT_DESKTOP_PREFS)
  })

  it('reads back what a save wrote, field for field', async () => {
    await savePrefs({
      wallpaper: 'builtin:teal-deep',
      carouselMinutes: 10,
      iconSize: 'lg',
      sortMode: 'name',
      wallpaperFit: 'tile',
      labelColor: '#0f7fff',
    })
    await expect(readPrefs()).resolves.toEqual({
      wallpaper: 'builtin:teal-deep',
      carouselMinutes: 10,
      iconSize: 'lg',
      sortMode: 'name',
      wallpaperFit: 'tile',
      labelColor: '#0f7fff',
    })
  })

  it('roundtrips every fill mode and both hex shapes', async () => {
    for (const fit of ['cover', 'contain', 'stretch', 'tile', 'center'] as const) {
      await savePrefs({ ...VALID_PREFS, wallpaperFit: fit })
      await expect(readPrefs()).resolves.toMatchObject({ wallpaperFit: fit })
    }
    for (const labelColor of ['#ffffff', '#333', '#0F7fFf']) {
      await savePrefs({ ...VALID_PREFS, labelColor })
      await expect(readPrefs()).resolves.toMatchObject({ labelColor })
    }
  })

  it('answers the defaults for a corrupt or shapeless file', async () => {
    for (const garbage of ['{oops', 'null', '"nope"', '42', '[]']) {
      writeFileSync(desktopPrefsPath(userData), garbage, 'utf8')
      await expect(readPrefs(), `${garbage} should read defaults`).resolves.toEqual(
        DEFAULT_DESKTOP_PREFS,
      )
    }
  })

  it('falls back per field for values that lost their shape', async () => {
    writeFileSync(
      desktopPrefsPath(userData),
      JSON.stringify({
        wallpaper: 'builtin:teal-deep',
        carouselMinutes: '10',
        iconSize: 'giant',
        sortMode: 'random',
        wallpaperFit: 'zoom',
        labelColor: 'red',
        foreign: true,
      }),
      'utf8',
    )
    await expect(readPrefs()).resolves.toEqual({
      wallpaper: 'builtin:teal-deep',
      carouselMinutes: 0,
      iconSize: 'md',
      sortMode: 'custom',
      wallpaperFit: 'cover',
      labelColor: '',
    })
  })

  it('normalizes anything but a #rgb/#rrggbb label color to the automatic empty string', async () => {
    for (const labelColor of ['ff0000', '#ff00', '#12345', '#gggggg', 'white', '#FF0000 ']) {
      writeFileSync(
        desktopPrefsPath(userData),
        JSON.stringify({ ...VALID_PREFS, labelColor }),
        'utf8',
      )
      await expect(readPrefs(), `${labelColor} should reset`).resolves.toMatchObject({
        labelColor: '',
      })
    }
  })

  it('clamps the carousel minutes into 0..1440 on read', async () => {
    writeFileSync(
      desktopPrefsPath(userData),
      JSON.stringify({
        wallpaper: 'default',
        carouselMinutes: 99_999,
        iconSize: 'md',
        sortMode: 'custom',
      }),
      'utf8',
    )
    await expect(readPrefs()).resolves.toMatchObject({ carouselMinutes: 1440 })
  })
})

describe('home:desktop-prefs-save', () => {
  const VALID = VALID_PREFS

  it('writes a pretty-printed prefs object to <userData>/desktop-prefs.json', async () => {
    await savePrefs(VALID)
    expect(JSON.parse(readFileSync(desktopPrefsPath(userData), 'utf8'))).toEqual(VALID)
    expect(readFileSync(desktopPrefsPath(userData), 'utf8')).toContain('\n  ')
  })

  it('refuses a payload that is not the prefs shape, leaving the file untouched', async () => {
    await savePrefs(VALID)
    const before = readFileSync(desktopPrefsPath(userData), 'utf8')
    for (const bad of [
      undefined,
      null,
      'prefs',
      42,
      [VALID],
      {},
      { wallpaper: 'x' },
      { ...VALID, wallpaper: '' },
      { ...VALID, wallpaper: 42 },
      { ...VALID, wallpaper: 'x'.repeat(201) },
      { ...VALID, carouselMinutes: -1 },
      { ...VALID, carouselMinutes: '10' },
      { ...VALID, carouselMinutes: Number.NaN },
      { ...VALID, iconSize: 'giant' },
      { ...VALID, iconSize: undefined },
      { ...VALID, sortMode: 'chaos' },
      { ...VALID, wallpaperFit: 'zoom' },
      { ...VALID, wallpaperFit: undefined },
      { ...VALID, labelColor: 'ff0000' },
      { ...VALID, labelColor: '#ff00' },
      { ...VALID, labelColor: '#12345' },
      { ...VALID, labelColor: 42 },
      { ...VALID, labelColor: undefined },
    ]) {
      await expect(savePrefs(bad), `${JSON.stringify(bad)} should refuse`).rejects.toThrow()
    }
    expect(readFileSync(desktopPrefsPath(userData), 'utf8')).toBe(before)
    expect(leftovers()).toEqual([])
  })

  it('stores exactly the contract shape, dropping anything extra the payload carries', async () => {
    await savePrefs({ ...VALID, note: 'foreign field' })
    expect(JSON.parse(readFileSync(desktopPrefsPath(userData), 'utf8'))).toEqual(VALID)
  })

  it('clamps the carousel minutes into 0..1440 instead of refusing', async () => {
    await savePrefs({ ...VALID, carouselMinutes: 99_999 })
    await expect(readPrefs()).resolves.toMatchObject({ carouselMinutes: 1440 })
  })

  it('keeps the last of two saves when they land in order', async () => {
    await savePrefs(VALID)
    await savePrefs({ ...VALID, iconSize: 'lg', sortMode: 'mtime', wallpaperFit: 'tile' })
    await expect(readPrefs()).resolves.toMatchObject({
      iconSize: 'lg',
      sortMode: 'mtime',
      wallpaperFit: 'tile',
    })
  })
})

// ── wallpapers: the fixture dirs the channels are re-registered against ──

let wallpapersDir = ''
let customDir = ''
const sourceDirs: string[] = []

const WALLPAPER_PNG = '\x89PNG\r\n\x1a\nwallpaper-bytes'
const WALLPAPER_JPEG = '\xff\xd8\xff\xe0jpeg-bytes'
const dataUrl = (bytes: string, mime: string) =>
  `data:${mime};base64,${Buffer.from(bytes, 'latin1').toString('base64')}`

/** a shipped-wallpapers fixture dir plus an empty custom dir, and the
 * channels re-registered against them; `manifest === undefined` leaves the
 * manifest out entirely */
function registerWallpaperFixture(
  manifest: unknown,
  files: Record<string, string>,
  pickImage?: () => Promise<string | null>,
): void {
  wallpapersDir = mkdtempSync(join(tmpdir(), 'genoffice-wp-'))
  customDir = mkdtempSync(join(tmpdir(), 'genoffice-wpcust-'))
  if (manifest !== undefined) {
    writeFileSync(join(wallpapersDir, 'wallpapers.json'), JSON.stringify(manifest))
  }
  for (const [name, body] of Object.entries(files)) {
    writeFileSync(join(wallpapersDir, name), body, 'latin1')
  }
  registerDesktopIpc({
    wallpapersDir: () => wallpapersDir,
    customWallpapersDir: () => customDir,
    pickImage,
  })
}

afterEach(() => {
  for (const dir of [wallpapersDir, customDir, ...sourceDirs]) {
    if (dir) rmSync(dir, { recursive: true, force: true })
  }
  wallpapersDir = ''
  customDir = ''
  sourceDirs.length = 0
})

describe('home:wallpaper-image', () => {
  const MANIFEST = [
    { id: 'brand-aurora', nameKey: 'wpNameAurora', filename: 'brand-aurora.png' },
    { id: 'teal-deep', nameKey: 'wpNameTeal', filename: 'teal-deep.png' },
  ]

  beforeEach(() => {
    registerWallpaperFixture(MANIFEST, {
      'brand-aurora.png': WALLPAPER_PNG,
      'teal-deep.png': WALLPAPER_PNG,
    })
  })

  it('serves a shipped builtin as a PNG data URL', async () => {
    await expect(wallpaperImage('builtin:teal-deep')).resolves.toBe(
      dataUrl(WALLPAPER_PNG, 'image/png'),
    )
  })

  it('serves a custom wallpaper with its own mime when it exists', async () => {
    writeFileSync(join(customDir, '1700000000-photo.jpeg'), WALLPAPER_JPEG, 'latin1')
    await expect(wallpaperImage('custom:1700000000-photo.jpeg')).resolves.toBe(
      dataUrl(WALLPAPER_JPEG, 'image/jpeg'),
    )
  })

  it("answers '' for the live aurora, unknown ids and non-strings", async () => {
    for (const id of [
      'default', // the CSS aurora has no file behind it
      '',
      'builtin:nope',
      'builtin:teal-deep.png', // the manifest holds ids, not file names
      'builtin:',
      'plain-id',
      42,
      null,
      undefined,
    ]) {
      await expect(wallpaperImage(id), `${String(id)} should be unknown`).resolves.toBe('')
    }
  })

  it("answers '' for path-shaped ids: a traversal has no syntax to ride in on", async () => {
    writeFileSync(join(customDir, 'teal-deep.png'), WALLPAPER_PNG, 'latin1')
    for (const id of [
      'builtin:../teal-deep',
      'builtin:..\\teal-deep',
      'builtin:teal/../teal-deep',
      'custom:../teal-deep.png',
      'custom:..\\teal-deep.png',
      'custom:sub/teal-deep.png',
      'custom:.',
      'custom:..',
      'custom:',
    ]) {
      await expect(wallpaperImage(id), `${id} should be refused`).resolves.toBe('')
    }
  })

  it("answers '' for a custom name that is not in the custom dir", async () => {
    await expect(wallpaperImage('custom:1700000000-photo.jpeg')).resolves.toBe('')
  })

  it('drops manifest entries that lost their shape, offering only the sound ones', async () => {
    registerWallpaperFixture(
      [
        { id: 'teal-deep', nameKey: 'wpNameTeal', filename: 'teal-deep.png' },
        { id: 'evil', nameKey: 'wpNameEvil', filename: '../outside.png' },
        { id: 'broken', nameKey: 'wpNameBroken' },
        'not an entry',
      ],
      { 'teal-deep.png': WALLPAPER_PNG },
    )
    await expect(wallpaperImage('builtin:teal-deep')).resolves.toBe(
      dataUrl(WALLPAPER_PNG, 'image/png'),
    )
    await expect(wallpaperImage('builtin:evil')).resolves.toBe('')
    await expect(wallpaperImage('builtin:broken')).resolves.toBe('')
  })

  it('is the empty string, never a throw, when the manifest is missing entirely', async () => {
    registerWallpaperFixture(undefined, {})
    await expect(wallpaperImage('builtin:teal-deep')).resolves.toBe('')
  })
})

describe('home:wallpaper-pick-custom', () => {
  /** a picked source image in its own folder, like a real dialog would return */
  const pickableImage = (name: string): string => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-wpsrc-'))
    sourceDirs.push(dir)
    const path = join(dir, name)
    writeFileSync(path, WALLPAPER_PNG, 'latin1')
    return path
  }

  it('copies the picked image into the custom dir and returns its id', async () => {
    const source = pickableImage('photo.png')
    registerWallpaperFixture([], {}, async () => source)
    const id = await pickCustom()
    expect(id).toMatch(/^custom:\d+-photo\.png$/)
    const name = id.slice('custom:'.length)
    expect(readFileSync(join(customDir, name), 'latin1')).toBe(WALLPAPER_PNG)
    // the returned id round-trips through wallpaperImage
    await expect(wallpaperImage(id)).resolves.toBe(dataUrl(WALLPAPER_PNG, 'image/png'))
  })

  it("answers '' when the dialog is cancelled", async () => {
    registerWallpaperFixture([], {}, async () => null)
    await expect(pickCustom()).resolves.toBe('')
    expect(readdirSync(customDir)).toEqual([])
  })

  it('never clobbers: the same instant and name land under distinct names', async () => {
    const source = pickableImage('photo.png')
    vi.spyOn(Date, 'now').mockReturnValue(1_700_000_000_000)
    registerWallpaperFixture([], {}, async () => source)
    const first = await pickCustom()
    const second = await pickCustom()
    expect(first).not.toBe(second)
    expect(readFileSync(join(customDir, first.slice('custom:'.length)), 'latin1')).toBe(
      WALLPAPER_PNG,
    )
    expect(readFileSync(join(customDir, second.slice('custom:'.length)), 'latin1')).toBe(
      WALLPAPER_PNG,
    )
    expect(readdirSync(customDir)).toHaveLength(2)
  })

  it('sanitizes hostile characters out of the picked name', async () => {
    vi.spyOn(Date, 'now').mockReturnValue(1_700_000_000_000)
    // the name builder is exercised directly: a real Windows filesystem
    // refuses to hold most of these characters in a source name at all
    expect(customWallpaperName('my:pho?to<>.png', 0)).toBe('1700000000000-my_pho_to__.png')
    expect(customWallpaperName('sub/dir/spaces in name.jpeg', 0)).toBe(
      '1700000000000-spaces in name.jpeg',
    )
    expect(customWallpaperName('C:\\imgs\\photo.webp', 0)).toBe('1700000000000-photo.webp')
    expect(customWallpaperName('???.png', 1)).toBe('1700000000000-2-___.png')
    expect(customWallpaperName('.png', 0)).toBe('1700000000000-.png.png')
  })
})
