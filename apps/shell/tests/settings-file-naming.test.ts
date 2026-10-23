import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { FILE_NAMING_PREF_KEY } from '../src/main/ai-naming'

/**
 * The model file-naming preference: the Settings toggle writes it, and the
 * naming policy reads it on every first save.
 *
 * Both halves live in the main process and `readAppSettings` is typed
 * `Record<string, unknown>`, so a writer and a reader that disagree about the
 * key is not a compile error — it is a toggle that animates, writes a real
 * file, and then reports the feature off on every launch after that. These
 * tests drive the actual IPC handlers against an actual settings file, so that
 * particular failure cannot pass again.
 */

type Handler = (event: unknown, ...args: unknown[]) => unknown

const handlers = new Map<string, Handler>()
let userData = ''

vi.mock('electron', () => ({
  app: { getPath: () => userData },
  ipcMain: {
    handle: (channel: string, fn: Handler) => {
      handlers.set(channel, fn)
    },
  },
}))

const call = (channel: string, ...args: unknown[]): unknown => {
  const fn = handlers.get(channel)
  if (!fn) throw new Error(`no handler registered for ${channel}`)
  return fn({}, ...args)
}

const SET = 'ai:set-file-naming'
const GET = 'ai:get-file-naming'

const onDisk = (): Record<string, unknown> =>
  JSON.parse(readFileSync(join(userData, 'app-settings.json'), 'utf8')) as Record<string, unknown>

beforeEach(async () => {
  vi.resetModules()
  handlers.clear()
  userData = mkdtempSync(join(tmpdir(), 'genoffice-naming-'))
  const { registerAiNamingIpc, readFileNamingPref } = await import('../src/main/ai-naming-service')
  // wired the way index.ts wires it: the reader the service exports, not a
  // second one spelled out at the call site
  registerAiNamingIpc(() => 'en', readFileNamingPref)
})

afterEach(() => {
  rmSync(userData, { recursive: true, force: true })
})

describe('the file-naming preference', () => {
  it('starts off, so a fresh launch never renames anything unasked', () => {
    expect(call(GET)).toBe(false)
  })

  it('reads back the value the toggle wrote', () => {
    expect(call(SET, true)).toBe(true)
    // through the handler the first-save path consults, not through the writer
    expect(call(GET)).toBe(true)
    // and through a fresh read of the file, which is what the next launch sees
    expect(call(GET)).toBe(true)
  })

  it('stores it under the key the reader looks for', () => {
    call(SET, true)
    expect(onDisk()[FILE_NAMING_PREF_KEY]).toBe(true)
  })

  it('turns back off', () => {
    call(SET, true)
    expect(call(SET, false)).toBe(false)
    expect(call(GET)).toBe(false)
    expect(onDisk()[FILE_NAMING_PREF_KEY]).toBe(false)
  })

  it('leaves the shell’s other preferences alone', async () => {
    const { writeAppSetting } = await import('../src/main/app-settings')
    writeAppSetting(join(userData, 'app-settings.json'), 'language', 'ar')
    call(SET, true)
    const settings = onDisk()
    expect(settings.language).toBe('ar')
    expect(settings[FILE_NAMING_PREF_KEY]).toBe(true)
  })

  it('ignores a payload that is not a boolean, and stores nothing', () => {
    expect(call(SET, 'yes')).toBe(false)
    expect(call(SET, 1)).toBe(false)
    expect(call(GET)).toBe(false)
    expect(() => onDisk()).toThrow()
  })
})
