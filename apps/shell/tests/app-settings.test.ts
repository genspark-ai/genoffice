import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { readAppSettings, writeAppSetting, writeAppSettings } from '../src/main/app-settings'

/**
 * userData/app-settings.json helpers (src/main/app-settings.ts): a flat JSON
 * object shared by the language preference and the first-run onboarding flag.
 */

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'app-settings-'))
  settingsPath = join(dir, 'app-settings.json')
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('readAppSettings', () => {
  it('returns an empty object when the file does not exist', () => {
    expect(readAppSettings(settingsPath)).toEqual({})
  })

  it('returns an empty object for invalid JSON', () => {
    writeFileSync(settingsPath, 'not json')
    expect(readAppSettings(settingsPath)).toEqual({})
  })

  it('returns an empty object when the JSON root is not an object', () => {
    writeFileSync(settingsPath, '[1, 2]')
    expect(readAppSettings(settingsPath)).toEqual({})
    writeFileSync(settingsPath, '"zh"')
    expect(readAppSettings(settingsPath)).toEqual({})
  })

  it('parses a valid settings object', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'zh', onboardingSeen: true }))
    expect(readAppSettings(settingsPath)).toEqual({ language: 'zh', onboardingSeen: true })
  })
})

describe('writeAppSetting', () => {
  it('creates the file with the single key on first write', () => {
    writeAppSetting(settingsPath, 'onboardingSeen', true)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ onboardingSeen: true })
  })

  it('preserves unrelated existing keys', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja' }))
    writeAppSetting(settingsPath, 'onboardingSeen', true)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      language: 'ja',
      onboardingSeen: true,
    })
  })

  it('overwrites the value of an existing key', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja' }))
    writeAppSetting(settingsPath, 'language', 'en')
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ language: 'en' })
  })

  it('recovers from a corrupt file by rewriting it', () => {
    writeFileSync(settingsPath, '{broken')
    writeAppSetting(settingsPath, 'language', 'en')
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ language: 'en' })
  })
})

describe('writeAppSettings', () => {
  it('persists onboarding completion and analytics choice together', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'en' }))
    writeAppSettings(settingsPath, { onboardingSeen: true, analyticsEnabled: false })
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      language: 'en',
      onboardingSeen: true,
      analyticsEnabled: false,
    })
  })
})
