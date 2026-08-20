import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { migrateUserDataDir } from '../src/index'

let appData: string

beforeEach(() => {
  appData = mkdtempSync(join(tmpdir(), 'genoffice-migrate-'))
})

afterEach(() => {
  rmSync(appData, { recursive: true, force: true })
})

describe('migrateUserDataDir', () => {
  it('copies the old directory into an empty new directory', () => {
    mkdirSync(join(appData, 'GenOffice Docs', 'ai-chat'), { recursive: true })
    writeFileSync(join(appData, 'GenOffice Docs', 'ai-settings.json'), '{"provider":"ollama"}')
    writeFileSync(join(appData, 'GenOffice Docs', 'ai-chat', 'docs.json'), '[]')

    migrateUserDataDir(appData, 'GenOffice Docs', 'KĀRYA Docs')

    expect(readFileSync(join(appData, 'KĀRYA Docs', 'ai-settings.json'), 'utf8')).toBe(
      '{"provider":"ollama"}',
    )
    expect(readFileSync(join(appData, 'KĀRYA Docs', 'ai-chat', 'docs.json'), 'utf8')).toBe('[]')
  })

  it('skips Chromium cache directories', () => {
    mkdirSync(join(appData, 'GenOffice', 'Cache'), { recursive: true })
    writeFileSync(join(appData, 'GenOffice', 'Cache', 'data_0'), 'cache')
    writeFileSync(join(appData, 'GenOffice', 'recent.json'), '[]')

    migrateUserDataDir(appData, 'GenOffice', 'KĀRYA')

    expect(readFileSync(join(appData, 'KĀRYA', 'recent.json'), 'utf8')).toBe('[]')
    expect(mkdirSync(join(appData, 'KĀRYA', 'Cache'), { recursive: true })).toBeDefined()
  })

  it('leaves a non-empty new directory alone (no duplicate migration)', () => {
    mkdirSync(join(appData, 'GenOffice'), { recursive: true })
    writeFileSync(join(appData, 'GenOffice', 'recent.json'), '["old"]')
    mkdirSync(join(appData, 'KĀRYA'), { recursive: true })
    writeFileSync(join(appData, 'KĀRYA', 'recent.json'), '["new"]')

    migrateUserDataDir(appData, 'GenOffice', 'KĀRYA')

    expect(readFileSync(join(appData, 'KĀRYA', 'recent.json'), 'utf8')).toBe('["new"]')
  })

  it('is a no-op when the old directory does not exist', () => {
    migrateUserDataDir(appData, 'GenOffice', 'KĀRYA')
    expect(mkdirSync(join(appData, 'KĀRYA'), { recursive: true })).toBeDefined()
  })

  it('is a no-op when old and new names are equal', () => {
    mkdirSync(join(appData, 'KĀRYA'), { recursive: true })
    writeFileSync(join(appData, 'KĀRYA', 'recent.json'), '[]')
    migrateUserDataDir(appData, 'KĀRYA', 'KĀRYA')
    expect(readFileSync(join(appData, 'KĀRYA', 'recent.json'), 'utf8')).toBe('[]')
  })
})
