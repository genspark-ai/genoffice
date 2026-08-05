import { afterEach, describe, expect, it } from 'vitest'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import {
  CredentialStorageUnavailableError,
  MemoryCredentialStore,
  VersionedAiSettingsStore,
  type PersistedAiSettings,
  type CredentialStore,
} from '../src/index.js'

function defaults(): PersistedAiSettings {
  return {
    version: 1,
    active: {
      chat: { providerId: 'openai', model: 'gpt-test' },
      image: { providerId: 'openai', model: 'image-test' },
    },
    providers: {
      openai: { providerId: 'openai', model: 'gpt-test', enabled: true },
      custom: { providerId: 'custom', model: 'local-test', enabled: true },
    },
    updatedAt: new Date(0).toISOString(),
  }
}

class UnavailableCredentialStore implements CredentialStore {
  get(): string | undefined {
    return undefined
  }

  has(): boolean {
    return false
  }

  set(): void {
    throw new CredentialStorageUnavailableError()
  }

  delete(): void {}
}

describe('VersionedAiSettingsStore', () => {
  const tempDirs: string[] = []

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
  })

  function makeStore() {
    const dir = mkdtempSync(join(tmpdir(), 'ai-electron-settings-'))
    tempDirs.push(dir)
    const credentials = new MemoryCredentialStore()
    const store = new VersionedAiSettingsStore({
      settingsPath: join(dir, 'ai-settings.json'),
      credentialStore: credentials,
      defaults: defaults(),
    })
    return { dir, credentials, store }
  }

  it('returns a secret-free public DTO and resolves credentials only in main', () => {
    const { dir, store } = makeStore()
    const publicSettings = store.updateProvider({
      providerId: 'openai',
      credential: 'sk-secret-value',
    })

    expect(publicSettings.providers.openai).toMatchObject({
      credentialConfigured: true,
      credentialHint: '••••alue',
    })
    expect(publicSettings.providers.openai).not.toHaveProperty('credentialId')
    expect(JSON.stringify(publicSettings)).not.toContain('sk-secret-value')
    expect(readFileSync(join(dir, 'ai-settings.json'), 'utf8')).not.toContain('sk-secret-value')
    expect(store.resolve('chat')).toMatchObject({ providerId: 'openai', apiKey: 'sk-secret-value' })
  })

  it('resolves a configured provider without changing the active provider', () => {
    const { store } = makeStore()
    store.updateProvider({
      providerId: 'custom',
      model: 'custom-model',
      credential: 'custom-secret',
    })

    expect(store.resolveProvider('custom')).toMatchObject({
      providerId: 'custom',
      model: 'custom-model',
      apiKey: 'custom-secret',
    })
    expect(store.load().active.chat.providerId).toBe('openai')
  })

  it('updates task selection without accepting renderer credentials', () => {
    const { store } = makeStore()
    store.updateProvider({ providerId: 'custom', model: 'local-tool-model' })
    const publicSettings = store.setActive({ task: 'chat', providerId: 'custom' })

    expect(publicSettings.active.chat).toEqual({ providerId: 'custom', model: 'local-tool-model' })
    expect(store.resolve('chat')).toMatchObject({ providerId: 'custom', model: 'local-tool-model' })
  })

  it('persists an arbitrary image model ID across a store reload', () => {
    const { dir, credentials, store } = makeStore()
    const settingsPath = join(dir, 'ai-settings.json')
    store.updateProvider({ providerId: 'runware', model: 'private:custom-image@7' })
    store.setActive({ task: 'image', providerId: 'runware', model: 'private:custom-image@7' })

    const reloaded = new VersionedAiSettingsStore({
      settingsPath,
      credentialStore: credentials,
      defaults: defaults(),
    })
    expect(reloaded.load().active.image).toEqual({
      providerId: 'runware',
      model: 'private:custom-image@7',
    })
    expect(reloaded.resolve('image')).toMatchObject({
      providerId: 'runware',
      model: 'private:custom-image@7',
    })
  })

  it('supports task-specific routing hints beyond chat and image', () => {
    const { store } = makeStore()
    store.updateProvider({ providerId: 'custom', model: 'slide-tool-model' })
    store.setActive({ task: 'slides-generation', providerId: 'custom' })

    expect(store.resolve('slides-generation')).toMatchObject({
      task: 'slides-generation',
      providerId: 'custom',
      model: 'slide-tool-model',
    })
  })

  it('migrates legacy apiKey values and rewrites sanitized versioned JSON', () => {
    const { dir, store, credentials } = makeStore()
    writeFileSync(
      join(dir, 'ai-settings.json'),
      JSON.stringify({
        provider: 'custom',
        model: 'legacy-model',
        apiKey: 'legacy-secret',
        baseUrl: 'http://127.0.0.1:1234/v1',
      }),
    )

    const publicSettings = store.load()
    const persisted = readFileSync(join(dir, 'ai-settings.json'), 'utf8')
    expect(publicSettings.active.chat).toEqual({ providerId: 'custom', model: 'legacy-model' })
    expect(publicSettings.providers.custom.credentialConfigured).toBe(true)
    expect(credentials.get('ai-provider:custom')).toBe('legacy-secret')
    expect(persisted).not.toContain('legacy-secret')
    expect(JSON.parse(persisted).version).toBe(1)
  })

  it('canonicalizes retired DeepSeek model aliases during migration', () => {
    const { dir, store } = makeStore()
    writeFileSync(
      join(dir, 'ai-settings.json'),
      JSON.stringify({
        version: 1,
        active: { chat: { providerId: 'deepseek', model: 'deepseek-chat' } },
        providers: {
          deepseek: {
            providerId: 'deepseek',
            model: 'deepseek-reasoner',
            enabled: true,
          },
        },
      }),
    )

    const settings = store.load()
    expect(settings.active.chat).toEqual({
      providerId: 'deepseek',
      model: 'deepseek-v4-flash',
    })
    expect(settings.providers.deepseek.model).toBe('deepseek-v4-flash')
  })

  it('migrates the legacy Genspark text model out of the image selection', () => {
    const { dir, store } = makeStore()
    writeFileSync(
      join(dir, 'ai-settings.json'),
      JSON.stringify({
        version: 1,
        active: {
          chat: { providerId: 'openai', model: 'gpt-current' },
          image: { providerId: 'genspark', model: 'claude-opus-4-7' },
        },
        providers: {
          openai: { providerId: 'openai', model: 'gpt-current', enabled: true },
          genspark: { providerId: 'genspark', model: 'claude-opus-4-7', enabled: true },
        },
      }),
    )

    expect(store.load().active.image).toEqual({
      providerId: 'genspark',
      model: 'nano-banana-2',
    })
  })

  it('migrates every legacy provider credential, including providers absent from defaults', () => {
    const { dir, store, credentials } = makeStore()
    writeFileSync(
      join(dir, 'ai-settings.json'),
      JSON.stringify({
        provider: 'custom',
        model: 'custom-chat',
        providers: {
          openai: { model: 'gpt-legacy', apiKey: 'openai-legacy-secret' },
          custom: {
            model: 'local-legacy',
            baseUrl: 'http://127.0.0.1:1234/v1',
            apiKey: 'custom-legacy-secret',
          },
          vendor: { model: 'vendor-legacy', apiKey: 'vendor-legacy-secret' },
        },
        imageProvider: 'vendor',
        imageModel: 'vendor-image',
      }),
    )

    const publicSettings = store.load()
    const persisted = readFileSync(join(dir, 'ai-settings.json'), 'utf8')
    expect(publicSettings.providers.vendor).toMatchObject({
      providerId: 'vendor',
      model: 'vendor-legacy',
      credentialConfigured: true,
    })
    expect(publicSettings.active.image).toEqual({ providerId: 'vendor', model: 'vendor-image' })
    expect(credentials.get('ai-provider:openai')).toBe('openai-legacy-secret')
    expect(credentials.get('ai-provider:custom')).toBe('custom-legacy-secret')
    expect(credentials.get('ai-provider:vendor')).toBe('vendor-legacy-secret')
    expect(persisted).not.toContain('legacy-secret')
  })

  it('repairs a versioned file that still contains plaintext provider keys', () => {
    const { dir, store, credentials } = makeStore()
    writeFileSync(
      join(dir, 'ai-settings.json'),
      JSON.stringify({
        version: 1,
        active: {
          chat: { providerId: 'openai', model: 'gpt-current' },
          image: { providerId: 'openai', model: 'image-current' },
        },
        providers: {
          openai: {
            providerId: 'openai',
            model: 'gpt-current',
            apiKey: 'plaintext-current-secret',
            enabled: true,
          },
          custom: { providerId: 'custom', model: 'local-current', enabled: true },
        },
        updatedAt: new Date(0).toISOString(),
      }),
    )

    expect(store.load().providers.openai.credentialConfigured).toBe(true)
    expect(credentials.get('ai-provider:openai')).toBe('plaintext-current-secret')
    expect(readFileSync(join(dir, 'ai-settings.json'), 'utf8')).not.toContain(
      'plaintext-current-secret',
    )
  })

  it('does not treat malformed or empty apiKey values as credentials', () => {
    const { dir, store, credentials } = makeStore()
    writeFileSync(
      join(dir, 'ai-settings.json'),
      JSON.stringify({
        provider: 'custom',
        model: 'legacy-model',
        providers: {
          custom: { model: 'legacy-model', apiKey: '' },
          openai: { model: 'gpt-legacy', apiKey: 12345 },
        },
      }),
    )

    expect(store.load().providers.custom.credentialConfigured).toBe(false)
    expect(credentials.has('ai-provider:custom')).toBe(false)
    expect(credentials.has('ai-provider:openai')).toBe(false)
  })

  it('leaves a legacy file untouched when encrypted credential storage is unavailable', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-electron-settings-unavailable-'))
    tempDirs.push(dir)
    const path = join(dir, 'ai-settings.json')
    const legacy = JSON.stringify({ provider: 'custom', model: 'legacy', apiKey: 'secret' })
    writeFileSync(path, legacy)
    const store = new VersionedAiSettingsStore({
      settingsPath: path,
      credentialStore: new UnavailableCredentialStore(),
      defaults: defaults(),
    })

    expect(() => store.load()).toThrow(CredentialStorageUnavailableError)
    expect(readFileSync(path, 'utf8')).toBe(legacy)
  })

  it('keeps a corrupt settings file from crashing and restores defaults atomically', () => {
    const { dir, store } = makeStore()
    writeFileSync(join(dir, 'ai-settings.json'), '{not-json')
    expect(store.load().active.chat.providerId).toBe('openai')
    expect(JSON.parse(readFileSync(join(dir, 'ai-settings.json'), 'utf8')).version).toBe(1)
  })
})
