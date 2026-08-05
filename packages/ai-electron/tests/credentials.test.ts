import { afterEach, describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import {
  CredentialStorageUnavailableError,
  EncryptedCredentialStore,
  type SafeStorageAdapter,
} from '../src/index.js'

class FakeSafeStorage implements SafeStorageAdapter {
  isEncryptionAvailable(): boolean {
    return true
  }

  encryptString(value: string): Uint8Array {
    return new TextEncoder().encode(`ciphertext:${value}`)
  }

  decryptString(value: Uint8Array): string {
    const decoded = new TextDecoder().decode(value)
    if (!decoded.startsWith('ciphertext:')) throw new Error('bad ciphertext')
    return decoded.slice('ciphertext:'.length)
  }
}

class UnavailableSafeStorage extends FakeSafeStorage {
  override isEncryptionAvailable(): boolean {
    return false
  }
}

describe('EncryptedCredentialStore', () => {
  const tempDirs: string[] = []

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
  })

  it('persists ciphertext and reloads the secret through the adapter', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-electron-credentials-'))
    tempDirs.push(dir)
    const path = join(dir, 'credentials.json')
    const first = new EncryptedCredentialStore(path, new FakeSafeStorage())
    first.set('openai', 'sk-secret-value')

    expect(first.get('openai')).toBe('sk-secret-value')
    expect(readFileSync(path, 'utf8')).not.toContain('sk-secret-value')
    expect(new EncryptedCredentialStore(path, new FakeSafeStorage()).get('openai')).toBe(
      'sk-secret-value',
    )
  })

  it('does not write a credential when encryption is unavailable', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-electron-credentials-'))
    tempDirs.push(dir)
    const path = join(dir, 'credentials.json')
    const store = new EncryptedCredentialStore(path, new UnavailableSafeStorage())

    expect(() => store.set('custom', 'secret')).toThrow(CredentialStorageUnavailableError)
    expect(existsSync(path)).toBe(false)
  })
})
