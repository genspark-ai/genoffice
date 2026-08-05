import { writeAtomicJson, readJsonFile } from './storage.js'

/** Minimal adapter around Electron's safeStorage API, kept injectable for tests. */
export interface SafeStorageAdapter {
  isEncryptionAvailable(): boolean
  encryptString(value: string): Uint8Array
  decryptString(value: Uint8Array): string
}

/**
 * Future-proof adapter shape for hosts that expose asynchronous keychain APIs.
 * Electron's current safeStorage methods are synchronous, so
 * EncryptedCredentialStore intentionally uses SafeStorageAdapter today.
 */
export interface AsyncSafeStorageAdapter {
  isEncryptionAvailable(): Promise<boolean>
  encryptString(value: string): Promise<Uint8Array>
  decryptString(value: Uint8Array): Promise<string>
}

export interface CredentialStore {
  get(id: string): string | undefined
  has(id: string): boolean
  set(id: string, secret: string): void
  delete(id: string): void
}

export class CredentialStorageUnavailableError extends Error {
  constructor() {
    super('The operating system credential store is unavailable')
    this.name = 'CredentialStorageUnavailableError'
  }
}

interface CredentialFile {
  version: 1
  values: Record<string, string>
}

/**
 * Stores safeStorage ciphertext in a separate JSON file. The adapter is passed
 * in by the Electron main process, so this package has no Electron dependency.
 */
export class EncryptedCredentialStore implements CredentialStore {
  private readonly values = new Map<string, string>()
  private loaded = false

  constructor(
    private readonly filePath: string,
    private readonly safeStorage: SafeStorageAdapter,
  ) {}

  get(id: string): string | undefined {
    this.ensureLoaded()
    const encoded = this.values.get(id)
    if (!encoded) return undefined
    try {
      return this.safeStorage.decryptString(Buffer.from(encoded, 'base64'))
    } catch {
      // Treat unreadable ciphertext as absent rather than returning corrupt data.
      return undefined
    }
  }

  has(id: string): boolean {
    return this.get(id) !== undefined
  }

  set(id: string, secret: string): void {
    if (!id) throw new Error('Credential id is required')
    if (!this.safeStorage.isEncryptionAvailable()) throw new CredentialStorageUnavailableError()
    this.ensureLoaded()
    const encrypted = this.safeStorage.encryptString(secret)
    this.values.set(id, Buffer.from(encrypted).toString('base64'))
    this.persist()
  }

  delete(id: string): void {
    this.ensureLoaded()
    if (!this.values.delete(id)) return
    this.persist()
  }

  private ensureLoaded(): void {
    if (this.loaded) return
    this.loaded = true
    const file = readJsonFile<CredentialFile>(this.filePath)
    if (!file || file.version !== 1 || !file.values || typeof file.values !== 'object') return
    for (const [id, value] of Object.entries(file.values)) {
      if (typeof value === 'string' && id) this.values.set(id, value)
    }
  }

  private persist(): void {
    const values = Object.fromEntries(this.values)
    writeAtomicJson(this.filePath, { version: 1, values } satisfies CredentialFile)
  }
}

/** Deterministic test double. It must never be used as a production fallback. */
export class MemoryCredentialStore implements CredentialStore {
  private readonly values = new Map<string, string>()

  get(id: string): string | undefined {
    return this.values.get(id)
  }

  has(id: string): boolean {
    return this.values.has(id)
  }

  set(id: string, secret: string): void {
    this.values.set(id, secret)
  }

  delete(id: string): void {
    this.values.delete(id)
  }
}
