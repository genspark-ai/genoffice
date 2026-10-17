/**
 * Credential-store audit pins (H2 hardening lane, spec C-1..C-11).
 *
 * Behavioral pins use synthetic example.com fixtures only. Tree scans
 * below read shipped sources; this test file is allowlisted out of the
 * caller scans (it builds fakes), while the secret-shape scans are
 * written so this file itself contributes zero matches.
 */

import { describe, expect, it } from 'vitest'
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Buffer } from 'node:buffer'
import {
  CredentialReauthRequiredError,
  CredentialCorruptError,
  CredentialStore,
  GMAIL_OAUTH_SCOPES,
  GRAPH_OAUTH_SCOPES,
  IMAP_OAUTH_SCOPES,
  NO_KEYRING_BANNER,
  UNPROTECTED_LABEL,
  connectSafeStorage,
  postureForBackend,
  scopesForProvider,
} from '../src/main/credential-store'
import type { PutCredentialInput, SafeStorageLike } from '../src/main/credential-store'

// ---------------------------------------------------------------- fakes

interface FakeVaultConfig {
  backend?: string
  available?: boolean
  failDecrypt?: boolean
}

function makeFakeVault(config: FakeVaultConfig = {}): SafeStorageLike & {
  calls: string[]
  config: FakeVaultConfig
} {
  const calls: string[] = []
  const backend = config.backend ?? 'gnome_libsecret'
  const available = config.available ?? true
  const encode = (plain: string): Uint8Array => Buffer.from(`vault:${plain}`, 'utf8')
  return {
    calls,
    config,
    encryptStringAsync: async (plain: string) => {
      calls.push('encryptStringAsync')
      return encode(plain)
    },
    decryptStringAsync: async (encrypted: Uint8Array) => {
      calls.push('decryptStringAsync')
      if (config.failDecrypt === true) throw new Error('keychain locked')
      const text = Buffer.from(encrypted).toString('utf8')
      if (!text.startsWith('vault:')) throw new Error('bad blob')
      return text.slice('vault:'.length)
    },
    isAsyncEncryptionAvailable: async () => {
      calls.push('isAsyncEncryptionAvailable')
      return available
    },
    getSelectedStorageBackend: async () => {
      calls.push('getSelectedStorageBackend')
      return backend
    },
  }
}

function freshRoot(): string {
  const dir = join(tmpdir(), `cred-audit-${Date.now()}-${Math.floor(Math.random() * 1e6)}`)
  mkdirSync(dir, { recursive: true })
  return dir
}

function makeStore(
  rootDir: string,
  vault: SafeStorageLike,
  ready: { current: boolean } = { current: true },
  extra: { forceSessionOnly?: boolean; allowUnprotectedPersistence?: boolean } = {},
): CredentialStore {
  return new CredentialStore({
    rootDir,
    safeStorage: vault,
    isAppReady: () => ready.current,
    clock: () => '2026-10-15T00:00:00.000Z',
    ...extra,
  })
}

function gmailInput(secret = 'synthetic-gmail-refresh-secret'): PutCredentialInput {
  return {
    accountId: 'acct-gmail-1',
    provider: 'gmail',
    address: 'user@example.com',
    scopesGranted: [...GMAIL_OAUTH_SCOPES],
    secret,
  }
}

function readRecordFile(rootDir: string, accountId: string): Record<string, unknown> {
  const files = readdirSync(join(rootDir, 'credentials'))
  expect(files).toHaveLength(1)
  const raw = readFileSync(join(rootDir, 'credentials', files[0] as string), 'utf8')
  void accountId
  return JSON.parse(raw) as Record<string, unknown>
}

// ---------------------------------------------------------------- tree scan helpers

// Vitest rewrites import.meta.url in this package (see mail-quarantine
// precedent), so tree scans resolve from the package working directory.
const SRC_ROOTS = [
  join(process.cwd(), 'src'),
  join(process.cwd(), '..', '..', 'packages', 'mail-protocols', 'src'),
]
const MODULE_PATH = join(process.cwd(), 'src', 'main', 'credential-store.ts')
const SELF_NAME = 'credential-store.test.ts'

function walkTsFiles(dir: string): string[] {
  const out: string[] = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name)
    if (entry.isDirectory()) out.push(...walkTsFiles(full))
    else if (entry.isFile() && (full.endsWith('.ts') || full.endsWith('.tsx'))) out.push(full)
  }
  return out
}

function shippedSources(): string[] {
  return SRC_ROOTS.flatMap(walkTsFiles).filter((file) => !file.endsWith(SELF_NAME))
}

// ---------------------------------------------------------------- C-1 single caller

describe('C-1 single safeStorage caller', () => {
  const callPattern = new RegExp(
    'safeStorage' +
      '\\s*\\.\\s*' +
      '(encryptString|decryptString|isAsyncEncryptionAvailable|isEncryptionAvailable|' +
      'getSelectedStorageBackend|setUsePlainText' +
      'Encryption)\\s*\\(',
  )
  const importPattern = new RegExp(
    'import[^;]*safeStorage[^;]*from\\s*[' + '\'"' + ']electron[' + '\'"' + ']',
  )

  it('no shipped file calls safeStorage methods except credential-store.ts', () => {
    const offenders = shippedSources().filter((file) => {
      if (file.replace(/\\/g, '/').endsWith('main/credential-store.ts')) return false
      return callPattern.test(readFileSync(file, 'utf8'))
    })
    expect(offenders).toEqual([])
  })

  it('no shipped file imports safeStorage from electron except credential-store.ts', () => {
    const offenders = shippedSources().filter((file) => {
      if (file.replace(/\\/g, '/').endsWith('main/credential-store.ts')) return false
      return importPattern.test(readFileSync(file, 'utf8'))
    })
    expect(offenders).toEqual([])
  })

  it('the credential module exposes the Electron wiring through one connector', () => {
    expect(typeof connectSafeStorage).toBe('function')
  })
})

// ---------------------------------------------------------------- C-2 async-only + ready gate

describe('C-2 async-only API behind the app-ready gate', () => {
  it('uses only the async safeStorage surface, never the sync one', () => {
    const src = readFileSync(MODULE_PATH, 'utf8')
    expect(src).toContain('encryptStringAsync')
    expect(src).toContain('decryptStringAsync')
    expect(src).toContain('isAsyncEncryptionAvailable')
    expect(src).toContain('getSelectedStorageBackend')
    expect(src).not.toMatch(/\.(encryptString|decryptString|isEncryptionAvailable)\s*\(/)
  })

  it('refuses every operation before the app is ready', async () => {
    const ready = { current: false }
    const store = makeStore(freshRoot(), makeFakeVault(), ready)
    await expect(store.putCredential(gmailInput())).rejects.toThrow('before app ready')
    await expect(store.getCredential('acct-gmail-1')).rejects.toThrow('before app ready')
    await expect(store.backendPosture()).rejects.toThrow('before app ready')
    expect(() => store.hasCredential('acct-gmail-1')).toThrow('before app ready')
    expect(() => store.isSessionOnlyAccount('acct-gmail-1')).toThrow('before app ready')
    expect(() => store.describe('acct-gmail-1')).toThrow('before app ready')
    expect(() => store.cacheAccessToken('a', 't')).toThrow('before app ready')
    ready.current = true
    await expect(store.putCredential(gmailInput())).resolves.toBeDefined()
  })
})

// ---------------------------------------------------------------- C-3 stored shapes

describe('C-3 one opaque encrypted blob per account', () => {
  it('persists metadata plus an opaque blob, then round-trips the secret', async () => {
    const root = freshRoot()
    const vault = makeFakeVault()
    const store = makeStore(root, vault)
    await store.putCredential(gmailInput())
    const record = readRecordFile(root, 'acct-gmail-1')
    expect(record['accountId']).toBe('acct-gmail-1')
    expect(record['provider']).toBe('gmail')
    expect(record['address']).toBe('user@example.com')
    expect(record['scopesGranted']).toEqual([...GMAIL_OAUTH_SCOPES])
    expect(typeof record['blob']).toBe('string')
    expect(record['blob']).not.toBe('synthetic-gmail-refresh-secret')
    // The fake vault applies a reversible transform (the real backend
    // uses OS-keyring crypto): the pin is that the blob went through
    // encryptStringAsync and is not the raw secret.
    const decoded = Buffer.from(record['blob'] as string, 'base64').toString('utf8')
    expect(decoded).not.toBe('synthetic-gmail-refresh-secret')
    expect(vault.calls).toContain('encryptStringAsync')
    expect(typeof record['backend']).toBe('string')
    expect(typeof record['updatedAt']).toBe('string')

    const cold = makeStore(root, makeFakeVault())
    await expect(cold.getCredential('acct-gmail-1')).resolves.toEqual({
      secret: 'synthetic-gmail-refresh-secret',
      sessionOnly: false,
    })
  })

  it('stores IMAP app passwords with an empty scope set', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    await store.putCredential({
      accountId: 'acct-imap-1',
      provider: 'imap',
      address: 'user@example.com',
      scopesGranted: [],
      secret: 'synthetic-app-password',
    })
    const record = readRecordFile(root, 'acct-imap-1')
    expect(record['provider']).toBe('imap')
    expect(record['scopesGranted']).toEqual([])
  })
})

// ---------------------------------------------------------------- C-4 access tokens memory-only

describe('C-4 access tokens never persist', () => {
  it('keeps cached tokens out of files, records, and cold starts', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    await store.putCredential(gmailInput())
    const marker = 'synthetic-access-marker-abc123'
    store.cacheAccessToken('acct-gmail-1', marker)
    expect(store.takeAccessToken('acct-gmail-1')).toBe(marker)

    const record = readRecordFile(root, 'acct-gmail-1')
    expect(JSON.stringify(record)).not.toContain(marker)
    expect(Object.keys(record).join(',')).not.toMatch(/access/i)
    const walk = (dir: string): string[] =>
      readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
        const full = join(dir, entry.name)
        return entry.isDirectory() ? walk(full) : [full]
      })
    for (const file of walk(root)) {
      expect(readFileSync(file, 'utf8')).not.toContain(marker)
    }

    const cold = makeStore(root, makeFakeVault())
    expect(cold.takeAccessToken('acct-gmail-1')).toBeUndefined()
    store.clearAccessToken('acct-gmail-1')
    expect(store.takeAccessToken('acct-gmail-1')).toBeUndefined()
  })

  it('grants no SQLite column to access tokens', () => {
    for (const rel of ['main/store/migrations.ts', 'main/store/store.ts']) {
      const src = readFileSync(join(process.cwd(), 'src', ...rel.split('/')), 'utf8')
      expect(src).not.toMatch(/access_?token/i)
    }
  })

  it('passes the bearer to googleapis only through the in-memory handoff', () => {
    const protoSrc = join(process.cwd(), '..', '..', 'packages', 'mail-protocols', 'src')
    const files = walkTsFiles(protoSrc)
    const hits = files.filter((file) => /access_?token/i.test(readFileSync(file, 'utf8')))
    expect(hits.length).toBeGreaterThan(0)
    // Allowlist: the driver in-memory handoff (transport lane owns that
    // file) plus this lane's memory-only token cache, pinned behaviorally
    // above. Anything else holding an access token fails the suite.
    const allowed = [
      /gmail\.ts$/,
      /gmail-transport\.ts$/,
      /graph-transport\.ts$/,
      /credential-store\.ts$/,
    ]
    for (const file of hits) {
      const normalized = file.replace(/\\/g, '/')
      expect(allowed.some((re) => re.test(normalized))).toBe(true)
    }
    // The driver handoff lines resolve a fresh token per call and the
    // driver sources own no persistence verbs.
    for (const name of ['gmail.ts', 'gmail-transport.ts', 'graph-transport.ts']) {
      const src = readFileSync(join(protoSrc, name), 'utf8')
      for (const line of src.split('\n')) {
        if (/access_?token/i.test(line)) {
          expect(line).toMatch(/getAccessToken|access_token: token|access token/i)
        }
      }
      expect(src).not.toMatch(/writeFileSync|INSERT INTO|localStorage|setItem/)
    }
    const transport = readFileSync(join(protoSrc, 'gmail-transport.ts'), 'utf8')
    expect(transport).toContain('oauth.setCredentials({ access_token: token })')
  })
})

// ---------------------------------------------------------------- C-5 Linux-no-keyring fallback

describe('C-5 no-keyring means session-only, never silent plaintext', () => {
  it('resolves posture from the backend probe', () => {
    expect(postureForBackend('basic_text', true)).toBe('session-only')
    expect(postureForBackend('basic_text', false)).toBe('session-only')
    expect(postureForBackend('gnome_libsecret', false)).toBe('session-only')
    expect(postureForBackend('', true)).toBe('session-only')
    expect(postureForBackend('gnome_libsecret', true)).toBe('protected-keyring')
    expect(postureForBackend('kwallet5', true)).toBe('protected-keyring')
  })

  it('keeps secrets in memory only with no keyring and signals re-auth', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault({ backend: 'basic_text' }))
    expect(await store.backendPosture()).toBe('session-only')
    expect(await store.requiresSignInEachLaunch()).toBe(true)
    const summary = await store.putCredential(gmailInput())
    expect(summary.sessionOnly).toBe(true)
    expect(existsSync(join(root, 'credentials'))).toBe(false)
    await expect(store.getCredential('acct-gmail-1')).resolves.toEqual({
      secret: 'synthetic-gmail-refresh-secret',
      sessionOnly: true,
    })
    const cold = makeStore(root, makeFakeVault({ backend: 'basic_text' }))
    await expect(cold.getCredential('acct-gmail-1')).rejects.toBeInstanceOf(
      CredentialReauthRequiredError,
    )
  })

  it('publishes the exact no-keyring banner text', () => {
    expect(NO_KEYRING_BANNER).toBe(
      'No OS keyring detected — accounts will need sign-in each launch.',
    )
  })

  it('persists unprotected blobs only under the explicit labeled opt-in', async () => {
    const root = freshRoot()
    const store = makeStore(
      root,
      makeFakeVault({ backend: 'basic_text' }),
      { current: true },
      {
        allowUnprotectedPersistence: true,
      },
    )
    expect(UNPROTECTED_LABEL).toBe('unprotected (obfuscated only)')
    const summary = await store.putCredential(gmailInput())
    expect(summary.protection).toBe('obfuscated-only')
    expect(existsSync(join(root, 'credentials'))).toBe(true)
  })

  it('stays file-free on forced session-only launches even with the opt-in', async () => {
    const root = freshRoot()
    const store = makeStore(
      root,
      makeFakeVault({ backend: 'gnome_libsecret' }),
      { current: true },
      { forceSessionOnly: true, allowUnprotectedPersistence: true },
    )
    expect(await store.backendPosture()).toBe('session-only')
    await store.putCredential(gmailInput())
    expect(existsSync(join(root, 'credentials'))).toBe(false)
  })
})

// ---------------------------------------------------------------- C-6 plain-text encryption forbidden

describe('C-6 setUsePlainTextEncryption is forbidden in shipped code', () => {
  it('finds no occurrence in shipped sources', () => {
    const needle = 'setUsePlainText' + 'Encryption'
    const offenders = shippedSources().filter((file) => readFileSync(file, 'utf8').includes(needle))
    expect(offenders).toEqual([])
  })
})

// ---------------------------------------------------------------- C-7 decrypt failure + rotation

describe('C-7 decrypt failure wipes the blob and requires re-auth', () => {
  it('wipes the stale blob on decrypt failure instead of weakening storage', async () => {
    const root = freshRoot()
    await makeStore(root, makeFakeVault()).putCredential(gmailInput())
    expect(existsSync(join(root, 'credentials'))).toBe(true)
    const locked = makeStore(root, makeFakeVault({ failDecrypt: true }))
    await expect(locked.getCredential('acct-gmail-1')).rejects.toBeInstanceOf(
      CredentialReauthRequiredError,
    )
    expect(existsSync(join(root, 'credentials', 'acct-gmail-1.cred.json'))).toBe(false)
    expect(locked.hasCredential('acct-gmail-1')).toBe(false)
  })

  it('drops corrupt records without touching crypto', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    await store.putCredential(gmailInput())
    writeFileSync(join(root, 'credentials', 'acct-gmail-1.cred.json'), '{broken', 'utf8')
    const cold = makeStore(root, makeFakeVault())
    await expect(cold.getCredential('acct-gmail-1')).rejects.toBeInstanceOf(CredentialCorruptError)
    expect(existsSync(join(root, 'credentials', 'acct-gmail-1.cred.json'))).toBe(false)
  })

  it('wipes on invalid_grant and never loops refresh', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    await store.putCredential(gmailInput())
    store.reportInvalidGrant('acct-gmail-1')
    expect(store.hasCredential('acct-gmail-1')).toBe(false)
    expect(existsSync(join(root, 'credentials', 'acct-gmail-1.cred.json'))).toBe(false)
    await expect(store.getCredential('acct-gmail-1')).rejects.toBeInstanceOf(
      CredentialReauthRequiredError,
    )
  })

  it('re-encrypts under the current backend after a backend change', async () => {
    const root = freshRoot()
    await makeStore(root, makeFakeVault({ backend: 'gnome_libsecret' })).putCredential(gmailInput())
    const moved = makeStore(root, makeFakeVault({ backend: 'kwallet5' }))
    await expect(moved.getCredential('acct-gmail-1')).resolves.toEqual({
      secret: 'synthetic-gmail-refresh-secret',
      sessionOnly: false,
    })
    const record = readRecordFile(root, 'acct-gmail-1')
    expect(record['backend']).toBe('kwallet5')
  })

  it('removes every trace of an account on removal', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    await store.putCredential(gmailInput())
    store.cacheAccessToken('acct-gmail-1', 'synthetic-token')
    store.removeCredential('acct-gmail-1')
    expect(store.hasCredential('acct-gmail-1')).toBe(false)
    expect(store.takeAccessToken('acct-gmail-1')).toBeUndefined()
    expect(store.isSessionOnlyAccount('acct-gmail-1')).toBe(false)
    expect(existsSync(join(root, 'credentials', 'acct-gmail-1.cred.json'))).toBe(false)
  })
})

// ---------------------------------------------------------------- C-8 least-privilege scopes

describe('C-8 least-privilege OAuth scopes', () => {
  it('pins the exact requested scope sets', () => {
    expect([...GMAIL_OAUTH_SCOPES]).toEqual([
      'openid',
      'email',
      'profile',
      'https://www.googleapis.com/auth/gmail.modify',
    ])
    expect([...GRAPH_OAUTH_SCOPES]).toEqual(['Mail.ReadWrite', 'offline_access', 'User.Read'])
    expect([...IMAP_OAUTH_SCOPES]).toEqual([])
    expect(scopesForProvider('gmail')).toBe(GMAIL_OAUTH_SCOPES)
    expect(scopesForProvider('graph')).toBe(GRAPH_OAUTH_SCOPES)
    expect(scopesForProvider('imap')).toBe(IMAP_OAUTH_SCOPES)
  })

  it('matches the driver scope constants, with no broader scope anywhere', () => {
    const gmail = readFileSync(
      join(process.cwd(), '..', '..', 'packages', 'mail-protocols', 'src', 'gmail.ts'),
      'utf8',
    )
    expect(gmail).toContain("GMAIL_MODIFY_SCOPE = 'https://www.googleapis.com/auth/gmail.modify'")
    expect(GMAIL_OAUTH_SCOPES).toContain('https://www.googleapis.com/auth/gmail.modify')
    for (const src of shippedSources().map((file) => readFileSync(file, 'utf8'))) {
      const code = src
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .split('\n')
        .filter((line) => !line.trim().startsWith('//') && !line.trim().startsWith('*'))
        .join('\n')
      expect(code).not.toMatch(/https:\/\/www\.googleapis\.com\/auth\/gmail(?!(\.modify))/m)
      expect(code).not.toMatch(/Mail\.ReadWrite\.All|\*\.(ReadWrite|Read)\.All/)
    }
  })

  it('refuses broader scopes and scoped IMAP secrets', async () => {
    const store = makeStore(freshRoot(), makeFakeVault())
    const wide: PutCredentialInput = {
      ...gmailInput(),
      scopesGranted: [...GMAIL_OAUTH_SCOPES, 'https://www.googleapis.com/auth/gmail'],
    }
    await expect(store.putCredential(wide)).rejects.toThrow('refusing broader OAuth scope')
    const graphWide: PutCredentialInput = {
      accountId: 'acct-graph-9',
      provider: 'graph',
      address: 'user@example.com',
      scopesGranted: ['Mail.ReadWrite', 'Mail.ReadWrite.All'],
      secret: 'synthetic-graph-refresh-secret',
    }
    await expect(store.putCredential(graphWide)).rejects.toThrow('refusing broader OAuth scope')
    await expect(
      store.putCredential({
        accountId: 'acct-imap-9',
        provider: 'imap',
        address: 'user@example.com',
        scopesGranted: ['Mail.ReadWrite'],
        secret: 'synthetic-app-password',
      }),
    ).rejects.toThrow('refusing broader OAuth scope')
  })
})

// ---------------------------------------------------------------- C-9 OAuth flow rules (no flow code in tree)

describe('C-9 OAuth flow invariants hold while no flow code exists', () => {
  it('keeps client secrets and code-exchange handling out of shipped code', () => {
    const needles = ['client' + '_secret', 'authorization' + '_code', 'clientSecret']
    const offenders = shippedSources().filter((file) => {
      const code = readFileSync(file, 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .split('\n')
        .filter((line) => !line.trim().startsWith('//') && !line.trim().startsWith('*'))
        .join('\n')
      return needles.some((needle) => code.includes(needle))
    })
    expect(offenders).toEqual([])
  })

  it('performs no embedded auth capture in the mail main process', () => {
    const offenders = shippedSources().filter((file) => {
      if (!file.includes('apps/mail/src')) return false
      const code = readFileSync(file, 'utf8')
      return /<webview|webviewTag|loadURL\(.*(oauth|authorize|login)/i.test(code)
    })
    expect(offenders).toEqual([])
  })
})

// ---------------------------------------------------------------- C-10 session-only account mode

describe('C-10 per-account session-only mode', () => {
  it('keeps flagged accounts memory-only on a healthy keyring', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    const summary = await store.putCredential({ ...gmailInput(), sessionOnly: true })
    expect(summary.sessionOnly).toBe(true)
    expect(store.isSessionOnlyAccount('acct-gmail-1')).toBe(true)
    expect(existsSync(join(root, 'credentials'))).toBe(false)
    await expect(store.getCredential('acct-gmail-1')).resolves.toEqual({
      secret: 'synthetic-gmail-refresh-secret',
      sessionOnly: true,
    })
    const summary2 = store.describe('acct-gmail-1')
    expect(summary2.sessionOnly).toBe(true)
    expect(summary2.address).toBe('user@example.com')
    const cold = makeStore(root, makeFakeVault())
    expect(cold.isSessionOnlyAccount('acct-gmail-1')).toBe(false)
    await expect(cold.getCredential('acct-gmail-1')).rejects.toBeInstanceOf(
      CredentialReauthRequiredError,
    )
  })

  it('marks ordinary keyring accounts as persistent', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    await store.putCredential(gmailInput())
    expect(store.isSessionOnlyAccount('acct-gmail-1')).toBe(false)
  })
})

// ---------------------------------------------------------------- C-11 secret-shape scans

describe('C-11 secret-shape scans stay out of shipped code', () => {
  const needles = [
    'refresh' + '_token',
    'client' + '_secret',
    'pass' + 'wd',
    '-----BEGIN ' + 'PRIVATE KEY-----',
    'ya29.',
    'xox' + 'b-',
  ]

  it('finds no secret shapes in shipped sources outside the redaction detectors', () => {
    const offenders = shippedSources().filter((file) => {
      if (file.replace(/\\/g, '/').endsWith('main/sync/redact.ts')) return false
      return needles.some((needle) => readFileSync(file, 'utf8').includes(needle))
    })
    expect(offenders).toEqual([])
  })

  it('confines test fixtures to synthetic hostile mail asserted as redacted', () => {
    const testDirs = [
      join(process.cwd(), 'tests'),
      join(process.cwd(), '..', '..', 'packages', 'mail-protocols', 'tests'),
    ]
    const hits: string[] = []
    for (const dir of testDirs) {
      for (const file of walkTsFiles(dir)) {
        if (file.endsWith(SELF_NAME)) continue
        if (needles.some((needle) => readFileSync(file, 'utf8').includes(needle))) {
          hits.push(file)
        }
      }
    }
    for (const file of hits) {
      const code = readFileSync(file, 'utf8')
      expect(code.toLowerCase()).toMatch(/synthetic|fake|example\.com|redact/)
    }
  })
})

// ---------------------------------------------------------------- log-safe summaries

describe('summaries never carry secrets or blobs', () => {
  it('exposes metadata only', async () => {
    const root = freshRoot()
    const store = makeStore(root, makeFakeVault())
    const summary = await store.putCredential(gmailInput('synthetic-top-secret-value'))
    expect(JSON.stringify(summary)).not.toContain('synthetic-top-secret-value')
    expect('secret' in summary).toBe(false)
    expect('blob' in summary).toBe(false)
    expect(summary.address).toBe('user@example.com')
    expect(summary.hasSecret).toBe(true)
  })
})
