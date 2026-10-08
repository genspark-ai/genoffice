import { X509Certificate, randomBytes } from 'node:crypto'
import { execFile } from 'node:child_process'
import { existsSync, readdirSync } from 'node:fs'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import type { CertificateErrorCode } from '../shared/ipc'
import { describeCertificate } from './pdf-signatures'

/**
 * Certificates that live in the operating system's own stores (the NSS database Chrome and
 * Firefox use on Linux, the Windows personal store, the macOS keychain). Private keys cannot be
 * read in place, so a chosen identity is exported through the platform's own tool into a
 * PKCS#12 that is protected by a random one-time password and kept in memory only; the signer
 * then treats it like any other .p12. Keys the store refuses to export are not supported.
 */

export type SystemCertificateRef =
  | { source: 'nss'; dir: string; nickname: string }
  | { source: 'windows'; thumbprint: string }
  | { source: 'macos'; sha1: string }

export interface SystemCertificate {
  ref: SystemCertificateRef
  label: string
  commonName: string
  issuerCommonName: string
  /** ISO 8601 */
  validTo: string
  expired: boolean
}

/** Why part of a store could not be read, so the UI can say what to do about it */
export type SystemStoreIssue = 'nss-tools-missing' | 'store-locked'

export interface SystemCertificateListing {
  certs: SystemCertificate[]
  issues: SystemStoreIssue[]
}

export interface ExportedIdentity {
  p12: Uint8Array
  password: string
  /** Lower-case SHA-1 of the wanted certificate, when the export holds more than one identity */
  sha1?: string
}

export class SystemStoreError extends Error {
  constructor(readonly code: CertificateErrorCode) {
    super(code)
  }
}

export interface RunResult {
  code: number
  stdout: string
  stderr: string
  /** The executable does not exist on this machine */
  missing: boolean
}

export type Runner = (
  file: string,
  args: string[],
  options?: { env?: Record<string, string> },
) => Promise<RunResult>

const COMMAND_TIMEOUT_MS = 60_000

/** execFile only: arguments never pass through a shell */
export const runCommand: Runner = (file, args, options) =>
  new Promise((resolve) => {
    const child = execFile(
      file,
      args,
      {
        timeout: COMMAND_TIMEOUT_MS,
        maxBuffer: 16 * 1024 * 1024,
        windowsHide: true,
        encoding: 'utf8',
        env: { ...process.env, ...options?.env },
      },
      (error, stdout, stderr) => {
        const failure = error as (NodeJS.ErrnoException & { code?: unknown }) | null
        resolve({
          code: !failure ? 0 : typeof failure.code === 'number' ? failure.code : 1,
          stdout: String(stdout ?? ''),
          stderr: String(stderr ?? ''),
          missing: failure?.code === 'ENOENT',
        })
      },
    )
    // Tools that want a password must fail instead of waiting on a terminal that is not there
    child.stdin?.end()
  })

export interface SystemStoreEnv {
  platform: NodeJS.Platform
  home: string
  run: Runner
  now: () => Date
}

const defaultEnv = (): SystemStoreEnv => ({
  platform: process.platform,
  home: homedir(),
  run: runCommand,
  now: () => new Date(),
})

/** Scratch directory readable by this user only, removed when `work` settles */
async function withPrivateDir<T>(work: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-cert-'))
  try {
    return await work(dir)
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

const newPassword = (): string => randomBytes(18).toString('base64url')

function toListed(
  ref: SystemCertificateRef,
  label: string,
  cert: X509Certificate,
  now: Date,
): SystemCertificate {
  const info = describeCertificate(cert)
  return {
    ref,
    label,
    commonName: info.commonName,
    issuerCommonName: info.issuerCommonName,
    validTo: info.validTo,
    expired: new Date(info.validTo).getTime() < now.getTime(),
  }
}

// ── Linux: NSS ───────────────────────────────────────────────────────────────

/** Databases worth looking at: the shared one Chrome uses, plus Firefox profiles */
export function nssDatabaseDirs(home: string): string[] {
  const dirs = [join(home, '.pki', 'nssdb')]
  const profileRoots = [
    join(home, '.mozilla', 'firefox'),
    join(home, 'snap', 'firefox', 'common', '.mozilla', 'firefox'),
    join(home, '.var', 'app', 'org.mozilla.firefox', '.mozilla', 'firefox'),
  ]
  for (const root of profileRoots) {
    try {
      for (const entry of readdirSync(root, { withFileTypes: true })) {
        if (entry.isDirectory()) dirs.push(join(root, entry.name))
      }
    } catch {
      /* the browser is not installed that way */
    }
  }
  return dirs.filter((dir) => existsSync(join(dir, 'cert9.db')))
}

/** `certutil -K` lines: `< 0> rsa  <key id>  <nickname>`; only RSA keys can sign here */
export function parseNssKeyList(output: string): string[] {
  const nicknames: string[] = []
  for (const line of output.split('\n')) {
    const match = /^<\s*\d+>\s+(\S+)\s+[0-9a-fA-F]{8,}\s+(.+?)\s*$/.exec(line)
    if (match && match[1] === 'rsa') nicknames.push(match[2]!)
  }
  return nicknames
}

const looksLocked = (stderr: string): boolean =>
  /password|login|authentication|SEC_ERROR_BAD_PASSWORD|token/i.test(stderr)

async function listNss(
  env: SystemStoreEnv,
  storePassword?: string,
): Promise<SystemCertificateListing> {
  const out: SystemCertificateListing = { certs: [], issues: [] }
  const dirs = nssDatabaseDirs(env.home)
  if (dirs.length === 0) return out
  await withPrivateDir(async (scratch) => {
    const pwArgs: string[] = []
    if (storePassword) {
      const file = join(scratch, 'store-pw')
      await writeFile(file, `${storePassword}\n`, { mode: 0o600 })
      pwArgs.push('-f', file)
    }
    for (const dir of dirs) {
      const db = `sql:${dir}`
      const keys = await env.run('certutil', ['-K', '-d', db, ...pwArgs])
      if (keys.missing) {
        out.issues.push('nss-tools-missing')
        return
      }
      if (keys.code !== 0) {
        if (looksLocked(keys.stderr)) out.issues.push('store-locked')
        continue
      }
      for (const nickname of parseNssKeyList(keys.stdout)) {
        const pem = await env.run('certutil', ['-L', '-d', db, '-n', nickname, '-a', ...pwArgs])
        if (pem.code !== 0) continue
        try {
          out.certs.push(
            toListed(
              { source: 'nss', dir, nickname },
              nickname,
              new X509Certificate(pem.stdout),
              env.now(),
            ),
          )
        } catch {
          /* a nickname that is not an X.509 certificate is not a signing identity */
        }
      }
    }
  })
  return out
}

async function exportNss(
  env: SystemStoreEnv,
  ref: Extract<SystemCertificateRef, { source: 'nss' }>,
  storePassword?: string,
): Promise<ExportedIdentity> {
  const password = newPassword()
  return withPrivateDir(async (scratch) => {
    const out = join(scratch, 'identity.p12')
    const exportPw = join(scratch, 'export-pw')
    await writeFile(exportPw, `${password}\n`, { mode: 0o600 })
    const args = ['-o', out, '-d', `sql:${ref.dir}`, '-n', ref.nickname, '-w', exportPw]
    if (storePassword) {
      const storePw = join(scratch, 'store-pw')
      await writeFile(storePw, `${storePassword}\n`, { mode: 0o600 })
      args.push('-k', storePw)
    }
    const result = await env.run('pk12util', args)
    if (result.missing) throw new SystemStoreError('cert-store-tool-missing')
    if (result.code !== 0) {
      throw new SystemStoreError(
        looksLocked(result.stderr) ? 'cert-store-locked' : 'cert-not-exportable',
      )
    }
    return { p12: new Uint8Array(await readFile(out)), password }
  })
}

// ── Windows: personal store ──────────────────────────────────────────────────

const POWERSHELL_UTF8 = '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; '

export const WINDOWS_LIST_SCRIPT =
  POWERSHELL_UTF8 +
  '@(Get-ChildItem Cert:\\CurrentUser\\My | Where-Object { $_.HasPrivateKey } | ForEach-Object { [pscustomobject]@{ ' +
  'Thumbprint = $_.Thumbprint; Subject = $_.Subject; Der = ([Convert]::ToBase64String($_.RawData)) } }) | ConvertTo-Json -Compress'

export const WINDOWS_EXPORT_SCRIPT =
  // Without 'Stop' a refused export is only a printed error and the process still exits 0
  "$ErrorActionPreference = 'Stop'; " +
  '$pw = ConvertTo-SecureString -String $env:GENOFFICE_PFX_PASSWORD -AsPlainText -Force; ' +
  "Export-PfxCertificate -Cert ('Cert:\\CurrentUser\\My\\' + $env:GENOFFICE_THUMBPRINT) " +
  '-FilePath $env:GENOFFICE_PFX_PATH -Password $pw -ChainOption BuildChain | Out-Null'

const THUMBPRINT = /^[0-9A-Fa-f]{40}$/

export function parseWindowsCertificates(json: string): { thumbprint: string; der: Buffer }[] {
  const text = json.replace(/^\uFEFF/, '').trim()
  if (!text) return []
  const parsed: unknown = JSON.parse(text)
  const rows = Array.isArray(parsed) ? parsed : [parsed]
  const out: { thumbprint: string; der: Buffer }[] = []
  for (const row of rows as { Thumbprint?: unknown; Der?: unknown }[]) {
    if (
      typeof row?.Thumbprint === 'string' &&
      typeof row.Der === 'string' &&
      THUMBPRINT.test(row.Thumbprint)
    ) {
      out.push({ thumbprint: row.Thumbprint.toUpperCase(), der: Buffer.from(row.Der, 'base64') })
    }
  }
  return out
}

async function listWindows(env: SystemStoreEnv): Promise<SystemCertificateListing> {
  const out: SystemCertificateListing = { certs: [], issues: [] }
  const result = await env.run('powershell.exe', [
    '-NoProfile',
    '-NonInteractive',
    '-Command',
    WINDOWS_LIST_SCRIPT,
  ])
  if (result.code !== 0) return out
  for (const row of parseWindowsCertificates(result.stdout)) {
    try {
      const cert = new X509Certificate(row.der)
      if (cert.publicKey.asymmetricKeyType !== 'rsa') continue
      const info = describeCertificate(cert)
      out.certs.push(
        toListed(
          { source: 'windows', thumbprint: row.thumbprint },
          info.commonName,
          cert,
          env.now(),
        ),
      )
    } catch {
      /* skip entries that are not plain X.509 */
    }
  }
  return out
}

async function exportWindows(
  env: SystemStoreEnv,
  ref: Extract<SystemCertificateRef, { source: 'windows' }>,
): Promise<ExportedIdentity> {
  if (!THUMBPRINT.test(ref.thumbprint)) throw new SystemStoreError('cert-invalid')
  const password = newPassword()
  return withPrivateDir(async (scratch) => {
    const out = join(scratch, 'identity.pfx')
    const result = await env.run(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-Command', WINDOWS_EXPORT_SCRIPT],
      {
        env: {
          GENOFFICE_PFX_PASSWORD: password,
          GENOFFICE_THUMBPRINT: ref.thumbprint,
          GENOFFICE_PFX_PATH: out,
        },
      },
    )
    if (result.missing) throw new SystemStoreError('cert-store-tool-missing')
    if (result.code !== 0) throw new SystemStoreError('cert-not-exportable')
    try {
      return { p12: new Uint8Array(await readFile(out)), password }
    } catch {
      // A zero exit with no file still means the store kept the key
      throw new SystemStoreError('cert-not-exportable')
    }
  })
}

// ── macOS: keychain ──────────────────────────────────────────────────────────

/** `security find-identity -v` lines: `  1) <SHA-1> "name"` */
export function parseMacIdentities(output: string): { sha1: string; name: string }[] {
  const out: { sha1: string; name: string }[] = []
  for (const line of output.split('\n')) {
    const match = /^\s*\d+\)\s+([0-9A-Fa-f]{40})\s+"(.*)"/.exec(line)
    if (match) out.push({ sha1: match[1]!.toLowerCase(), name: match[2]! })
  }
  return out
}

/** `security find-certificate -a -Z -p`: a `SHA-1 hash:` line precedes each PEM block */
export function parseMacCertificates(output: string): Map<string, string> {
  const byHash = new Map<string, string>()
  const pattern =
    /SHA-1 hash:\s*([0-9A-Fa-f]{40})[\s\S]*?(-----BEGIN CERTIFICATE-----[\s\S]*?-----END CERTIFICATE-----)/g
  for (const match of output.matchAll(pattern)) byHash.set(match[1]!.toLowerCase(), match[2]!)
  return byHash
}

async function listMac(env: SystemStoreEnv): Promise<SystemCertificateListing> {
  const out: SystemCertificateListing = { certs: [], issues: [] }
  const identities = await env.run('security', ['find-identity', '-v'])
  if (identities.code !== 0) return out
  const pems = parseMacCertificates(
    (await env.run('security', ['find-certificate', '-a', '-Z', '-p'])).stdout,
  )
  for (const identity of parseMacIdentities(identities.stdout)) {
    const pem = pems.get(identity.sha1)
    if (!pem) continue
    try {
      const cert = new X509Certificate(pem)
      if (cert.publicKey.asymmetricKeyType !== 'rsa') continue
      out.certs.push(
        toListed({ source: 'macos', sha1: identity.sha1 }, identity.name, cert, env.now()),
      )
    } catch {
      /* skip */
    }
  }
  return out
}

async function exportMac(
  env: SystemStoreEnv,
  ref: Extract<SystemCertificateRef, { source: 'macos' }>,
): Promise<ExportedIdentity> {
  const password = newPassword()
  return withPrivateDir(async (scratch) => {
    const out = join(scratch, 'identities.p12')
    // `security` can only export every identity of a keychain; the signer picks ours by SHA-1
    const result = await env.run('security', [
      'export',
      '-t',
      'identities',
      '-f',
      'pkcs12',
      '-P',
      password,
      '-o',
      out,
    ])
    if (result.missing) throw new SystemStoreError('cert-store-tool-missing')
    if (result.code !== 0) throw new SystemStoreError('cert-not-exportable')
    return { p12: new Uint8Array(await readFile(out)), password, sha1: ref.sha1 }
  })
}

// ── entry points ─────────────────────────────────────────────────────────────

export async function listSystemCertificates(
  storePassword?: string,
  env: SystemStoreEnv = defaultEnv(),
): Promise<SystemCertificateListing> {
  const listing =
    env.platform === 'win32'
      ? await listWindows(env)
      : env.platform === 'darwin'
        ? await listMac(env)
        : await listNss(env, storePassword)
  listing.certs.sort(
    (a, b) => Number(a.expired) - Number(b.expired) || a.label.localeCompare(b.label),
  )
  return listing
}

export async function exportSystemCertificate(
  ref: SystemCertificateRef,
  storePassword?: string,
  env: SystemStoreEnv = defaultEnv(),
): Promise<ExportedIdentity> {
  switch (ref.source) {
    case 'nss':
      return exportNss(env, ref, storePassword)
    case 'windows':
      return exportWindows(env, ref)
    case 'macos':
      return exportMac(env, ref)
  }
}
