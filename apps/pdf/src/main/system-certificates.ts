import { X509Certificate, randomBytes } from 'node:crypto'
import { execFile } from 'node:child_process'
import { existsSync, readdirSync } from 'node:fs'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import type { CertificateErrorCode } from '../shared/ipc'
import type { ExternalSigner } from './pdf-sign'
import { describeCertificate } from './pdf-signatures'

/**
 * Certificates that live in the operating system's own stores (the NSS database Chrome and
 * Firefox use on Linux, the Windows personal store). Private keys cannot be read in place, so a
 * chosen identity is exported through the platform's own tool into a PKCS#12 that is protected
 * by a random one-time password and kept in memory only; the signer then treats it like any
 * other .p12. Keys the store refuses to export are not supported there.
 *
 * Windows never exports: the personal store signs in place (CNG/CryptoAPI `SignHash`), so keys
 * flagged non-exportable, and keys on smart cards, work the same way they do in Adobe Reader.
 *
 * macOS is not offered: `security export` can only dump every identity of a keychain, prompts
 * per key and fails on any non-exportable one. The follow-up is a per-identity SecItemExport
 * path (native addon); until then Mac users sign from a .p12/.pfx file.
 */

export type SystemCertificateRef =
  { source: 'nss'; dir: string; nickname: string } | { source: 'windows'; thumbprint: string }

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
export type SystemStoreIssue = 'nss-tools-missing' | 'store-locked' | 'platform-unsupported'

export interface SystemCertificateListing {
  certs: SystemCertificate[]
  issues: SystemStoreIssue[]
}

export interface ExportedIdentity {
  p12: Uint8Array
  password: string
}

/** What a chosen system identity signs with: an exported PKCS#12, or the OS store itself */
export type SystemIdentity =
  ({ kind: 'p12' } & ExportedIdentity) | { kind: 'in-place'; signer: ExternalSigner }

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

/** Signer first, then the rest of the chain Windows can build; no revocation lookups */
export const WINDOWS_CHAIN_SCRIPT =
  POWERSHELL_UTF8 +
  "$ErrorActionPreference = 'Stop'; " +
  "$c = Get-Item ('Cert:\\CurrentUser\\My\\' + $env:GENOFFICE_THUMBPRINT); " +
  '$chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain; ' +
  "$chain.ChainPolicy.RevocationMode = 'NoCheck'; [void]$chain.Build($c); " +
  '$certs = @($c) + @($chain.ChainElements | ForEach-Object { $_.Certificate } | Select-Object -Skip 1); ' +
  '@($certs | ForEach-Object { [Convert]::ToBase64String($_.RawData) }) | ConvertTo-Json -Compress'

/** The private key stays in the store: Windows signs the digest and returns only the signature */
export const WINDOWS_SIGN_SCRIPT =
  "$ErrorActionPreference = 'Stop'; " +
  "$c = Get-Item ('Cert:\\CurrentUser\\My\\' + $env:GENOFFICE_THUMBPRINT); " +
  '$rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($c); ' +
  "if (-not $rsa) { throw 'no RSA private key' } " +
  '$hex = $env:GENOFFICE_DIGEST; ' +
  '$hash = [byte[]]::new($hex.Length / 2); ' +
  'for ($i = 0; $i -lt $hash.Length; $i++) { $hash[$i] = [Convert]::ToByte($hex.Substring(2 * $i, 2), 16) } ' +
  '[Convert]::ToBase64String($rsa.SignHash($hash, ' +
  '[System.Security.Cryptography.HashAlgorithmName]::SHA256, ' +
  '[System.Security.Cryptography.RSASignaturePadding]::Pkcs1))'

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

export function parseWindowsChain(json: string): Buffer[] {
  const text = json.replace(/^\uFEFF/, '').trim()
  if (!text) return []
  const parsed: unknown = JSON.parse(text)
  const rows = Array.isArray(parsed) ? parsed : [parsed]
  return rows
    .filter((row): row is string => typeof row === 'string')
    .map((b64) => Buffer.from(b64, 'base64'))
}

async function openWindowsSigner(
  env: SystemStoreEnv,
  ref: Extract<SystemCertificateRef, { source: 'windows' }>,
): Promise<SystemIdentity> {
  if (!THUMBPRINT.test(ref.thumbprint)) throw new SystemStoreError('cert-invalid')
  const thumbprintEnv = { GENOFFICE_THUMBPRINT: ref.thumbprint }
  const result = await env.run(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-Command', WINDOWS_CHAIN_SCRIPT],
    { env: thumbprintEnv },
  )
  if (result.missing) throw new SystemStoreError('cert-store-tool-missing')
  if (result.code !== 0) throw new SystemStoreError('cert-invalid')
  let chainDer: Buffer[]
  try {
    chainDer = parseWindowsChain(result.stdout)
  } catch {
    throw new SystemStoreError('cert-invalid')
  }
  if (chainDer.length === 0) throw new SystemStoreError('cert-invalid')
  return {
    kind: 'in-place',
    signer: {
      chainDer,
      async signDigest(digest) {
        const signed = await env.run(
          'powershell.exe',
          ['-NoProfile', '-NonInteractive', '-Command', WINDOWS_SIGN_SCRIPT],
          { env: { ...thumbprintEnv, GENOFFICE_DIGEST: Buffer.from(digest).toString('hex') } },
        )
        if (signed.missing) throw new SystemStoreError('cert-store-tool-missing')
        const signature = Buffer.from(signed.stdout.trim(), 'base64')
        if (signed.code !== 0 || signature.length === 0)
          throw new SystemStoreError('cert-not-exportable')
        return signature
      },
    },
  }
}

// ── entry points ─────────────────────────────────────────────────────────────

export async function listSystemCertificates(
  storePassword?: string,
  env: SystemStoreEnv = defaultEnv(),
): Promise<SystemCertificateListing> {
  if (env.platform === 'darwin') return { certs: [], issues: ['platform-unsupported'] }
  const listing =
    env.platform === 'win32' ? await listWindows(env) : await listNss(env, storePassword)
  listing.certs.sort(
    (a, b) => Number(a.expired) - Number(b.expired) || a.label.localeCompare(b.label),
  )
  return listing
}

export async function openSystemSigner(
  ref: SystemCertificateRef,
  storePassword?: string,
  env: SystemStoreEnv = defaultEnv(),
): Promise<SystemIdentity> {
  if (env.platform === 'darwin') throw new SystemStoreError('cert-store-unsupported')
  switch (ref.source) {
    case 'nss':
      return { kind: 'p12', ...(await exportNss(env, ref, storePassword)) }
    case 'windows':
      return openWindowsSigner(env, ref)
  }
}
