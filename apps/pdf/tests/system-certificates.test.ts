// @vitest-environment node
import { execFileSync, spawnSync } from 'node:child_process'
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import forge from 'node-forge'
import { PDFDocument } from 'pdf-lib'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { openExternalCertificate, signPdf } from '../src/main/pdf-sign'
import { readPdfSignatures } from '../src/main/pdf-signatures'
import {
  SystemStoreError,
  openSystemSigner,
  listSystemCertificates,
  parseNssKeyList,
  parseWindowsCertificates,
  runCommand,
} from '../src/main/system-certificates'
import type { RunResult, Runner, SystemStoreEnv } from '../src/main/system-certificates'

const ok = (stdout: string): RunResult => ({ code: 0, stdout, stderr: '', missing: false })

function selfSigned(commonName: string, notAfter: Date) {
  const keys = forge.pki.rsa.generateKeyPair(2048)
  const cert = forge.pki.createCertificate()
  cert.publicKey = keys.publicKey
  cert.serialNumber = '0a'
  cert.validity.notBefore = new Date(Date.now() - 86_400_000)
  cert.validity.notAfter = notAfter
  const attrs = [
    {
      name: 'commonName',
      value: forge.util.encodeUtf8(commonName),
      valueTagClass: forge.asn1.Type.UTF8,
    },
  ]
  cert.setSubject(attrs)
  cert.setIssuer(attrs)
  cert.sign(keys.privateKey, forge.md.sha256.create())
  return { keys, cert, pem: forge.pki.certificateToPem(cert) }
}

describe('store output parsers', () => {
  it('reads RSA nicknames from `certutil -K`, keeping spaces and skipping other key types', () => {
    const output = [
      'certutil: Checking token "NSS Certificate DB" in slot "NSS User Private Key and Certificate Services"',
      '< 0> rsa      87580f5bc3d2587d52666d585574952c35a89a61   My Work Cert',
      '< 1> ec       aabbccddeeff00112233445566778899aabbccdd   Phone key',
      '< 2> rsa      00112233445566778899aabbccddeeff00112233   second',
      '',
    ].join('\n')
    expect(parseNssKeyList(output)).toEqual(['My Work Cert', 'second'])
  })

  it('reads one or many certificates from PowerShell JSON', () => {
    const der = Buffer.from('abc').toString('base64')
    const one = JSON.stringify({ Thumbprint: 'A'.repeat(40), Der: der })
    expect(parseWindowsCertificates(`\uFEFF${one}`)).toHaveLength(1)
    const many = JSON.stringify([
      { Thumbprint: 'a'.repeat(40), Der: der },
      { Thumbprint: 'not-a-thumbprint', Der: der },
    ])
    expect(parseWindowsCertificates(many).map((c) => c.thumbprint)).toEqual(['A'.repeat(40)])
    expect(parseWindowsCertificates('')).toEqual([])
  })
})

describe('listing and exporting with a stubbed platform', () => {
  it('lists Windows identities, drops non-RSA ones and marks expired ones', async () => {
    const live = selfSigned('Live', new Date(Date.now() + 86_400_000))
    const old = selfSigned('Old', new Date(Date.now() - 3_600_000))
    const der = (c: forge.pki.Certificate) =>
      Buffer.from(forge.asn1.toDer(forge.pki.certificateToAsn1(c)).getBytes(), 'binary').toString(
        'base64',
      )
    const run: Runner = async () =>
      ok(
        JSON.stringify([
          { Thumbprint: '1'.repeat(40), Der: der(old.cert) },
          { Thumbprint: '2'.repeat(40), Der: der(live.cert) },
        ]),
      )
    const env: SystemStoreEnv = { platform: 'win32', home: '/h', run, now: () => new Date() }
    const listing = await listSystemCertificates(undefined, env)
    expect(listing.certs.map((c) => [c.commonName, c.expired])).toEqual([
      ['Live', false],
      ['Old', true],
    ])
    expect(listing.certs[0]!.ref).toEqual({ source: 'windows', thumbprint: '2'.repeat(40) })
  })

  it('refuses a malformed Windows thumbprint instead of passing it to PowerShell', async () => {
    const env: SystemStoreEnv = {
      platform: 'win32',
      home: '/h',
      run: async () => {
        throw new Error('must not run')
      },
      now: () => new Date(),
    }
    await expect(
      openSystemSigner({ source: 'windows', thumbprint: '1"; calc; "' }, undefined, env),
    ).rejects.toMatchObject({ code: 'cert-invalid' })
  })

  it('signs with a Windows key that cannot be exported, without ever reading the key', async () => {
    const made = selfSigned('Non Exportable', new Date(Date.now() + 86_400_000))
    const der = forge.asn1.toDer(forge.pki.certificateToAsn1(made.cert)).getBytes()
    const thumbprint = forge.md.sha1.create().update(der).digest().toHex().toUpperCase()
    const scripts: string[] = []
    const env: SystemStoreEnv = {
      platform: 'win32',
      home: '/h',
      now: () => new Date(),
      run: async (_file, args, options) => {
        const script = args[args.length - 1]!
        scripts.push(script)
        expect(script).not.toContain('Export-PfxCertificate')
        expect(options?.env?.GENOFFICE_THUMBPRINT).toBe(thumbprint)
        if (script.includes('X509Chain')) {
          return ok(JSON.stringify([Buffer.from(der, 'binary').toString('base64')]))
        }
        // Stand-in for CNG: sign the SHA-256 digest with PKCS#1 v1.5
        const digest = Buffer.from(options!.env!.GENOFFICE_DIGEST!, 'hex')
        const fakeMd = {
          algorithm: 'sha256',
          digest: () => forge.util.createBuffer(digest.toString('binary')),
        }
        const raw = made.keys.privateKey.sign(fakeMd as never)
        return ok(Buffer.from(raw, 'binary').toString('base64'))
      },
    }
    const identity = await openSystemSigner({ source: 'windows', thumbprint }, undefined, env)
    if (identity.kind !== 'in-place') throw new Error('Windows signs in place')
    const opened = openExternalCertificate(identity.signer)
    if (!opened.ok) throw new Error(opened.error)
    const pdf = await PDFDocument.create()
    pdf.addPage([200, 200])
    const signed = await signPdf(await pdf.save(), {
      external: { signer: identity.signer, loaded: opened.loaded },
    })
    const [sig] = await readPdfSignatures(signed)
    expect(sig!.signer?.commonName).toBe('Non Exportable')
    expect(sig!.problems).toEqual(['self-signed'])
    expect(scripts.length).toBeGreaterThanOrEqual(2)
  }, 60_000)

  it('maps a tool that is not installed and a store that will not export', async () => {
    const missing: RunResult = { code: 1, stdout: '', stderr: '', missing: true }
    const refused: RunResult = {
      code: 1,
      stdout: '',
      stderr: 'PKCS12 Export: Unable to export private key',
      missing: false,
    }
    const make = (result: RunResult): SystemStoreEnv => ({
      platform: 'linux',
      home: '/h',
      run: async () => result,
      now: () => new Date(),
    })
    const ref = { source: 'nss' as const, dir: '/h/.pki/nssdb', nickname: 'x' }
    await expect(openSystemSigner(ref, undefined, make(missing))).rejects.toMatchObject({
      code: 'cert-store-tool-missing',
    })
    await expect(openSystemSigner(ref, undefined, make(refused))).rejects.toMatchObject({
      code: 'cert-not-exportable',
    })
    expect(new SystemStoreError('cert-store-locked').code).toBe('cert-store-locked')
  })

  it('offers no system store on macOS', async () => {
    const env: SystemStoreEnv = {
      platform: 'darwin',
      home: '/h',
      run: async () => {
        throw new Error('no tool may run')
      },
      now: () => new Date(),
    }
    expect(await listSystemCertificates(undefined, env)).toEqual({
      certs: [],
      issues: ['platform-unsupported'],
    })
    const ref = { source: 'nss' as const, dir: '/h/.pki/nssdb', nickname: 'x' }
    await expect(openSystemSigner(ref, undefined, env)).rejects.toMatchObject({
      code: 'cert-store-unsupported',
    })
  })
})

/** The tools print usage and exit non-zero for --help; only a missing executable matters */
const installed = (tool: string): boolean =>
  spawnSync(tool, ['--help'], { stdio: 'ignore' }).error === undefined
const hasNss = installed('certutil') && installed('pk12util')

describe.skipIf(!hasNss || process.platform !== 'linux')('against a real NSS database', () => {
  const home = mkdtempSync(join(tmpdir(), 'nss-home-'))
  const db = join(home, '.pki', 'nssdb')
  const PASSWORD = 'import-pw'

  beforeAll(() => {
    const made = selfSigned('\u6d4b\u8bd5 NSS User', new Date(Date.now() + 86_400_000))
    mkdirSync(db, { recursive: true })
    const p12 = join(home, 'in.p12')
    const asn1 = forge.pkcs12.toPkcs12Asn1(made.keys.privateKey, [made.cert], PASSWORD, {
      friendlyName: 'work cert',
      algorithm: '3des',
    })
    writeFileSync(p12, Buffer.from(forge.asn1.toDer(asn1).getBytes(), 'binary'))
    execFileSync('certutil', ['-N', '-d', `sql:${db}`, '--empty-password'])
    execFileSync('pk12util', ['-i', p12, '-d', `sql:${db}`, '-W', PASSWORD])
  }, 60_000)

  afterAll(() => rmSync(home, { recursive: true, force: true }))

  const env = (): SystemStoreEnv => ({
    platform: 'linux',
    home,
    run: runCommand,
    now: () => new Date(),
  })

  it('lists the identity, exports it and signs a PDF that verifies', async () => {
    const listing = await listSystemCertificates(undefined, env())
    expect(listing.issues).toEqual([])
    expect(listing.certs).toHaveLength(1)
    expect(listing.certs[0]!.commonName).toBe('\u6d4b\u8bd5 NSS User')

    const identity = await openSystemSigner(listing.certs[0]!.ref, undefined, env())
    if (identity.kind !== 'p12') throw new Error('NSS identities are exported')
    const pdf = await PDFDocument.create()
    pdf.addPage([200, 200])
    const signed = await signPdf(await pdf.save(), {
      p12: identity.p12,
      password: identity.password,
    })
    const [sig] = await readPdfSignatures(signed)
    expect(sig!.signer?.commonName).toBe('\u6d4b\u8bd5 NSS User')
    expect(sig!.problems).toEqual(['self-signed'])
  }, 60_000)

  it('leaves no key material behind in the temp directory', async () => {
    const before = new Set(
      (await import('node:fs'))
        .readdirSync(tmpdir())
        .filter((n) => n.startsWith('genoffice-cert-')),
    )
    const listing = await listSystemCertificates(undefined, env())
    await openSystemSigner(listing.certs[0]!.ref, undefined, env())
    const after = (await import('node:fs'))
      .readdirSync(tmpdir())
      .filter((n) => n.startsWith('genoffice-cert-') && !before.has(n))
    expect(after).toEqual([])
  }, 60_000)
})
