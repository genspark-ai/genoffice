// @vitest-environment node
import { randomBytes } from 'node:crypto'
import forge from 'node-forge'
import { describe, expect, it } from 'vitest'

type KeyFn = (
  password: string | null,
  salt: forge.util.ByteStringBuffer,
  id: number,
  iter: number,
  n: number,
  md?: forge.md.MessageDigest,
) => forge.util.ByteStringBuffer

// Capture forge's own implementation before the native one is installed
const reference = (forge.pki as unknown as { pbe: { generatePkcs12Key: KeyFn } }).pbe
  .generatePkcs12Key
const { nativePkcs12Key } = await import('../src/main/forge-native')

describe('native PKCS#12 key derivation', () => {
  const hashes: [string, () => forge.md.MessageDigest][] = [
    ['sha1', () => forge.md.sha1.create()],
    ['sha256', () => forge.md.sha256.create()],
    ['sha512', () => forge.md.sha512.create()],
  ]

  for (const [name, make] of hashes) {
    it(`matches forge's JavaScript result for ${name}`, () => {
      for (const password of ['', 'secret', 'pw-测试 ünï', 'x'.repeat(80)]) {
        for (const [id, iter, n, saltLength] of [
          [1, 1, 24, 8],
          [2, 7, 8, 16],
          [3, 50, 32, 20],
          [1, 3, 80, 0],
        ] as const) {
          const salt = randomBytes(saltLength).toString('binary')
          const want = reference(password, forge.util.createBuffer(salt), id, iter, n, make())
          const got = nativePkcs12Key(password, forge.util.createBuffer(salt), id, iter, n, make())
          expect(got.toHex()).toBe(want.toHex())
        }
      }
    })
  }

  it('refuses an absurd iteration count instead of spinning', () => {
    expect(() =>
      nativePkcs12Key('x', forge.util.createBuffer('saltsalt'), 1, 2_000_000_000, 24),
    ).toThrow(/iteration/)
  })

  it('reads a 600 000-iteration PBES2 bundle through the native derivation', async () => {
    const { execFileSync } = await import('node:child_process')
    let p12: Buffer
    try {
      const dir = (await import('node:fs')).mkdtempSync(
        (await import('node:path')).join((await import('node:os')).tmpdir(), 'iter-'),
      )
      const run = (file: string, args: string[]) =>
        execFileSync(file, args, { cwd: dir, stdio: 'pipe' })
      run('openssl', [
        'req',
        '-x509',
        '-newkey',
        'rsa:2048',
        '-nodes',
        '-keyout',
        'k.pem',
        '-out',
        'c.pem',
        '-days',
        '2',
        '-subj',
        '/CN=iter',
      ])
      run('openssl', [
        'pkcs12',
        '-export',
        '-inkey',
        'k.pem',
        '-in',
        'c.pem',
        '-out',
        'a.p12',
        '-passout',
        'pass:pw',
        '-iter',
        '600000',
        '-macalg',
        'sha256',
      ])
      p12 = (await import('node:fs')).readFileSync((await import('node:path')).join(dir, 'a.p12'))
    } catch {
      return // no openssl on this machine; the equivalence tests above carry the guarantee
    }
    const { loadCertificate } = await import('../src/main/pdf-sign')
    expect(loadCertificate(p12, 'pw').certificate.subject.getField('CN').value).toBe('iter')
    // Wall-clock limits flake under load; what matters is that forge no longer uses its JS loop
    expect(
      (forge.pki as unknown as { pbe: { generatePkcs12Key: unknown } }).pbe.generatePkcs12Key,
    ).toBe(nativePkcs12Key)
    expect((forge.pkcs12 as unknown as { generateKey: unknown }).generateKey).toBe(nativePkcs12Key)
  }, 30_000)
})
