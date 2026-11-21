// @vitest-environment node
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { setImmediate } from 'node:timers/promises'
import forge from 'node-forge'
import { afterEach, expect, it } from 'vitest'
import { listSystemCertificates } from '../src/main/system-certificates'
import type {
  RunResult,
  SystemCertificateListing,
  SystemStoreEnv,
} from '../src/main/system-certificates'

const homes: string[] = []
afterEach(() => {
  for (const home of homes.splice(0)) rmSync(home, { recursive: true, force: true })
})
const ok = (stdout: string): RunResult => ({ code: 0, stdout, stderr: '', missing: false })
const failed = (stderr = ''): RunResult => ({ code: 1, stdout: '', stderr, missing: false })
const keyList = (names: string[]) =>
  names.map((name, index) => `< ${index}> rsa aabbccddeeff0011 ${name}`).join('\n')

// Parseable synthetic ASN.1 only: dummy signature and public modulus, no private key or signing.
function certificate(name: string): string {
  const cert = forge.pki.createCertificate()
  cert.publicKey = forge.pki.setRsaPublicKey(
    new forge.jsbn.BigInteger('f'.repeat(256), 16),
    new forge.jsbn.BigInteger('65537'),
  )
  cert.serialNumber = '01'
  cert.validity.notBefore = new Date('2020-01-01')
  cert.validity.notAfter = new Date('2099-01-01')
  cert.setSubject([{ name: 'commonName', value: name }])
  cert.setIssuer([{ name: 'commonName', value: 'Synthetic issuer' }])
  cert.signatureOid = forge.pki.oids.sha256WithRSAEncryption
  cert.siginfo.algorithmOid = cert.signatureOid
  cert.signature = 'fixture'
  return forge.pki.certificateToPem(cert)
}
function homeWithDatabases(profiles: string[] = []) {
  const home = mkdtempSync(join(tmpdir(), 'nss-enumeration-fixture-'))
  homes.push(home)
  const dirs = [
    join(home, '.pki', 'nssdb'),
    ...profiles.map((name) => join(home, '.mozilla', 'firefox', name)),
  ]
  for (const dir of dirs) {
    mkdirSync(dir, { recursive: true })
    writeFileSync(join(dir, 'cert9.db'), '')
  }
  return { home, dirs }
}
type Pending = { finish: () => void }
function runner(
  home: string,
  keys: (db: string) => RunResult,
  reads: (index: number, db: string, nickname: string) => RunResult,
) {
  const pending: Pending[] = [],
    commands: string[][] = [],
    completed: number[] = []
  let active = 0,
    maximum = 0,
    next = 0
  const env: SystemStoreEnv = {
    platform: 'linux',
    home,
    now: () => new Date('2030-01-01'),
    run: async (file, args) => {
      expect(file).toBe('certutil')
      commands.push(args)
      const db = args[args.indexOf('-d') + 1]!
      if (args[0] === '-K') {
        expect(active).toBe(0)
        return keys(db)
      }
      expect(args[0]).toBe('-L')
      expect(args).toContain('-a')
      const index = next++,
        nickname = args[args.indexOf('-n') + 1]!
      active++
      maximum = Math.max(maximum, active)
      await new Promise<void>((resolve) => pending.push({ finish: resolve }))
      try {
        completed.push(index)
        return reads(index, db, nickname)
      } finally {
        active--
      }
    },
  }
  return { env, pending, commands, completed, maximum: () => maximum }
}
async function drain(work: Promise<SystemCertificateListing>, pending: Pending[]) {
  let done = false,
    value: SystemCertificateListing | undefined,
    error: unknown
  void work.then(
    (result) => {
      value = result
      done = true
    },
    (reason) => {
      error = reason
      done = true
    },
  )
  let waves = 0
  while (!done) {
    await setImmediate()
    if (pending.length) {
      waves++
      for (const item of pending.splice(0).reverse()) item.finish()
    }
  }
  if (error) throw error
  return { listing: value!, waves }
}

it.each([0, 1, 7, 8, 9, 32])(
  'bounds parallel reads and preserves every identity for %i keys',
  async (count) => {
    const { home } = homeWithDatabases()
    const names = Array.from(
      { length: count },
      (_, index) => `Identity ${index.toString().padStart(2, '0')}`,
    )
    const pem = certificate('Synthetic identity')
    const fake = runner(
      home,
      () => ok(keyList(names)),
      () => ok(pem),
    )
    const { listing, waves } = await drain(
      listSystemCertificates(undefined, fake.env),
      fake.pending,
    )
    expect(fake.maximum()).toBe(Math.min(8, count))
    expect(waves).toBe(Math.ceil(count / 8))
    expect(listing.certs.map((cert) => cert.label)).toEqual(names)
    expect(listing.issues).toEqual([])
    expect(fake.commands).toHaveLength(1 + count)
  },
)

it('merges out-of-order completions stably and skips individual read or PEM failures', async () => {
  const { home } = homeWithDatabases()
  // Equal labels make the public final sort retain, rather than mask, collection order.
  const names = Array.from({ length: 9 }, () => 'Same nickname')
  const pems = names.map((_, index) => certificate(`Synthetic ${index}`))
  const fake = runner(
    home,
    () => ok(keyList(names)),
    (index) =>
      index === 2
        ? failed('read failed')
        : index === 5
          ? ok('not a certificate')
          : ok(pems[index]!),
  )
  const { listing } = await drain(listSystemCertificates(undefined, fake.env), fake.pending)
  expect(fake.completed.slice(0, 8)).toEqual([7, 6, 5, 4, 3, 2, 1, 0])
  expect(listing.certs.map((cert) => cert.commonName)).toEqual(
    [0, 1, 3, 4, 6, 7, 8].map((index) => `Synthetic ${index}`),
  )
  expect(listing.issues).toEqual([])
})

it('keeps database order, locked/missing-tool issues and the global listing bound', async () => {
  const { home, dirs } = homeWithDatabases(['a-profile', 'b-profile', 'c-profile'])
  const pem = certificate('Synthetic identity')
  const fake = runner(
    home,
    (db) =>
      db === `sql:${dirs[1]}`
        ? failed('SEC_ERROR_BAD_PASSWORD')
        : db === `sql:${dirs[3]}`
          ? { ...failed(), missing: true }
          : ok(keyList(['same', 'same'])),
    () => ok(pem),
  )
  const { listing } = await drain(
    listSystemCertificates('fixture-password', fake.env),
    fake.pending,
  )
  expect(listing.certs.map((cert) => cert.ref)).toEqual(
    [dirs[0], dirs[0], dirs[2], dirs[2]].map((dir) => ({ source: 'nss', dir, nickname: 'same' })),
  )
  expect(listing.issues).toEqual(['store-locked', 'nss-tools-missing'])
  expect(fake.maximum()).toBe(2)
  expect(fake.commands.every((args) => args.includes('-f'))).toBe(true)
})

it('does not invoke a tool when no synthetic database exists', async () => {
  const home = mkdtempSync(join(tmpdir(), 'nss-empty-fixture-'))
  homes.push(home)
  const fake = runner(
    home,
    () => {
      throw new Error('no database')
    },
    () => {
      throw new Error('no keys')
    },
  )
  expect(await listSystemCertificates(undefined, fake.env)).toEqual({ certs: [], issues: [] })
  expect(fake.commands).toEqual([])
})

it('drains outstanding reads before cleaning up when a runner unexpectedly rejects', async () => {
  const { home } = homeWithDatabases()
  const failure = new Error('synthetic runner failure')
  const pem = certificate('Synthetic identity')
  const fake = runner(
    home,
    () => ok(keyList(Array.from({ length: 8 }, () => 'same'))),
    (index) => {
      if (index === 0) throw failure
      return ok(pem)
    },
  )
  let settled = false
  const observed = listSystemCertificates('fixture-password', fake.env).then(
    () => {
      settled = true
      return undefined
    },
    (error) => {
      settled = true
      return error
    },
  )
  while (!fake.pending.length) await setImmediate()
  const batch = fake.pending.splice(0)
  try {
    expect(batch).toHaveLength(8)
    batch[0]!.finish()
    await setImmediate()
    await setImmediate()
    expect(settled).toBe(false)
  } finally {
    for (const item of batch) item.finish()
    expect(await observed).toBe(failure)
  }
})
