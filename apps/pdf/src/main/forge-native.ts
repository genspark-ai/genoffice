import { createHash } from 'node:crypto'
import forge from 'node-forge'

/**
 * forge derives PKCS#12 keys in pure JavaScript. Stores such as NSS now write 600 000
 * iterations, which takes seconds and freezes the whole app while a certificate is read.
 * This swaps the iteration loop for Node's native hashes (same algorithm, RFC 7292
 * appendix B) and lets forge's PBKDF2 use the native implementation it already has.
 */

const BLOCK_BYTES: Record<string, number> = {
  md5: 64,
  sha1: 64,
  sha256: 64,
  sha384: 128,
  sha512: 128,
}

/** Far above anything real (NSS: 600 000) but low enough that a hostile file cannot spin forever */
const MAX_ITERATIONS = 10_000_000

type Pkcs12KeyFn = (
  password: string | null | undefined,
  salt: forge.util.ByteStringBuffer,
  id: number,
  iter: number,
  n: number,
  md?: forge.md.MessageDigest | null,
) => forge.util.ByteStringBuffer

const repeatTo = (source: Buffer, length: number): Buffer => {
  if (source.length === 0) return Buffer.alloc(0)
  const out = Buffer.alloc(length)
  for (let i = 0; i < length; i++) out[i] = source[i % source.length]!
  return out
}

export function nativePkcs12Key(
  password: string | null | undefined,
  salt: forge.util.ByteStringBuffer,
  id: number,
  iter: number,
  n: number,
  md?: forge.md.MessageDigest | null,
): forge.util.ByteStringBuffer {
  const algorithm = md?.algorithm ?? 'sha1'
  const v = BLOCK_BYTES[algorithm]
  if (!v) throw new Error(`unsupported PKCS#12 hash: ${algorithm}`)
  if (!Number.isInteger(iter) || iter < 1 || iter > MAX_ITERATIONS) {
    throw new Error('unsupported PKCS#12 iteration count')
  }
  const u = createHash(algorithm).digest().length

  // Password as UTF-16BE with a terminating zero, as the specification asks
  const units: number[] = []
  if (password !== null && password !== undefined) {
    for (let i = 0; i < password.length; i++) units.push(password.charCodeAt(i))
    units.push(0)
  }
  const passBytes = Buffer.alloc(units.length * 2)
  units.forEach((unit, i) => passBytes.writeUInt16BE(unit, i * 2))

  const saltBytes = Buffer.from(salt.bytes(), 'binary')
  const s = saltBytes.length
  const p = passBytes.length
  const diversifier = Buffer.alloc(v, id)
  let inputs = Buffer.concat([
    repeatTo(saltBytes, v * Math.ceil(s / v)),
    repeatTo(passBytes, v * Math.ceil(p / v)),
  ])

  const blocks = Math.ceil(n / u)
  const chunks: Buffer[] = []
  for (let i = 1; i <= blocks; i++) {
    let a = createHash(algorithm).update(diversifier).update(inputs).digest()
    for (let round = 1; round < iter; round++) a = createHash(algorithm).update(a).digest()
    chunks.push(a)
    const b = repeatTo(a, v)
    const next = Buffer.from(inputs)
    const k = Math.ceil(s / v) + Math.ceil(p / v)
    for (let j = 0; j < k; j++) {
      let x = 0x1ff
      for (let l = v - 1; l >= 0; l--) {
        x = x >> 8
        x += b[l]! + next[j * v + l]!
        next[j * v + l] = x & 0xff
      }
    }
    inputs = next
  }
  return forge.util.createBuffer(Buffer.concat(chunks).subarray(0, n).toString('binary'))
}

let installed = false

export function useNativePkcs12Crypto(): void {
  if (installed) return
  installed = true
  const pbe = (forge.pki as unknown as { pbe: { generatePkcs12Key: Pkcs12KeyFn } }).pbe
  const pkcs12 = forge.pkcs12 as unknown as { generateKey: Pkcs12KeyFn }
  pbe.generatePkcs12Key = nativePkcs12Key
  pkcs12.generateKey = nativePkcs12Key

  // PBES2 hands PBKDF2 a digest object, which forces forge's JS path; its name takes the native one
  const pkcs5 = forge.pkcs5 as unknown as {
    pbkdf2: (...args: unknown[]) => unknown
  }
  const original = pkcs5.pbkdf2
  pkcs5.pbkdf2 = (password, salt, count, keyLength, md, ...rest) =>
    original(
      password,
      salt,
      count,
      keyLength,
      md && typeof md === 'object' && 'algorithm' in md
        ? (md as { algorithm: string }).algorithm
        : md,
      ...rest,
    )
}
