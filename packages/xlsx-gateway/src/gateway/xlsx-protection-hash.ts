/// Worksheet protection password hashes: the legacy 16-bit `password=`
/// attribute and the ECMA-376 iterated hash (`algorithmName/hashValue/
/// saltValue/spinCount`, SHA-512 as Excel writes it). WebCrypto only, so the
/// same code runs in the renderer and in Node.

export type SheetPasswordHash =
  | { readonly legacy: string }
  | {
      readonly algorithmName: string
      readonly hashValue: string
      readonly saltValue: string
      readonly spinCount: number
    }

export const DEFAULT_SPIN_COUNT = 100_000
/// Untrusted spinCount values from files are clamped here (CPU DoS guard).
export const MAX_SPIN_COUNT = 1_000_000

/// Excel's legacy worksheet password hash (ECMA-376 Part 1, 18.3.1.85 note);
/// four uppercase hex digits.
export function legacySheetPasswordHash(password: string): string {
  let hash = 0
  for (let index = password.length - 1; index >= 0; index -= 1) {
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)
    hash ^= password.charCodeAt(index)
  }
  hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)
  hash ^= password.length
  hash ^= 0xce4b
  return hash.toString(16).toUpperCase().padStart(4, '0')
}

function toBase64(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary)
}

function fromBase64(text: string): Uint8Array | null {
  try {
    const binary = atob(text)
    const out = new Uint8Array(binary.length)
    for (let index = 0; index < binary.length; index += 1) out[index] = binary.charCodeAt(index)
    return out
  } catch {
    return null
  }
}

function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2)
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    out[index * 2] = code & 0xff
    out[index * 2 + 1] = code >> 8
  }
  return out
}

function concat(left: Uint8Array, right: Uint8Array): Uint8Array {
  const out = new Uint8Array(left.length + right.length)
  out.set(left, 0)
  out.set(right, left.length)
  return out
}

const DIGEST_BY_ALGORITHM: Record<string, string> = {
  'SHA-512': 'SHA-512',
  'SHA-384': 'SHA-384',
  'SHA-256': 'SHA-256',
  'SHA-1': 'SHA-1',
}

async function iteratedHash(
  algorithm: string,
  password: string,
  salt: Uint8Array,
  spinCount: number,
): Promise<Uint8Array> {
  const subtle = globalThis.crypto.subtle
  const digest = async (bytes: Uint8Array) =>
    new Uint8Array(await subtle.digest(algorithm, bytes as unknown as ArrayBuffer))
  let hash = await digest(concat(salt, utf16le(password)))
  const iteration = new Uint8Array(4)
  const view = new DataView(iteration.buffer)
  for (let index = 0; index < spinCount; index += 1) {
    view.setUint32(0, index, true)
    hash = await digest(concat(hash, iteration))
  }
  return hash
}

export async function hashSheetPassword(
  password: string,
  spinCount = DEFAULT_SPIN_COUNT,
  salt: Uint8Array = globalThis.crypto.getRandomValues(new Uint8Array(16)),
): Promise<SheetPasswordHash> {
  const count = Math.min(Math.max(0, Math.trunc(spinCount)), MAX_SPIN_COUNT)
  const hash = await iteratedHash('SHA-512', password, salt, count)
  return {
    algorithmName: 'SHA-512',
    hashValue: toBase64(hash),
    saltValue: toBase64(salt),
    spinCount: count,
  }
}

/// Wrong password, unknown algorithm and malformed base64 all fail closed.
export async function verifySheetPassword(
  password: string,
  stored: SheetPasswordHash,
): Promise<boolean> {
  if ('legacy' in stored) {
    return legacySheetPasswordHash(password) === stored.legacy.toUpperCase().padStart(4, '0')
  }
  const algorithm = DIGEST_BY_ALGORITHM[stored.algorithmName.toUpperCase()]
  if (!algorithm) return false
  const salt = fromBase64(stored.saltValue)
  const expected = fromBase64(stored.hashValue)
  if (!salt || !expected) return false
  const count = Math.min(Math.max(0, Math.trunc(stored.spinCount)), MAX_SPIN_COUNT)
  const actual = await iteratedHash(algorithm, password, salt, count)
  if (actual.length !== expected.length) return false
  let diff = 0
  for (let index = 0; index < actual.length; index += 1) diff |= actual[index]! ^ expected[index]!
  return diff === 0
}
