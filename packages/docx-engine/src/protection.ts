import type { DocProtection } from './types'

/**
 * w:documentProtection password hash (Word 2013+ scheme, ECMA-376 iterated hash):
 *   h0 = H(salt || UTF-16LE(password))
 *   hi = H(h(i-1) || LE32(i))   i = 0..spinCount-1
 * sid 14 = SHA-512 (the only implemented algorithm). Uses WebCrypto, so it works in
 * both the renderer process and Node 20+.
 */

// Word's default iteration count for newly generated hashes.
export const DEFAULT_SPIN_COUNT = 100000
// Upper bound for untrusted spinCount values read from document XML.
// Each iteration costs one Subtle.digest call, so an unbounded value
// from a hostile file would stall the opener (CPU DoS). Both the pure
// helper below and the verify path clamp to this limit before any
// hashing starts, so there is exactly one policy for untrusted counts.
export const MAX_SPIN_COUNT = 1000000

/**
 * Validate and normalize an untrusted spinCount value.
 * Returns DEFAULT_SPIN_COUNT when the value is absent, returns the value
 * itself when it is a safe non-negative integer, and clamps values above
 * MAX_SPIN_COUNT down to MAX_SPIN_COUNT so callers never loop millions
 * of times on attacker-controlled input. Throws on negative, non-integer,
 * or non-numeric values so bad input fails fast without any hashing.
 */
export function resolveSpinCount(raw: unknown): number {
  if (raw === undefined || raw === null || raw === '') return DEFAULT_SPIN_COUNT
  const value = typeof raw === 'string' ? Number(raw.trim()) : (raw as number)
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`invalid spinCount: ${String(raw)}`)
  }
  if (!Number.isInteger(value)) {
    throw new Error(`invalid spinCount: ${String(raw)}`)
  }
  if (value < 0) {
    throw new Error(`invalid spinCount: ${String(raw)}`)
  }
  // DoS guard: cap attacker-controlled iteration counts.
  if (value > MAX_SPIN_COUNT) return MAX_SPIN_COUNT
  return value
}

function toBase64(bytes: Uint8Array): string {
  let bin = ''
  for (const b of bytes) bin += String.fromCharCode(b)
  return btoa(bin)
}

function fromBase64(b64: string): Uint8Array {
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}

function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2)
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i)
    out[i * 2] = code & 0xff
    out[i * 2 + 1] = code >> 8
  }
  return out
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length)
  out.set(a, 0)
  out.set(b, a.length)
  return out
}

async function iteratedSha512(
  password: string,
  salt: Uint8Array,
  spinCount: number,
): Promise<Uint8Array> {
  const subtle = globalThis.crypto.subtle
  const digest = async (bytes: Uint8Array) =>
    new Uint8Array(await subtle.digest('SHA-512', bytes as unknown as ArrayBuffer))
  let hash = await digest(concat(salt, utf16le(password)))
  const iter = new Uint8Array(4)
  const view = new DataView(iter.buffer)
  for (let i = 0; i < spinCount; i++) {
    view.setUint32(0, i, true)
    hash = await digest(concat(hash, iter))
  }
  return hash
}

/** Generate the protection password hash (random 16-byte salt, 100000 iterations by default) */
export async function hashProtectionPassword(
  password: string,
  spinCount = DEFAULT_SPIN_COUNT,
): Promise<{ hash: string; salt: string; spinCount: number; algorithmSid: number }> {
  // Validate the requested count through the same guard so a bad caller
  // cannot trigger an unbounded digest loop either.
  const resolved = resolveSpinCount(spinCount)
  const salt = globalThis.crypto.getRandomValues(new Uint8Array(16))
  const hash = await iteratedSha512(password, salt, resolved)
  return { hash: toBase64(hash), salt: toBase64(salt), spinCount: resolved, algorithmSid: 14 }
}

/** Check whether the password matches the hash in documentProtection (no hash = no password, always true) */
export async function verifyProtectionPassword(
  password: string,
  protection: Pick<DocProtection, 'hash' | 'salt' | 'spinCount' | 'algorithmSid'>,
): Promise<boolean> {
  if (!protection.hash) return true
  if ((protection.algorithmSid ?? 14) !== 14) return false // only SHA-512 is supported
  // Fail closed on missing credentials instead of hashing with an empty
  // salt. Callers treat false as wrong-password, and neither App.tsx nor
  // ProtectDialog catches, so this must never throw.
  if (!protection.salt) {
    return false
  }
  // One policy for untrusted counts: creation and normalization clamp via
  // resolveSpinCount, while verify fail-closes without hashing when the
  // stored count is above MAX_SPIN_COUNT. Hashing a clamped count to the
  // end would burn up to a million digests just to return false (measured
  // past the 20s test timeout), and a stored count that high can only come
  // from a hostile or hand-crafted file. Callers treat false as
  // wrong-password, and neither App.tsx nor ProtectDialog catches, so this
  // must never throw.
  const rawSpin = protection.spinCount ?? DEFAULT_SPIN_COUNT
  const numericSpin = typeof rawSpin === 'string' ? Number(String(rawSpin).trim()) : rawSpin
  if (
    typeof numericSpin === 'number' &&
    Number.isFinite(numericSpin) &&
    numericSpin > MAX_SPIN_COUNT
  ) {
    return false
  }
  const spinCount = resolveSpinCount(rawSpin)
  const salt = fromBase64(protection.salt)
  const hash = await iteratedSha512(password, salt, spinCount)
  return toBase64(hash) === protection.hash
}
