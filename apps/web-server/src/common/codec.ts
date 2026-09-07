/**
 * Shared IPC transport codec — encodes/decodes binary values across the JSON
 * HTTP boundary. Mirrors the tagged-base64 contract used by
 * `@genoffice/ipc-bridge` so web-client transports can speak the same wire
 * format as Electron main-process handlers.
 */

export const BYTES_TAG = '__ipcBytes'

const TYPED_ARRAY_CTORS = {
  i8: Int8Array,
  u8: Uint8Array,
  u8c: Uint8ClampedArray,
  i16: Int16Array,
  u16: Uint16Array,
  i32: Int32Array,
  u32: Uint32Array,
  f32: Float32Array,
  f64: Float64Array,
  bi64: BigInt64Array,
  bu64: BigUint64Array,
} as const

export type TypedArrayTag = keyof typeof TYPED_ARRAY_CTORS

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('base64')
  }
  let binary = ''
  const CHUNK = 0x8000
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  }
  return btoa(binary)
}

function base64ToBytes(b64: string): Uint8Array {
  if (typeof Buffer !== 'undefined') {
    return new Uint8Array(Buffer.from(b64, 'base64'))
  }
  const binary = atob(b64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

function isTypedArrayTag(tag: string): tag is TypedArrayTag {
  return tag in TYPED_ARRAY_CTORS
}

export function encodeTransportValue(value: unknown, depth = 0): unknown {
  if (value === null || typeof value !== 'object') return value
  if (depth > 64) throw new Error('IPC payload nesting exceeds the transport limit (64)')
  if (ArrayBuffer.isView(value)) {
    const view = value as ArrayBufferView
    const bytes = new Uint8Array(view.buffer, view.byteOffset, view.byteLength)
    if (value instanceof DataView) {
      return { [BYTES_TAG]: 'dv', b64: bytesToBase64(bytes) }
    }
    const tag = (Object.keys(TYPED_ARRAY_CTORS) as TypedArrayTag[]).find(
      (key) => view instanceof TYPED_ARRAY_CTORS[key],
    )
    if (!tag) throw new Error(`IPC payload carries an unsupported binary view: ${view.constructor?.name}`)
    return { [BYTES_TAG]: tag, b64: bytesToBase64(bytes) }
  }
  if (value instanceof ArrayBuffer) {
    return { [BYTES_TAG]: 'ab', b64: bytesToBase64(new Uint8Array(value)) }
  }
  if (Array.isArray(value)) return value.map((item) => encodeTransportValue(item, depth + 1))
  const out: Record<string, unknown> = {}
  for (const [key, item] of Object.entries(value)) out[key] = encodeTransportValue(item, depth + 1)
  return out
}

export function decodeTransportValue(value: unknown, depth = 0): unknown {
  if (value === null || typeof value !== 'object') return value
  if (depth > 64) throw new Error('IPC payload nesting exceeds the transport limit (64)')
  if (Array.isArray(value)) return value.map((item) => decodeTransportValue(item, depth + 1))
  const record = value as Record<string, unknown>
  const tag = record[BYTES_TAG]
  if (typeof tag === 'string' && typeof record.b64 === 'string') {
    if (tag === 'ab') return base64ToBytes(record.b64).buffer
    if (tag === 'dv') return new DataView(base64ToBytes(record.b64).buffer)
    if (isTypedArrayTag(tag)) {
      const Ctor = TYPED_ARRAY_CTORS[tag]
      return new Ctor(base64ToBytes(record.b64).buffer as ArrayBuffer)
    }
  }
  const out: Record<string, unknown> = {}
  for (const [key, item] of Object.entries(record)) out[key] = decodeTransportValue(item, depth + 1)
  return out
}
