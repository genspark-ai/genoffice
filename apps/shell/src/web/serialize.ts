/**
 * Wire codec for the web bridge. Electron's IPC moves values by structured
 * clone; the web bridge moves them as JSON text frames, so binary values
 * travel as base64 markers. Shared by the server dispatch and the browser
 * shim — isomorphic: no node/electron builtins (btoa/atob are globals in
 * Node 22 and in every browser this targets).
 *
 * Semantics mirror structured clone as Electron delivers it: binaries arrive
 * as typed-array views (Uint8Array), which is exactly what a main-process
 * handler receives for a renderer Buffer or ArrayBuffer over real IPC.
 *
 * Values are walked manually rather than through a JSON.stringify replacer:
 * Buffer's toJSON() would intercept binary payloads before any replacer ran,
 * and unsupported mutable types must fail loudly instead of silently losing
 * their contents.
 */

const BYTES_TAG = '__genofficeBytes'
const VIEW_TAG = '__genofficeBytesView'

interface BytesMarker {
  [BYTES_TAG]: string
  [VIEW_TAG]?: true
}

/** base64 chunks must split on group boundaries: 32766 bytes is 3-aligned,
 *  and the decode side slices the text on 4-aligned lengths */
const BYTE_CHUNK = 0x7ffe
const B64_CHUNK = (BYTE_CHUNK * 4) / 3

function bytesToBase64(bytes: Uint8Array): string {
  let out = ''
  for (let i = 0; i < bytes.length; i += BYTE_CHUNK) {
    out += btoa(String.fromCharCode(...bytes.subarray(i, i + BYTE_CHUNK)))
  }
  return out
}

function base64ToBytes(b64: string): Uint8Array {
  const out = new Uint8Array((b64.length * 3) / 4)
  let written = 0
  for (let i = 0; i < b64.length; i += B64_CHUNK) {
    const bin = atob(b64.slice(i, i + B64_CHUNK))
    for (let j = 0; j < bin.length; j++) out[written++] = bin.charCodeAt(j)
  }
  return out.subarray(0, written)
}

export function encodeDeep(value: unknown): unknown {
  if (value instanceof ArrayBuffer) {
    return { [BYTES_TAG]: bytesToBase64(new Uint8Array(value)) } satisfies BytesMarker
  }
  if (ArrayBuffer.isView(value)) {
    const view = value as ArrayBufferView
    // Buffer shares its backing store; copy the live byte range only
    return {
      [BYTES_TAG]: bytesToBase64(new Uint8Array(view.buffer, view.byteOffset, view.byteLength)),
      [VIEW_TAG]: true,
    } satisfies BytesMarker
  }
  if (value === null || typeof value !== 'object') return value
  if (Array.isArray(value)) return value.map(encodeDeep)
  if (value instanceof Date || value instanceof Map || value instanceof Set) {
    // structured clone keeps these; the JSON wire would mangle them — say so
    throw new Error(`web bridge cannot carry a ${value.constructor.name} over IPC`)
  }
  const out: Record<string, unknown> = {}
  for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
    out[key] = encodeDeep(entry)
  }
  return out
}

function isBytesMarker(value: unknown): value is BytesMarker {
  return (
    typeof value === 'object' &&
    value !== null &&
    BYTES_TAG in value &&
    typeof (value as BytesMarker)[BYTES_TAG] === 'string'
  )
}

/** JSON.parse reviver for frames encoded by encodeValue. Every binary arrives
 *  as a Uint8Array view: Electron's IPC delivers a renderer ArrayBuffer to
 *  main-process handlers the same way, and the handlers' Buffer.from(data)
 *  calls depend on it (Buffer.from(arrayBuffer) one-arg throws). */
export function decodeValue(_key: string, value: unknown): unknown {
  if (isBytesMarker(value)) {
    return base64ToBytes(value[BYTES_TAG])
  }
  return value
}

export function encodeFrame(payload: unknown): string {
  return JSON.stringify(encodeDeep(payload))
}

export function decodeFrame<T>(text: string): T {
  return JSON.parse(text, decodeValue) as T
}
