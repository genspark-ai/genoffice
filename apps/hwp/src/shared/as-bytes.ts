/** IPC may deliver a Node Buffer JSON, a view, or a raw ArrayBuffer. */
export function rawFileBytes(raw: unknown): Uint8Array | null {
  if (raw instanceof Uint8Array) return raw
  if (raw instanceof ArrayBuffer) return new Uint8Array(raw)
  if (ArrayBuffer.isView(raw)) {
    const view = raw as ArrayBufferView
    return new Uint8Array(view.buffer, view.byteOffset, view.byteLength)
  }
  if (
    raw &&
    typeof raw === 'object' &&
    (raw as { type?: string }).type === 'Buffer' &&
    Array.isArray((raw as { data?: unknown }).data)
  ) {
    return Uint8Array.from((raw as { data: number[] }).data)
  }
  return null
}

export function asBytes(raw: unknown): Uint8Array {
  const bytes = rawFileBytes(raw)
  if (!bytes) throw new Error('hwp: expected file bytes')
  return bytes
}
