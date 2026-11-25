/** Bytes a renderer sent: IPC hands them over as an ArrayBuffer or as a view of one (a Buffer). */
export function ipcBytes(data: unknown): Buffer | null {
  if (data instanceof ArrayBuffer) return Buffer.from(data)
  if (ArrayBuffer.isView(data)) return Buffer.from(data.buffer, data.byteOffset, data.byteLength)
  return null
}
