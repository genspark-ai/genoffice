/** Fetch the bytes behind a one-shot handoff URL from the main process (see main/byte-handoff.ts). */
export async function fetchDocBytes(url: string): Promise<Uint8Array> {
  const res = await fetch(url)
  if (!res.ok) throw new Error(`document bytes unavailable (${res.status})`)
  return new Uint8Array(await res.arrayBuffer())
}
