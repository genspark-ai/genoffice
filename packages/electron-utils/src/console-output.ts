import { TextDecoder } from 'node:util'

/**
 * Decode output captured from a child process.
 *
 * Windows console programs (reg.exe, powershell.exe) write the active OEM code
 * page -- GBK on Chinese systems -- rather than UTF-8. Reading those bytes as
 * UTF-8 turns localised text (registry values, certificate subjects, paths)
 * into U+FFFD replacement characters. GB18030 is a superset of GBK and agrees
 * with UTF-8 on ASCII, so it is a safe fallback when a strict UTF-8 decode
 * fails. PowerShell can also emit UTF-16 when redirected, so honour BOMs.
 *
 * Non-Windows output stays UTF-8.
 */
export function decodeCommandOutput(
  output: Buffer | string,
  platform: NodeJS.Platform = process.platform,
): string {
  if (typeof output === 'string') return output
  if (output.length >= 2 && output[0] === 0xff && output[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(output).replace(/^\uFEFF/, '')
  }
  if (output.length >= 2 && output[0] === 0xfe && output[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(output).replace(/^\uFEFF/, '')
  }
  if (platform !== 'win32') return output.toString('utf8')
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(output)
  } catch {
    return new TextDecoder('gb18030').decode(output)
  }
}

/**
 * Join stream chunks BEFORE decoding. A multi-byte character can straddle two
 * 'data' events, so decoding each chunk on its own silently corrupts it.
 */
export function decodeChunks(
  chunks: readonly (Buffer | string)[],
  platform: NodeJS.Platform = process.platform,
): string {
  if (chunks.length === 0) return ''
  if (chunks.every((c) => typeof c === 'string')) return (chunks as string[]).join('')
  const buf = Buffer.concat(chunks.map((c) => (typeof c === 'string' ? Buffer.from(c, 'utf8') : c)))
  return decodeCommandOutput(buf, platform)
}
